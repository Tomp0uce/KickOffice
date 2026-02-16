import { Router } from 'express'

import { buildChatBody, isChatGptModel, isGpt5Model, LLM_API_BASE_URL, LLM_API_KEY, models } from '../config/models.js'
import { validateMaxTokens, validateTemperature, validateTools } from '../middleware/validate.js'
import { fetchWithTimeout, logAndRespond } from '../utils/http.js'

const chatRouter = Router()

function requiresReasoningSafeParams(modelConfig) {
  return isGpt5Model(modelConfig.id) && modelConfig.reasoningEffort !== 'none'
}

function getChatTimeoutMs(modelTier) {
  if (modelTier === 'reasoning') return 300_000
  return 120_000
}

chatRouter.post('/', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens } = req.body

  if (!messages || !Array.isArray(messages)) {
    return logAndRespond(res, 400, { error: 'messages array is required' }, 'POST /api/chat')
  }

  const modelConfig = models[modelTier]
  if (!modelConfig) {
    return logAndRespond(res, 400, { error: `Unknown model tier: ${modelTier}` }, 'POST /api/chat')
  }

  if (modelConfig.type === 'image') {
    return logAndRespond(res, 400, { error: 'Use /api/image for image generation' }, 'POST /api/chat')
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return logAndRespond(res, 400, { error: parsedTemperature.error }, 'POST /api/chat')
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return logAndRespond(res, 400, { error: parsedMaxTokens.error }, 'POST /api/chat')
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return logAndRespond(res, 400, {
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    }, 'POST /api/chat')
  }

  if (requiresReasoningSafeParams(modelConfig) && temperature !== undefined) {
    return logAndRespond(res, 400, {
      error: 'temperature is only supported for GPT-5 models when reasoning effort is none',
    }, 'POST /api/chat')
  }

  try {
    const body = buildChatBody({
      modelTier,
      modelConfig,
      messages,
      temperature,
      maxTokens,
      stream: true,
    })

    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    }, getChatTimeoutMs(modelTier))

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`LLM API error ${response.status}:`, errorText)
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat')
    }

    res.setHeader('Content-Type', 'text/event-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.setHeader('Connection', 'keep-alive')

    const reader = response.body.getReader()
    const decoder = new TextDecoder()

    try {
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        const chunk = decoder.decode(value, { stream: true })
        res.write(chunk)
      }
    } catch (streamError) {
      console.error('Stream error:', streamError)
    } finally {
      res.end()
    }
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat')
    }
    console.error('Chat proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat')
  }
})

chatRouter.post('/sync', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body

  if (!messages || !Array.isArray(messages)) {
    return logAndRespond(res, 400, { error: 'messages array is required' }, 'POST /api/chat/sync')
  }

  const modelConfig = models[modelTier]
  if (!modelConfig || modelConfig.type === 'image') {
    return logAndRespond(res, 400, { error: `Invalid model tier for chat: ${modelTier}` }, 'POST /api/chat/sync')
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return logAndRespond(res, 400, { error: parsedTemperature.error }, 'POST /api/chat/sync')
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return logAndRespond(res, 400, { error: parsedMaxTokens.error }, 'POST /api/chat/sync')
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return logAndRespond(res, 400, {
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    }, 'POST /api/chat/sync')
  }

  if (requiresReasoningSafeParams(modelConfig) && temperature !== undefined) {
    return logAndRespond(res, 400, {
      error: 'temperature is only supported for GPT-5 models when reasoning effort is none',
    }, 'POST /api/chat/sync')
  }

  const parsedTools = validateTools(tools)
  if (parsedTools.error) {
    return logAndRespond(res, 400, { error: parsedTools.error }, 'POST /api/chat/sync')
  }

  try {
    const body = buildChatBody({
      modelTier,
      modelConfig,
      messages,
      temperature,
      maxTokens,
      stream: false,
      tools: parsedTools.value,
    })

    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    }, getChatTimeoutMs(modelTier))

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`LLM API error ${response.status}:`, errorText)
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat/sync')
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat/sync')
    }
    console.error('Chat sync proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat/sync')
  }
})

export {
  chatRouter,
}
