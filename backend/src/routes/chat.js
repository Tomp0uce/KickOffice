import { Router } from 'express'

import { buildChatBody } from '../config/models.js'
import { validateChatRequest } from '../middleware/validate.js'
import { chatCompletion, handleErrorResponse } from '../services/llmClient.js'
import { logAndRespond } from '../utils/http.js'
import { systemLog } from '../utils/logger.js'

const chatRouter = Router()
const VERBOSE_LOGGING_ENABLED = process.env.VERBOSE_LOGGING === 'true'
const verboseLog = VERBOSE_LOGGING_ENABLED ? console.info.bind(console, '[KO-CHAT]') : () => {}

chatRouter.post('/', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body
  verboseLog(` /api/chat incoming`, {
    modelTier,
    messageCount: Array.isArray(messages) ? messages.length : 0,
    hasTemperature: temperature !== undefined,
    hasMaxTokens: maxTokens !== undefined,
  })

  const validation = validateChatRequest(req.body, 'POST /api/chat')
  if (validation.error) {
    return logAndRespond(res, 400, { error: validation.error }, 'POST /api/chat')
  }

  const { modelConfig, parsedTools } = validation

  try {
    const body = buildChatBody({
      modelTier,
      modelConfig,
      messages,
      temperature,
      maxTokens,
      stream: true,
      tools: parsedTools.value,
    })

    verboseLog(` /api/chat upstream payload`, {
      model: body.model,
      stream: body.stream,
      messageCount: body.messages?.length || 0,
      hasReasoningEffort: Object.prototype.hasOwnProperty.call(body, 'reasoning_effort'),
      hasTemperature: Object.prototype.hasOwnProperty.call(body, 'temperature'),
      hasMaxTokens: Object.prototype.hasOwnProperty.call(body, 'max_tokens') || Object.prototype.hasOwnProperty.call(body, 'max_completion_tokens'),
    })

    systemLog('INFO', `POST /api/chat upstream request initiated`, {
      url: '/v1/chat/completions',
      body,
    })

    const response = await chatCompletion({
      body,
      userCredentials: req.userCredentials,
      modelTier,
    })

    if (!response.ok) {
      await handleErrorResponse(response, '/api/chat')
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat')
    }

    res.setHeader('Content-Type', 'text/event-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.setHeader('Connection', 'keep-alive')

    const reader = response.body.getReader()
    const decoder = new TextDecoder()
    let clientDisconnected = false

    // Track client disconnection
    res.on('close', () => { clientDisconnected = true })

    try {
      while (true) {
        if (clientDisconnected) break
        const { done, value } = await reader.read()
        if (done) break
        const chunk = decoder.decode(value, { stream: true })

        // Check write result for backpressure
        const canContinue = res.write(chunk)
        if (!canContinue && !clientDisconnected) {
          // Wait for drain before continuing
          await new Promise(resolve => res.once('drain', resolve))
        }
      }
      // Flush any remaining bytes in the decoder
      const finalChunk = decoder.decode()
      if (finalChunk && !clientDisconnected) {
        res.write(finalChunk)
      }
      systemLog('INFO', 'POST /api/chat stream completed successfully')
    } catch (streamError) {
      if (!clientDisconnected) {
        systemLog('ERROR', 'POST /api/chat stream error', streamError)
        console.error('Stream error:', streamError)
      }
    } finally {
      if (!res.writableEnded) {
        res.end()
      }
    }
  } catch (error) {
    if (error.name === 'AbortError') {
      systemLog('ERROR', 'POST /api/chat LLM API request timeout')
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat')
    }
    systemLog('ERROR', 'POST /api/chat Chat proxy error', error)
    console.error('Chat proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat')
  }
})

chatRouter.post('/sync', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body
  verboseLog(` /api/chat/sync incoming`, {
    modelTier,
    messageCount: Array.isArray(messages) ? messages.length : 0,
    toolCount: Array.isArray(tools) ? tools.length : 0,
    hasTemperature: temperature !== undefined,
    hasMaxTokens: maxTokens !== undefined,
  })

  const validation = validateChatRequest(req.body, 'POST /api/chat/sync')
  if (validation.error) {
    return logAndRespond(res, 400, { error: validation.error }, 'POST /api/chat/sync')
  }

  const { modelConfig, parsedTools } = validation

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

    verboseLog(` /api/chat/sync upstream payload`, {
      model: body.model,
      stream: body.stream,
      messageCount: body.messages?.length || 0,
      toolCount: body.tools?.length || 0,
      hasReasoningEffort: Object.prototype.hasOwnProperty.call(body, 'reasoning_effort'),
      hasTemperature: Object.prototype.hasOwnProperty.call(body, 'temperature'),
      hasMaxTokens: Object.prototype.hasOwnProperty.call(body, 'max_tokens') || Object.prototype.hasOwnProperty.call(body, 'max_completion_tokens'),
    })

    systemLog('INFO', `POST /api/chat/sync upstream request initiated`, {
      url: '/v1/chat/completions',
      body,
    })

    const response = await chatCompletion({
      body,
      userCredentials: req.userCredentials,
      modelTier,
    })

    if (!response.ok) {
      await handleErrorResponse(response, '/api/chat/sync')
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat/sync')
    }

    const data = await response.json()

    // Validate upstream response structure
    if (!data || typeof data !== 'object') {
      console.error('LLM API returned invalid response format', { type: typeof data })
      return logAndRespond(res, 502, {
        error: 'The AI service returned an invalid response format.',
      }, 'POST /api/chat/sync')
    }

    if (!Array.isArray(data.choices) || data.choices.length === 0) {
      console.error('LLM API returned no choices', { data: JSON.stringify(data).slice(0, 500) })
      return logAndRespond(res, 502, {
        error: 'The AI service returned an empty response.',
      }, 'POST /api/chat/sync')
    }

    const firstChoice = data.choices[0]
    if (!firstChoice.message || typeof firstChoice.message !== 'object') {
      console.error('LLM API returned invalid choice structure', { choice: firstChoice })
      return logAndRespond(res, 502, {
        error: 'The AI service returned an invalid response structure.',
      }, 'POST /api/chat/sync')
    }

    verboseLog(` /api/chat/sync upstream response`, {
      id: data?.id,
      model: data?.model,
      choiceCount: data?.choices?.length || 0,
      hasFirstChoice: !!data?.choices?.[0],
      finishReason: data?.choices?.[0]?.finish_reason ?? null,
      hasContent: !!data?.choices?.[0]?.message?.content,
      toolCallCount: data?.choices?.[0]?.message?.tool_calls?.length || 0,
    })
    
    systemLog('INFO', 'POST /api/chat/sync upstream response completed', data)
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      systemLog('ERROR', 'POST /api/chat/sync LLM API request timeout')
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat/sync')
    }
    systemLog('ERROR', 'POST /api/chat/sync Chat sync proxy error', error)
    console.error('Chat sync proxy error', { modelTier, error })
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat/sync')
  }
})

export {
  chatRouter,
}
