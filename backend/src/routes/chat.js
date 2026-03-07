import { Router } from 'express'

import { buildChatBody } from '../config/models.js'
import { ErrorCodes } from '../config/errorCodes.js'
import { validateChatRequest } from '../middleware/validate.js'
import { chatCompletion, handleErrorResponse } from '../services/llmClient.js'
import { logAndRespond } from '../utils/http.js'
import logger from '../utils/logger.js'

const chatRouter = Router()

chatRouter.post('/', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body
  req.logger.debug(` /api/chat incoming`, {
    modelTier,
    messageCount: Array.isArray(messages) ? messages.length : 0,
    hasTemperature: temperature !== undefined,
    hasMaxTokens: maxTokens !== undefined,
  })

  const validation = validateChatRequest(req.body)
  if (validation.error) {
    return logAndRespond(res, 400, { code: ErrorCodes.VALIDATION_ERROR, error: validation.error }, 'POST /api/chat')
  }

  const { modelConfig, parsedTools, temperature: validTemp, maxTokens: validMaxTokens } = validation

  try {
    const body = buildChatBody({
      modelTier,
      modelConfig,
      messages,
      temperature: validTemp,
      maxTokens: validMaxTokens,
      stream: true,
      tools: parsedTools.value,
    })

    req.logger.debug(` /api/chat upstream payload`, {
      model: body.model,
      stream: body.stream,
      messageCount: body.messages?.length || 0,
      hasReasoningEffort: Object.prototype.hasOwnProperty.call(body, 'reasoning_effort'),
      hasTemperature: Object.prototype.hasOwnProperty.call(body, 'temperature'),
      hasMaxTokens: Object.prototype.hasOwnProperty.call(body, 'max_tokens') || Object.prototype.hasOwnProperty.call(body, 'max_completion_tokens'),
    })

    req.logger.info(`POST /api/chat upstream request initiated`, {
      traffic: 'llm',
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
        code: ErrorCodes.LLM_UPSTREAM_ERROR,
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat')
    }

    res.setHeader('Content-Type', 'text/event-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.setHeader('Connection', 'keep-alive')

    const reader = response.body.getReader()
    const decoder = new TextDecoder()
    let clientDisconnected = false

    // Track client disconnection and cancel upstream reader
    res.on('close', () => {
      clientDisconnected = true
      // Cancel upstream reader to stop draining the response
      reader.cancel().catch(() => {
        // Ignore cancel errors - connection is already closed
      })
    })

    let streamContent = ''

    try {
      while (true) {
        if (clientDisconnected) break

        // Add read timeout to prevent hanging requests
        const readPromise = reader.read()
        const timeoutPromise = new Promise((_, reject) => {
          setTimeout(() => reject(new Error('Read timeout')), 30000) // 30s timeout
        })

        let readResult
        try {
          readResult = await Promise.race([readPromise, timeoutPromise])
        } catch (readError) {
          if (clientDisconnected) break
          throw readError
        }

        const { done, value } = readResult
        if (done) break
        const chunk = decoder.decode(value, { stream: true })
        streamContent += chunk

        // Check if client disconnected before writing
        if (clientDisconnected) break

        // Check write result for backpressure
        let canContinue
        try {
          canContinue = res.write(chunk)
        } catch (writeError) {
          // Client disconnected during write
          if (clientDisconnected) break
          throw writeError
        }

        if (!canContinue && !clientDisconnected) {
          // Wait for drain before continuing
          await new Promise(resolve => {
            const onDrain = () => { res.removeListener('close', onClose); resolve() }
            const onClose = () => { res.removeListener('drain', onDrain); clientDisconnected = true; resolve() }
            res.once('drain', onDrain)
            res.once('close', onClose)
          })
        }
      }
      // Flush any remaining bytes in the decoder
      const finalChunk = decoder.decode()
      if (finalChunk && !clientDisconnected) {
        try {
          res.write(finalChunk)
        } catch {
          // Ignore write errors if client disconnected
        }
      }
      req.logger.info('POST /api/chat stream completed successfully', {
        traffic: 'llm',
        responseLength: streamContent.length,
        responseContent: streamContent
      })
    } catch (streamError) {
      if (!clientDisconnected) {
        req.logger.error('POST /api/chat stream error', { error: streamError, traffic: 'system' })
      }
    } finally {
      // Cancel reader if still active
      try {
        await reader.cancel()
      } catch {
        // Ignore cancel errors
      }

      if (!res.writableEnded) {
        res.end()
      }
    }
  } catch (error) {
    if (res.headersSent) {
      req.logger.error('POST /api/chat proxy error during stream', { error, traffic: 'system' })
      if (!res.writableEnded) {
        res.write(`data: ${JSON.stringify({ error: 'Internal server error during stream processing' })}\n\n`)
        res.end()
      }
      return
    }
    if (error.name === 'AbortError') {
      req.logger.error('POST /api/chat LLM API request timeout', { traffic: 'system' })
      return logAndRespond(res, 504, { code: ErrorCodes.LLM_TIMEOUT, error: 'LLM API request timeout' }, 'POST /api/chat')
    }
    req.logger.error('POST /api/chat Chat proxy error', { error, traffic: 'system' })
    return logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error' }, 'POST /api/chat')
  }
})

chatRouter.post('/sync', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body
  req.logger.debug(` /api/chat/sync incoming`, {
    modelTier,
    messageCount: Array.isArray(messages) ? messages.length : 0,
    toolCount: Array.isArray(tools) ? tools.length : 0,
    hasTemperature: temperature !== undefined,
    hasMaxTokens: maxTokens !== undefined,
  })

  const validation = validateChatRequest(req.body)
  if (validation.error) {
    return logAndRespond(res, 400, { code: ErrorCodes.VALIDATION_ERROR, error: validation.error }, 'POST /api/chat/sync')
  }

  const { modelConfig, parsedTools, temperature: validTemp, maxTokens: validMaxTokens } = validation

  try {
    const body = buildChatBody({
      modelTier,
      modelConfig,
      messages,
      temperature: validTemp,
      maxTokens: validMaxTokens,
      stream: false,
      tools: parsedTools.value,
    })

    req.logger.debug(` /api/chat/sync upstream payload`, {
      model: body.model,
      stream: body.stream,
      messageCount: body.messages?.length || 0,
      toolCount: body.tools?.length || 0,
      hasReasoningEffort: Object.prototype.hasOwnProperty.call(body, 'reasoning_effort'),
      hasTemperature: Object.prototype.hasOwnProperty.call(body, 'temperature'),
      hasMaxTokens: Object.prototype.hasOwnProperty.call(body, 'max_tokens') || Object.prototype.hasOwnProperty.call(body, 'max_completion_tokens'),
    })

    req.logger.info(`POST /api/chat/sync upstream request initiated`, {
      traffic: 'llm',
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
        code: ErrorCodes.LLM_UPSTREAM_ERROR,
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat/sync')
    }

    const data = await response.json()

    // Validate upstream response structure
    if (!data || typeof data !== 'object') {
      req.logger.error('LLM API returned invalid response format', { type: typeof data, traffic: 'system' })
      return logAndRespond(res, 502, {
        code: ErrorCodes.LLM_INVALID_JSON,
        error: 'The AI service returned an invalid response format.',
      }, 'POST /api/chat/sync')
    }

    if (!Array.isArray(data.choices) || data.choices.length === 0) {
      req.logger.error('LLM API returned no choices', { data: JSON.stringify(data).slice(0, 500), traffic: 'system' })
      return logAndRespond(res, 502, {
        code: ErrorCodes.LLM_NO_CHOICES,
        error: 'The AI service returned an empty response.',
      }, 'POST /api/chat/sync')
    }

    const firstChoice = data.choices[0]
    if (!firstChoice.message || typeof firstChoice.message !== 'object') {
      req.logger.error('LLM API returned invalid choice structure', { choice: firstChoice, traffic: 'system' })
      return logAndRespond(res, 502, {
        code: ErrorCodes.LLM_INVALID_JSON,
        error: 'The AI service returned an invalid response structure.',
      }, 'POST /api/chat/sync')
    }

    if (!firstChoice.message.content && (!firstChoice.message.tool_calls || firstChoice.message.tool_calls.length === 0)) {
      req.logger.error('LLM API returned empty content without tool calls', { choice: firstChoice, traffic: 'system' })
      return logAndRespond(res, 502, {
        code: ErrorCodes.LLM_EMPTY_RESPONSE,
        error: 'The AI service returned an empty response.',
      }, 'POST /api/chat/sync')
    }

    req.logger.debug(` /api/chat/sync upstream response`, {
      id: data?.id,
      model: data?.model,
      choiceCount: data?.choices?.length || 0,
      hasFirstChoice: !!data?.choices?.[0],
      finishReason: data?.choices?.[0]?.finish_reason ?? null,
      hasContent: !!data?.choices?.[0]?.message?.content,
      toolCallCount: data?.choices?.[0]?.message?.tool_calls?.length || 0,
    })
    
    req.logger.info('POST /api/chat/sync upstream response completed', { traffic: 'llm', response: data })
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      req.logger.error('POST /api/chat/sync LLM API request timeout', { traffic: 'system' })
      return logAndRespond(res, 504, { code: ErrorCodes.LLM_TIMEOUT, error: 'LLM API request timeout' }, 'POST /api/chat/sync')
    }
    req.logger.error('POST /api/chat/sync Chat sync proxy error', { error, traffic: 'system' })
    return logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error' }, 'POST /api/chat/sync')
  }
})

export {
  chatRouter,
}
