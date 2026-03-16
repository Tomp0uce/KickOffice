import { Router } from 'express'

import { buildChatBody } from '../config/models.js'
import { ErrorCodes } from '../config/errorCodes.js'
import { validateChatRequest } from '../middleware/validate.js'
import { chatCompletion, handleErrorResponse, RateLimitError } from '../services/llmClient.js'
import { logAndRespond } from '../utils/http.js'
import logger from '../utils/logger.js'
import { logToolUsage, logChatRequest } from '../utils/toolUsageLogger.js'

const chatRouter = Router()

/**
 * ERR-M1: Shared error handler for chat endpoints
 *
 * Handles common error types:
 * - AbortError: LLM API timeout
 * - RateLimitError: Upstream rate limiting
 * - Generic errors: Internal server errors
 *
 * @param {Response} res - Express response object
 * @param {Error} error - Error to handle
 * @param {Object} req - Express request object (for logger)
 * @param {string} endpoint - Endpoint name for logging (e.g., 'POST /api/chat')
 * @param {boolean} isStreaming - Whether this is a streaming endpoint
 * @returns {void}
 */
function handleChatError(res, error, req, endpoint, isStreaming = false) {
  // Special handling for streaming endpoints that already sent headers
  if (isStreaming && res.headersSent) {
    req.logger.error(`${endpoint} proxy error during stream`, { error, traffic: 'system' })
    if (!res.writableEnded) {
      res.write(`data: ${JSON.stringify({ error: 'Internal server error during stream processing' })}\n\n`)
      res.end()
    }
    return
  }

  // AbortError: LLM API timeout
  if (error.name === 'AbortError') {
    req.logger.error(`${endpoint} LLM API request timeout`, { traffic: 'system' })
    return logAndRespond(res, 504, { code: ErrorCodes.LLM_TIMEOUT, error: 'LLM API request timeout' }, endpoint)
  }

  // RateLimitError: Upstream rate limiting
  if (error instanceof RateLimitError) {
    req.logger.warn(`${endpoint} rate limited by upstream`, { retryAfterMs: error.retryAfterMs, traffic: 'system' })
    return logAndRespond(res, 429, { code: ErrorCodes.RATE_LIMITED, error: 'Rate limit exceeded. Please wait before retrying.' }, endpoint)
  }

  // Generic internal error
  const errorType = isStreaming ? 'proxy error' : 'Chat sync proxy error'
  req.logger.error(`${endpoint} ${errorType}`, { error, traffic: 'system' })
  return logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error' }, endpoint)
}

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

  // FB-M1: Log chat request for history tracking
  try {
    const userId = req.logger.defaultMeta?.userId || 'anonymous'
    const host = req.logger.defaultMeta?.host || 'unknown'
    logChatRequest(userId, host, '/api/chat', messages?.length || 0)
  } catch (logError) {
    req.logger.warn('Failed to log chat request', { error: logError, traffic: 'system' })
  }

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
      model: body.model,
      messageCount: body.messages?.length || 0,
      tools: body.tools?.map(t => t.function?.name) || [],
    })

    const response = await chatCompletion({
      body,
      userCredentials: req.userCredentials,
      modelTier,
    })

    if (!response.ok) {
      const { status: upstreamStatus, rawMessage } = await handleErrorResponse(response, '/api/chat')
      // For 4xx errors from LiteLLM (e.g. invalid image data, bad model), forward the
      // sanitized detail so the UI can display a specific actionable message.
      if (upstreamStatus >= 400 && upstreamStatus < 500) {
        return logAndRespond(res, 400, {
          code: ErrorCodes.LLM_BAD_REQUEST,
          error: 'The AI service rejected the request.',
          detail: rawMessage,
        }, 'POST /api/chat')
      }
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
    let toolCalls = [] // LOG-H1: Accumulate tool calls from stream

    try {
      while (true) {
        if (clientDisconnected) break

        // Add read timeout to prevent hanging requests
        const readPromise = reader.read()
        let timeoutId
        const timeoutPromise = new Promise((_, reject) => {
          timeoutId = setTimeout(() => reject(new Error('Read timeout')), 30000) // 30s timeout
        })

        let readResult
        try {
          readResult = await Promise.race([readPromise, timeoutPromise])
        } catch (readError) {
          if (clientDisconnected) break
          // ERR-M5: Cancel the upstream reader so the LLM provider stops streaming to nobody
          reader.cancel().catch(() => {})
          throw readError
        } finally {
          clearTimeout(timeoutId)
        }

        const { done, value } = readResult
        if (done) break
        const chunk = decoder.decode(value, { stream: true })
        streamContent += chunk

        // LOG-H1: Parse SSE chunks for tool calls
        try {
          const lines = chunk.split('\n').filter(line => line.startsWith('data: '))
          for (const line of lines) {
            const jsonStr = line.slice(6) // Remove 'data: ' prefix
            if (jsonStr === '[DONE]') continue
            try {
              const parsed = JSON.parse(jsonStr)
              const deltaToolCalls = parsed?.choices?.[0]?.delta?.tool_calls
              if (deltaToolCalls && Array.isArray(deltaToolCalls)) {
                toolCalls.push(...deltaToolCalls)
              }
            } catch (parseErr) {
              // ERR-C1: Log parse failures so tool call drops are visible in server logs
              req.logger.warn('POST /api/chat SSE chunk JSON parse failure', {
                traffic: 'system',
                rawChunk: jsonStr.slice(0, 200),
                error: parseErr?.message,
              })
            }
          }
        } catch {
          // Ignore parsing errors
        }

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
        toolCallCount: toolCalls.length,
      })

      // LOG-H1: Log tool usage if tool calls were detected in stream
      if (toolCalls.length > 0) {
        try {
          const userId = req.logger.defaultMeta?.userId || 'anonymous'
          const host = req.logger.defaultMeta?.host || 'unknown'
          logToolUsage(userId, host, toolCalls)
        } catch (logError) {
          req.logger.warn('Failed to log tool usage', { error: logError, traffic: 'system' })
        }
      }
    } catch (streamError) {
      if (!clientDisconnected) {
        req.logger.error('POST /api/chat stream error', { error: streamError, traffic: 'system' })
        // ERR-C2: Deliver error frame to client so it knows the stream was interrupted
        if (!res.writableEnded) {
          try {
            res.write(`data: ${JSON.stringify({ error: 'stream_interrupted' })}\n\n`)
          } catch {
            // Client disconnected between the check and the write — ignore
          }
        }
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
    // ERR-M1: Use shared error handler
    handleChatError(res, error, req, 'POST /api/chat', true)
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

  // FB-M1: Log chat request for history tracking
  try {
    const userId = req.logger.defaultMeta?.userId || 'anonymous'
    const host = req.logger.defaultMeta?.host || 'unknown'
    logChatRequest(userId, host, '/api/chat/sync', messages?.length || 0)
  } catch (logError) {
    req.logger.warn('Failed to log chat request', { error: logError, traffic: 'system' })
  }

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
      const { status: upstreamStatus, rawMessage } = await handleErrorResponse(response, '/api/chat/sync')
      if (upstreamStatus >= 400 && upstreamStatus < 500) {
        return logAndRespond(res, 400, {
          code: ErrorCodes.LLM_BAD_REQUEST,
          error: 'The AI service rejected the request.',
          detail: rawMessage,
        }, 'POST /api/chat/sync')
      }
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

    // QUAL-L1: Log summary only — full response body can be thousands of chars.
    req.logger.info('POST /api/chat/sync upstream response completed', {
      traffic: 'llm',
      model: data?.model,
      usage: data?.usage,
      finish_reason: data?.choices?.[0]?.finish_reason ?? null,
      tool_calls: data?.choices?.[0]?.message?.tool_calls?.map(tc => tc?.function?.name) ?? [],
    })

    // LOG-H1: Log tool usage if tool calls present
    const toolCalls = data?.choices?.[0]?.message?.tool_calls
    if (toolCalls && toolCalls.length > 0) {
      try {
        const userId = req.logger.defaultMeta?.userId || 'anonymous'
        const host = req.logger.defaultMeta?.host || 'unknown'
        logToolUsage(userId, host, toolCalls)
      } catch (logError) {
        req.logger.warn('Failed to log tool usage', { error: logError, traffic: 'system' })
      }
    }

    res.json(data)
  } catch (error) {
    // ERR-M1: Use shared error handler
    handleChatError(res, error, req, 'POST /api/chat/sync', false)
  }
})

export {
  chatRouter,
}
