/**
 * Centralized LLM API client service.
 * Handles all HTTP communication with the upstream LLM provider.
 */

import { LLM_API_BASE_URL, LLM_API_KEY } from '../config/models.js'
import { fetchWithTimeout, sanitizeErrorText } from '../utils/http.js'
import logger from '../utils/logger.js'

// Centralized timeout configuration (in milliseconds)
const TIMEOUTS = {
  CHAT_STANDARD: 300_000,   // 5 minutes for standard chat (large files)
  CHAT_REASONING: 300_000,  // 5 minutes for reasoning models
  IMAGE: 180_000,           // 3 minutes for image generation
}

/**
 * Error thrown when the upstream LLM API rate-limits us and all retries are exhausted.
 */
export class RateLimitError extends Error {
  constructor(retryAfterMs) {
    super(`Rate limit exceeded. Retry after ${retryAfterMs}ms.`)
    this.name = 'RateLimitError'
    this.retryAfterMs = retryAfterMs
  }
}

/**
 * Gets the appropriate timeout for a chat request based on model tier.
 */
function getChatTimeoutMs(modelTier) {
  if (modelTier === 'reasoning') return TIMEOUTS.CHAT_REASONING
  return TIMEOUTS.CHAT_STANDARD
}

/**
 * Gets the timeout for image generation requests.
 */
function getImageTimeoutMs() {
  return TIMEOUTS.IMAGE
}

/**
 * Retries a fetch factory function with exponential backoff on transient errors.
 * Retries on 429 (rate-limit) and 5xx responses, or on network-level failures.
 * @param {() => Promise<Response>} fetchFn - Factory returning a fetch Promise
 * @param {number} [maxAttempts=3] - Maximum number of attempts
 * @returns {Promise<Response>}
 */
async function withRetry(fetchFn, maxAttempts = 3) {
  let lastError
  let lastRateLimitMs = 0
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = await fetchFn()
      // Retry on rate-limit or server-side transient errors
      if ((response.status === 429 || response.status >= 500) && attempt < maxAttempts) {
        let delay = Math.min(1000 * 2 ** (attempt - 1), 8000) // 1s, 2s, 4s … capped at 8s
        if (response.status === 429) {
          const retryAfter = response.headers.get('Retry-After')
          if (retryAfter) {
            const parsed = parseFloat(retryAfter)
            // Retry-After may be seconds (float) or an HTTP-date
            const retryMs = isFinite(parsed) ? parsed * 1000 : new Date(retryAfter).getTime() - Date.now()
            if (retryMs > 0) delay = Math.min(retryMs, 60_000) // cap at 60s
          }
          lastRateLimitMs = delay
        }
        logger.warn(`[llmClient] HTTP ${response.status} on attempt ${attempt}/${maxAttempts}, retrying in ${delay}ms`, { traffic: 'system' })
        await new Promise(resolve => setTimeout(resolve, delay))
        continue
      }
      if (response.status === 429) {
        // All retries exhausted on rate limit
        const retryAfter = response.headers.get('Retry-After')
        let retryMs = lastRateLimitMs || 60_000
        if (retryAfter) {
          const parsed = parseFloat(retryAfter)
          retryMs = isFinite(parsed) ? parsed * 1000 : Math.max(retryMs, new Date(retryAfter).getTime() - Date.now())
        }
        throw new RateLimitError(retryMs)
      }
      return response
    } catch (err) {
      if (err instanceof RateLimitError) throw err
      lastError = err
      if (attempt < maxAttempts) {
        const delay = Math.min(1000 * 2 ** (attempt - 1), 8000)
        logger.warn(`[llmClient] Network error on attempt ${attempt}/${maxAttempts}, retrying in ${delay}ms`, { error: err, traffic: 'system' })
        await new Promise(resolve => setTimeout(resolve, delay))
      }
    }
  }
  throw lastError
}

/**
 * Strips header injection characters (\r, \n, non-printable) from a header value.
 */
function sanitizeHeaderValue(value) {
  if (typeof value !== 'string') return ''
  return value.replace(/[\r\n\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
}

/**
 * Builds common headers for LLM API requests.
 */
function buildHeaders(userCredentials) {
  return {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${LLM_API_KEY}`,
    'X-User-Key': sanitizeHeaderValue(userCredentials.userKey),
    'X-OpenWebUi-User-Email': sanitizeHeaderValue(userCredentials.userEmail),
  }
}

/**
 * Makes a chat completion request to the LLM API.
 * @param {Object} options - Request options
 * @param {Object} options.body - The request body (model, messages, etc.)
 * @param {Object} options.userCredentials - User credentials for headers
 * @param {string} options.modelTier - Model tier for timeout selection
 * @returns {Promise<Response>} The fetch response
 */
export async function chatCompletion({ body, userCredentials, modelTier }) {
  const timeoutMs = getChatTimeoutMs(modelTier)
  return withRetry(() => fetchWithTimeout(
    `${LLM_API_BASE_URL}/chat/completions`,
    {
      method: 'POST',
      headers: buildHeaders(userCredentials),
      body: JSON.stringify(body),
    },
    timeoutMs
  ))
}

/**
 * Makes an image generation request to the LLM API.
 * @param {Object} options - Request options
 * @param {Object} options.body - The request body (model, prompt, etc.)
 * @param {Object} options.userCredentials - User credentials for headers
 * @returns {Promise<Response>} The fetch response
 */
export async function imageGeneration({ body, userCredentials }) {
  const timeoutMs = getImageTimeoutMs()
  return withRetry(() => fetchWithTimeout(
    `${LLM_API_BASE_URL}/images/generations`,
    {
      method: 'POST',
      headers: buildHeaders(userCredentials),
      body: JSON.stringify(body),
    },
    timeoutMs
  ))
}

/**
 * Handles an error response from the LLM API.
 * Extracts and sanitizes error text for logging.
 * @param {Response} response - The failed response
 * @param {string} context - Context string for logging
 */
/**
 * Extracts and sanitizes error text for logging.
 * Returns the sanitized error text so callers can forward it to the client.
 * @param {Response} response - The failed response
 * @param {string} context - Context string for logging
 * @returns {{ status: number, rawMessage: string }} Sanitized error info
 */
export async function handleErrorResponse(response, context) {
  const errorText = await response.text()
  const sanitized = sanitizeErrorText(errorText)
  logger.error(`LLM API error on ${context}`, {
    status: response.status,
    errorText: sanitized,
    traffic: 'system'
  })

  // Try to extract a user-readable message from the LiteLLM JSON error body
  // (e.g. { "error": { "message": "litellm.BadRequestError: ..." } })
  let rawMessage = sanitized
  try {
    const parsed = JSON.parse(sanitized)
    if (parsed?.error?.message) rawMessage = parsed.error.message
    else if (typeof parsed?.error === 'string') rawMessage = parsed.error
    else if (parsed?.message) rawMessage = parsed.message
  } catch {
    // Not JSON — keep sanitized raw text
  }

  return { status: response.status, rawMessage }
}
