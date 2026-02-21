/**
 * Centralized LLM API client service.
 * Handles all HTTP communication with the upstream LLM provider.
 */

import { LLM_API_BASE_URL, LLM_API_KEY } from '../config/models.js'
import { fetchWithTimeout, sanitizeErrorText } from '../utils/http.js'

// Centralized timeout configuration (in milliseconds)
export const TIMEOUTS = {
  CHAT_STANDARD: 120_000,   // 2 minutes for standard chat
  CHAT_REASONING: 300_000,  // 5 minutes for reasoning models
  IMAGE: 180_000,           // 3 minutes for image generation
}

/**
 * Gets the appropriate timeout for a chat request based on model tier.
 */
export function getChatTimeoutMs(modelTier) {
  if (modelTier === 'reasoning') return TIMEOUTS.CHAT_REASONING
  return TIMEOUTS.CHAT_STANDARD
}

/**
 * Gets the timeout for image generation requests.
 */
export function getImageTimeoutMs() {
  return TIMEOUTS.IMAGE
}

/**
 * Builds common headers for LLM API requests.
 */
function buildHeaders(userCredentials) {
  return {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${LLM_API_KEY}`,
    'X-User-Key': userCredentials.userKey,
    'X-OpenWebUi-User-Email': userCredentials.userEmail,
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
  return fetchWithTimeout(
    `${LLM_API_BASE_URL}/chat/completions`,
    {
      method: 'POST',
      headers: buildHeaders(userCredentials),
      body: JSON.stringify(body),
    },
    timeoutMs
  )
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
  return fetchWithTimeout(
    `${LLM_API_BASE_URL}/images/generations`,
    {
      method: 'POST',
      headers: buildHeaders(userCredentials),
      body: JSON.stringify(body),
    },
    timeoutMs
  )
}

/**
 * Handles an error response from the LLM API.
 * Extracts and sanitizes error text for logging.
 * @param {Response} response - The failed response
 * @param {string} context - Context string for logging
 * @returns {Promise<string>} Sanitized error text
 */
export async function handleErrorResponse(response, context) {
  const errorText = await response.text()
  const sanitized = sanitizeErrorText(errorText)
  console.error(`LLM API error on ${context}`, {
    status: response.status,
    errorText: sanitized,
  })
  return sanitized
}
