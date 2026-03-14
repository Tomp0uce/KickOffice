/**
 * Chat request validation
 * ARCH-M2: Extracted from validate.js
 */

import { models } from '../../config/models.js'
import { validateTools } from './toolValidator.js'

const MAX_MESSAGES = 500
const VALID_ROLES = new Set(['system', 'user', 'assistant', 'tool'])

/**
 * Validates temperature parameter
 * @param {unknown} value
 * @returns {{ value: number|undefined }|{ error: string }}
 */
function validateTemperature(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isFinite(value)) return { error: 'temperature must be a finite number' }
  if (value < 0 || value > 2) return { error: 'temperature must be between 0 and 2' }
  return { value }
}

/**
 * Validates maxTokens parameter
 * @param {unknown} value
 * @returns {{ value: number|undefined }|{ error: string }}
 */
function validateMaxTokens(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isInteger(value)) return { error: 'maxTokens must be an integer' }
  if (value < 1 || value > 128000) return { error: 'maxTokens must be between 1 and 128000' }
  return { value }
}

/**
 * Validates individual message structure
 * @param {object} message
 * @param {number} index
 * @returns {{ error: string }|{ valid: true }}
 */
function validateMessage(message, index) {
  if (!message || typeof message !== 'object') {
    return { error: `messages[${index}] must be an object` }
  }

  const { role, content, tool_calls, tool_call_id } = message

  if (typeof role !== 'string' || !VALID_ROLES.has(role)) {
    return { error: `messages[${index}].role must be one of: ${[...VALID_ROLES].join(', ')}` }
  }

  // Content validation depends on role
  if (role === 'tool') {
    // Tool messages require tool_call_id and content
    if (typeof tool_call_id !== 'string' || !tool_call_id) {
      return { error: `messages[${index}].tool_call_id is required for tool messages` }
    }
    if (content !== undefined && content !== null && typeof content !== 'string') {
      return { error: `messages[${index}].content must be a string or null` }
    }
  } else if (role === 'assistant') {
    // Assistant messages may have content and/or tool_calls
    if (content !== undefined && content !== null && typeof content !== 'string') {
      return { error: `messages[${index}].content must be a string or null` }
    }
    if (tool_calls !== undefined && !Array.isArray(tool_calls)) {
      return { error: `messages[${index}].tool_calls must be an array` }
    }
  } else {
    // System and user messages require content
    if (content === undefined || content === null) {
      return { error: `messages[${index}].content is required` }
    }
    if (typeof content !== 'string' && !Array.isArray(content)) {
      return { error: `messages[${index}].content must be a string or array` }
    }
  }

  return { valid: true }
}

/**
 * Validates complete chat request (shared between /api/chat and /api/chat/sync)
 * @param {object} body - Request body
 * @returns {{ error: string }|{ modelConfig: object, parsedTools: object, temperature: number|undefined, maxTokens: number|undefined }}
 */
function validateChatRequest(body) {
  const { messages, modelTier, tools, temperature, maxTokens } = body

  // Check required fields
  if (!messages || !Array.isArray(messages)) {
    return { error: 'messages is required and must be an array' }
  }

  if (messages.length === 0) {
    return { error: 'messages array cannot be empty' }
  }

  if (messages.length > MAX_MESSAGES) {
    return { error: `messages array cannot exceed ${MAX_MESSAGES} messages` }
  }

  // Validate each message structure
  for (let i = 0; i < messages.length; i++) {
    const validation = validateMessage(messages[i], i)
    if (validation.error) {
      return { error: validation.error }
    }
  }

  const modelConfig = models[modelTier]
  if (!modelConfig) {
    return { error: `Unknown model tier: ${modelTier}` }
  }

  if (modelConfig.type === 'image') {
    return { error: 'Use /api/image for image generation' }
  }

  const parsedTools = validateTools(tools)
  if (parsedTools.error) {
    return { error: parsedTools.error }
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return { error: parsedTemperature.error }
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return { error: parsedMaxTokens.error }
  }

  return {
    modelConfig,
    parsedTools,
    temperature: parsedTemperature.value,
    maxTokens: parsedMaxTokens.value,
  }
}

export { validateChatRequest }
