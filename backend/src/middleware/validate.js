import { isChatGptModel, isGpt5Model, MAX_TOOLS, models } from '../config/models.js'

/** @param {unknown} value @returns {boolean} */
function isPlainObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
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
 * @param {unknown} value
 * @returns {{ value: number|undefined }|{ error: string }}
 */
function validateMaxTokens(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isInteger(value)) return { error: 'maxTokens must be an integer' }
  if (value < 1 || value > 32768) return { error: 'maxTokens must be between 1 and 32768' }
  return { value }
}

/**
 * @param {unknown} tools
 * @returns {{ value: Array<object>|undefined }|{ error: string }}
 */
function validateTools(tools) {
  if (tools === undefined) return { value: undefined }
  if (!Array.isArray(tools)) return { error: 'tools must be an array' }
  if (tools.length === 0) return { value: undefined }
  if (tools.length > MAX_TOOLS) return { error: `tools supports at most ${MAX_TOOLS} entries` }

  const sanitizedTools = []
  for (const tool of tools) {
    if (!isPlainObject(tool)) return { error: 'each tool must be an object' }
    if (tool.type !== 'function') return { error: 'only function tools are supported' }
    if (!isPlainObject(tool.function)) return { error: 'tool.function must be an object' }

    const { name, description, parameters, strict } = tool.function
    if (typeof name !== 'string' || !/^[a-zA-Z0-9_-]{1,64}$/.test(name)) {
      return { error: 'tool.function.name must match /^[a-zA-Z0-9_-]{1,64}$/' }
    }
    if (description !== undefined && typeof description !== 'string') {
      return { error: 'tool.function.description must be a string' }
    }
    if (!isPlainObject(parameters)) {
      return { error: 'tool.function.parameters must be a JSON schema object' }
    }
    if (strict !== undefined && typeof strict !== 'boolean') {
      return { error: 'tool.function.strict must be a boolean' }
    }

    sanitizedTools.push({
      type: 'function',
      function: {
        name,
        ...(description !== undefined ? { description } : {}),
        parameters,
        ...(strict !== undefined ? { strict } : {}),
      },
    })
  }

  return { value: sanitizedTools }
}

/**
 * @param {{ prompt?: unknown, size?: unknown, quality?: unknown, n?: unknown }} [payload]
 * @returns {{ value: { prompt: string, size: string, quality: string, n: number } }|{ error: string }}
 */
function validateImagePayload(payload = {}) {
  const { prompt, size = '1024x1024', quality = 'auto', n = 1 } = payload

  const allowedSizes = new Set(['1024x1024', '1024x1536', '1536x1024'])
  const allowedQualities = new Set(['low', 'medium', 'high', 'auto'])
  const maxPromptLength = 4000

  if (!prompt) {
    return { error: 'prompt is required' }
  }
  if (typeof prompt !== 'string') {
    return { error: 'prompt must be a string' }
  }
  if (prompt.length > maxPromptLength) {
    return { error: `prompt must be <= ${maxPromptLength} characters` }
  }
  if (typeof size !== 'string' || !allowedSizes.has(size)) {
    return { error: `size must be one of: ${[...allowedSizes].join(', ')}` }
  }
  if (typeof quality !== 'string' || !allowedQualities.has(quality)) {
    return { error: `quality must be one of: ${[...allowedQualities].join(', ')}` }
  }
  if (!Number.isInteger(n) || n < 1 || n > 4) {
    return { error: 'n must be an integer between 1 and 4' }
  }

  return { value: { prompt, size, quality, n } }
}

const MAX_MESSAGES = 200
const VALID_ROLES = new Set(['system', 'user', 'assistant', 'tool'])

/**
 * Validates individual message structure.
 * Returns { error: string } on failure, or { valid: true } on success.
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
 * Validates chat request parameters (shared between /api/chat and /api/chat/sync).
 * Returns { error: string } on failure, or { modelConfig, parsedTools, temperature, maxTokens } on success.
 */
const validateChatRequest = (body) => {
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
    maxTokens: parsedMaxTokens.value
  }
}

export {
  validateChatRequest,
  validateImagePayload
}
