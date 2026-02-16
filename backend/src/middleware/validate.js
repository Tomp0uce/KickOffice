import { MAX_TOOLS } from '../config/models.js'

function isPlainObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateTemperature(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isFinite(value)) return { error: 'temperature must be a finite number' }
  if (value < 0 || value > 2) return { error: 'temperature must be between 0 and 2' }
  return { value }
}

function validateMaxTokens(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isInteger(value)) return { error: 'maxTokens must be an integer' }
  if (value < 1 || value > 32768) return { error: 'maxTokens must be between 1 and 32768' }
  return { value }
}

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

export {
  validateImagePayload,
  validateMaxTokens,
  validateTemperature,
  validateTools,
}
