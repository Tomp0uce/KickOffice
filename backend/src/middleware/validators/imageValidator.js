/**
 * Image generation validation
 * ARCH-M2: Extracted from validate.js
 */

/**
 * Validates image generation request payload
 * @param {{ prompt?: unknown, size?: unknown, quality?: unknown, n?: unknown }} [payload]
 * @returns {{ value: { prompt: string, size: string, quality: string, n: number } }|{ error: string }}
 */
function validateImagePayload(payload = {}) {
  const { prompt, size = '1792x1024', quality = 'auto', n = 1 } = payload

  const allowedSizes = new Set(['1024x1024', '1024x1536', '1536x1024', '1792x1024'])
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

export { validateImagePayload }
