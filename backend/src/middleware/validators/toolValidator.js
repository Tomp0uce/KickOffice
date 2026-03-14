/**
 * Tool validation
 * ARCH-M2: Extracted from validate.js
 */

import { MAX_TOOLS } from '../../config/models.js'
import { isPlainObject, getObjectDepth } from './common.js'

/**
 * Validates tools array for function calling
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
    if (getObjectDepth(parameters) > 20) {
      return { error: 'tool.function.parameters schema exceeds maximum allowed depth' }
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

export { validateTools }
