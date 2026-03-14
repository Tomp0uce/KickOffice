/**
 * Common validation utilities
 * ARCH-M2: Extracted from validate.js for better separation of concerns
 */

/**
 * Check if a value is a plain object (not array, not null)
 */
function isPlainObject(value) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const proto = Object.getPrototypeOf(value)
  return proto === Object.prototype || proto === null
}

/**
 * Calculate the depth of an object or array structure
 */
function getObjectDepth(obj, currentDepth = 1, maxDepth = 20) {
  if (currentDepth > maxDepth) return currentDepth
  if (!isPlainObject(obj) && !Array.isArray(obj)) return currentDepth

  let maxChildDepth = currentDepth
  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const depth = getObjectDepth(obj[key], currentDepth + 1, maxDepth)
      if (depth > maxChildDepth) maxChildDepth = depth
    }
  }
  return maxChildDepth
}

export { isPlainObject, getObjectDepth }
