// Sensitive header names that should be redacted from error logs
const SENSITIVE_HEADERS = [
  'x-user-key',
  'x-user-email',
  'x-openwebui-user-email',
  'authorization',
  'api-key',
  'x-api-key',
]

/**
 * Sanitizes error text by redacting known sensitive header values.
 * Prevents credential leakage in log aggregation systems.
 */
function sanitizeErrorText(errorText) {
  if (typeof errorText !== 'string') return errorText
  let sanitized = errorText
  for (const header of SENSITIVE_HEADERS) {
    // Match header patterns like "X-User-Key: value" or "x-user-key":"value"
    const headerRegex = new RegExp(`(["']?${header}["']?\\s*[:=]\\s*["']?)([^"'\\s,}]+)(["']?)`, 'gi')
    sanitized = sanitized.replace(headerRegex, '$1[REDACTED]$3')
  }
  return sanitized
}

async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController()
  const timeoutHandle = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fetch(url, {
      ...options,
      signal: controller.signal,
    })
  } finally {
    clearTimeout(timeoutHandle)
  }
}

function logAndRespond(res, status, errorObj, context = 'API') {
  if (status >= 400) {
    const message = typeof errorObj?.error === 'string' ? errorObj.error : 'Unhandled error'
    const logPrefix = `[${context}] ${status} ${message}`
    if (status >= 500) {
      console.error(logPrefix)
    } else {
      console.warn(logPrefix)
    }
  }
  return res.status(status).json(errorObj)
}

export {
  fetchWithTimeout,
  logAndRespond,
  sanitizeErrorText,
}
