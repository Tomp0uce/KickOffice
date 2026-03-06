// Sensitive header names that should be redacted from error logs
const SENSITIVE_HEADERS = [
  'x-user-key',
  'x-user-email',
  'x-openwebui-user-email',
  'authorization',
  'api-key',
  'x-api-key',
]

const SENSITIVE_HEADERS_PATTERN = SENSITIVE_HEADERS.join('|')
const SENSITIVE_HEADER_REGEX = new RegExp(`(["']?(?:${SENSITIVE_HEADERS_PATTERN})["']?\\s*[:=]\\s*["']?)([^"'\\s,}]+)(["']?)`, 'gi')

/**
 * Sanitizes error text by redacting known sensitive header values.
 * Prevents credential leakage in log aggregation systems.
 */
function sanitizeErrorText(errorText) {
  if (typeof errorText !== 'string') return errorText
  return errorText.replace(SENSITIVE_HEADER_REGEX, '$1[REDACTED]$3')
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
