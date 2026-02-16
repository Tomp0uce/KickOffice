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
}
