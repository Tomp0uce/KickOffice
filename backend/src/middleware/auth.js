import { LLM_API_KEY } from '../config/models.js'
import { logAndRespond } from '../utils/http.js'

// Basic email format validation (not exhaustive, but catches obvious issues)
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
const MIN_KEY_LENGTH = 8

function ensureLlmApiKey(req, res, next) {
  if (!LLM_API_KEY) {
    return logAndRespond(
      res,
      500,
      { error: 'LLM API key not configured on server' },
      `${req.method} ${req.originalUrl}`,
    )
  }
  return next()
}

function ensureUserCredentials(req, res, next) {
  const userKey = req.headers['x-user-key']
  const userEmail = req.headers['x-user-email']

  if (!userKey || !userEmail) {
    return logAndRespond(
      res,
      401,
      { error: 'LiteLLM user credentials required (X-User-Key and X-User-Email headers)' },
      `${req.method} ${req.originalUrl}`,
    )
  }

  // Validate email format
  if (typeof userEmail !== 'string' || !EMAIL_REGEX.test(userEmail)) {
    return logAndRespond(
      res,
      400,
      { error: 'Invalid email format in X-User-Email header' },
      `${req.method} ${req.originalUrl}`,
    )
  }

  // Validate key length
  if (typeof userKey !== 'string' || userKey.length < MIN_KEY_LENGTH) {
    return logAndRespond(
      res,
      400,
      { error: `X-User-Key must be at least ${MIN_KEY_LENGTH} characters` },
      `${req.method} ${req.originalUrl}`,
    )
  }

  req.userCredentials = { userKey, userEmail }
  return next()
}

export {
  ensureLlmApiKey,
  ensureUserCredentials,
}
