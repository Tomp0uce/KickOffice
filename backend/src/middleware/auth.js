import { LLM_API_KEY } from '../config/models.js'
import { logAndRespond } from '../utils/http.js'

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
  req.userCredentials = { userKey, userEmail }
  return next()
}

export {
  ensureLlmApiKey,
  ensureUserCredentials,
}
