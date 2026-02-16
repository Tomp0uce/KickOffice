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

export {
  ensureLlmApiKey,
}
