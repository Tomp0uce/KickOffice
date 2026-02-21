const PORT = process.env.PORT || 3003
const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:3002'
const PUBLIC_FRONTEND_URL = process.env.PUBLIC_FRONTEND_URL

/**
 * Parses an integer from environment variable with validation.
 * Throws on NaN or negative values.
 */
function parsePositiveInt(envVar, defaultValue, name) {
  const raw = process.env[envVar]
  if (raw === undefined) return defaultValue
  const parsed = parseInt(raw, 10)
  if (Number.isNaN(parsed)) {
    throw new Error(`Invalid ${name}: "${raw}" is not a valid integer`)
  }
  if (parsed < 0) {
    throw new Error(`Invalid ${name}: must be a positive integer`)
  }
  return parsed
}

const CHAT_RATE_LIMIT_WINDOW_MS = parsePositiveInt('CHAT_RATE_LIMIT_WINDOW_MS', 60000, 'CHAT_RATE_LIMIT_WINDOW_MS')
const CHAT_RATE_LIMIT_MAX = parsePositiveInt('CHAT_RATE_LIMIT_MAX', 20, 'CHAT_RATE_LIMIT_MAX')
const IMAGE_RATE_LIMIT_WINDOW_MS = parsePositiveInt('IMAGE_RATE_LIMIT_WINDOW_MS', 60000, 'IMAGE_RATE_LIMIT_WINDOW_MS')
const IMAGE_RATE_LIMIT_MAX = parsePositiveInt('IMAGE_RATE_LIMIT_MAX', 5, 'IMAGE_RATE_LIMIT_MAX')

export {
  CHAT_RATE_LIMIT_MAX,
  CHAT_RATE_LIMIT_WINDOW_MS,
  FRONTEND_URL,
  IMAGE_RATE_LIMIT_MAX,
  IMAGE_RATE_LIMIT_WINDOW_MS,
  PORT,
  PUBLIC_FRONTEND_URL,
}
