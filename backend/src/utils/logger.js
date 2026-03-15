import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import winston from 'winston'
import DailyRotateFile from 'winston-daily-rotate-file'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const LOGS_DIR = path.join(__dirname, '../../logs')
if (!fs.existsSync(LOGS_DIR)) {
  fs.mkdirSync(LOGS_DIR, { recursive: true })
}

const isProduction = process.env.NODE_ENV === 'production'

// Visual traffic prefix for console output
function trafficPrefix(traffic) {
  switch (traffic) {
    case 'llm':    return '⇄ LLM'
    case 'user':   return '→ USER'
    case 'system': return '⚙  SYS'
    case 'auto':   return '∿ POLL'
    default:       return '·     '
  }
}

const consoleFormat = winston.format.combine(
  winston.format.colorize(),
  winston.format.timestamp({ format: 'HH:mm:ss' }),
  winston.format.printf(({ timestamp, level, message, traffic, source, host, userId, reqId, body, responseContent, ...meta }) => {
    const prefix = trafficPrefix(traffic)
    const ctx = []
    if (host && host !== 'unknown') ctx.push(host)
    if (userId && userId !== 'anonymous') ctx.push(userId.split('@')[0]) // short user
    if (reqId) ctx.push(reqId.slice(0, 8)) // first 8 chars of UUID

    let msg = `[${timestamp}] [${level}] ${prefix}  `
    if (ctx.length) msg += `(${ctx.join('|')}) `
    msg += message

    // Avoid double logging of error stacks if present in meta
    if (meta.error && meta.error.stack) {
      msg += `\n${meta.error.stack}`
      delete meta.error
    }

    // Omit full LLM request/response bodies from console (too verbose)
    // They are still written to file logs in full JSON
    const metaStr = Object.keys(meta).length ? JSON.stringify(meta) : ''
    return msg + (metaStr ? ` ${metaStr}` : '')
  })
)

const fileFormat = winston.format.combine(
  winston.format.timestamp(),
  winston.format.json()
)

const logger = winston.createLogger({
  level: isProduction ? 'info' : 'debug',
  levels: {
    error: 0,
    warn: 1,
    info: 2,
    http: 3,
    debug: 4
  },
  format: fileFormat,
  defaultMeta: { source: 'backend' },
  transports: [
    new DailyRotateFile({
      dirname: LOGS_DIR,
      filename: 'kickoffice-%DATE%.log',
      datePattern: 'YYYY-MM-DD',
      zippedArchive: true,
      maxSize: '10m',
      maxFiles: '7d'
    })
  ]
})

if (!isProduction) {
  logger.add(new winston.transports.Console({
    format: consoleFormat
  }))
} else {
  logger.add(new winston.transports.Console({
    format: fileFormat
  }))
}

export default logger
