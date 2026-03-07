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

const consoleFormat = winston.format.combine(
  winston.format.colorize(),
  winston.format.timestamp(),
  winston.format.printf(({ timestamp, level, message, traffic, source, host, sessionId, userId, reqId, ...meta }) => {
    let msg = `[${timestamp}] [${level}] `
    const ctx = []
    if (traffic) ctx.push(`traffic:${traffic}`)
    if (host) ctx.push(`host:${host}`)
    if (userId) ctx.push(`user:${userId}`)
    if (reqId) ctx.push(`reqId:${reqId}`)
    
    if (ctx.length) msg += `(${ctx.join('|')}) `
    msg += message
    
    // Avoid double logging of error stacks if present in meta
    if (meta.error && meta.error.stack) {
       msg += `\n${meta.error.stack}`
       delete meta.error
    }
    
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
