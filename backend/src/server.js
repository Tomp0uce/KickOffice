import cors from 'cors'
import 'dotenv/config'
import express from 'express'
import rateLimit from 'express-rate-limit'
import helmet from 'helmet'
import crypto from 'crypto'
import logger from './utils/logger.js'

import { models } from './config/models.js'
import {
  CHAT_RATE_LIMIT_MAX,
  CHAT_RATE_LIMIT_WINDOW_MS,
  FRONTEND_URL,
  IMAGE_RATE_LIMIT_MAX,
  IMAGE_RATE_LIMIT_WINDOW_MS,
  PORT,
  PUBLIC_FRONTEND_URL,
} from './config/env.js'

import { ensureLlmApiKey, ensureUserCredentials } from './middleware/auth.js'
import { chatRouter } from './routes/chat.js'
import { healthRouter } from './routes/health.js'
import { imageRouter } from './routes/image.js'
import { modelsRouter } from './routes/models.js'
import { uploadRouter } from './routes/upload.js'
import { feedbackRouter } from './routes/feedback.js'
import { logsRouter } from './routes/logs.js'
import { plotDigitizerRouter } from './routes/plotDigitizer.js'
import { filesRouter } from './routes/files.js'
import { logAndRespond } from './utils/http.js'

const isProduction = process.env.NODE_ENV === 'production'
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '600000', 10) // 10 minutes default

const app = express()

// Trust proxy for Synology/nginx reverse proxy compatibility
// Allows express-rate-limit to correctly identify client IPs via X-Forwarded-For
app.set('trust proxy', true)

const chatLimiter = rateLimit({
  windowMs: CHAT_RATE_LIMIT_WINDOW_MS,
  max: CHAT_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many chat requests.' },
})

const imageLimiter = rateLimit({
  windowMs: IMAGE_RATE_LIMIT_WINDOW_MS,
  max: IMAGE_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many image requests.' },
})

// Rate limiter for lightweight info endpoints (health, models)
const infoLimiter = rateLimit({
  windowMs: 60_000, // 1 minute
  max: 120, // generous limit for info endpoints
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many requests.' },
})

// Rate limiter for frontend log submission
const logsLimiter = rateLimit({
  windowMs: 60_000, // 1 minute
  max: 20, // max 20 log batches per minute per IP
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many log requests. Please try again later.' },
})

// Rate limiter for file upload endpoint to prevent memory exhaustion
const uploadLimiter = rateLimit({
  windowMs: 60_000, // 1 minute
  max: 10, // max 10 uploads per minute per IP
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many upload requests. Please try again later.' },
})

const allowedOrigins = [FRONTEND_URL]
if (PUBLIC_FRONTEND_URL) {
  allowedOrigins.push(PUBLIC_FRONTEND_URL)
}

app.use(cors({
  origin: allowedOrigins,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-User-Key', 'X-User-Email', 'x-csrf-token', 'X-Office-Host', 'X-Session-Id'],
  exposedHeaders: ['X-Request-Id'],
  credentials: true,
}))

app.use(helmet({
  contentSecurityPolicy: false,
  crossOriginEmbedderPolicy: false,
  // Office add-ins load inside an iframe inside Outlook/Word/Excel/PowerPoint.
  // helmet's default frameguard sets X-Frame-Options: SAMEORIGIN which blocks that iframe.
  // framing protection is handled by CSP frame-ancestors in nginx.conf instead.
  frameguard: false,
  // Enable HSTS in production for HTTPS enforcement
  strictTransportSecurity: isProduction ? {
    maxAge: 31536000, // 1 year
    includeSubDomains: true,
    preload: true,
  } : false,
}))

// Request timeout middleware - prevents hanging requests
app.use((req, res, next) => {
  req.setTimeout(REQUEST_TIMEOUT_MS, () => {
    if (!res.headersSent) {
      res.status(408).json({ error: 'Request timeout' })
    }
  })
  next()
})


// Add request ID and context logger
app.use((req, res, next) => {
  res.locals.reqId = crypto.randomUUID()
  res.locals.userId = req.headers['x-user-email'] || 'anonymous'
  res.locals.host = req.headers['x-office-host'] || 'unknown'

  res.setHeader('X-Request-Id', res.locals.reqId)

  req.logger = logger.child({
    reqId: res.locals.reqId,
    userId: res.locals.userId,
    host: res.locals.host
  })
  next()
})

app.use(express.json({ limit: '32mb' }))

// HTTP Request Logger
app.use((req, res, next) => {
  const start = Date.now()
  res.on('finish', () => {
    const duration = Date.now() - start
    const isHealth = req.path.startsWith('/api/health')
    const traffic = isHealth ? 'auto' : 
                   (req.path.startsWith('/api/chat') || req.path.startsWith('/api/image')) ? 'llm' : 'user'
    
    // Fall back to main logger if req.logger isn't available
    const reqLogger = req.logger || logger
    const logLevel = isHealth ? 'debug' : 'http'
    
    reqLogger.log(logLevel, `${req.method} ${req.originalUrl} ${res.statusCode} ${duration}ms`, {
      traffic,
      method: req.method,
      url: req.originalUrl,
      status: res.statusCode,
      duration
    })
  })
  next()
})

// Reject POST/PUT/PATCH requests that don't declare application/json to avoid silent empty bodies
app.use((req, res, next) => {
  if (req.path === '/api/upload' || req.path.startsWith('/api/files')) return next()
  if (['POST', 'PUT', 'PATCH'].includes(req.method) && !req.is('application/json')) {
    return res.status(415).json({ error: 'Content-Type must be application/json' })
  }
  next()
})

// CSRF Protection Mechanism
app.use((req, res, next) => {
  // Exempt healthcheck from CSRF if we want, but letting it generate a cookie is fine.
  let token = null
  if (req.headers.cookie) {
    const match = req.headers.cookie.match(/(?:^| )csrf_token=([^;]+)/)
    if (match) token = match[1]
  }

  if (!token) {
    token = crypto.randomUUID()
    res.cookie('csrf_token', token, {
      httpOnly: false,
      secure: isProduction,
      sameSite: 'none'
    })
  }

  if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(req.method)) {
    // Validate origin for CSRF protection
    const origin = req.headers.origin || req.headers.referer
    const isValidOrigin = origin && allowedOrigins.some(allowed => origin.startsWith(allowed))

    // Skip CSRF check for requests authenticated via the X-User-Key header
    // (Office add-in requests can't reliably carry cookies cross-origin)
    const isKeyAuthenticated = !!req.headers['x-user-key']

    if (!isKeyAuthenticated) {
      // Enforce origin validation for non-key-authenticated requests
      if (!isValidOrigin) {
        return res.status(403).json({ error: 'Invalid origin' })
      }

      const headerToken = req.headers['x-csrf-token']
      if (!headerToken || headerToken !== token) {
        return res.status(403).json({ error: 'Invalid CSRF token' })
      }
    }
  }

  next()
})



app.use(infoLimiter, healthRouter)
app.use(infoLimiter, modelsRouter)
app.use('/api/chat', ensureLlmApiKey, ensureUserCredentials, chatLimiter, chatRouter)
app.use('/api/image', ensureLlmApiKey, ensureUserCredentials, imageLimiter, imageRouter)
app.use('/api/upload', ensureUserCredentials, uploadLimiter, uploadRouter)
app.use('/api/feedback', ensureUserCredentials, feedbackRouter)
app.use('/api/logs', ensureUserCredentials, logsLimiter, logsRouter)
app.use('/api/chart-extract', ensureUserCredentials, uploadLimiter, plotDigitizerRouter)
app.use('/api/files', ensureLlmApiKey, ensureUserCredentials, uploadLimiter, filesRouter)

app.use((req, res) => {
  return logAndRespond(res, 404, { error: 'Route not found' }, `${req.method} ${req.originalUrl}`)
})

app.use((err, req, res, next) => {
  if (res.headersSent) {
    return next(err)
  }
  logger.error(`Unhandled error: ${err.message}`, {
    error: {
      name: err.name,
      message: err.message,
      stack: err.stack
    },
    traffic: 'system'
  })
  return logAndRespond(res, 500, { error: 'Internal server error' }, 'SERVER')
})

const server = app.listen(PORT, '0.0.0.0', () => {
  logger.info(`KickOffice backend running on port ${PORT}`, { traffic: 'system' })
  logger.info('Models configured:', { traffic: 'system' })
  for (const [tier, config] of Object.entries(models)) {
    logger.info(`  ${tier}: ${config.id} (${config.label})`, { traffic: 'system' })
  }
})

// Graceful shutdown: stop accepting new connections and wait for in-flight requests
function shutdown(signal) {
  logger.info(`${signal} received — shutting down gracefully`, { traffic: 'system' })
  server.close(() => {
    logger.info('All connections closed. Exiting.', { traffic: 'system' })
    process.exit(0)
  })
  // Force exit after 30s if connections don't drain
  setTimeout(() => {
    logger.error('Graceful shutdown timeout — forcing exit', { traffic: 'system' })
    process.exit(1)
  }, 30_000).unref()
}

process.on('SIGTERM', () => shutdown('SIGTERM'))
process.on('SIGINT', () => shutdown('SIGINT'))
