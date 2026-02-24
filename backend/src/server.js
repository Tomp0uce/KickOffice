import cors from 'cors'
import 'dotenv/config'
import express from 'express'
import rateLimit from 'express-rate-limit'
import helmet from 'helmet'
import morgan from 'morgan'

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

const isProduction = process.env.NODE_ENV === 'production'
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '600000', 10) // 10 minutes default
import { ensureLlmApiKey, ensureUserCredentials } from './middleware/auth.js'
import { chatRouter } from './routes/chat.js'
import { healthRouter } from './routes/health.js'
import { imageRouter } from './routes/image.js'
import { modelsRouter } from './routes/models.js'
import { logAndRespond } from './utils/http.js'

const app = express()

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

const allowedOrigins = [FRONTEND_URL]
if (PUBLIC_FRONTEND_URL) {
  allowedOrigins.push(PUBLIC_FRONTEND_URL)
}

app.use(cors({
  origin: allowedOrigins,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-User-Key', 'X-User-Email'],
}))

app.use(helmet({
  contentSecurityPolicy: false,
  crossOriginEmbedderPolicy: false,
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

app.use(express.json({ limit: '4mb' }))

// Reject POST/PUT/PATCH requests that don't declare application/json to avoid silent empty bodies
app.use((req, res, next) => {
  if (['POST', 'PUT', 'PATCH'].includes(req.method) && !req.is('application/json')) {
    return res.status(415).json({ error: 'Content-Type must be application/json' })
  }
  next()
})

app.use(morgan(':method :url :status :res[content-length] - :response-time ms'))

app.use(infoLimiter, healthRouter)
app.use(infoLimiter, modelsRouter)
app.use('/api/chat', ensureLlmApiKey, ensureUserCredentials, chatLimiter, chatRouter)
app.use('/api/image', ensureLlmApiKey, ensureUserCredentials, imageLimiter, imageRouter)

app.use((req, res) => {
  return logAndRespond(res, 404, { error: 'Route not found' }, `${req.method} ${req.originalUrl}`)
})

app.use((err, req, res, next) => {
  if (res.headersSent) {
    return next(err)
  }
  console.error('Unhandled error:', err)
  return logAndRespond(res, 500, { error: 'Internal server error' }, 'SERVER')
})

app.listen(PORT, '0.0.0.0', () => {
  console.log(`KickOffice backend running on port ${PORT}`)
  console.log('Models configured:')
  for (const [tier, config] of Object.entries(models)) {
    console.log(`  ${tier}: ${config.id} (${config.label})`)
  }
})
