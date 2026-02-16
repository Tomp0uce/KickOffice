import cors from 'cors'
import 'dotenv/config'
import express from 'express'
import rateLimit from 'express-rate-limit'
import morgan from 'morgan'

const app = express()
const PORT = process.env.PORT || 3003
const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:3002'

// --- Configuration ---
const LLM_API_BASE_URL = process.env.LLM_API_BASE_URL || 'https://api.openai.com/v1'
const LLM_API_KEY = process.env.LLM_API_KEY || ''
const MAX_TOOLS = parseInt(process.env.MAX_TOOLS || '128', 10)
const CHAT_RATE_LIMIT_WINDOW_MS = parseInt(process.env.CHAT_RATE_LIMIT_WINDOW_MS || '60000', 10)
const CHAT_RATE_LIMIT_MAX = parseInt(process.env.CHAT_RATE_LIMIT_MAX || '20', 10)
const IMAGE_RATE_LIMIT_WINDOW_MS = parseInt(process.env.IMAGE_RATE_LIMIT_WINDOW_MS || '60000', 10)
const IMAGE_RATE_LIMIT_MAX = parseInt(process.env.IMAGE_RATE_LIMIT_MAX || '5', 10)

const models = {
  nano: {
    id: process.env.MODEL_NANO || 'gpt-5-nano',
    label: process.env.MODEL_NANO_LABEL || 'Nano (rapide)',
    maxTokens: parseInt(process.env.MODEL_NANO_MAX_TOKENS || '4096', 10),
    temperature: parseFloat(process.env.MODEL_NANO_TEMPERATURE || '0.7'),
    type: 'chat',
  },
  standard: {
    id: process.env.MODEL_STANDARD || 'gpt-5-mini',
    label: process.env.MODEL_STANDARD_LABEL || 'Standard',
    maxTokens: parseInt(process.env.MODEL_STANDARD_MAX_TOKENS || '4096', 10),
    temperature: parseFloat(process.env.MODEL_STANDARD_TEMPERATURE || '0.7'),
    type: 'chat',
  },
  reasoning: {
    id: process.env.MODEL_REASONING || 'gpt-5.2',
    label: process.env.MODEL_REASONING_LABEL || 'Raisonnement',
    maxTokens: parseInt(process.env.MODEL_REASONING_MAX_TOKENS || '8192', 10),
    temperature: parseFloat(process.env.MODEL_REASONING_TEMPERATURE || '1'),
    type: 'chat',
  },
  image: {
    id: process.env.MODEL_IMAGE || 'gpt-image-1.5',
    label: process.env.MODEL_IMAGE_LABEL || 'Image',
    type: 'image',
  },
}


function isGpt5Model(modelId = '') {
  return modelId.toLowerCase().startsWith('gpt-5')
}

function isChatGptModel(modelId = '') {
  return modelId.toLowerCase().startsWith('chatgpt-')
}

function isPlainObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateTemperature(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isFinite(value)) return { error: 'temperature must be a finite number' }
  if (value < 0 || value > 2) return { error: 'temperature must be between 0 and 2' }
  return { value }
}

function validateMaxTokens(value) {
  if (value === undefined) return { value: undefined }
  if (!Number.isInteger(value)) return { error: 'maxTokens must be an integer' }
  if (value < 1 || value > 32768) return { error: 'maxTokens must be between 1 and 32768' }
  return { value }
}

function validateTools(tools) {
  if (tools === undefined) return { value: undefined }
  if (!Array.isArray(tools)) return { error: 'tools must be an array' }
  if (tools.length === 0) return { value: undefined }
  if (tools.length > MAX_TOOLS) return { error: `tools supports at most ${MAX_TOOLS} entries` }

  const sanitizedTools = []
  for (const tool of tools) {
    if (!isPlainObject(tool)) return { error: 'each tool must be an object' }
    if (tool.type !== 'function') return { error: 'only function tools are supported' }
    if (!isPlainObject(tool.function)) return { error: 'tool.function must be an object' }
    const { name, description, parameters, strict } = tool.function
    if (typeof name !== 'string' || !/^[a-zA-Z0-9_-]{1,64}$/.test(name)) {
      return { error: 'tool.function.name must match /^[a-zA-Z0-9_-]{1,64}$/' }
    }
    if (description !== undefined && typeof description !== 'string') {
      return { error: 'tool.function.description must be a string' }
    }
    if (!isPlainObject(parameters)) {
      return { error: 'tool.function.parameters must be a JSON schema object' }
    }
    if (strict !== undefined && typeof strict !== 'boolean') {
      return { error: 'tool.function.strict must be a boolean' }
    }

    sanitizedTools.push({
      type: 'function',
      function: {
        name,
        ...(description !== undefined ? { description } : {}),
        parameters,
        ...(strict !== undefined ? { strict } : {}),
      },
    })
  }

  return { value: sanitizedTools }
}

function getChatTimeoutMs(modelTier) {
  if (modelTier === 'nano') return 60_000
  if (modelTier === 'reasoning') return 300_000
  return 120_000
}

function getImageTimeoutMs() {
  return 180_000
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

function buildChatBody({ modelConfig, messages, temperature, maxTokens, stream, tools }) {
  const modelId = modelConfig.id
  const supportsLegacyParams = !isChatGptModel(modelId)
  const body = {
    model: modelId,
    messages,
    stream,
  }

  if (supportsLegacyParams) {
    const resolvedMaxTokens = maxTokens ?? modelConfig.maxTokens
    if (resolvedMaxTokens) {
      if (isGpt5Model(modelId)) {
        body.max_completion_tokens = resolvedMaxTokens
      } else {
        body.max_tokens = resolvedMaxTokens
      }
    }
  }

  if (supportsLegacyParams) {
    const resolvedTemperature = temperature ?? modelConfig.temperature
    if (!isGpt5Model(modelId) && Number.isFinite(resolvedTemperature)) {
      body.temperature = resolvedTemperature
    }
  }

  if (tools && tools.length > 0) {
    body.tools = tools
    body.tool_choice = 'auto'
  }

  return body
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

const chatLimiter = rateLimit({
  windowMs: CHAT_RATE_LIMIT_WINDOW_MS,
  max: CHAT_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many chat requests, please try again later.' },
})

const imageLimiter = rateLimit({
  windowMs: IMAGE_RATE_LIMIT_WINDOW_MS,
  max: IMAGE_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many image requests, please try again later.' },
})

// --- Middleware ---
app.use(cors({
  origin: FRONTEND_URL,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}))
app.use(express.json({ limit: '4mb' }))
app.use(morgan(':method :url :status :res[content-length] - :response-time ms'))
app.use('/api/chat', chatLimiter)
app.use('/api/image', imageLimiter)

// --- Health Check ---
app.get('/health', (_req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    version: '1.0.0',
  })
})

// --- Get available models (no secrets exposed) ---
app.get('/api/models', (_req, res) => {
  const publicModels = {}
  for (const [tier, config] of Object.entries(models)) {
    publicModels[tier] = {
      id: config.id,
      label: config.label,
      type: config.type,
    }
  }
  res.json(publicModels)
})

// --- Chat completion proxy (streaming) ---
app.post('/api/chat', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens } = req.body

  if (!messages || !Array.isArray(messages)) {
    return logAndRespond(res, 400, { error: 'messages array is required' }, 'POST /api/chat')
  }

  const modelConfig = models[modelTier]
  if (!modelConfig) {
    return logAndRespond(res, 400, { error: `Unknown model tier: ${modelTier}` }, 'POST /api/chat')
  }

  if (modelConfig.type === 'image') {
    return logAndRespond(res, 400, { error: 'Use /api/image for image generation' }, 'POST /api/chat')
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return logAndRespond(res, 400, { error: parsedTemperature.error }, 'POST /api/chat')
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return logAndRespond(res, 400, { error: parsedMaxTokens.error }, 'POST /api/chat')
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return logAndRespond(res, 400, {
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    }, 'POST /api/chat')
  }

  if (!LLM_API_KEY) {
    return logAndRespond(res, 500, { error: 'LLM API key not configured on server' }, 'POST /api/chat')
  }

  try {
    const body = buildChatBody({
      modelConfig,
      messages,
      temperature,
      maxTokens,
      stream: true,
    })

    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    }, getChatTimeoutMs(modelTier))

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`LLM API error ${response.status}:`, errorText)
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat')
    }

    // Stream the response
    res.setHeader('Content-Type', 'text/event-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.setHeader('Connection', 'keep-alive')

    const reader = response.body.getReader()
    const decoder = new TextDecoder()

    try {
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        const chunk = decoder.decode(value, { stream: true })
        res.write(chunk)
      }
    } catch (streamError) {
      console.error('Stream error:', streamError)
    } finally {
      res.end()
    }
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat')
    }
    console.error('Chat proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat')
  }
})

// --- Chat completion proxy (non-streaming, for agent mode) ---
app.post('/api/chat/sync', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body

  if (!messages || !Array.isArray(messages)) {
    return logAndRespond(res, 400, { error: 'messages array is required' }, 'POST /api/chat/sync')
  }

  const modelConfig = models[modelTier]
  if (!modelConfig || modelConfig.type === 'image') {
    return logAndRespond(res, 400, { error: `Invalid model tier for chat: ${modelTier}` }, 'POST /api/chat/sync')
  }

  if (!LLM_API_KEY) {
    return logAndRespond(res, 500, { error: 'LLM API key not configured on server' }, 'POST /api/chat/sync')
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return logAndRespond(res, 400, { error: parsedTemperature.error }, 'POST /api/chat/sync')
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return logAndRespond(res, 400, { error: parsedMaxTokens.error }, 'POST /api/chat/sync')
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return logAndRespond(res, 400, {
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    }, 'POST /api/chat/sync')
  }

  const parsedTools = validateTools(tools)
  if (parsedTools.error) {
    return logAndRespond(res, 400, { error: parsedTools.error }, 'POST /api/chat/sync')
  }

  try {
    const body = buildChatBody({
      modelConfig,
      messages,
      temperature,
      maxTokens,
      stream: false,
      tools: parsedTools.value,
    })

    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    }, getChatTimeoutMs(modelTier))

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`LLM API error ${response.status}:`, errorText)
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/chat/sync')
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'LLM API request timeout' }, 'POST /api/chat/sync')
    }
    console.error('Chat sync proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/chat/sync')
  }
})

// --- Image generation proxy ---
app.post('/api/image', async (req, res) => {
  const { prompt, size = '1024x1024', quality = 'auto', n = 1 } = req.body

  const allowedSizes = new Set(['1024x1024', '1024x1536', '1536x1024'])
  const allowedQualities = new Set(['low', 'medium', 'high', 'auto'])
  const maxPromptLength = 4000

  if (!prompt) {
    return logAndRespond(res, 400, { error: 'prompt is required' }, 'POST /api/image')
  }
  if (typeof prompt !== 'string') {
    return logAndRespond(res, 400, { error: 'prompt must be a string' }, 'POST /api/image')
  }
  if (prompt.length > maxPromptLength) {
    return logAndRespond(res, 400, { error: `prompt must be <= ${maxPromptLength} characters` }, 'POST /api/image')
  }
  if (typeof size !== 'string' || !allowedSizes.has(size)) {
    return logAndRespond(res, 400, { error: `size must be one of: ${[...allowedSizes].join(', ')}` }, 'POST /api/image')
  }
  if (typeof quality !== 'string' || !allowedQualities.has(quality)) {
    return logAndRespond(res, 400, { error: `quality must be one of: ${[...allowedQualities].join(', ')}` }, 'POST /api/image')
  }
  if (!Number.isInteger(n) || n < 1 || n > 4) {
    return logAndRespond(res, 400, { error: 'n must be an integer between 1 and 4' }, 'POST /api/image')
  }

  const imageModel = models.image
  if (!imageModel) {
    return logAndRespond(res, 500, { error: 'Image model not configured' }, 'POST /api/image')
  }

  if (!LLM_API_KEY) {
    return logAndRespond(res, 500, { error: 'LLM API key not configured on server' }, 'POST /api/image')
  }

  try {
    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/images/generations`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify({
        model: imageModel.id,
        prompt,
        size,
        quality,
        n,
      }),
    }, getImageTimeoutMs())

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`Image API error ${response.status}:`, errorText)
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/image')
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'Image API request timeout' }, 'POST /api/image')
    }
    console.error('Image proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/image')
  }
})

// --- Fallback + error middleware ---
app.use((req, res) => {
  return logAndRespond(res, 404, { error: 'Route not found' }, `${req.method} ${req.originalUrl}`)
})

app.use((err, req, res, next) => {
  console.error(`[${req.method} ${req.originalUrl}] Unhandled server error:`, err)
  if (res.headersSent) {
    return next(err)
  }
  return logAndRespond(res, 500, { error: 'Internal server error' }, `${req.method} ${req.originalUrl}`)
})

// --- Start server ---
app.listen(PORT, '0.0.0.0', () => {
  console.log(`KickOffice backend running on port ${PORT}`)
  console.log(`CORS allowed origin: ${FRONTEND_URL}`)
  console.log(`LLM API base URL: ${LLM_API_BASE_URL}`)
  console.log(`Models configured:`)
  for (const [tier, config] of Object.entries(models)) {
    console.log(`  ${tier}: ${config.id} (${config.label})`)
  }
})
