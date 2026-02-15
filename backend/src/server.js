import cors from 'cors'
import 'dotenv/config'
import express from 'express'

const app = express()
const PORT = process.env.PORT || 3003
const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:3002'

// --- Configuration ---
const LLM_API_BASE_URL = process.env.LLM_API_BASE_URL || 'https://api.openai.com/v1'
const LLM_API_KEY = process.env.LLM_API_KEY || ''

const models = {
  nano: {
    id: process.env.MODEL_NANO || 'gpt-4.1-nano',
    label: process.env.MODEL_NANO_LABEL || 'Nano (rapide)',
    maxTokens: parseInt(process.env.MODEL_NANO_MAX_TOKENS || '1024', 10),
    temperature: parseFloat(process.env.MODEL_NANO_TEMPERATURE || '0.7'),
    type: 'chat',
  },
  standard: {
    id: process.env.MODEL_STANDARD || 'gpt-4.1',
    label: process.env.MODEL_STANDARD_LABEL || 'Standard',
    maxTokens: parseInt(process.env.MODEL_STANDARD_MAX_TOKENS || '4096', 10),
    temperature: parseFloat(process.env.MODEL_STANDARD_TEMPERATURE || '0.7'),
    type: 'chat',
  },
  reasoning: {
    id: process.env.MODEL_REASONING || 'o3',
    label: process.env.MODEL_REASONING_LABEL || 'Raisonnement',
    maxTokens: parseInt(process.env.MODEL_REASONING_MAX_TOKENS || '8192', 10),
    temperature: parseFloat(process.env.MODEL_REASONING_TEMPERATURE || '1'),
    type: 'chat',
  },
  image: {
    id: process.env.MODEL_IMAGE || 'gpt-image-1',
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
  if (tools.length > 32) return { error: 'tools supports at most 32 entries' }

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

// --- Middleware ---
app.use(cors({
  origin: FRONTEND_URL,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}))
app.use(express.json({ limit: '4mb' }))

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
    return res.status(400).json({ error: 'messages array is required' })
  }

  const modelConfig = models[modelTier]
  if (!modelConfig) {
    return res.status(400).json({ error: `Unknown model tier: ${modelTier}` })
  }

  if (modelConfig.type === 'image') {
    return res.status(400).json({ error: 'Use /api/image for image generation' })
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return res.status(400).json({ error: parsedTemperature.error })
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return res.status(400).json({ error: parsedMaxTokens.error })
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return res.status(400).json({
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    })
  }

  if (!LLM_API_KEY) {
    return res.status(500).json({ error: 'LLM API key not configured on server' })
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
      return res.status(response.status).json({
        error: `LLM API error: ${response.status}`,
        details: errorText,
      })
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
      return res.status(504).json({ error: 'LLM API request timeout' })
    }
    console.error('Chat proxy error:', error)
    res.status(500).json({ error: 'Internal server error' })
  }
})

// --- Chat completion proxy (non-streaming, for agent mode) ---
app.post('/api/chat/sync', async (req, res) => {
  const { messages, modelTier = 'standard', temperature, maxTokens, tools } = req.body

  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ error: 'messages array is required' })
  }

  const modelConfig = models[modelTier]
  if (!modelConfig || modelConfig.type === 'image') {
    return res.status(400).json({ error: `Invalid model tier for chat: ${modelTier}` })
  }

  if (!LLM_API_KEY) {
    return res.status(500).json({ error: 'LLM API key not configured on server' })
  }

  const parsedTemperature = validateTemperature(temperature)
  if (parsedTemperature.error) {
    return res.status(400).json({ error: parsedTemperature.error })
  }

  const parsedMaxTokens = validateMaxTokens(maxTokens)
  if (parsedMaxTokens.error) {
    return res.status(400).json({ error: parsedMaxTokens.error })
  }

  if (isChatGptModel(modelConfig.id) && (temperature !== undefined || maxTokens !== undefined)) {
    return res.status(400).json({
      error: 'temperature and maxTokens are not supported for ChatGPT models',
    })
  }

  const parsedTools = validateTools(tools)
  if (parsedTools.error) {
    return res.status(400).json({ error: parsedTools.error })
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
      return res.status(response.status).json({
        error: `LLM API error: ${response.status}`,
        details: errorText,
      })
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return res.status(504).json({ error: 'LLM API request timeout' })
    }
    console.error('Chat sync proxy error:', error)
    res.status(500).json({ error: 'Internal server error' })
  }
})

// --- Image generation proxy ---
app.post('/api/image', async (req, res) => {
  const { prompt, size = '1024x1024', quality = 'auto', n = 1 } = req.body

  const allowedSizes = new Set(['1024x1024', '1024x1536', '1536x1024'])
  const allowedQualities = new Set(['low', 'medium', 'high', 'auto'])
  const maxPromptLength = 4000

  if (!prompt) {
    return res.status(400).json({ error: 'prompt is required' })
  }
  if (typeof prompt !== 'string') {
    return res.status(400).json({ error: 'prompt must be a string' })
  }
  if (prompt.length > maxPromptLength) {
    return res.status(400).json({ error: `prompt must be <= ${maxPromptLength} characters` })
  }
  if (typeof size !== 'string' || !allowedSizes.has(size)) {
    return res.status(400).json({ error: `size must be one of: ${[...allowedSizes].join(', ')}` })
  }
  if (typeof quality !== 'string' || !allowedQualities.has(quality)) {
    return res.status(400).json({ error: `quality must be one of: ${[...allowedQualities].join(', ')}` })
  }
  if (!Number.isInteger(n) || n < 1 || n > 4) {
    return res.status(400).json({ error: 'n must be an integer between 1 and 4' })
  }

  const imageModel = models.image
  if (!imageModel) {
    return res.status(500).json({ error: 'Image model not configured' })
  }

  if (!LLM_API_KEY) {
    return res.status(500).json({ error: 'LLM API key not configured on server' })
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
      return res.status(response.status).json({
        error: `Image API error: ${response.status}`,
        details: errorText,
      })
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return res.status(504).json({ error: 'Image API request timeout' })
    }
    console.error('Image proxy error:', error)
    res.status(500).json({ error: 'Internal server error' })
  }
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
