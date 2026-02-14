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

// --- Middleware ---
app.use(cors({
  origin: FRONTEND_URL,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}))
app.use(express.json({ limit: '10mb' }))

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

  if (!LLM_API_KEY) {
    return res.status(500).json({ error: 'LLM API key not configured on server' })
  }

  try {
    const body = {
      model: modelConfig.id,
      messages,
      temperature: temperature ?? modelConfig.temperature,
      max_tokens: maxTokens ?? modelConfig.maxTokens,
      stream: true,
    }

    const response = await fetch(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    })

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

  try {
    const body = {
      model: modelConfig.id,
      messages,
      temperature: temperature ?? modelConfig.temperature,
      max_tokens: maxTokens ?? modelConfig.maxTokens,
      stream: false,
    }

    if (tools && tools.length > 0) {
      body.tools = tools
      body.tool_choice = 'auto'
    }

    const response = await fetch(`${LLM_API_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify(body),
    })

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
    console.error('Chat sync proxy error:', error)
    res.status(500).json({ error: 'Internal server error' })
  }
})

// --- Image generation proxy ---
app.post('/api/image', async (req, res) => {
  const { prompt, size = '1024x1024', quality = 'auto', n = 1 } = req.body

  if (!prompt) {
    return res.status(400).json({ error: 'prompt is required' })
  }

  const imageModel = models.image
  if (!imageModel) {
    return res.status(500).json({ error: 'Image model not configured' })
  }

  if (!LLM_API_KEY) {
    return res.status(500).json({ error: 'LLM API key not configured on server' })
  }

  try {
    const response = await fetch(`${LLM_API_BASE_URL}/images/generations`, {
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
        response_format: 'b64_json',
      }),
    })

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
