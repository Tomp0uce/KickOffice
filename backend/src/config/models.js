const MAX_TOOLS = parseInt(process.env.MAX_TOOLS || '128', 10)
const LLM_API_BASE_URL = process.env.LLM_API_BASE_URL || 'https://api.openai.com/v1'
const LLM_API_KEY = process.env.LLM_API_KEY || ''

const models = {
  standard: {
    id: process.env.MODEL_STANDARD || 'gpt-5.2',
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
    reasoningEffort: process.env.MODEL_REASONING_EFFORT || 'high',
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

function getPublicModels() {
  const publicModels = {}
  for (const [tier, config] of Object.entries(models)) {
    publicModels[tier] = {
      id: config.id,
      label: config.label,
      type: config.type,
    }
  }
  return publicModels
}

function buildChatBody({ modelTier, modelConfig, messages, temperature, maxTokens, stream, tools }) {
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

  if (modelTier === 'reasoning' && isGpt5Model(modelId)) {
    body.reasoning_effort = modelConfig.reasoningEffort
  }

  return body
}

export {
  buildChatBody,
  getPublicModels,
  isChatGptModel,
  LLM_API_BASE_URL,
  LLM_API_KEY,
  models,
  MAX_TOOLS,
}
