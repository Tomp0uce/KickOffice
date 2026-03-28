/** @typedef {'standard'|'reasoning'|'image'} ModelTier */

/**
 * @typedef {Object} ModelConfig
 * @property {string} id - Model identifier
 * @property {string} label - Display label
 * @property {'chat'|'image'} type - Model type
 * @property {number} [maxTokens] - Max output tokens
 * @property {number} [temperature] - Sampling temperature
 * @property {string|undefined} [reasoningEffort] - Reasoning effort level (low/medium/high)
 */

/**
 * @typedef {Object} ChatBodyParams
 * @property {ModelTier} modelTier
 * @property {ModelConfig} modelConfig
 * @property {Array<object>} messages
 * @property {number|undefined} [temperature]
 * @property {number|undefined} [maxTokens]
 * @property {boolean} stream
 * @property {Array<object>|undefined} [tools]
 */

import logger from '../utils/logger.js';
import { parsePositiveInt } from './env.js';

// QUAL-M5: Use parsePositiveInt so invalid env vars throw at startup rather than silently produce NaN.
const MAX_TOOLS = parsePositiveInt('MAX_TOOLS', 128, 'MAX_TOOLS');
const LLM_API_BASE_URL = process.env.LLM_API_BASE_URL || 'https://litellm.kickmaker.net/v1';
const LLM_API_KEY = process.env.LLM_API_KEY || '';

// Validate required configuration at startup
const isProduction = process.env.NODE_ENV === 'production';
if (!LLM_API_KEY) {
  if (isProduction) {
    throw new Error('FATAL: LLM_API_KEY environment variable is required in production');
  } else {
    logger.warn('[Config] WARNING: LLM_API_KEY is not set. API requests will fail.');
  }
}

const models = {
  standard: {
    id: process.env.MODEL_STANDARD || 'gpt-5.2',
    label: process.env.MODEL_STANDARD_LABEL || 'Standard',
    maxTokens: parsePositiveInt('MODEL_STANDARD_MAX_TOKENS', 32000, 'MODEL_STANDARD_MAX_TOKENS'),
    contextWindow: parsePositiveInt(
      'MODEL_STANDARD_CONTEXT_WINDOW',
      400000,
      'MODEL_STANDARD_CONTEXT_WINDOW',
    ),
    temperature: parseFloat(process.env.MODEL_STANDARD_TEMPERATURE || '0.7'),
    reasoningEffort: process.env.MODEL_STANDARD_REASONING_EFFORT || undefined,
    type: 'chat',
  },
  reasoning: {
    id: process.env.MODEL_REASONING || 'gpt-5.2',
    label: process.env.MODEL_REASONING_LABEL || 'Reasoning',
    maxTokens: parsePositiveInt('MODEL_REASONING_MAX_TOKENS', 65000, 'MODEL_REASONING_MAX_TOKENS'),
    contextWindow: parsePositiveInt(
      'MODEL_REASONING_CONTEXT_WINDOW',
      400000,
      'MODEL_REASONING_CONTEXT_WINDOW',
    ),
    temperature: parseFloat(process.env.MODEL_REASONING_TEMPERATURE || '1'),
    reasoningEffort: process.env.MODEL_REASONING_EFFORT || 'high',
    type: 'chat',
  },
  image: {
    id: process.env.MODEL_IMAGE || 'gpt-image-1',
    label: process.env.MODEL_IMAGE_LABEL || 'Image',
    type: 'image',
  },
};

/** @param {string} modelId @returns {boolean} */
function isGpt5Model(modelId = '') {
  return modelId.toLowerCase().startsWith('gpt-5');
}

/** @param {string} modelId @returns {boolean} */
function isChatGptModel(modelId = '') {
  return modelId.toLowerCase().startsWith('chatgpt-');
}

/** @returns {Record<string, {id: string, label: string, type: string}>} */
function getPublicModels() {
  const publicModels = {};
  for (const [tier, config] of Object.entries(models)) {
    publicModels[tier] = {
      id: config.id,
      label: config.label,
      type: config.type,
      ...(config.contextWindow ? { contextWindow: config.contextWindow } : {}),
    };
  }
  return publicModels;
}

/**
 * Sanitizes messages to remove empty tool_calls arrays (Azure/LiteLLM rejects them).
 * @param {Array<object>} messages
 * @returns {Array<object>}
 */
function sanitizeMessages(messages) {
  return messages.map(msg => {
    if (msg.role === 'assistant' && Array.isArray(msg.tool_calls) && msg.tool_calls.length === 0) {
      const { tool_calls, ...rest } = msg;
      return rest;
    }
    return msg;
  });
}

/**
 * Builds the request body for the LLM API.
 * @param {ChatBodyParams} params
 * @returns {Record<string, unknown>}
 */
function buildChatBody({
  modelTier,
  modelConfig,
  messages,
  temperature,
  maxTokens,
  stream,
  tools,
}) {
  const modelId = modelConfig.id;
  const sanitizedMessages = sanitizeMessages(messages);
  const supportsLegacyParams = !isChatGptModel(modelId);
  const reasoningEffort = isGpt5Model(modelId)
    ? modelConfig.reasoningEffort || undefined
    : undefined;
  const canUseSamplingParams = !isGpt5Model(modelId);
  const body = {
    model: modelId,
    messages: sanitizedMessages,
    stream,
  };

  if (stream) {
    body.stream_options = { include_usage: true };
  }

  if (supportsLegacyParams) {
    const resolvedMaxTokens = maxTokens ?? modelConfig.maxTokens;
    if (resolvedMaxTokens) {
      if (isGpt5Model(modelId)) {
        body.max_completion_tokens = resolvedMaxTokens;
      } else {
        body.max_tokens = resolvedMaxTokens;
      }
    }
  }

  if (supportsLegacyParams) {
    const resolvedTemperature = temperature ?? modelConfig.temperature;
    if (canUseSamplingParams && Number.isFinite(resolvedTemperature)) {
      body.temperature = resolvedTemperature;
    }
  }

  if (tools && tools.length > 0) {
    body.tools = tools;
    // gpt-5.2 on Azure does not support explicit 'tool_choice'
    if (!modelId.toLowerCase().startsWith('gpt-5.2')) {
      body.tool_choice = 'auto';
    }
  }

  if (modelTier !== 'image' && isGpt5Model(modelId) && reasoningEffort) {
    body.reasoning_effort = reasoningEffort;
  }

  return body;
}

export {
  buildChatBody,
  getPublicModels,
  isChatGptModel,
  isGpt5Model,
  LLM_API_BASE_URL,
  LLM_API_KEY,
  models,
  MAX_TOOLS,
};
