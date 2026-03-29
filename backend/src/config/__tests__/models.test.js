import { describe, it, expect, vi } from 'vitest';

// Mock the global logger to prevent console output and DailyRotateFile initialization.
// models.js calls logger.warn() at import time when LLM_API_KEY is not set.
vi.mock('../../utils/logger.js', () => ({
  default: { error: vi.fn(), warn: vi.fn(), info: vi.fn(), debug: vi.fn() },
}));

import { buildChatBody, isGpt5Model, isChatGptModel } from '../models.js';

// ─── isGpt5Model ───────────────────────────────────────────────────────────────

describe('isGpt5Model', () => {
  it('returns true for gpt-5 prefixed models', () => {
    expect(isGpt5Model('gpt-5.1')).toBe(true);
    expect(isGpt5Model('gpt-5.2-turbo')).toBe(true);
    expect(isGpt5Model('GPT-5.1')).toBe(true);
  });

  it('returns false for non gpt-5 models', () => {
    expect(isGpt5Model('gpt-4o')).toBe(false);
    expect(isGpt5Model('chatgpt-4o-latest')).toBe(false);
    expect(isGpt5Model('claude-3.5-sonnet')).toBe(false);
    expect(isGpt5Model('')).toBe(false);
    expect(isGpt5Model()).toBe(false);
  });
});

// ─── isChatGptModel ────────────────────────────────────────────────────────────

describe('isChatGptModel', () => {
  it('returns true for chatgpt- prefixed models', () => {
    expect(isChatGptModel('chatgpt-4o-latest')).toBe(true);
    expect(isChatGptModel('ChatGPT-4o')).toBe(true);
  });

  it('returns false for other models', () => {
    expect(isChatGptModel('gpt-5.1')).toBe(false);
    expect(isChatGptModel('gpt-4o')).toBe(false);
    expect(isChatGptModel('')).toBe(false);
  });
});

// ─── buildChatBody ─────────────────────────────────────────────────────────────

const baseMessages = [
  { role: 'system', content: 'You are a helpful assistant.' },
  { role: 'user', content: 'Hello' },
];

const standardConfig = {
  id: 'gpt-4o',
  label: 'Standard',
  maxTokens: 16000,
  temperature: 0.7,
  type: 'chat',
};

const gpt5Config = {
  id: 'gpt-5.1',
  label: 'GPT-5',
  maxTokens: 32000,
  temperature: 1,
  reasoningEffort: 'high',
  type: 'chat',
};

const chatGptConfig = {
  id: 'chatgpt-4o-latest',
  label: 'ChatGPT',
  type: 'chat',
};

const gpt52Config = {
  id: 'gpt-5.2-turbo',
  label: 'GPT-5.2',
  maxTokens: 65000,
  temperature: 1,
  reasoningEffort: 'medium',
  type: 'chat',
};

describe('buildChatBody', () => {
  it('builds basic body for standard model', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: false,
    });

    expect(body.model).toBe('gpt-4o');
    expect(body.messages).toEqual(baseMessages);
    expect(body.stream).toBe(false);
    expect(body.max_tokens).toBe(16000);
    expect(body.temperature).toBe(0.7);
    expect(body.stream_options).toBeUndefined();
  });

  it('adds stream_options when streaming', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: true,
    });

    expect(body.stream_options).toEqual({ include_usage: true });
  });

  it('uses max_completion_tokens for GPT-5 models', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: gpt5Config,
      messages: baseMessages,
      stream: false,
    });

    expect(body.max_completion_tokens).toBe(32000);
    expect(body.max_tokens).toBeUndefined();
  });

  it('does not set temperature for GPT-5 models', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: gpt5Config,
      messages: baseMessages,
      stream: false,
    });

    expect(body.temperature).toBeUndefined();
  });

  it('sets reasoning_effort for GPT-5 models on non-image tiers', () => {
    const body = buildChatBody({
      modelTier: 'reasoning',
      modelConfig: gpt5Config,
      messages: baseMessages,
      stream: false,
    });

    expect(body.reasoning_effort).toBe('high');
  });

  it('does not set reasoning_effort for image tier', () => {
    const body = buildChatBody({
      modelTier: 'image',
      modelConfig: gpt5Config,
      messages: baseMessages,
      stream: false,
    });

    expect(body.reasoning_effort).toBeUndefined();
  });

  it('does not set max_tokens or temperature for ChatGPT models', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: chatGptConfig,
      messages: baseMessages,
      stream: false,
    });

    expect(body.max_tokens).toBeUndefined();
    expect(body.max_completion_tokens).toBeUndefined();
    expect(body.temperature).toBeUndefined();
  });

  it('adds tools and tool_choice when tools provided', () => {
    const tools = [{ type: 'function', function: { name: 'test' } }];
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: false,
      tools,
    });

    expect(body.tools).toBe(tools);
    expect(body.tool_choice).toBe('auto');
  });

  it('does not set tool_choice for gpt-5.2 models', () => {
    const tools = [{ type: 'function', function: { name: 'test' } }];
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: gpt52Config,
      messages: baseMessages,
      stream: false,
      tools,
    });

    expect(body.tools).toBe(tools);
    expect(body.tool_choice).toBeUndefined();
  });

  it('does not add tools when empty array provided', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: false,
      tools: [],
    });

    expect(body.tools).toBeUndefined();
    expect(body.tool_choice).toBeUndefined();
  });

  it('strips empty tool_calls arrays from messages', () => {
    const messagesWithEmptyToolCalls = [
      { role: 'system', content: 'sys' },
      { role: 'assistant', content: 'hi', tool_calls: [] },
      { role: 'user', content: 'hello' },
    ];

    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: messagesWithEmptyToolCalls,
      stream: false,
    });

    const assistantMsg = body.messages.find(m => m.role === 'assistant');
    expect(assistantMsg.tool_calls).toBeUndefined();
    expect(assistantMsg.content).toBe('hi');
  });

  it('preserves non-empty tool_calls arrays', () => {
    const toolCalls = [
      { id: 'tc1', type: 'function', function: { name: 'test', arguments: '{}' } },
    ];
    const messagesWithToolCalls = [{ role: 'assistant', content: '', tool_calls: toolCalls }];

    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: messagesWithToolCalls,
      stream: false,
    });

    expect(body.messages[0].tool_calls).toBe(toolCalls);
  });

  it('overrides maxTokens with caller-provided value', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: false,
      maxTokens: 4096,
    });

    expect(body.max_tokens).toBe(4096);
  });

  it('overrides temperature with caller-provided value', () => {
    const body = buildChatBody({
      modelTier: 'standard',
      modelConfig: standardConfig,
      messages: baseMessages,
      stream: false,
      temperature: 0.2,
    });

    expect(body.temperature).toBe(0.2);
  });
});
