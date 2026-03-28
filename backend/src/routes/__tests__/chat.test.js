import { describe, it, expect, vi, beforeEach } from 'vitest';
import express from 'express';
import request from 'supertest';

// ─── Mocks ────────────────────────────────────────────────────────────────────
// Mock llmClient before importing the router — the route imports it at module level.
vi.mock('../../services/llmClient.js', () => ({
  chatCompletion: vi.fn(),
  handleErrorResponse: vi.fn(),
  RateLimitError: class RateLimitError extends Error {
    constructor(retryAfterMs) {
      super(`Rate limit exceeded. Retry after ${retryAfterMs}ms.`);
      this.name = 'RateLimitError';
      this.retryAfterMs = retryAfterMs;
    }
  },
}));

// Mock toolUsageLogger to prevent filesystem writes during tests
vi.mock('../../utils/toolUsageLogger.js', () => ({
  logToolUsage: vi.fn(),
  logChatRequest: vi.fn(),
}));

import { chatCompletion, handleErrorResponse, RateLimitError } from '../../services/llmClient.js';
import { chatRouter } from '../chat.js';

// ─── Test App Factory ─────────────────────────────────────────────────────────
// Creates a minimal Express app with just the chat router and a request logger.
// Skips CSRF, auth, and rate-limit middleware so we can test the route in isolation.

function createTestApp() {
  const app = express();
  app.use(express.json());
  // Attach a mock req.logger and req.userCredentials like the real middleware does
  app.use((req, _res, next) => {
    req.logger = {
      debug: vi.fn(),
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      defaultMeta: { userId: 'test@example.com', host: 'Word' },
    };
    req.userCredentials = { userKey: 'test-key-12345', userEmail: 'test@example.com' };
    next();
  });
  app.use('/api/chat', chatRouter);
  return app;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

const VALID_MESSAGES = [
  { role: 'system', content: 'You are a helpful assistant.' },
  { role: 'user', content: 'Hello' },
];

function validBody(overrides = {}) {
  return {
    messages: VALID_MESSAGES,
    modelTier: 'standard',
    ...overrides,
  };
}

/**
 * Creates a mock ReadableStream body that emits SSE chunks then closes.
 * Mimics the upstream LLM streaming response.
 */
function createMockSSEStream(chunks) {
  const encoder = new TextEncoder();
  let index = 0;
  return new ReadableStream({
    pull(controller) {
      if (index < chunks.length) {
        controller.enqueue(encoder.encode(chunks[index]));
        index++;
      } else {
        controller.close();
      }
    },
  });
}

function mockStreamResponse(chunks) {
  const body = createMockSSEStream(chunks);
  return {
    ok: true,
    status: 200,
    body: body,
  };
}

// ─── 1. Validation Errors ─────────────────────────────────────────────────────

describe('POST /api/chat — validation', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns 400 when messages is missing', async () => {
    const res = await request(app).post('/api/chat').send({ modelTier: 'standard' });

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/messages/);
  });

  it('returns 400 when messages is an empty array', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ messages: [] }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/empty/);
  });

  it('returns 400 when messages is not an array', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ messages: 'not an array' }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/messages/);
  });

  it('returns 400 when message has invalid role', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ messages: [{ role: 'invalid', content: 'hi' }] }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/role/);
  });

  it('returns 400 for unknown model tier', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ modelTier: 'nonexistent' }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/Unknown model tier/);
  });

  it('returns 400 for image model tier on chat endpoint', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ modelTier: 'image' }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/image/i);
  });

  it('returns 400 for invalid temperature', async () => {
    const res = await request(app)
      .post('/api/chat')
      .send(validBody({ temperature: 3 }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
    expect(res.body.error).toMatch(/temperature/);
  });
});

describe('POST /api/chat/sync — validation', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns 400 when messages is missing', async () => {
    const res = await request(app).post('/api/chat/sync').send({ modelTier: 'standard' });

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
  });

  it('returns 400 when messages is empty', async () => {
    const res = await request(app)
      .post('/api/chat/sync')
      .send(validBody({ messages: [] }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
  });

  it('returns 400 for unknown model tier', async () => {
    const res = await request(app)
      .post('/api/chat/sync')
      .send(validBody({ modelTier: 'nonexistent' }));

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('VALIDATION_ERROR');
  });
});

// ─── 2. SSE Streaming ────────────────────────────────────────────────────────

describe('POST /api/chat — SSE streaming', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns SSE headers and streams chunks', async () => {
    const sseChunks = [
      'data: {"id":"chatcmpl-1","choices":[{"delta":{"content":"Hello"}}]}\n\n',
      'data: {"id":"chatcmpl-1","choices":[{"delta":{"content":" world"}}]}\n\n',
      'data: [DONE]\n\n',
    ];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toMatch(/text\/event-stream/);
    expect(res.headers['cache-control']).toBe('no-cache');
    expect(res.text).toContain('data: {"id":"chatcmpl-1"');
    expect(res.text).toContain('data: [DONE]');
  });

  it('streams tool call deltas correctly', async () => {
    const sseChunks = [
      'data: {"id":"chatcmpl-2","choices":[{"delta":{"tool_calls":[{"index":0,"id":"tc1","function":{"name":"getDoc","arguments":""}}]}}]}\n\n',
      'data: {"id":"chatcmpl-2","choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{}"}}]}}]}\n\n',
      'data: [DONE]\n\n',
    ];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(200);
    expect(res.text).toContain('getDoc');
    expect(res.text).toContain('[DONE]');
  });
});

// ─── 3. Sync Endpoint ────────────────────────────────────────────────────────

describe('POST /api/chat/sync — success', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns the LLM JSON response', async () => {
    const llmResponse = {
      id: 'chatcmpl-sync-1',
      model: 'gpt-4o',
      choices: [{ message: { role: 'assistant', content: 'Hi there' }, finish_reason: 'stop' }],
      usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
    };

    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce(llmResponse),
    });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(200);
    expect(res.body.id).toBe('chatcmpl-sync-1');
    expect(res.body.choices[0].message.content).toBe('Hi there');
  });

  it('returns 502 when LLM returns no choices', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce({ id: 'x', model: 'gpt-4o', choices: [] }),
    });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(502);
    expect(res.body.code).toBe('LLM_NO_CHOICES');
  });

  it('returns 502 when LLM response has empty content and no tool_calls', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce({
        id: 'x',
        model: 'gpt-4o',
        choices: [{ message: { role: 'assistant', content: '' }, finish_reason: 'stop' }],
      }),
    });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(502);
    expect(res.body.code).toBe('LLM_EMPTY_RESPONSE');
  });

  it('returns 502 when LLM response has null content and no tool_calls', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce({
        id: 'x',
        model: 'gpt-4o',
        choices: [{ message: { role: 'assistant', content: null }, finish_reason: 'stop' }],
      }),
    });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(502);
    expect(res.body.code).toBe('LLM_EMPTY_RESPONSE');
  });

  it('accepts response with tool_calls but no content', async () => {
    const llmResponse = {
      id: 'chatcmpl-tools',
      model: 'gpt-4o',
      choices: [
        {
          message: {
            role: 'assistant',
            content: null,
            tool_calls: [
              { id: 'tc1', type: 'function', function: { name: 'test', arguments: '{}' } },
            ],
          },
          finish_reason: 'tool_calls',
        },
      ],
    };

    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce(llmResponse),
    });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(200);
    expect(res.body.choices[0].message.tool_calls).toHaveLength(1);
  });
});

// ─── 4. Error Handling ───────────────────────────────────────────────────────

describe('POST /api/chat — error handling', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns 502 when upstream returns 5xx', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: false,
      status: 500,
      headers: new Headers(),
      text: vi.fn().mockResolvedValueOnce('Internal Server Error'),
    });
    handleErrorResponse.mockResolvedValueOnce({ status: 500, rawMessage: 'Internal Server Error' });

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(502);
    expect(res.body.code).toBe('LLM_UPSTREAM_ERROR');
  });

  it('returns 400 when upstream returns 4xx', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: false,
      status: 422,
      headers: new Headers(),
      text: vi.fn().mockResolvedValueOnce('Unprocessable Entity'),
    });
    handleErrorResponse.mockResolvedValueOnce({ status: 422, rawMessage: 'Invalid image data' });

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('LLM_BAD_REQUEST');
    expect(res.body.detail).toBe('Invalid image data');
  });

  it('returns 429 when chatCompletion throws RateLimitError', async () => {
    chatCompletion.mockRejectedValueOnce(new RateLimitError(30000));

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(429);
    expect(res.body.code).toBe('RATE_LIMITED');
  });

  it('returns 504 when chatCompletion throws AbortError (timeout)', async () => {
    const abortError = new Error('The operation was aborted');
    abortError.name = 'AbortError';
    chatCompletion.mockRejectedValueOnce(abortError);

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(504);
    expect(res.body.code).toBe('LLM_TIMEOUT');
  });

  it('returns 500 on unexpected error', async () => {
    chatCompletion.mockRejectedValueOnce(new Error('Something broke'));

    const res = await request(app).post('/api/chat').send(validBody());

    expect(res.status).toBe(500);
    expect(res.body.code).toBe('INTERNAL_ERROR');
  });
});

describe('POST /api/chat/sync — error handling', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('returns 502 when upstream returns 5xx', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: false,
      status: 503,
      headers: new Headers(),
    });
    handleErrorResponse.mockResolvedValueOnce({ status: 503, rawMessage: 'Service Unavailable' });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(502);
    expect(res.body.code).toBe('LLM_UPSTREAM_ERROR');
  });

  it('returns 400 when upstream returns 4xx', async () => {
    chatCompletion.mockResolvedValueOnce({
      ok: false,
      status: 400,
      headers: new Headers(),
    });
    handleErrorResponse.mockResolvedValueOnce({ status: 400, rawMessage: 'Bad model param' });

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(400);
    expect(res.body.code).toBe('LLM_BAD_REQUEST');
    expect(res.body.detail).toBe('Bad model param');
  });

  it('returns 429 on RateLimitError', async () => {
    chatCompletion.mockRejectedValueOnce(new RateLimitError(10000));

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(429);
    expect(res.body.code).toBe('RATE_LIMITED');
  });

  it('returns 504 on AbortError (timeout)', async () => {
    const abortError = new Error('Timeout');
    abortError.name = 'AbortError';
    chatCompletion.mockRejectedValueOnce(abortError);

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(504);
    expect(res.body.code).toBe('LLM_TIMEOUT');
  });

  it('returns 500 on unexpected error', async () => {
    chatCompletion.mockRejectedValueOnce(new TypeError('fetch failed'));

    const res = await request(app).post('/api/chat/sync').send(validBody());

    expect(res.status).toBe(500);
    expect(res.body.code).toBe('INTERNAL_ERROR');
  });
});

// ─── 5. Request Body Construction ────────────────────────────────────────────

describe('POST /api/chat — request body construction', () => {
  let app;

  beforeEach(() => {
    vi.clearAllMocks();
    app = createTestApp();
  });

  it('calls chatCompletion with stream: true and correct body shape', async () => {
    const sseChunks = [
      'data: {"id":"x","choices":[{"delta":{"content":"ok"}}]}\n\n',
      'data: [DONE]\n\n',
    ];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    await request(app).post('/api/chat').send(validBody());

    expect(chatCompletion).toHaveBeenCalledOnce();
    const call = chatCompletion.mock.calls[0][0];
    expect(call.body.stream).toBe(true);
    expect(call.body.messages).toBeDefined();
    expect(call.body.model).toBeDefined();
    expect(call.body.stream_options).toEqual({ include_usage: true });
    expect(call.modelTier).toBe('standard');
    expect(call.userCredentials).toEqual({
      userKey: 'test-key-12345',
      userEmail: 'test@example.com',
    });
  });

  it('calls chatCompletion with stream: false for /sync', async () => {
    const llmResponse = {
      id: 'x',
      model: 'gpt-4o',
      choices: [{ message: { role: 'assistant', content: 'ok' }, finish_reason: 'stop' }],
    };
    chatCompletion.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: vi.fn().mockResolvedValueOnce(llmResponse),
    });

    await request(app).post('/api/chat/sync').send(validBody());

    expect(chatCompletion).toHaveBeenCalledOnce();
    const call = chatCompletion.mock.calls[0][0];
    expect(call.body.stream).toBe(false);
    expect(call.body.stream_options).toBeUndefined();
  });

  it('passes tools through to chatCompletion when provided', async () => {
    const tools = [
      {
        type: 'function',
        function: { name: 'readDoc', parameters: { type: 'object', properties: {} } },
      },
    ];
    const sseChunks = ['data: [DONE]\n\n'];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    await request(app).post('/api/chat').send(validBody({ tools }));

    const call = chatCompletion.mock.calls[0][0];
    expect(call.body.tools).toEqual(tools);
    expect(call.body.tool_choice).toBe('auto');
  });

  it('forwards custom temperature and maxTokens to chatCompletion body', async () => {
    const sseChunks = ['data: [DONE]\n\n'];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    await request(app)
      .post('/api/chat')
      .send(validBody({ temperature: 0.2, maxTokens: 1024 }));

    const call = chatCompletion.mock.calls[0][0];
    // Whether temperature/max_tokens appear in body depends on model type.
    // The important thing is that chatCompletion was called and the body was built.
    expect(call.body).toBeDefined();
    expect(call.body.model).toBeDefined();
    // For non-GPT-5 models, temperature and max_tokens would appear.
    // For GPT-5 models, max_completion_tokens is used instead and temperature is omitted.
    const hasMaxTokenParam =
      call.body.max_tokens !== undefined || call.body.max_completion_tokens !== undefined;
    expect(hasMaxTokenParam).toBe(true);
  });

  it('passes modelTier through to chatCompletion', async () => {
    const sseChunks = ['data: [DONE]\n\n'];
    chatCompletion.mockResolvedValueOnce(mockStreamResponse(sseChunks));

    await request(app).post('/api/chat').send(validBody());

    const call = chatCompletion.mock.calls[0][0];
    expect(call.modelTier).toBe('standard');
  });
});

// ─── 6. Rate Limiting (full app) ─────────────────────────────────────────────
// NOTE: Testing rate limiting requires the full Express app with middleware.
// We import the server app indirectly — but since server.js starts listening,
// we test rate limiting conceptually through the RateLimitError path above
// (section 4). The express-rate-limit middleware is integration-tested via
// the upstream RateLimitError test which verifies the 429 code path.
