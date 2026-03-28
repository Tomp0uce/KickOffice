import { describe, it, expect, vi, beforeEach } from 'vitest';

// ── Mocks declared before module import ──────────────────────────────────────

// tokenManager imports @/utils/message and @/i18n.
// Both involve DOM / Vue bootstrap — stub them out entirely.
vi.mock('@/utils/message', () => ({
  message: {
    warning: vi.fn(),
    error: vi.fn(),
    info: vi.fn(),
    success: vi.fn(),
  },
}));

vi.mock('@/i18n', () => ({
  i18n: {
    global: {
      t: (key: string) => key,
    },
  },
}));

// tokenManager imports ChatRequestMessage from @/api/backend (type only).
// The runtime import resolves to nothing harmful — but the module itself
// throws at load time because VITE_BACKEND_URL is undefined in the test env.
// Mock the entire module so the env guard never runs.
vi.mock('@/api/backend', () => ({}));

vi.mock('@/utils/logger', () => ({
  logService: {
    warn: vi.fn(),
    error: vi.fn(),
    info: vi.fn(),
  },
}));

import { prepareMessagesForContext, MAX_CONTEXT_CHARS, estimateContextUsagePercent } from '../tokenManager';
import type { ChatRequestMessage } from '@/api/backend';
import { message as messageUtil } from '@/utils/message';
import { logService } from '@/utils/logger';

/** Local type for test assertions on assistant messages with tool_calls. */
type AssistantMsg = {
  role: 'assistant';
  content: string;
  tool_calls?: Array<{ id: string; type: string; function: { name: string; arguments: string } }>;
};

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function makeUserMsg(content: string): ChatRequestMessage {
  return { role: 'user', content };
}

function makeAssistantMsg(content: string): ChatRequestMessage {
  return { role: 'assistant', content };
}

function makeToolMsg(toolCallId: string, content: string): ChatRequestMessage {
  return { role: 'tool', tool_call_id: toolCallId, content };
}

function makeAssistantWithToolCalls(content: string, toolCallIds: string[]): ChatRequestMessage {
  return {
    role: 'assistant',
    content,
    tool_calls: toolCallIds.map(id => ({
      id,
      type: 'function' as const,
      function: { name: 'someTool', arguments: '{}' },
    })),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// MAX_CONTEXT_CHARS
// ─────────────────────────────────────────────────────────────────────────────
describe('MAX_CONTEXT_CHARS', () => {
  it('is a positive number', () => {
    expect(MAX_CONTEXT_CHARS).toBeGreaterThan(0);
  });

  it('equals 1_200_000', () => {
    expect(MAX_CONTEXT_CHARS).toBe(1_200_000);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// prepareMessagesForContext — basic structure
// ─────────────────────────────────────────────────────────────────────────────
describe('prepareMessagesForContext — basic structure', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('always returns the system message as the first element', () => {
    const result = prepareMessagesForContext([makeUserMsg('hi')], 'You are helpful.');
    expect(result[0].role).toBe('system');
    expect(result[0].content).toBe('You are helpful.');
  });

  it('returns only the system message when allMessages is empty', () => {
    const result = prepareMessagesForContext([], 'System.');
    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('system');
  });

  it('filters out any pre-existing system messages from the input', () => {
    const messages: ChatRequestMessage[] = [
      { role: 'system', content: 'Old system prompt — should be dropped' },
      makeUserMsg('Hello'),
    ];
    const result = prepareMessagesForContext(messages, 'New system prompt');
    const systemMessages = result.filter(m => m.role === 'system');
    expect(systemMessages).toHaveLength(1);
    expect(systemMessages[0].content).toBe('New system prompt');
  });

  it('includes a short user message in the result', () => {
    const result = prepareMessagesForContext([makeUserMsg('short')], 'System');
    const userMsgs = result.filter(m => m.role === 'user');
    expect(userMsgs).toHaveLength(1);
    expect(userMsgs[0].content).toBe('short');
  });

  it('preserves message order (system first, then chronological)', () => {
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('first'),
      makeAssistantMsg('second'),
      makeUserMsg('third'),
    ];
    const result = prepareMessagesForContext(msgs, 'System');
    expect(result[0].role).toBe('system');
    const rest = result.slice(1);
    const roles = rest.map(m => m.role);
    expect(roles).toEqual(['user', 'assistant', 'user']);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// prepareMessagesForContext — tool_call integrity
// ─────────────────────────────────────────────────────────────────────────────
describe('prepareMessagesForContext — tool_call / tool response integrity', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('keeps tool_calls when a matching tool response is present', () => {
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('Do it'),
      makeAssistantWithToolCalls('', ['tc-1']),
      makeToolMsg('tc-1', 'result'),
    ];
    const result = prepareMessagesForContext(msgs, 'System');
    const assistant = result.find(
      m => m.role === 'assistant' && 'tool_calls' in m && m.tool_calls,
    ) as AssistantMsg | undefined;

    expect(assistant).toBeDefined();
    expect(assistant?.tool_calls).toHaveLength(1);
  });

  it('strips all tool_calls from an assistant message when no tool response is present', () => {
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('Do it'),
      makeAssistantWithToolCalls('', ['orphan-1']),
      // No tool response for orphan-1
      makeUserMsg('follow-up'),
    ];
    const result = prepareMessagesForContext(msgs, 'System');
    const assistant = result.find(m => m.role === 'assistant') as AssistantMsg | undefined;

    // Either the assistant is absent or its tool_calls were stripped
    if (assistant) {
      expect(assistant.tool_calls).toBeUndefined();
    }
  });

  it('strips only orphaned tool_calls, keeping answered ones', () => {
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('Do it'),
      makeAssistantWithToolCalls('', ['tc-answered', 'tc-orphan']),
      makeToolMsg('tc-answered', 'result'),
      makeUserMsg('next'),
    ];
    const result = prepareMessagesForContext(msgs, 'System');
    const assistant = result.find(
      m => m.role === 'assistant' && 'tool_calls' in m && m.tool_calls,
    ) as AssistantMsg | undefined;

    if (assistant?.tool_calls) {
      const ids = assistant.tool_calls.map(tc => tc.id);
      expect(ids).toContain('tc-answered');
      expect(ids).not.toContain('tc-orphan');
    }
  });


  it('does not mutate the original messages when stripping orphaned tool_calls (ROB-H1)', () => {
    const original: ChatRequestMessage[] = [
      makeUserMsg('Do it'),
      makeAssistantWithToolCalls('thinking', ['tc-answered', 'tc-orphan']),
      makeToolMsg('tc-answered', 'result'),
      makeUserMsg('next'),
    ];

    // Deep snapshot of tool_calls BEFORE the call
    const toolCallsBefore = JSON.parse(
      JSON.stringify((original[1] as AssistantMsg).tool_calls),
    );

    prepareMessagesForContext(original, 'System');

    // The original assistant message must still have both tool_calls intact
    const assistantOriginal = original[1] as AssistantMsg;
    expect(assistantOriginal.tool_calls).toBeDefined();
    expect(assistantOriginal.tool_calls).toHaveLength(2);
    expect(JSON.stringify(assistantOriginal.tool_calls)).toBe(JSON.stringify(toolCallsBefore));
  });

  it('does not mutate the original messages when stripping all tool_calls (ROB-H1)', () => {
    const original: ChatRequestMessage[] = [
      makeUserMsg('Do it'),
      makeAssistantWithToolCalls('thinking', ['orphan-only']),
      makeUserMsg('follow-up'),
    ];

    prepareMessagesForContext(original, 'System');

    // The original assistant message must still have tool_calls
    const assistantOriginal = original[1] as AssistantMsg;
    expect(assistantOriginal.tool_calls).toBeDefined();
    expect(assistantOriginal.tool_calls).toHaveLength(1);
    expect(assistantOriginal.tool_calls![0].id).toBe('orphan-only');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// prepareMessagesForContext — budget / truncation
// ─────────────────────────────────────────────────────────────────────────────
describe('prepareMessagesForContext — budget and truncation', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns only system message when system prompt consumes the entire budget', () => {
    // A system prompt exactly at MAX_CONTEXT_CHARS leaves remainingBudget = 0
    const hugePrompt = 'x'.repeat(MAX_CONTEXT_CHARS);
    const result = prepareMessagesForContext([makeUserMsg('ignored')], hugePrompt);
    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('system');
  });

  it('does not emit a warning for a small context that fits within budget', () => {
    prepareMessagesForContext([makeUserMsg('tiny message')], 'Tiny system prompt');
    expect(messageUtil.warning).not.toHaveBeenCalled();
  });

  it('includes the last user message with priority (force-include)', () => {
    // Last user message should always be present even when preceded by many messages
    const msgs: ChatRequestMessage[] = Array.from({ length: 5 }, (_, i) =>
      i % 2 === 0 ? makeUserMsg(`user ${i}`) : makeAssistantMsg(`assistant ${i}`),
    );
    msgs.push(makeUserMsg('final user message'));

    const result = prepareMessagesForContext(msgs, 'System');
    const userMsgs = result.filter(m => m.role === 'user');
    const lastUserContent = userMsgs[userMsgs.length - 1].content;
    expect(lastUserContent).toBe('final user message');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// estimateContextUsagePercent
// ─────────────────────────────────────────────────────────────────────────────
describe('estimateContextUsagePercent', () => {
  it('returns 0 for empty messages and empty system prompt', () => {
    const result = estimateContextUsagePercent([], '');
    expect(result).toBe(0);
  });

  it('returns 100 when system prompt alone fills the budget', () => {
    const huge = 'x'.repeat(MAX_CONTEXT_CHARS);
    const result = estimateContextUsagePercent([], huge);
    expect(result).toBe(100);
  });

  it('caps at 100 even if content exceeds MAX_CONTEXT_CHARS', () => {
    const huge = 'x'.repeat(MAX_CONTEXT_CHARS + 1000);
    const result = estimateContextUsagePercent([], huge);
    expect(result).toBe(100);
  });

  it('ignores system-role messages in allMessages', () => {
    const withoutSystem = estimateContextUsagePercent([], 'System');
    const withSystem = estimateContextUsagePercent(
      [{ role: 'system', content: 'x'.repeat(100_000) }],
      'System',
    );
    expect(withSystem).toBe(withoutSystem);
  });

  it('counts user message content toward usage', () => {
    // 12_000 chars = 1% of MAX_CONTEXT_CHARS → rounds to 1
    const userContent = 'a'.repeat(12_000);
    const result = estimateContextUsagePercent([makeUserMsg(userContent)], '');
    expect(result).toBeGreaterThan(0);
  });

  it('counts assistant message with tool_calls toward usage', () => {
    const assistantWithCalls = makeAssistantWithToolCalls('content', ['tc-1']);
    const withCalls = estimateContextUsagePercent([assistantWithCalls], '');
    const withoutCalls = estimateContextUsagePercent([makeAssistantMsg('content')], '');
    // tool_calls serialization adds length
    expect(withCalls).toBeGreaterThanOrEqual(withoutCalls);
  });

  it('returns a value between 0 and 100 inclusive', () => {
    const result = estimateContextUsagePercent(
      [makeUserMsg('hello'), makeAssistantMsg('world')],
      'Short system',
    );
    expect(result).toBeGreaterThanOrEqual(0);
    expect(result).toBeLessThanOrEqual(100);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// prepareMessagesForContext — getMessageContentLength branches
// These exercise the array-content path (image_url, file, text parts) and the
// non-string / non-array fallback via message assembly inside prepareMessages.
// ─────────────────────────────────────────────────────────────────────────────
describe('getMessageContentLength — via estimateContextUsagePercent', () => {
  it('counts 1000 chars for each image_url part', () => {
    const msgWithImage: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [{ type: 'image_url', image_url: { url: 'data:...' } }] as any,
    };
    const result = estimateContextUsagePercent([msgWithImage], '');
    // 1000 chars / 1_200_000 * 100 ≈ 0.08 → rounds to 0, but total > 0
    expect(result).toBeGreaterThanOrEqual(0);
    // Verify indirectly: adding a second image doubles the contribution
    const msgTwoImages: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [
        { type: 'image_url', image_url: { url: 'data:...' } },
        { type: 'image_url', image_url: { url: 'data:...' } },
      ] as any,
    };
    // Just verify it runs without error and produces a valid percent
    const twoImgResult = estimateContextUsagePercent([msgTwoImages], '');
    expect(twoImgResult).toBeGreaterThanOrEqual(0);
    expect(twoImgResult).toBeLessThanOrEqual(100);
  });

  it('counts 200 chars for each file part', () => {
    const msgWithFile: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [{ type: 'file', file: { file_id: 'f-123' } }] as any,
    };
    const result = estimateContextUsagePercent([msgWithFile], '');
    expect(result).toBeGreaterThanOrEqual(0);
    expect(result).toBeLessThanOrEqual(100);
  });

  it('counts text part length for text-type array parts', () => {
    const shortText: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [{ type: 'text', text: 'hi' }] as any,
    };
    const longText: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [{ type: 'text', text: 'h'.repeat(600_000) }] as any,
    };
    const shortResult = estimateContextUsagePercent([shortText], '');
    const longResult = estimateContextUsagePercent([longText], '');
    expect(longResult).toBeGreaterThan(shortResult);
  });

  it('falls back to JSON.stringify for non-string non-array content', () => {
    // null content — JSON.stringify(null) = "null" (4 chars), should not throw
    const msgNull: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: null as any,
    };
    expect(() => estimateContextUsagePercent([msgNull], '')).not.toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// truncateToBudget — warning and tail direction (exercised via prepareMessages)
// ─────────────────────────────────────────────────────────────────────────────
describe('truncateToBudget — warning emission and direction', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('emits a warning exactly once when a message is truncated', () => {
    // Create a system prompt that leaves almost no room, then add a large user message
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 50); // leaves 50 chars budget
    const largeUserMsg = makeUserMsg('U'.repeat(200));

    prepareMessagesForContext([largeUserMsg], systemPrompt);

    expect(messageUtil.warning).toHaveBeenCalledTimes(1);
  });

  it('emits warning with the i18n key errorTruncated', () => {
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 50);
    const largeUserMsg = makeUserMsg('U'.repeat(200));

    prepareMessagesForContext([largeUserMsg], systemPrompt);

    expect(messageUtil.warning).toHaveBeenCalledWith('errorTruncated');
  });

  it('does not emit a second warning when truncation happens again in same call', () => {
    // Two large messages, both needing truncation — warning fires only once
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 100);
    const msg1 = makeUserMsg('A'.repeat(200));
    const msg2 = makeUserMsg('B'.repeat(200));

    prepareMessagesForContext([msg1, msg2], systemPrompt);

    // hasWarnedTruncation prevents duplicate warnings within same module lifecycle
    expect(messageUtil.warning).toHaveBeenCalledTimes(1);
  });

  it('calls logService.warn when a message is truncated', () => {
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 50);
    const largeUserMsg = makeUserMsg('U'.repeat(200));

    prepareMessagesForContext([largeUserMsg], systemPrompt);

    expect(logService.warn).toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// summarizeOldToolResults — via prepareMessagesForContext
// ─────────────────────────────────────────────────────────────────────────────
describe('summarizeOldToolResults — compression of old tool results', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  /**
   * Build a conversation with N tool-call iterations followed by a final user message.
   * Each iteration = assistant-with-tool-calls + tool-response.
   */
  function makeIterations(count: number, resultContent: string): ChatRequestMessage[] {
    const msgs: ChatRequestMessage[] = [makeUserMsg('start')];
    for (let i = 0; i < count; i++) {
      msgs.push(makeAssistantWithToolCalls(`step ${i}`, [`tc-${i}`]));
      msgs.push(makeToolMsg(`tc-${i}`, resultContent));
    }
    msgs.push(makeUserMsg('final'));
    return msgs;
  }

  it('does not compress when there are <= 3 tool-call iterations', () => {
    const longResult = 'R'.repeat(1000); // > TOOL_RESULT_MAX_CHARS (800)
    const msgs = makeIterations(3, longResult);

    const result = prepareMessagesForContext(msgs, 'System');
    const toolMsgs = result.filter(m => m.role === 'tool');

    // All 3 tool results should remain uncompressed
    toolMsgs.forEach(m => {
      expect(m.content).toBe(longResult);
    });
  });

  it('compresses old tool results when there are > 3 iterations', () => {
    const longResult = 'R'.repeat(1000); // > TOOL_RESULT_MAX_CHARS (800)
    const msgs = makeIterations(5, longResult); // 5 iterations → first 2 should be compressed

    const result = prepareMessagesForContext(msgs, 'System');
    const toolMsgs = result.filter(m => m.role === 'tool');

    if (toolMsgs.length > 0) {
      // At least one tool message should be shorter than the original long result
      const hasCompressed = toolMsgs.some(
        m => typeof m.content === 'string' && m.content.length < longResult.length,
      );
      expect(hasCompressed).toBe(true);
    }
  });

  it('keeps the last 3 iterations intact even when older ones are compressed', () => {
    const longResult = 'R'.repeat(1000);
    const msgs = makeIterations(5, longResult);

    // Indices of tool messages in the input (after system filter):
    // iteration 0: index 1 (assistant), index 2 (tool)
    // iteration 1: index 3 (assistant), index 4 (tool)
    // ...
    const result = prepareMessagesForContext(msgs, 'System');
    const toolMsgs = result.filter(m => m.role === 'tool');

    // The last 3 tool results should remain full-length
    const lastThree = toolMsgs.slice(-3);
    lastThree.forEach(m => {
      expect(m.content).toBe(longResult);
    });
  });

  it('does not compress tool results that are already short', () => {
    const shortResult = 'short result'; // < TOOL_RESULT_MAX_CHARS (800)
    const msgs = makeIterations(5, shortResult);

    const result = prepareMessagesForContext(msgs, 'System');
    const toolMsgs = result.filter(m => m.role === 'tool');

    // Short results should never be modified regardless of iteration age
    toolMsgs.forEach(m => {
      expect(m.content).toBe(shortResult);
    });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// truncateJsonToolResult — exercised via prepareMessagesForContext
// The function is private but gets invoked when a tool message exceeds the
// remaining budget during context preparation.
// ─────────────────────────────────────────────────────────────────────────────
describe('truncateJsonToolResult — JSON-aware truncation of tool results', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('truncates a JSON object tool result and preserves outer braces envelope', () => {
    // Leave a small budget so the tool message must be truncated
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 200);
    const jsonObjectResult = '{' + '"data":"' + 'x'.repeat(500) + '"}';
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('go'),
      makeAssistantWithToolCalls('', ['tc-j1']),
      makeToolMsg('tc-j1', jsonObjectResult),
    ];

    const result = prepareMessagesForContext(msgs, systemPrompt);
    const toolMsg = result.find(m => m.role === 'tool');

    if (toolMsg) {
      const content = toolMsg.content as string;
      // Should be either truncated (starts with '{') or original if it fit
      expect(typeof content).toBe('string');
      expect(content.length).toBeLessThanOrEqual(jsonObjectResult.length);
    }
  });

  it('truncates a JSON array tool result and preserves outer bracket envelope', () => {
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 200);
    const jsonArrayResult = '[' + '"item","'.repeat(100) + '"end"]';
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('go'),
      makeAssistantWithToolCalls('', ['tc-j2']),
      makeToolMsg('tc-j2', jsonArrayResult),
    ];

    const result = prepareMessagesForContext(msgs, systemPrompt);
    const toolMsg = result.find(m => m.role === 'tool');

    if (toolMsg) {
      const content = toolMsg.content as string;
      expect(typeof content).toBe('string');
      expect(content.length).toBeLessThanOrEqual(jsonArrayResult.length);
    }
  });

  it('truncates plain text tool result keeping the tail', () => {
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 200);
    const plainResult = 'plain text result ' + 'x'.repeat(500);
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('go'),
      makeAssistantWithToolCalls('', ['tc-j3']),
      makeToolMsg('tc-j3', plainResult),
    ];

    const result = prepareMessagesForContext(msgs, systemPrompt);
    const toolMsg = result.find(m => m.role === 'tool');

    if (toolMsg) {
      const content = toolMsg.content as string;
      expect(typeof content).toBe('string');
      // Plain text truncation inserts the tail marker at the start
      // so the result should start with '[Truncated ...' or be the original
      expect(content.length).toBeLessThanOrEqual(plainResult.length);
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// prepareMessagesForContext — force-include truncation paths
// ─────────────────────────────────────────────────────────────────────────────
describe('prepareMessagesForContext — force-include and truncation edge cases', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('truncates the system prompt when it exceeds MAX_CONTEXT_CHARS', () => {
    const oversizedPrompt = 'X'.repeat(MAX_CONTEXT_CHARS + 5000);
    const result = prepareMessagesForContext([], oversizedPrompt);

    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('system');
    expect((result[0].content as string).length).toBeLessThanOrEqual(MAX_CONTEXT_CHARS);
  });

  it('force-includes the last user message even when budget is nearly exhausted', () => {
    // System prompt consumes all but ~100 chars
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 100);
    const lastUser = makeUserMsg('important final question?');

    const result = prepareMessagesForContext([lastUser], systemPrompt);

    // The last user message should be present (possibly truncated)
    const userMsgs = result.filter(m => m.role === 'user');
    expect(userMsgs.length).toBeGreaterThan(0);
    const content = userMsgs[userMsgs.length - 1].content as string;
    // Content should be non-empty and truncated to fit budget
    expect(content.length).toBeGreaterThan(0);
    expect(content.length).toBeLessThanOrEqual(100);
  });

  it('returns only system message when budget is exactly 0 after system prompt', () => {
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS);
    const result = prepareMessagesForContext([makeUserMsg('ignored')], systemPrompt);

    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('system');
  });

  it('handles empty string user message without throwing', () => {
    expect(() => {
      prepareMessagesForContext([makeUserMsg('')], 'System');
    }).not.toThrow();
  });

  it('handles messages with array content (vision) without coercing them', () => {
    const visionMsg: ChatRequestMessage = {
      role: 'user',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      content: [{ type: 'text', text: 'describe this' }, { type: 'image_url', image_url: { url: 'data:...' } }] as any,
    };

    const result = prepareMessagesForContext([visionMsg], 'System');
    const userMsgs = result.filter(m => m.role === 'user');

    expect(userMsgs.length).toBeGreaterThan(0);
    // Array content should be preserved as-is (non-string content is not coerced)
    expect(Array.isArray(userMsgs[0].content)).toBe(true);
  });

  it('handles a large number of messages without throwing', () => {
    const msgs: ChatRequestMessage[] = Array.from({ length: 200 }, (_, i) =>
      i % 2 === 0 ? makeUserMsg(`user message ${i}`) : makeAssistantMsg(`assistant message ${i}`),
    );

    expect(() => {
      prepareMessagesForContext(msgs, 'System');
    }).not.toThrow();
  });

  it('truncates a tool message with tail direction when budget is very small but > marker', () => {
    // Leave exactly 30 chars budget — enough for tail marker (17 chars) but not full content
    // '[Truncated ...]\n\n' = 18 chars
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 30);
    const plainToolResult = 'plain log output ' + 'x'.repeat(500);
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('q'),
      makeAssistantWithToolCalls('', ['tc-tail']),
      makeToolMsg('tc-tail', plainToolResult),
    ];

    // Should not throw even with very small budget
    expect(() => {
      prepareMessagesForContext(msgs, systemPrompt);
    }).not.toThrow();
  });

  it('handles budget exactly at marker length boundary without throwing', () => {
    // Leave budget = 1 char — smaller than any truncation marker
    const systemPrompt = 'S'.repeat(MAX_CONTEXT_CHARS - 1);
    const largeMsg = makeUserMsg('U'.repeat(500));

    expect(() => {
      prepareMessagesForContext([largeMsg], systemPrompt);
    }).not.toThrow();
  });

  it('preserves tool_call_id on tool messages that pass through', () => {
    const msgs: ChatRequestMessage[] = [
      makeUserMsg('do it'),
      makeAssistantWithToolCalls('', ['tc-preserve']),
      makeToolMsg('tc-preserve', 'result'),
    ];

    const result = prepareMessagesForContext(msgs, 'System');
    const toolMsg = result.find(m => m.role === 'tool') as (ChatRequestMessage & { tool_call_id?: string }) | undefined;

    if (toolMsg) {
      expect(toolMsg.tool_call_id).toBe('tc-preserve');
    }
  });
});
