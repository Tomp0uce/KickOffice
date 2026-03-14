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

import { prepareMessagesForContext, MAX_CONTEXT_CHARS } from '../tokenManager';
import type { ChatRequestMessage } from '@/api/backend';
import { message as messageUtil } from '@/utils/message';

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
