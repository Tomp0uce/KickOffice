import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ref } from 'vue';
import type { DisplayMessage } from '@/types/chat';

vi.mock('@/api/backend', () => ({
  chatStream: vi.fn(),
}));

import { useAgentStream } from '@/composables/useAgentStream';
import { chatStream } from '@/api/backend';

beforeEach(() => {
  vi.clearAllMocks();
});

describe('useAgentStream', () => {
  const baseMessages = [
    { role: 'system' as const, content: 'sys' },
    { role: 'user' as const, content: 'hello' },
  ];

  it('returns response with assistant content from stream', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onStream('Hello world');
    });

    const { executeStream } = useAgentStream();
    const { response } = await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
    });

    expect(response.choices[0].message.content).toBe('Hello world');
    expect(response.choices[0].message.role).toBe('assistant');
  });

  it('updates currentAssistantMessage during streaming', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onStream('chunk1');
      opts.onStream('chunk1 chunk2');
    });

    const assistantMsg: DisplayMessage = { id: '1', role: 'assistant', content: '' };
    const { executeStream } = useAgentStream();

    await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
      currentAssistantMessage: assistantMsg,
    });

    expect(assistantMsg.content).toBe('chunk1 chunk2');
  });

  it('clears currentAction on stream data', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onStream('data');
    });

    const currentAction = ref('thinking...');
    const { executeStream } = useAgentStream();

    await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
      currentAction,
    });

    expect(currentAction.value).toBe('');
  });

  it('accumulates tool call deltas', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onToolCallDelta([
        {
          index: 0,
          id: 'tc1',
          function: { name: 'myTool', arguments: '{"a":' },
        },
      ]);
      opts.onToolCallDelta([
        {
          index: 0,
          id: 'tc1',
          function: { arguments: '1}' },
        },
      ]);
    });

    const { executeStream } = useAgentStream();
    const { response } = await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
    });

    expect(response.choices[0].message.tool_calls).toHaveLength(1);
    expect(response.choices[0].message.tool_calls[0].id).toBe('tc1');
    expect(response.choices[0].message.tool_calls[0].function.name).toBe('myTool');
    expect(response.choices[0].message.tool_calls[0].function.arguments).toBe('{"a":1}');
  });

  it('detects length truncation', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onStream('truncated');
      opts.onFinishReason('length');
    });

    const { executeStream } = useAgentStream();
    const { truncatedByLength } = await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
    });

    expect(truncatedByLength).toBe(true);
  });

  it('returns truncatedByLength=false for normal completion', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onStream('done');
      opts.onFinishReason('stop');
    });

    const { executeStream } = useAgentStream();
    const { truncatedByLength } = await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
    });

    expect(truncatedByLength).toBe(false);
  });

  it('filters out sparse tool_calls array entries', async () => {
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      // Create a sparse array by only setting index 2
      opts.onToolCallDelta([
        {
          index: 2,
          id: 'tc3',
          function: { name: 'tool3', arguments: '{}' },
        },
      ]);
    });

    const { executeStream } = useAgentStream();
    const { response } = await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
    });

    // Sparse entries (undefined at index 0, 1) should be filtered out
    const toolCalls = response.choices[0].message.tool_calls;
    expect(toolCalls.every(tc => tc !== undefined)).toBe(true);
    expect(toolCalls).toHaveLength(1);
    expect(toolCalls[0].id).toBe('tc3');
  });

  it('passes tools and abortSignal to chatStream', async () => {
    vi.mocked(chatStream).mockImplementation(async () => {});

    const controller = new AbortController();
    const tools = [{ type: 'function' as const, function: { name: 'test', parameters: {} } }];

    const { executeStream } = useAgentStream();
    await executeStream({
      messages: baseMessages,
      modelTier: 'reasoning',
      tools,
      abortSignal: controller.signal,
    });

    expect(chatStream).toHaveBeenCalledWith(
      expect.objectContaining({
        messages: baseMessages,
        modelTier: 'reasoning',
        tools,
        abortSignal: controller.signal,
      }),
    );
  });

  it('calls accumulateUsage callback', async () => {
    const usage = { prompt_tokens: 10, completion_tokens: 20, total_tokens: 30 };
    vi.mocked(chatStream).mockImplementation(async (opts: any) => {
      opts.onUsage(usage);
    });

    const accumulateUsage = vi.fn();
    const { executeStream } = useAgentStream();
    await executeStream({
      messages: baseMessages,
      modelTier: 'standard',
      accumulateUsage,
    });

    expect(accumulateUsage).toHaveBeenCalledWith(usage);
  });
});
