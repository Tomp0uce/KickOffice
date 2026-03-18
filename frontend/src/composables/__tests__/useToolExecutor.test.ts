import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ref } from 'vue';

// ── Mocks must be declared before the module import ───────────────────────────

vi.mock('@/utils/logger', () => ({
  logService: { warn: vi.fn(), error: vi.fn(), info: vi.fn() },
}));

import { executeAgentToolCall } from '../useToolExecutor';
import type { ToolCall } from '../useAgentStream';
import type { ToolDefinition } from '@/types';
import type { DisplayMessage } from '@/types/chat';

function makeToolCall(name: string, args: unknown = {}, id = 'tc1'): ToolCall {
  return {
    id,
    type: 'function',
    function: { name, arguments: JSON.stringify(args) },
  };
}

function makeTool(name: string, execute: (args: any) => Promise<string>): ToolDefinition {
  return { name, description: '', parameters: {}, execute } as unknown as ToolDefinition;
}

function makeAssistantMessage(): DisplayMessage {
  return { id: 'msg1', role: 'assistant', content: '' };
}

describe('executeAgentToolCall', () => {
  const currentActionRef = ref('');
  const scrollFn = vi.fn().mockResolvedValue(undefined);
  const getActionLabel = vi.fn().mockReturnValue('Running tool...');

  beforeEach(() => {
    currentActionRef.value = '';
    vi.clearAllMocks();
    scrollFn.mockResolvedValue(undefined);
  });

  // ─── Argument parsing ────────────────────────────────────────────────────

  it('returns an error result when tool arguments are malformed JSON', async () => {
    const badCall: ToolCall = {
      id: 'tc_bad',
      type: 'function',
      function: { name: 'someTool', arguments: '{invalid json}' },
    };
    const msg = makeAssistantMessage();
    const result = await executeAgentToolCall(
      badCall, [], msg, currentActionRef, getActionLabel, scrollFn,
    );
    expect(result.success).toBe(false);
    expect(result.content).toContain('malformed tool arguments');
    expect(msg.toolCalls?.[0].status).toBe('error');
  });

  // ─── Tool lookup ─────────────────────────────────────────────────────────

  it('returns an error result when the tool is not found in enabledToolDefs', async () => {
    const msg = makeAssistantMessage();
    const result = await executeAgentToolCall(
      makeToolCall('unknownTool'), [], msg, currentActionRef, getActionLabel, scrollFn,
    );
    expect(result.success).toBe(false);
    expect(result.content).toContain('unknownTool not found');
    expect(msg.toolCalls?.[0].status).toBe('error');
  });

  // ─── Successful execution ────────────────────────────────────────────────

  it('returns a success result and includes the signature', async () => {
    const tool = makeTool('myTool', async () => 'done');
    const msg = makeAssistantMessage();
    const result = await executeAgentToolCall(
      makeToolCall('myTool', { x: 1 }),
      [tool],
      msg,
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(result.success).toBe(true);
    expect(result.content).toBe('done');
    expect(result.signature).toContain('myTool');
    expect(msg.toolCalls?.[0].status).toBe('complete');
    expect(msg.toolCalls?.[0].result).toBe('done');
  });

  it('sets and then clears currentActionRef during execution', async () => {
    let actionDuringExecution = '';
    const tool = makeTool('actionTool', async () => {
      actionDuringExecution = currentActionRef.value;
      return 'ok';
    });
    await executeAgentToolCall(
      makeToolCall('actionTool'),
      [tool],
      makeAssistantMessage(),
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(actionDuringExecution).toBe('Running tool...');
    expect(currentActionRef.value).toBe('');
  });

  it('calls scrollToBottomFn during execution', async () => {
    const tool = makeTool('scrollTool', async () => 'ok');
    await executeAgentToolCall(
      makeToolCall('scrollTool'),
      [tool],
      makeAssistantMessage(),
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(scrollFn).toHaveBeenCalledOnce();
  });

  // ─── Tool execution failure ──────────────────────────────────────────────

  it('returns an error result when the tool throws', async () => {
    const tool = makeTool('failTool', async () => {
      throw new Error('Office not ready');
    });
    const msg = makeAssistantMessage();
    const result = await executeAgentToolCall(
      makeToolCall('failTool'),
      [tool],
      msg,
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(result.success).toBe(false);
    expect(result.content).toContain('Office not ready');
    expect(msg.toolCalls?.[0].status).toBe('error');
    expect(msg.toolCalls?.[0].result).toBe('Office not ready');
    // currentActionRef must be cleared even on failure
    expect(currentActionRef.value).toBe('');
  });

  // ─── Screenshot detection ────────────────────────────────────────────────

  it('detects a screenshot result and exposes base64 data', async () => {
    const screenshotPayload = JSON.stringify({
      __screenshot__: true,
      base64: 'abc123',
      mimeType: 'image/png',
      description: 'slide 1',
    });
    const tool = makeTool('screenshotTool', async () => screenshotPayload);
    const msg = makeAssistantMessage();
    const result = await executeAgentToolCall(
      makeToolCall('screenshotTool'),
      [tool],
      msg,
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(result.screenshotBase64).toBe('abc123');
    expect(result.screenshotMimeType).toBe('image/png');
    // The content returned to the LLM should be the normalised success message, not the raw payload
    const parsed = JSON.parse(result.content);
    expect(parsed.success).toBe(true);
    // screenshotSrc set on tool call for UI display
    expect(msg.toolCalls?.[0].screenshotSrc).toContain('data:image/png;base64,abc123');
  });

  // ─── No assistantMessage ──────────────────────────────────────────────────

  it('works correctly when assistantMessage is undefined (no crash)', async () => {
    const tool = makeTool('safeTool', async () => 'result');
    const result = await executeAgentToolCall(
      makeToolCall('safeTool'),
      [tool],
      undefined,
      currentActionRef,
      getActionLabel,
      scrollFn,
    );
    expect(result.success).toBe(true);
    expect(result.content).toBe('result');
  });
});
