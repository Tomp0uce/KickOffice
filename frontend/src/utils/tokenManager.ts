import type { ChatRequestMessage } from '@/api/backend';

import { message as messageUtil } from '@/utils/message';
import { logService } from '@/utils/logger';
import { i18n } from '@/i18n';

// GPT-5.2: 400k token context window × 3 chars/token ≈ 1.2M chars (conservative ratio)
export const MAX_CONTEXT_CHARS = 1_200_000;

// Phase 7A: Tool result summarization — heuristic approach
// Tool results older than TOOL_RESULT_KEEP_FULL_COUNT iterations get compressed
const TOOL_RESULT_KEEP_FULL_COUNT = 3; // Keep last 3 tool result groups in full
const TOOL_RESULT_MAX_CHARS = 800; // Compress results longer than 800 chars

const TRUNCATION_MARKER_HEAD = '\n\n[... Truncated]';
const TRUNCATION_MARKER_TAIL = '[Truncated ...]\n\n';
let hasWarnedTruncation = false;

/**
 * QUAL-M3: JSON-aware truncation for tool results.
 * If content starts with '{' or '[', preserve the outer structure so the LLM
 * can still parse the opening and closing brackets of the JSON envelope.
 * Falls back to plain tail-truncation for non-JSON content.
 */
function truncateJsonToolResult(content: string, budget: number): string {
  if (budget <= 0) return '';
  if (content.length <= budget) return content;

  const isObject = content.trimStart().startsWith('{');
  const isArray = !isObject && content.trimStart().startsWith('[');

  if (!isObject && !isArray) {
    // Plain text tool result — keep tail (conclusion)
    const marker = TRUNCATION_MARKER_TAIL;
    if (budget <= marker.length) return marker.slice(0, budget);
    return `${marker}${content.slice(-(budget - marker.length))}`;
  }

  const open = isObject ? '{' : '[';
  const close = isObject ? '}' : ']';
  const inner = `${open} ...[${content.length - budget} chars truncated]... ${close}`;
  if (budget >= inner.length) return inner;
  // budget too small even for the envelope — fall back to plain truncation
  return content.slice(0, budget);
}

/**
 * Phase 7A: Heuristic tool result summarization.
 * Groups messages by tool-call iteration (assistant w/ tool_calls + its tool responses).
 * Keeps the last TOOL_RESULT_KEEP_FULL_COUNT iterations intact.
 * Compresses older tool results to TOOL_RESULT_MAX_CHARS chars.
 */
function summarizeOldToolResults(messages: ChatRequestMessage[]): ChatRequestMessage[] {
  // Find iteration boundaries: assistant messages that issued tool calls
  const iterationStartIndices: number[] = [];
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    if (msg.role === 'assistant' && 'tool_calls' in msg && msg.tool_calls && msg.tool_calls.length > 0) {
      iterationStartIndices.push(i);
    }
  }

  if (iterationStartIndices.length <= TOOL_RESULT_KEEP_FULL_COUNT) return messages;

  // First index of the Nth-from-last iteration — everything before it gets compressed
  const keepFromIndex =
    iterationStartIndices[iterationStartIndices.length - TOOL_RESULT_KEEP_FULL_COUNT];

  return messages.map((msg, i) => {
    if (msg.role !== 'tool' || i >= keepFromIndex) return msg;
    if (typeof msg.content !== 'string' || msg.content.length <= TOOL_RESULT_MAX_CHARS) return msg;

    // QUAL-M3: Use JSON-aware truncation so the LLM receives a valid envelope.
    const compressed = truncateJsonToolResult(msg.content, TOOL_RESULT_MAX_CHARS);
    return { ...msg, content: compressed };
  });
}

/**
 * Truncates content to fit within a character budget.
 * @param direction 'head' keeps the beginning (cuts tail) — best for documents, mails, code.
 *                  'tail' keeps the end (cuts beginning) — best for tool results, logs.
 */
function truncateToBudget(content: any, budget: number, direction: 'head' | 'tail' = 'head'): any {
  if (typeof content !== 'string') return content; // L4 fix: Implicit coercion protection for vision arrays
  if (budget <= 0) return '';
  if (content.length <= budget) {
    hasWarnedTruncation = false; // Reset per full fit
    return content;
  }

  if (!hasWarnedTruncation) {
    messageUtil.warning(
      (i18n.global.t as any)('errorTruncated') ?? 'Message was truncated due to context limits',
    );
    hasWarnedTruncation = true;
  }

  const marker = direction === 'head' ? TRUNCATION_MARKER_HEAD : TRUNCATION_MARKER_TAIL;

  if (budget <= marker.length) {
    logService.warn(`[tokenManager] Message truncated entirely (budget: ${budget})`);
    return marker.slice(0, budget);
  }

  const kept = budget - marker.length;
  logService.warn(
    `[tokenManager] Message truncated (${direction}) by ${content.length - kept} chars`,
  );

  if (direction === 'tail') {
    return `${marker}${content.slice(-kept)}`;
  }
  return `${content.slice(0, kept)}${marker}`;
}

function getMessageContentLength(message: ChatRequestMessage): number {
  let length = 0;
  if (typeof message.content === 'string') {
    length = message.content.length;
  } else if (Array.isArray(message.content)) {
    // CORRECTION (Step 2): Avoid JSON.stringify massif on Base64 strings
    for (const part of message.content) {
      if (part.type === 'text' && part.text) {
        length += part.text.length;
      } else if (part.type === 'image_url') {
        // Images have a fixed token cost regardless of Base64 length
        length += 1000;
      } else if (part.type === 'file') {
        // File references (/v1/files) have a minimal fixed cost (just metadata, content stored server-side)
        length += 200;
      }
    }
  } else {
    length = JSON.stringify(message.content).length;
  }

  if ('tool_calls' in message && message.tool_calls) {
    length += JSON.stringify(message.tool_calls).length;
  }
  return length;
}

/**
 * Estimate context usage percentage for a given set of messages.
 * Uses the same char-based heuristic as prepareMessagesForContext.
 */
export function estimateContextUsagePercent(
  allMessages: ChatRequestMessage[],
  systemPrompt: string,
): number {
  let total = Math.min(systemPrompt.length, MAX_CONTEXT_CHARS);
  const nonSystem = allMessages.filter(m => m.role !== 'system');
  for (const msg of nonSystem) {
    total += getMessageContentLength(msg);
  }
  return Math.min(100, Math.round((total / MAX_CONTEXT_CHARS) * 100));
}

export function prepareMessagesForContext(
  allMessages: ChatRequestMessage[],
  systemPrompt: string,
): ChatRequestMessage[] {
  const safeSystemPrompt = truncateToBudget(systemPrompt, MAX_CONTEXT_CHARS);
  const systemMessage: ChatRequestMessage = { role: 'system', content: safeSystemPrompt };
  const nonSystemMessages = summarizeOldToolResults(
    allMessages.filter(message => message.role !== 'system'),
  );

  let remainingBudget = MAX_CONTEXT_CHARS - safeSystemPrompt.length;
  if (remainingBudget <= 0) return [systemMessage];

  const selectedMessages: Array<{ index: number; message: ChatRequestMessage }> = [];
  const selectedIndices = new Set<number>();

  const addMessageWithBudget = (
    index: number,
    message: ChatRequestMessage,
    forceInclude: boolean = false,
  ): void => {
    if (remainingBudget <= 0 || selectedIndices.has(index)) return;

    const messageLength = getMessageContentLength(message);
    if (messageLength <= remainingBudget) {
      selectedMessages.push({ index, message });
      selectedIndices.add(index);
      remainingBudget -= messageLength;
      return;
    }

    if (message.role === 'tool' || forceInclude) {
      // Tool results: JSON-aware truncation preserves the outer structure.
      // User/assistant: keep the beginning (structure/intent).
      const truncatedContent =
        message.role === 'tool' && typeof message.content === 'string'
          ? truncateJsonToolResult(message.content, remainingBudget)
          : truncateToBudget(message.content, remainingBudget, 'head');
      if (!truncatedContent && !forceInclude) return;

      selectedMessages.push({ index, message: { ...message, content: truncatedContent } });
      selectedIndices.add(index);
      remainingBudget = 0;
    }
  };

  // First priority: System prompt (already added)
  // Second priority: The latest user message and its immediate tool context

  const findLastIndexByRole = (role: ChatRequestMessage['role']): number => {
    for (let index = nonSystemMessages.length - 1; index >= 0; index -= 1) {
      if (nonSystemMessages[index].role === role) return index;
    }
    return -1;
  };

  const lastUserIndex = findLastIndexByRole('user');
  if (lastUserIndex >= 0) {
    addMessageWithBudget(lastUserIndex, nonSystemMessages[lastUserIndex], true);
  }

  // Iterate backwards to add messages, ensuring tool/tool_calls pairs are kept together
  for (let index = nonSystemMessages.length - 1; index >= 0; index -= 1) {
    if (selectedIndices.has(index)) continue;

    const message = nonSystemMessages[index];
    const messageLength = getMessageContentLength(message);

    // If it's a tool call or tool response, we try to include the whole block if it fits
    if (message.role === 'tool' || message.role === 'assistant') {
      if (messageLength > remainingBudget) break;
    } else {
      // Normal user messages can break if they exceed budget, or we can truncate them
      if (messageLength > remainingBudget) break;
    }

    selectedMessages.push({ index, message });
    selectedIndices.add(index);
    remainingBudget -= messageLength;
  }

  selectedMessages.sort((a, b) => a.index - b.index);

  // Ensure tool_calls logic integrity per individual tool_call_id
  // If an assistant message has tool_calls, strip only the ones with no matching tool response
  const finalMessages = selectedMessages.map(entry => entry.message);

  // Collect all tool_call_ids that have a matching 'tool' response in the selected messages
  const respondedToolCallIds = new Set<string>();
  for (const msg of finalMessages) {
    if (msg.role === 'tool' && 'tool_call_id' in msg) {
      respondedToolCallIds.add(msg.tool_call_id);
    }
  }

  for (let i = 0; i < finalMessages.length; i++) {
    const msg = finalMessages[i];
    if (msg.role === 'assistant' && msg.tool_calls && msg.tool_calls.length > 0) {
      // Filter out only the tool_calls that have no matching tool response
      const answeredCalls = msg.tool_calls.filter(tc => respondedToolCallIds.has(tc.id));
      if (answeredCalls.length === 0) {
        // No responses at all → strip all tool_calls
        delete msg.tool_calls;
      } else if (answeredCalls.length < msg.tool_calls.length) {
        // Some orphaned tool_calls → strip only those
        msg.tool_calls = answeredCalls;
      }
      // If all answered → keep as-is
    }
  }

  return [systemMessage, ...selectedMessages.map(entry => entry.message)];
}
