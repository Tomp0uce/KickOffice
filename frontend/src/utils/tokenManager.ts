import type { ChatRequestMessage } from '@/api/backend'

import { message as messageUtil } from '@/utils/message'
import { i18n } from '@/i18n'

// GPT-5.2: 400k token context window × 3 chars/token ≈ 1.2M chars (conservative ratio)
export const MAX_CONTEXT_CHARS = 1_200_000

const TRUNCATION_MARKER_HEAD = '\n\n[... Truncated]'
const TRUNCATION_MARKER_TAIL = '[Truncated ...]\n\n'
let hasWarnedTruncation = false

/**
 * Truncates content to fit within a character budget.
 * @param direction 'head' keeps the beginning (cuts tail) — best for documents, mails, code.
 *                  'tail' keeps the end (cuts beginning) — best for tool results, logs.
 */
function truncateToBudget(content: any, budget: number, direction: 'head' | 'tail' = 'head'): any {
  if (typeof content !== 'string') return content // L4 fix: Implicit coercion protection for vision arrays
  if (budget <= 0) return ''
  if (content.length <= budget) {
    hasWarnedTruncation = false // Reset per full fit
    return content
  }

  if (!hasWarnedTruncation) {
    messageUtil.warning((i18n.global.t as any)('errorTruncated') ?? 'Message was truncated due to context limits')
    hasWarnedTruncation = true
  }

  const marker = direction === 'head' ? TRUNCATION_MARKER_HEAD : TRUNCATION_MARKER_TAIL

  if (budget <= marker.length) {
    console.warn(`[tokenManager] Message truncated entirely (budget: ${budget})`)
    return marker.slice(0, budget)
  }

  const kept = budget - marker.length
  console.warn(`[tokenManager] Message truncated (${direction}) by ${content.length - kept} chars`)

  if (direction === 'tail') {
    return `${marker}${content.slice(-kept)}`
  }
  return `${content.slice(0, kept)}${marker}`
}

function getMessageContentLength(message: ChatRequestMessage): number {
  let length = 0
  if (typeof message.content === 'string') {
    length = message.content.length
  } else {
    length = JSON.stringify(message.content).length
  }
  if ('tool_calls' in message && message.tool_calls) {
    length += JSON.stringify(message.tool_calls).length
  }
  return length
}

export function prepareMessagesForContext(allMessages: ChatRequestMessage[], systemPrompt: string): ChatRequestMessage[] {
  const safeSystemPrompt = truncateToBudget(systemPrompt, MAX_CONTEXT_CHARS)
  const systemMessage: ChatRequestMessage = { role: 'system', content: safeSystemPrompt }
  const nonSystemMessages = allMessages.filter(message => message.role !== 'system')

  let remainingBudget = MAX_CONTEXT_CHARS - safeSystemPrompt.length
  if (remainingBudget <= 0) return [systemMessage]

  const selectedMessages: Array<{ index: number, message: ChatRequestMessage }> = []
  const selectedIndices = new Set<number>()

  const addMessageWithBudget = (index: number, message: ChatRequestMessage, forceInclude: boolean = false): void => {
    if (remainingBudget <= 0 || selectedIndices.has(index)) return

    const messageLength = getMessageContentLength(message)
    if (messageLength <= remainingBudget) {
      selectedMessages.push({ index, message })
      selectedIndices.add(index)
      remainingBudget -= messageLength
      return
    }

    if (message.role === 'tool' || forceInclude) {
      // Tool results: keep the end (conclusion/result). User/assistant: keep the beginning (structure/intent).
      const dir = message.role === 'tool' ? 'tail' : 'head'
      const truncatedContent = truncateToBudget(message.content, remainingBudget, dir)
      if (!truncatedContent && !forceInclude) return

      selectedMessages.push({ index, message: { ...message, content: truncatedContent } })
      selectedIndices.add(index)
      remainingBudget = 0
    }
  }

  // First priority: System prompt (already added)
  // Second priority: The latest user message and its immediate tool context
  
  const findLastIndexByRole = (role: ChatRequestMessage['role']): number => {
    for (let index = nonSystemMessages.length - 1; index >= 0; index -= 1) {
      if (nonSystemMessages[index].role === role) return index
    }
    return -1
  }

  const lastUserIndex = findLastIndexByRole('user')
  if (lastUserIndex >= 0) {
    addMessageWithBudget(lastUserIndex, nonSystemMessages[lastUserIndex], true)
  }

  // Iterate backwards to add messages, ensuring tool/tool_calls pairs are kept together
  for (let index = nonSystemMessages.length - 1; index >= 0; index -= 1) {
    if (selectedIndices.has(index)) continue

    const message = nonSystemMessages[index]
    const messageLength = getMessageContentLength(message)
    
    // If it's a tool call or tool response, we try to include the whole block if it fits
    if (message.role === 'tool' || message.role === 'assistant') {
      if (messageLength > remainingBudget) break
    } else {
      // Normal user messages can break if they exceed budget, or we can truncate them
      if (messageLength > remainingBudget) break
    }

    selectedMessages.push({ index, message })
    selectedIndices.add(index)
    remainingBudget -= messageLength
  }

  selectedMessages.sort((a, b) => a.index - b.index)

  // Ensure tool_calls logic integrity per individual tool_call_id
  // If an assistant message has tool_calls, strip only the ones with no matching tool response
  const finalMessages = selectedMessages.map(entry => entry.message)

  // Collect all tool_call_ids that have a matching 'tool' response in the selected messages
  const respondedToolCallIds = new Set<string>()
  for (const msg of finalMessages) {
    if (msg.role === 'tool' && 'tool_call_id' in msg) {
      respondedToolCallIds.add(msg.tool_call_id)
    }
  }

  for (let i = 0; i < finalMessages.length; i++) {
    const msg = finalMessages[i]
    if (msg.role === 'assistant' && msg.tool_calls && msg.tool_calls.length > 0) {
      // Filter out only the tool_calls that have no matching tool response
      const answeredCalls = msg.tool_calls.filter(tc => respondedToolCallIds.has(tc.id))
      if (answeredCalls.length === 0) {
        // No responses at all → strip all tool_calls
        delete msg.tool_calls
      } else if (answeredCalls.length < msg.tool_calls.length) {
        // Some orphaned tool_calls → strip only those
        msg.tool_calls = answeredCalls
      }
      // If all answered → keep as-is
    }
  }

  return [systemMessage, ...selectedMessages.map(entry => entry.message)]
}
