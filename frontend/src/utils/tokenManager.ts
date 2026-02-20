import type { ChatRequestMessage } from '@/api/backend'

export const MAX_CONTEXT_CHARS = 100_000

const TRUNCATION_MARKER = '\n\n[... Truncated]'

function truncateToBudget(content: string, budget: number): string {
  if (budget <= 0) return ''
  if (content.length <= budget) return content
  if (budget <= TRUNCATION_MARKER.length) {
    return TRUNCATION_MARKER.slice(0, budget)
  }

  const headLength = budget - TRUNCATION_MARKER.length
  return `${content.slice(0, headLength)}${TRUNCATION_MARKER}`
}

function getMessageContentLength(message: ChatRequestMessage): number {
  return message.content.length
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
      const truncatedContent = truncateToBudget(message.content, remainingBudget)
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

  // Ensure tool_calls logic integrity
  // If an assistant message has tool_calls, but the subsequent tool messages were pruned,
  // we should remove those tool_calls from the assistant message to prevent API errors.
  const finalMessages = selectedMessages.map(entry => entry.message)
  for (let i = 0; i < finalMessages.length; i++) {
    const msg = finalMessages[i]
    if (msg.role === 'assistant' && msg.tool_calls && msg.tool_calls.length > 0) {
      // Check if the next messages are the tool responses
      const hasMatchingTools = finalMessages.slice(i + 1).some(m => m.role === 'tool')
      if (!hasMatchingTools) {
        // Strip tool_calls if we pruned the tool responses
        delete msg.tool_calls
      }
    }
  }

  return [systemMessage, ...selectedMessages.map(entry => entry.message)]
}
