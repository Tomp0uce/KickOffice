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

  const lastToolIndex = findLastIndexByRole('tool')
  if (lastToolIndex >= 0) {
    addMessageWithBudget(lastToolIndex, nonSystemMessages[lastToolIndex], false)
  }

  for (let index = nonSystemMessages.length - 1; index >= 0; index -= 1) {
    if (selectedIndices.has(index)) continue

    const message = nonSystemMessages[index]
    const messageLength = getMessageContentLength(message)
    if (messageLength > remainingBudget) break

    selectedMessages.push({ index, message })
    selectedIndices.add(index)
    remainingBudget -= messageLength
  }

  selectedMessages.sort((a, b) => a.index - b.index)

  return [systemMessage, ...selectedMessages.map(entry => entry.message)]
}
