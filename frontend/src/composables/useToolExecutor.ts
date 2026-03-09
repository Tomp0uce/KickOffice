import type { Ref } from 'vue'
import type { DisplayMessage } from '@/types/chat'
import type { ToolCategory } from '@/types'
import { logService } from '@/utils/logger'

/**
 * Safe JSON stringify with depth and circular reference checks
 * Prevents DoS attacks via deeply nested objects
 */
function safeStringify(obj: unknown, maxDepth = 10): string {
  const seen = new WeakSet()

  function stringify(value: unknown, depth: number): string {
    if (depth > maxDepth) {
      return '"[Max depth exceeded]"'
    }

    if (value === null) return 'null'
    if (value === undefined) return 'undefined'
    if (typeof value === 'string') return JSON.stringify(value)
    if (typeof value === 'number' || typeof value === 'boolean') return String(value)

    if (Array.isArray(value)) {
      if (seen.has(value)) return '"[Circular]"'
      seen.add(value)

      const items = value.map(item => stringify(item, depth + 1))
      return `[${items.join(',')}]`
    }

    if (typeof value === 'object') {
      if (seen.has(value)) return '"[Circular]"'
      seen.add(value)

      const entries = Object.entries(value).map(([k, v]) => {
        return `${JSON.stringify(k)}:${stringify(v, depth + 1)}`
      })
      return `{${entries.join(',')}}`
    }

    return '"[Unsupported type]"'
  }

  return stringify(obj, 0)
}

export async function executeAgentToolCall(
  toolCall: any,
  enabledToolDefs: any[],
  assistantMessage: DisplayMessage | undefined,
  currentActionRef: Ref<string>,
  getActionLabelForCategory: (cat?: ToolCategory) => string,
  scrollToBottomFn: () => Promise<void>
) {
  const toolName = toolCall.function.name
  let toolArgs: Record<string, any> = {}
  try {
    const parsed = JSON.parse(toolCall.function.arguments)
    toolArgs = typeof parsed === 'object' && parsed !== null ? parsed : {}
  } catch (parseErr) {
    logService.error('[AgentLoop] Failed to parse tool call arguments', parseErr, { toolName, arguments: toolCall.function.arguments, traffic: 'user' })
    if (assistantMessage) {
      if (!assistantMessage.toolCalls) assistantMessage.toolCalls = []
      assistantMessage.toolCalls.push({ id: toolCall.id, name: toolName, args: {}, status: 'error', result: 'Malformed tool arguments — JSON parse failed' })
    }
    return { tool_call_id: toolCall.id, content: `Error in ${toolName}: malformed tool arguments — JSON parse failed`, success: false }
  }

  const toolDef = enabledToolDefs.find(tool => tool.name === toolName)
  if (!toolDef) {
    if (assistantMessage) {
      if (!assistantMessage.toolCalls) assistantMessage.toolCalls = []
      assistantMessage.toolCalls.push({ id: toolCall.id, name: toolName, args: toolArgs, status: 'error', result: `Tool ${toolName} not found` })
    }
    return { tool_call_id: toolCall.id, content: `Error: Tool ${toolName} not found`, success: false }
  }

  const signature = `${toolName}${safeStringify(toolArgs)}`
  let result = ''
  let success = false

  if (assistantMessage) {
    if (!assistantMessage.toolCalls) assistantMessage.toolCalls = []
    assistantMessage.toolCalls.push({ id: toolCall.id, name: toolName, args: toolArgs, status: 'running' })
  }

  currentActionRef.value = getActionLabelForCategory(toolDef.category)
  await scrollToBottomFn()
  let screenshotBase64: string | undefined
  let screenshotMimeType: string | undefined
  const executionStartTime = Date.now()
  try {
    result = await toolDef.execute(toolArgs)
    // Detect screenshot results
    try {
      const parsed = JSON.parse(result)
      if (parsed && parsed.__screenshot__ === true) {
        screenshotBase64 = parsed.base64
        screenshotMimeType = parsed.mimeType || 'image/png'
        result = JSON.stringify({ success: true, message: parsed.description || 'Screenshot captured. Image injected into vision context for your next response.' })
      }
    } catch {
      // Not JSON or not a screenshot — keep result as is
    }
    success = true
    logService.info('[AgentLoop] tool execution succeeded', 'user', { toolName, duration: Date.now() - executionStartTime })
    if (assistantMessage?.toolCalls) {
      const idx = assistantMessage.toolCalls.findIndex(t => t.id === toolCall.id)
      if (idx !== -1) assistantMessage.toolCalls[idx].status = 'complete'
    }
  } catch (err: unknown) {
    logService.error('[AgentLoop] tool execution failed', err, { toolName, toolArgs, traffic: 'user' })
    const errorMessage = err instanceof Error ? err.message : String(err)
    result = `Error in ${toolName}: ${errorMessage}`
    if (assistantMessage?.toolCalls) {
      const idx = assistantMessage.toolCalls.findIndex(t => t.id === toolCall.id)
      if (idx !== -1) {
         assistantMessage.toolCalls[idx].status = 'error'
         assistantMessage.toolCalls[idx].result = errorMessage
      }
    }
  }
  currentActionRef.value = ''

  let safeContent = ''
  if (result === null || result === undefined) {
    safeContent = ''
  } else if (typeof result === 'object') {
    safeContent = JSON.stringify(result)
  } else {
    safeContent = String(result)
  }

  if (assistantMessage?.toolCalls) {
    const idx = assistantMessage.toolCalls.findIndex(t => t.id === toolCall.id)
    if (idx !== -1 && success) {
      assistantMessage.toolCalls[idx].result = safeContent
    }
  }

  return { tool_call_id: toolCall.id, content: safeContent, success, signature, screenshotBase64, screenshotMimeType }
}
