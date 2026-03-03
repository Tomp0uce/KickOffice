import type { ModelTier, ModelInfo, ToolCategory } from '@/types'
import { nextTick, ref, type Ref } from 'vue'

import { type ChatMessage, type ChatRequestMessage, type TokenUsage, chatStream, generateImage, uploadFile, categorizeError } from '@/api/backend'
import { GLOBAL_STYLE_INSTRUCTIONS, builtInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt, getOutlookBuiltInPrompt, getPowerPointBuiltInPrompt, outlookBuiltInPrompt, powerPointBuiltInPrompt } from '@/utils/constant'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { message as messageUtil } from '@/utils/message'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointToolDefinitions } from '@/utils/powerpointTools'
import { prepareMessagesForContext } from '@/utils/tokenManager'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { getEnabledToolNamesFromStorage } from '@/utils/toolStorage'
import { extractTextFromHtml, reassembleWithFragments, getPreservationInstruction, type RichContentContext } from '@/utils/richContentPreserver'
import { applyInheritedStyles, renderOfficeCommonApiHtml } from '@/utils/markdown'
import { useAgentPrompts } from '@/composables/useAgentPrompts'
import { useOfficeSelection } from '@/composables/useOfficeSelection'
import {
  getExcelDocumentContext,
  getPowerPointDocumentContext,
  getOutlookDocumentContext,
  getWordDocumentContext,
} from '@/utils/officeDocumentContext'

import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, OutlookQuickAction, QuickAction, ToolCallPart } from '@/types/chat'


interface ToolCallFunction {
  name: string
  arguments: string
}

interface ToolCall {
  id: string
  type: 'function'
  function: ToolCallFunction
}

interface AssistantMessage {
  role: 'assistant'
  content: string
  tool_calls: ToolCall[]
}

interface StreamResponseChoice {
  message: AssistantMessage
  finish_reason?: string | null
}

interface StreamResponse {
  choices: StreamResponseChoice[]
}

interface AgentLoopRefs {
  history: Ref<DisplayMessage[]>
  userInput: Ref<string>
  loading: Ref<boolean>
  imageLoading: Ref<boolean>
  backendOnline: Ref<boolean>
  abortController: Ref<AbortController | null>
  inputTextarea: Ref<HTMLTextAreaElement | undefined>
  draftFocusGlow: Ref<boolean>
}

interface AgentLoopModels {
  availableModels: Ref<Record<string, ModelInfo>>
  selectedModelTier: Ref<ModelTier>
  selectedModelInfo: Ref<ModelInfo | undefined>
  firstChatModelTier: Ref<ModelTier>
}

interface AgentLoopHost {
  isOutlook: boolean
  isPowerPoint: boolean
  isExcel: boolean
  isWord: boolean
}

interface AgentLoopSettings {
  customSystemPrompt: Ref<string>
  agentMaxIterations: Ref<number>
  useSelectedText: Ref<boolean>
  excelFormulaLanguage: Ref<'en' | 'fr'>
  userGender: Ref<string>
  userFirstName: Ref<string>
  userLastName: Ref<string>
}

interface AgentLoopActions {
  quickActions: Ref<QuickAction[]>
  outlookQuickActions?: Ref<OutlookQuickAction[]>
  excelQuickActions: Ref<ExcelQuickAction[]>
  powerPointQuickActions: Ref<PowerPointQuickAction[]>
}

interface AgentLoopHelpers {
  createDisplayMessage: (role: DisplayMessage['role'], content: string, imageSrc?: string) => DisplayMessage
  adjustTextareaHeight: () => void
  scrollToBottom: () => Promise<void>
  scrollToMessageTop?: () => Promise<void>
  scrollToVeryBottom?: () => Promise<void>
}

interface UseAgentLoopOptions {
  t: (key: string) => string
  refs: AgentLoopRefs
  models: AgentLoopModels
  host: AgentLoopHost
  settings: AgentLoopSettings
  actions: AgentLoopActions
  helpers: AgentLoopHelpers
}

/**
 * Check if an error is a 401 credential error from LiteLLM
 */
function isCredentialError(error: unknown): boolean {
  if (!error) return false
  const errObj = error as Record<string, any>
  const message = (errObj.message || String(error)).toString()
  return (
    message.includes('401') ||
    message.includes('LiteLLM user credentials') ||
    message.includes('X-User-Key') ||
    message.includes('X-User-Email')
  )
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
    toolArgs = JSON.parse(toolCall.function.arguments)
  } catch (parseErr) {
    console.error('[AgentLoop] Failed to parse tool call arguments', { toolName, arguments: toolCall.function.arguments, error: parseErr })
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

  const signature = `${toolName}${JSON.stringify(toolArgs)}`
  let result = ''
  let success = false

  if (assistantMessage) {
    if (!assistantMessage.toolCalls) assistantMessage.toolCalls = []
    assistantMessage.toolCalls.push({ id: toolCall.id, name: toolName, args: toolArgs, status: 'running' })
  }

  currentActionRef.value = getActionLabelForCategory(toolDef.category)
  await scrollToBottomFn()
  try {
    result = await toolDef.execute(toolArgs)
    success = true
    if (assistantMessage?.toolCalls) {
      const idx = assistantMessage.toolCalls.findIndex(t => t.id === toolCall.id)
      if (idx !== -1) assistantMessage.toolCalls[idx].status = 'complete'
    }
  } catch (err: unknown) {
    console.error('[AgentLoop] tool execution failed', { toolName, toolArgs, error: err })
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

  return { tool_call_id: toolCall.id, content: safeContent, success, signature }
}

export function useAgentLoop(options: UseAgentLoopOptions) {
  const { t, refs, models, host, settings, actions, helpers } = options

  // Destructure refs
  const {
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    abortController,
    inputTextarea,
    draftFocusGlow,
  } = refs

  // Destructure models
  const {
    availableModels,
    selectedModelTier,
    selectedModelInfo,
    firstChatModelTier,
  } = models

  // Destructure host flags (aliased to match existing code)
  const {
    isOutlook: hostIsOutlook,
    isPowerPoint: hostIsPowerPoint,
    isExcel: hostIsExcel,
  } = host

  // Destructure settings
  const {
    customSystemPrompt,
    agentMaxIterations,
    useSelectedText,
    excelFormulaLanguage,
    userGender,
    userFirstName,
    userLastName,
  } = settings

  // Destructure actions
  const {
    quickActions,
    outlookQuickActions,
    excelQuickActions,
    powerPointQuickActions,
  } = actions

  // Destructure helpers
  const {
    createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop = scrollToBottom, // fallback to scrollToBottom if not provided
    scrollToVeryBottom = scrollToBottom, // fallback to scrollToBottom if not provided
  } = helpers

  const currentAction = ref('')
  const pendingSmartReply = ref(false)

  const sessionStats = ref({
    inputTokens: 0,
    outputTokens: 0,
    totalTokens: 0,
  })

  function resetSessionStats() {
    sessionStats.value = { inputTokens: 0, outputTokens: 0, totalTokens: 0 }
  }

  function accumulateUsage(usage: TokenUsage) {
    sessionStats.value.inputTokens += usage.promptTokens
    sessionStats.value.outputTokens += usage.completionTokens
    sessionStats.value.totalTokens += usage.totalTokens
  }

  const getActionLabelForCategory = (category?: ToolCategory) => {
    switch (category) {
      case 'read':
        return t('agentActionReading')
      case 'format':
        return t('agentActionFormatting')
      case 'write':
      default:
        return t('agentActionRunning')
    }
  }

  const { agentPrompt } = useAgentPrompts({
    t,
    userGender,
    userFirstName,
    userLastName,
    excelFormulaLanguage,
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
  })

  function buildChatMessages(systemPrompt: string): ChatMessage[] {
    const msgs: ChatRequestMessage[] = [{ role: 'system', content: systemPrompt }]
    for (const m of history.value) {
      if (m.rawMessages && m.rawMessages.length > 0) {
        // This includes the assistant message and all tool responses from that turn
        msgs.push(...m.rawMessages)
      } else {
        msgs.push({ role: m.role, content: m.content })
      }
    }
    return msgs as ChatMessage[]
  }

  const { getOfficeSelection, getOfficeSelectionAsHtml } = useOfficeSelection({
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
  })

  const resolveChatModelTier = (): ModelTier => (
    selectedModelInfo.value?.type === 'image' ? firstChatModelTier.value : selectedModelTier.value
  )





async function runAgentLoop(messages: ChatMessage[], modelTier: ModelTier) {
    let appToolDefs = getWordToolDefinitions()
    if (hostIsOutlook) appToolDefs = getOutlookToolDefinitions()
    else if (hostIsPowerPoint) appToolDefs = getPowerPointToolDefinitions()
    else if (hostIsExcel) appToolDefs = getExcelToolDefinitions()
    
    const generalToolDefs = getGeneralToolDefinitions()
    const allToolDefs = [...generalToolDefs, ...appToolDefs]
    const enabledToolNames = getEnabledToolNamesFromStorage(allToolDefs.map(def => def.name))
    const enabledToolDefs = allToolDefs.filter(def => enabledToolNames.has(def.name))
    const tools = enabledToolDefs.map(def => ({ type: 'function' as const, function: { name: def.name, description: def.description, parameters: def.inputSchema as Record<string, any> } }))
    let iteration = 0
    const maxIter = Number(agentMaxIterations.value) || 10
    const startTime = Date.now()
    const timeoutMs = maxIter * 60 * 1000 // up to 1 minute per iteration allowed
    let currentMessages: ChatRequestMessage[] = [...messages]
    // Sliding window of last N signatures to detect repetitive loops (P6)
    const LOOP_WINDOW_SIZE = 5
    const LOOP_REPEAT_THRESHOLD = 2
    const recentSignatures: string[] = []
    let toolsWereExecuted = false // Track if any tools were successfully executed
    currentAction.value = t('agentAnalyzing')
    history.value.push(createDisplayMessage('assistant', ''))
    await scrollToMessageTop() // Scroll to show start of assistant response
    const currentAssistantMessage = history.value[history.value.length - 1]
    let abortedByUser = false
    while (Date.now() - startTime < timeoutMs) {
      if (abortController.value?.signal.aborted) {
        abortedByUser = true
        break
      }

      iteration++
      currentAction.value = t('agentAnalyzing')
      const currentSystemPrompt = messages[0]?.role === 'system' ? messages[0].content : ''
      const contextSafeMessages = prepareMessagesForContext(currentMessages, currentSystemPrompt)
      let response: StreamResponse = { choices: [{ message: { role: 'assistant', content: '', tool_calls: [] } }] }
      let truncatedByLength = false
      try {
        let streamStarted = false
        await chatStream({
          messages: contextSafeMessages,
          modelTier,
          tools,
          abortSignal: abortController.value?.signal,
          onStream: (text) => {
            if (!streamStarted) {
              streamStarted = true
              currentAction.value = ''
            }
            currentAssistantMessage.content = text
            response.choices[0].message.content = text
            scrollToBottom().catch(console.error)
          },
          onToolCallDelta: (toolCallDeltas) => {
            if (!streamStarted) {
              streamStarted = true
              currentAction.value = ''
            }
            for (const delta of toolCallDeltas) {
              const idx = delta.index
              if (!response.choices[0].message.tool_calls[idx]) {
                response.choices[0].message.tool_calls[idx] = { id: delta.id, type: 'function', function: { name: delta.function.name || '', arguments: '' } }
              }
              if (delta.function?.arguments) {
                response.choices[0].message.tool_calls[idx].function.arguments += delta.function.arguments
              }
            }
          },
          onFinishReason: (finishReason) => {
            if (finishReason === 'length') truncatedByLength = true
          },
          onUsage: accumulateUsage,
        })
        response.choices[0].message.tool_calls = response.choices[0].message.tool_calls.filter(Boolean)
      } catch (err: unknown) {
        if ((err instanceof Error && err.name === 'AbortError') || abortController.value?.signal.aborted) {
          abortedByUser = true
          break
        }
        console.error('[AgentLoop] chatStream failed', {
          host: hostIsOutlook ? 'outlook' : hostIsPowerPoint ? 'powerpoint' : hostIsExcel ? 'excel' : 'word',
          modelTier,
          iteration,
          messageCount: currentMessages.length,
          error: err,
        })
        const errInfo = categorizeError(err)
        if (errInfo.type === 'auth') {
          currentAssistantMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
        } else {
          currentAssistantMessage.content = t(errInfo.i18nKey)
        }
        currentAction.value = ''
        break
      }

      // Handle finish_reason: "length" — model was cut off mid-response (P7)
      if (truncatedByLength) {
        currentAction.value = ''
        if (!currentAssistantMessage.content?.trim()) {
          currentAssistantMessage.content = t('errorTruncated')
        } else {
          // Append warning to existing content
          currentAssistantMessage.content += `\n\n${t('errorTruncated')}`
        }
        break
      }

      const choice = response.choices?.[0]
      if (!choice) break
      const assistantMsg = choice.message
      const assistantMsgForHistory: ChatRequestMessage = {
        role: 'assistant',
        content: assistantMsg.content || '',
      }
      // Only include tool_calls if non-empty (Azure/LiteLLM rejects empty arrays)
      if (assistantMsg.tool_calls?.length) {
        assistantMsgForHistory.tool_calls = assistantMsg.tool_calls
      }
      currentMessages.push(assistantMsgForHistory)
      if (assistantMsg.content) currentAssistantMessage.content = assistantMsg.content
      if (!assistantMsg.tool_calls?.length) {
        currentAction.value = ''
        break
      }
      // Collect all tool results before adding to messages (atomic update)
      const toolResults: { tool_call_id: string; content: string }[] = []
      let toolLoopAborted = false

      for (const toolCall of assistantMsg.tool_calls) {
        // Check abort before each tool execution
        if (abortController.value?.signal.aborted) {
          toolLoopAborted = true
          break
        }

        const toolResult = await executeAgentToolCall(toolCall, enabledToolDefs, currentAssistantMessage, currentAction, getActionLabelForCategory, scrollToBottom)
        const sig = toolResult.signature

        // Sliding window loop detection (P6)
        if (sig) {
          recentSignatures.push(sig)
          if (recentSignatures.length > LOOP_WINDOW_SIZE) recentSignatures.shift()
          const sigCount = recentSignatures.filter(s => s === sig).length
          if (sigCount >= LOOP_REPEAT_THRESHOLD) {
            toolResults.push({ tool_call_id: toolCall.id, content: 'Error: You have called this exact tool with the same arguments multiple times in a row. This is a loop. Stop repeating and try a different approach.' })
            continue
          }
        }

        if (toolResult.success) toolsWereExecuted = true
        toolResults.push({ tool_call_id: toolResult.tool_call_id, content: toolResult.content })
      }

      // If aborted mid-tool-loop, rollback partial state by removing incomplete assistant message
      if (toolLoopAborted) {
        // Remove the last assistant message with tool_calls since we didn't complete all tools
        const lastMsgIdx = currentMessages.length - 1
        if (lastMsgIdx >= 0 && currentMessages[lastMsgIdx].role === 'assistant') {
          currentMessages.pop()
        }
        abortedByUser = true
        break
      }

      // Atomically add all tool results now that loop completed successfully
      for (const toolResult of toolResults) {
        currentMessages.push({ role: 'tool', tool_call_id: toolResult.tool_call_id, content: toolResult.content })
      }
      currentAction.value = t('agentAnalyzing')
    }

    // P8: Persist full tool call sequence in history so subsequent turns have context
    const initialMsgCount = messages.length
    const newMessages = currentMessages.slice(initialMsgCount)
    if (newMessages.length > 0) {
      currentAssistantMessage.rawMessages = newMessages
    }

    if (abortedByUser) {
      currentAction.value = ''
      history.value.push(createDisplayMessage('system', t('agentStoppedByUser')))
      return
    }

    const assistantContent = currentAssistantMessage?.content?.trim() || ''
    if (!assistantContent) {
      // If tools were executed successfully but no text response, that's OK (e.g., proofreading with comments)
      if (toolsWereExecuted) {
        currentAssistantMessage.content = t('toolsExecutedSuccessfully')
      } else {
        currentAssistantMessage.content = t('noModelResponse')
      }
    }

    if (Date.now() - startTime >= timeoutMs) messageUtil.warning(t('recursionLimitExceeded'))
    currentAction.value = ''
  }

  async function handleSmartReply(userMessage: string) {
    pendingSmartReply.value = false
    const replyIntent = userMessage
    // Fetch the full email body for context
    let emailBody = ''
    try {
      emailBody = await getOfficeSelection({ actionKey: 'reply' })
    } catch (err) {
      console.warn('[AgentLoop] Failed to fetch email body for smart reply', err)
    }
    if (!emailBody) {
      messageUtil.error(t('selectEmailPrompt'))
      return
    }
    const lang = localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français'
    const replyPrompt = getOutlookBuiltInPrompt()['reply']
    const systemMsg = replyPrompt.system(lang) + `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`
    const sanitizedEmail = '\\n<email_content>\\n' + emailBody.replace(new RegExp('</?email_content>', 'g'), '') + '\\n<'+'/email_content>\\n'
    const sanitizedIntent = '\\n<user_intent>\\n' + replyIntent.replace(new RegExp('</?user_intent>', 'g'), '') + '\\n<'+'/user_intent>\\n'
    const userMsg = replyPrompt.user(sanitizedEmail, lang).replace('[REPLY_INTENT]', sanitizedIntent)
    history.value.push(createDisplayMessage('assistant', ''))
    await scrollToMessageTop()
    try {
      await chatStream({
        messages: [{ role: 'system', content: systemMsg }, { role: 'user', content: userMsg }],
        modelTier: resolveChatModelTier(),
        onStream: async (text: string) => {
          const message = history.value[history.value.length - 1]
          message.role = 'assistant'
          message.content = text
          await scrollToBottom()
        },
        onUsage: accumulateUsage,
        abortSignal: abortController.value?.signal,
      })
      const lastMessage = history.value[history.value.length - 1]
      if (!lastMessage?.content?.trim()) {
        lastMessage.content = t('noModelResponse')
      }
    } catch (err: unknown) {
      if (err instanceof Error && err.name === 'AbortError') return
      console.error('[AgentLoop] Smart reply chatStream failed', err)
      const lastMessage = history.value[history.value.length - 1]
      const errInfo = categorizeError(err)
      if (errInfo.type === 'auth') {
        lastMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
      } else {
        lastMessage.content = t(errInfo.i18nKey)
      }
    }
  }

  async function fetchSelectionWithTimeout() {
    let timeoutId: ReturnType<typeof setTimeout> | null = null
    let localSelectedText = ''
    try {
      const timeoutPromise = new Promise<string>((_, reject) => {
        timeoutId = setTimeout(() => reject(new Error('getOfficeSelection timeout')), 3000)
      }).catch(() => '') as Promise<string>
      
      if (!hostIsExcel) {
        // F1: Extract formatted HTML natively and convert to markdown to preserve styling (Word, PPT, Outlook)
        const htmlPromise = new Promise<string>((_, reject) => {
          timeoutId = setTimeout(() => reject(new Error('getOfficeSelectionAsHtml timeout')), 3000)
        }).catch(() => '') as Promise<string>
        
        try {
          const htmlContent = await Promise.race([getOfficeSelectionAsHtml({ includeOutlookSelectedText: true }), htmlPromise])
          if (htmlContent) {
             const richContext = extractTextFromHtml(htmlContent)
             localSelectedText = richContext.cleanText || localSelectedText
          } else {
             localSelectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
          }
        } catch {
          localSelectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
        }
      } else {
        localSelectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
      }
    } catch (error) {
      console.warn('[AgentLoop] Failed to fetch selection before sending message', error)
    } finally {
      if (timeoutId) clearTimeout(timeoutId)
    }
    return localSelectedText
  }

  async function processChat(userMessage: string, visionImages?: Array<{ filename: string; dataUri: string }>) {
    const modelConfig = availableModels.value[selectedModelTier.value]
    if (modelConfig?.type === 'image') {
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')))
      await scrollToMessageTop() // Scroll to top of assistant message
      imageLoading.value = true
      try {
        const imageSrc = await generateImage({ prompt: userMessage })
        const message = history.value[history.value.length - 1]
        message.role = 'assistant'; message.content = ''; message.imageSrc = imageSrc
      } catch (err: unknown) {
        console.error('[AgentLoop] image generation failed', err)
        const message = history.value[history.value.length - 1]
        message.role = 'assistant'; message.content = t('imageError'); message.imageSrc = undefined
      } finally {
        imageLoading.value = false
      }
      await scrollToBottom() // Final scroll after image loads
      return
    }
    const systemPrompt = customSystemPrompt.value || agentPrompt(localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français')
    const messages = buildChatMessages(systemPrompt)
    const modelTier = resolveChatModelTier()

    // Inject document context into the last user message (not shown in UI — messages is a new array copy)
    try {
      let docContextJson = ''
      if (hostIsExcel) docContextJson = await getExcelDocumentContext()
      else if (hostIsPowerPoint) docContextJson = await getPowerPointDocumentContext()
      else if (hostIsOutlook) docContextJson = await getOutlookDocumentContext()
      else docContextJson = await getWordDocumentContext()

      if (docContextJson) {
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
        if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
          messages[lastUserIdx].content += `\n\n<doc_context>\n${docContextJson}\n</doc_context>`
        }
      }
    } catch (ctxErr) {
      console.warn('[AgentLoop] Failed to fetch document context', ctxErr)
    }

    // Inject vision images as multipart content into the last user message
    if (visionImages && visionImages.length > 0) {
      const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
      if (lastUserIdx !== -1) {
        const textContent = messages[lastUserIdx].content || userMessage
        const parts: any[] = [{ type: 'text', text: textContent }]
        for (const img of visionImages) {
          parts.push({ type: 'image_url', image_url: { url: img.dataUri } })
        }
        ;(messages[lastUserIdx] as any).content = parts
      }
    }

    await runAgentLoop(messages, modelTier)
  }

  async function sendMessage(payload?: string, files?: File[]) {
    let textToSend = ''

    if (payload) {
      textToSend = payload
    } else if (userInput.value && typeof userInput.value === 'string') {
      textToSend = userInput.value
    }

    textToSend = textToSend?.trim() || ''

    if (!textToSend) {
      if (availableModels.value[selectedModelTier.value]?.type !== 'image') {
        return
      }
    }

    if (loading.value) {
      return
    }
    
    loading.value = true

    if (!backendOnline.value) {
      loading.value = false
      return messageUtil.error(t('backendOffline'))
    }

    if (userInput.value.trim() === textToSend) {
      userInput.value = ''
      adjustTextareaHeight()
    }

    const userMessage = textToSend

    let isImageFromSelection = false
    let selectedText = ''
    
    // For direct image generation from selection
    if (!userMessage && availableModels.value[selectedModelTier.value]?.type === 'image') {
      try {
        selectedText = await getOfficeSelection()
      } catch (err) {
        console.warn('[AgentLoop] Failed to fetch selection for image generation', err)
      }
      const wordCount = selectedText.trim().split(/\s+/).filter(w => w.length > 0).length
      if (wordCount < 5) {
        loading.value = false
        return messageUtil.error(t('fileExtractError'))
      }
      isImageFromSelection = true
    }

    abortController.value = new AbortController()

    // If it's pure selection image, we show the selection as the user message bubble
    const displayMessageText = isImageFromSelection ? selectedText : userMessage
    history.value.push(createDisplayMessage('user', displayMessageText))
    await scrollToVeryBottom() // Scroll to very bottom after user message

    try {
      // Smart reply interception: when user sends after clicking "Reply" quick action
      if (pendingSmartReply.value && hostIsOutlook) {
        await handleSmartReply(userMessage)
        return
      }

      // If we haven't fetched it yet and it's enabled
      if (useSelectedText.value && !isImageFromSelection) {
        selectedText = await fetchSelectionWithTimeout()
      }

      let fullMessage = displayMessageText
      let extractedFilesContext = ''
      // Images uploaded are sent as vision content (base64 data-URIs)
      const uploadedImages: Array<{ filename: string; dataUri: string }> = []

      if (files && files.length > 0) {
        currentAction.value = t('agentUploadingFiles') || 'Extraction des fichiers...'
        try {
           for (const file of files) {
             const result = await uploadFile(file)
             if (result.imageBase64) {
               // Image file: store for vision injection
               uploadedImages.push({ filename: result.filename, dataUri: result.imageBase64 })
               // Show a preview thumbnail in the user message bubble
               history.value[history.value.length - 1].imageSrc = result.imageBase64
             } else {
               extractedFilesContext += `\n\n[Contenu extrait du fichier "${result.filename}"]:\n${result.extractedText}\n[Fin du fichier]`
             }
           }
        } catch (uploadObjErr: unknown) {
           console.error('[AgentLoop] File upload/extraction failed', uploadObjErr)
           messageUtil.error(t('somethingWentWrong'))
           return
        }
      }

      // Only append context to standard text chats, not pure image generations
      if (isImageFromSelection) {
        fullMessage = t('imageGenerationPrompt').replace('{text}', selectedText)
      } else {
        if (selectedText) {
           const selectionLabel = hostIsOutlook ? 'Selected text' : hostIsPowerPoint ? 'Selected slide text' : hostIsExcel ? 'Selected cells' : 'Selected text'
           const sanitizedText = '\\n<document_content>\\n' + selectedText.replace(new RegExp('</?document_content>', 'g'), '') + '\\n<'+'/document_content>\\n'
           fullMessage += '\\n\\n[' + selectionLabel + ']: ' + sanitizedText
        }
        if (extractedFilesContext) {
           fullMessage += extractedFilesContext
        }
        history.value[history.value.length - 1].content = fullMessage.trim()
      }

      await processChat(fullMessage.trim(), uploadedImages.length > 0 ? uploadedImages : undefined)
    } catch (error: unknown) {
      if (!(error instanceof Error) || error.name !== 'AbortError') {
        console.error('[AgentLoop] sendMessage failed', error)
        const errInfo = categorizeError(error)
        messageUtil.error(t(errInfo.i18nKey))

      }
    } finally {
      currentAction.value = ''
      loading.value = false
      abortController.value = null
    }
  }

  async function applyQuickAction(actionKey: string) {
    if (!backendOnline.value) return messageUtil.error(t('backendOffline'))
    const selectedQuickAction = hostIsExcel
      ? excelQuickActions.value.find(a => a.key === actionKey)
      : hostIsPowerPoint
        ? powerPointQuickActions.value.find(a => a.key === actionKey)
        : hostIsOutlook && outlookQuickActions
          ? outlookQuickActions.value?.find(a => a.key === actionKey)
          : quickActions.value.find(a => a.key === actionKey)

    const selectedExcelQuickAction = hostIsExcel ? selectedQuickAction as ExcelQuickAction | undefined : undefined
    const selectedPowerPointQuickAction = hostIsPowerPoint ? selectedQuickAction as PowerPointQuickAction | undefined : undefined
    const selectedOutlookQuickAction = hostIsOutlook ? selectedQuickAction as OutlookQuickAction | undefined : undefined

    if (actionKey === 'visual' && hostIsPowerPoint) {
      const imageModelTier = Object.entries(availableModels.value).find(([_, info]) => info.type === 'image')?.[0] as ModelTier
      if (!imageModelTier) {
        return messageUtil.error(t('imageError') || 'No image model configured.')
      }

      // Clear input so `sendMessage` detects `isImageFromSelection = true`
      userInput.value = ''
      adjustTextareaHeight()

      const previousModelTier = selectedModelTier.value
      selectedModelTier.value = imageModelTier
      try {
        await sendMessage()
      } finally {
        selectedModelTier.value = previousModelTier
      }
      return
    }

    if (selectedOutlookQuickAction?.mode === 'smart-reply') {
      pendingSmartReply.value = true
      userInput.value = selectedOutlookQuickAction.prefix || ''
      adjustTextareaHeight()
      draftFocusGlow.value = true
      setTimeout(() => { draftFocusGlow.value = false; }, 1500)
      await nextTick()
      const el = inputTextarea.value
      if (el) {
        el.focus()
        const len = userInput.value.length
        el.setSelectionRange(len, len)
      }
      return
    }
    if (selectedOutlookQuickAction?.mode === 'draft') {
      userInput.value = selectedOutlookQuickAction.prefix || ''
      adjustTextareaHeight()
      draftFocusGlow.value = true
      setTimeout(() => { draftFocusGlow.value = false; }, 1500)
      await nextTick()
      const el = inputTextarea.value
      if (el) {
        el.focus()
        const len = userInput.value.length
        el.setSelectionRange(len, len)
      }
      return
    }
    if (selectedExcelQuickAction?.mode === 'draft') {
      userInput.value = selectedExcelQuickAction.prefix || ''
      adjustTextareaHeight()
      draftFocusGlow.value = true
      setTimeout(() => { draftFocusGlow.value = false; }, 1000)
      await nextTick()
      const el = inputTextarea.value
      if (el) {
        el.focus()
        const len = userInput.value.length
        el.setSelectionRange(len, len)
      }
      return
    }
    if (loading.value) return
    loading.value = true
    abortController.value = new AbortController()

    try {
      const selectedText = await getOfficeSelection({ includeOutlookSelectedText: true, actionKey })
      if (!selectedText) {
        messageUtil.error(t(hostIsOutlook ? 'selectEmailPrompt' : hostIsPowerPoint ? 'selectSlideTextPrompt' : hostIsExcel ? 'selectCellsPrompt' : 'selectTextPrompt'))
        return
      }

      // F1: Try to get HTML selection for rich content preservation (Word, Outlook)
      let richContext: RichContentContext | null = null
      const isTextModifyingAction = !selectedQuickAction?.executeWithAgent && !hostIsExcel
      if (isTextModifyingAction) {
        try {
          const htmlContent = await getOfficeSelectionAsHtml({ includeOutlookSelectedText: true, actionKey })
          if (htmlContent) {
            richContext = extractTextFromHtml(htmlContent)
          }
        } catch (err) {
          console.warn('[AgentLoop] Failed to get HTML selection for rich content preservation', err)
        }
      }

      // Use Markdown text if HTML was parsed successfully, otherwise fallback to plain text selection
      const rawTextForLlm = richContext ? richContext.cleanText : selectedText
      const textForLlm = '\\n<document_content>\\n' + rawTextForLlm.replace(new RegExp('</?document_content>', 'g'), '') + '\\n<'+'/document_content>\\n'

      let action: { system: (lang: string) => string, user: (text: string, lang: string) => string } | undefined
      let systemMsg = ''
      let userMsg = ''
      if (hostIsOutlook) {
        action = getOutlookBuiltInPrompt()[actionKey as keyof typeof outlookBuiltInPrompt] || getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt]
      } else if (hostIsPowerPoint) {
        action = getPowerPointBuiltInPrompt()[actionKey as keyof typeof powerPointBuiltInPrompt] || getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt]
      } else if (hostIsExcel) {
        if (selectedExcelQuickAction?.mode === 'immediate' && selectedExcelQuickAction.systemPrompt) {
          systemMsg = selectedExcelQuickAction.systemPrompt
          userMsg = `Selection:\n${selectedText}`
        } else action = getExcelBuiltInPrompt()[actionKey as keyof typeof excelBuiltInPrompt]
      } else action = getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt]
      if (!systemMsg || !userMsg) {
        if (!action) action = getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt]
        if (!action) return
        const lang = localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français'
        systemMsg = action.system(lang)
        userMsg = action.user(textForLlm, lang)
      }

      // Enforce global formatting constraints on all Quick Actions
      systemMsg += `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`

      // F1: Add preservation instruction if rich content was detected
      if (richContext?.hasRichContent) {
        systemMsg += getPreservationInstruction(richContext)
      }

      const actionLabel = selectedQuickAction?.label || t(actionKey)
      history.value.push(createDisplayMessage('user', `[${actionLabel}] ${selectedText.substring(0, 100)}...`))

      if (selectedQuickAction?.executeWithAgent) {
        await runAgentLoop([{ role: 'system', content: systemMsg }, { role: 'user', content: userMsg }], resolveChatModelTier())
      } else {
        history.value.push(createDisplayMessage('assistant', ''))
        await scrollToMessageTop() // Scroll to show start of assistant response
        try {
          await chatStream({
            messages: [{ role: 'system', content: systemMsg }, { role: 'user', content: userMsg }],
            modelTier: resolveChatModelTier(),
            onStream: async (text: string) => {
              const message = history.value[history.value.length - 1]
              message.role = 'assistant'
              message.content = text
              await scrollToBottom()
            },
            abortSignal: abortController.value?.signal,
          })
          // Check for empty response
          const lastMessage = history.value[history.value.length - 1]
          if (!lastMessage?.content?.trim()) {
            lastMessage.content = t('noModelResponse')
          }
          // F1: Reassemble rich content with preserved fragments and inject native styles
          if (lastMessage?.content) {
            let finalHtml = ''
            if (richContext?.hasRichContent) {
              finalHtml = reassembleWithFragments(lastMessage.content, richContext)
            }
            if (richContext?.extractedStyles && hostIsOutlook) {
              if (!finalHtml) finalHtml = renderOfficeCommonApiHtml(lastMessage.content)
              finalHtml = applyInheritedStyles(finalHtml, richContext.extractedStyles)
            }
            if (finalHtml) {
              lastMessage.richHtml = finalHtml
            }
          }
        } catch (err: any) {
          if (err.name === 'AbortError') return
          console.error('[AgentLoop] Quick action chatStream failed', err)
          const lastMessage = history.value[history.value.length - 1]
          const errInfo = categorizeError(err)
          if (errInfo.type === 'auth') {
            lastMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
            messageUtil.warning(t('credentialsRequired'))
          } else {
            lastMessage.content = t(errInfo.i18nKey)
            messageUtil.error(t(errInfo.i18nKey))
          }
        }
      }
    } finally {
      loading.value = false
      abortController.value = null
    }
  }

  return { sendMessage, applyQuickAction, runAgentLoop, getOfficeSelection, currentAction, sessionStats, resetSessionStats }
}
