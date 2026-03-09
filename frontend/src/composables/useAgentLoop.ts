import type { ModelTier, ModelInfo, ToolCategory } from '@/types'
import { nextTick, ref, type Ref, type ComputedRef } from 'vue'

import { type ChatMessage, type ChatRequestMessage, type TokenUsage, chatStream, generateImage, uploadFile, uploadFileToPlatform, categorizeError } from '@/api/backend'
import { GLOBAL_STYLE_INSTRUCTIONS, builtInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt, getOutlookBuiltInPrompt, getPowerPointBuiltInPrompt, outlookBuiltInPrompt, powerPointBuiltInPrompt } from '@/utils/constant'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { message as messageUtil } from '@/utils/message'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointToolDefinitions, setCurrentSlideSpeakerNotes, powerpointImageRegistry } from '@/utils/powerpointTools'
import { prepareMessagesForContext } from '@/utils/tokenManager'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { getEnabledToolNamesFromStorage } from '@/utils/toolStorage'
import { extractTextFromHtml, reassembleWithFragments, getPreservationInstruction, type RichContentContext } from '@/utils/richContentPreserver'
import { applyInheritedStyles, renderOfficeCommonApiHtml } from '@/utils/markdown'
import { useAgentPrompts } from '@/composables/useAgentPrompts'
import { useOfficeSelection } from '@/composables/useOfficeSelection'
import { setLastRichContext, clearLastRichContext, getLastRichContext } from '@/utils/richContextStore'
import {
  getExcelDocumentContext,
  getPowerPointDocumentContext,
  getOutlookDocumentContext,
  getWordDocumentContext,
} from '@/utils/officeDocumentContext'
import { areCredentialsConfigured } from '@/utils/credentialStorage'
import { logService } from '@/utils/logger'

import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, OutlookQuickAction, QuickAction } from '@/types/chat'

import { useAgentStream } from './useAgentStream'
import { executeAgentToolCall } from './useToolExecutor'
import { useLoopDetection } from './useLoopDetection'
interface AgentLoopRefs {
  history: Ref<DisplayMessage[]>
  userInput: Ref<string>
  loading: Ref<boolean>
  imageLoading: Ref<boolean>
  backendOnline: Ref<boolean>
  abortController: Ref<AbortController | null>
  inputTextarea: Ref<HTMLTextAreaElement | undefined>
  isDraftFocusGlowing: Ref<boolean>
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
  excelFormulaLanguage: Ref<'en' | 'fr'>
  userGender: Ref<string>
  userFirstName: Ref<string>
  userLastName: Ref<string>
}

interface AgentLoopActions {
  quickActions: ComputedRef<QuickAction[] | undefined>
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


export function useAgentLoop(options: UseAgentLoopOptions) {
  const { t, refs, models, host, settings, actions, helpers } = options

  const { executeStream } = useAgentStream()
  const { addSignatureAndCheckLoop, clearSignatures } = useLoopDetection(5, 2)

  // Step 4: Persistent session memory for uploaded content (Point 2)
  const sessionUploadedFiles = ref<{ filename: string; content: string; fileId?: string }[]>([])
  const sessionUploadedImages = ref<{ filename: string; dataUri: string; imageId?: string }[]>([])


  // Destructure refs
  const {
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    abortController,
    inputTextarea,
    isDraftFocusGlowing,
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
      let contentToKeep = m.content;
      // If the assistant message only had tool calls and no content, ensure it's not totally empty
      if (m.role === 'assistant' && !contentToKeep?.trim() && m.rawMessages && m.rawMessages.length > 0) {
        contentToKeep = `[Tools executed internally]`;
      }
      msgs.push({ role: m.role, content: contentToKeep || '' })
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

    // Add preservation instruction to system prompt if we have rich content
    const richContext = getLastRichContext()
    if (richContext?.hasRichContent && messages[0]?.role === 'system') {
      messages[0].content += getPreservationInstruction(richContext)
    }

    let iteration = 0
    const maxIter = Number(agentMaxIterations.value) || 10
    const startTime = Date.now()
    const timeoutMs = maxIter * 60 * 1000 // up to 1 minute per iteration allowed
    let currentMessages: ChatRequestMessage[] = [...messages]
    // Sliding window loop detection (P6) uses useLoopDetection composable
    clearSignatures()
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

      // Enforce max iterations limit
      if (iteration > maxIter) {
        currentAssistantMessage.content += `\n\n⚠️ ${t('agentMaxIterationsReached')}`
        break
      }
      
      // H11: Show "agentAnalyzing" initially, or "agentWaitingForLLM" if tools were just executed and we are generating a response
      const llmWaitLabel = iteration === 1 ? t('agentAnalyzing') : t('agentWaitingForLLM')
      currentAction.value = llmWaitLabel
      const llmWaitStart = Date.now()
      const llmWaitTimer = setInterval(() => {
        const elapsed = Math.round((Date.now() - llmWaitStart) / 1000)
        currentAction.value = `${llmWaitLabel} (${elapsed}s)`
      }, 1000)

      const currentSystemPrompt = messages[0]?.role === 'system' ? (typeof messages[0].content === 'string' ? messages[0].content : '') : ''
      const contextSafeMessages = prepareMessagesForContext(currentMessages, currentSystemPrompt)
      logService.info('llm_request', 'llm', { model: modelTier, messageCount: contextSafeMessages.length })

      let response: any
      let truncatedByLength = false

      try {
        const streamResult = await executeStream({
          messages: contextSafeMessages,
          modelTier,
          tools,
          abortSignal: abortController.value?.signal || undefined,
          currentAction,
          currentAssistantMessage,
          scrollToBottom,
          accumulateUsage
        })
        
        clearInterval(llmWaitTimer)
        response = streamResult.response
        truncatedByLength = streamResult.truncatedByLength
        logService.info('llm_response_complete', 'llm', { tokensUsed: sessionStats.value.totalTokens })
      } catch (err: unknown) {
        clearInterval(llmWaitTimer)
        if ((err instanceof Error && err.name === 'AbortError') || abortController.value?.signal.aborted) {
          abortedByUser = true
          break
        }
        logService.error('[AgentLoop] chatStream failed', err, {
          host: hostIsOutlook ? 'outlook' : hostIsPowerPoint ? 'powerpoint' : hostIsExcel ? 'excel' : 'word',
          modelTier,
          iteration,
          messageCount: currentMessages.length,
          traffic: 'system'
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

        // Sliding window loop detection (P6) — same signature repeated
        if (sig && addSignatureAndCheckLoop(sig)) {
          toolResults.push({ tool_call_id: toolCall.id, content: 'Error: You have called this exact tool with the same arguments multiple times in a row. This is a loop. Stop repeating and try a different approach.' })
          continue
        }


        if (toolResult.success) toolsWereExecuted = true
        if (toolResult.screenshotBase64) {
          const mimeType = toolResult.screenshotMimeType || 'image/png'
          const dataUri = `data:${mimeType};base64,${toolResult.screenshotBase64}`
          const filename = `screenshot_${Date.now()}.png`
          sessionUploadedImages.value.push({ filename, dataUri })
        }
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
      
      // H11: Switch status from tool execution to waiting for LLM response
      currentAction.value = t('agentWaitingForLLM')
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

    await nextTick()
    await scrollToVeryBottom()
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
    // M2: Centralized LocalStorage Language Preference with validation
    const storedLang = localStorage.getItem('localLanguage')
    const validLangs = ['en', 'fr']
    const langKey = validLangs.includes(storedLang || '') ? storedLang : 'fr' // Default to fr safely
    const lang = langKey === 'en' ? 'English' : 'Français'
    
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
        onStream: (text: string) => {
          const message = history.value[history.value.length - 1]
          message.role = 'assistant'
          message.content = text
          // No auto-scroll during streaming: user can freely scroll.
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
             // Store rich context globally so tools can access it (especially for Outlook image preservation)
             if (richContext.hasRichContent) {
               setLastRichContext(richContext)
             }
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

  async function processChat(
    userMessage: string,
    visionImages?: Array<{ filename: string; dataUri: string; imageId?: string }>,
    injectedContext?: string,
    selectionContext?: string,
    uploadedFiles?: Array<{ filename: string; content: string; fileId?: string }>,
  ) {
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
        const errInfo = categorizeError(err)
        const baseMsg = t(errInfo.i18nKey)
        const detail = err instanceof Error ? err.message : String(err)
        message.role = 'assistant'; message.content = `${baseMsg}\n\n${detail}`; message.imageSrc = undefined
      } finally {
        imageLoading.value = false
      }
      await scrollToBottom() // Final scroll after image loads
      return
    }

    // M2: Centralized LocalStorage Language Preference with validation
    const storedLang = localStorage.getItem('localLanguage')
    const langKey = ['en', 'fr'].includes(storedLang || '') ? storedLang : 'fr'
    const lang = langKey === 'en' ? 'English' : 'Français'
    const systemPrompt = customSystemPrompt.value || agentPrompt(lang)
    
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
      if (injectedContext) {
        // Legacy string-based injection (kept for any call sites that still pass it)
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
        if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
          messages[lastUserIdx].content += `\n\n<attached_files>\n${injectedContext}\n</attached_files>`
        }
      }
      if (uploadedFiles && uploadedFiles.length > 0) {
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
        if (lastUserIdx !== -1) {
          const hasFileRefs = uploadedFiles.some(f => f.fileId)
          if (hasFileRefs && typeof messages[lastUserIdx].content === 'string') {
            // Convert to content array: text + file references + inline fallback for files without fileId
            const parts: any[] = [{ type: 'text', text: messages[lastUserIdx].content as string }]
            for (const f of uploadedFiles) {
              if (f.fileId) {
                parts.push({ type: 'file', file: { file_id: f.fileId } })
              } else {
                parts.push({ type: 'text', text: `\n\n[Contenu du fichier "${f.filename}"]:\n${f.content}\n[Fin du fichier]` })
              }
            }
            messages[lastUserIdx].content = parts
          } else if (typeof messages[lastUserIdx].content === 'string') {
            // All inline fallback
            const inlineText = uploadedFiles
              .map(f => `\n\n[Contenu du fichier "${f.filename}"]:\n${f.content}\n[Fin du fichier]`)
              .join('')
            messages[lastUserIdx].content += `\n\n<attached_files>${inlineText}\n</attached_files>`
          }
        }
      }
      if (selectionContext) {
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
        if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
          const selectionLabel = hostIsOutlook ? 'Selected text' : hostIsPowerPoint ? 'Selected slide text' : hostIsExcel ? 'Selected cells' : 'Selected text'
          const sanitizedSelection = selectionContext.replace(new RegExp('</?document_content>', 'g'), '')
          messages[lastUserIdx].content += `\n\nHere is the current context from the user's document (${selectionLabel}). IMPORTANT: First evaluate if this context is relevant to the user's query. If it is not relevant, ignore it completely and answer the query normally.\n\n<document_content>\n${sanitizedSelection}\n</document_content>`
        }
      }
    } catch (ctxErr) {
      console.warn('[AgentLoop] Failed to fetch document context', ctxErr)
    }

    // Inject vision images as multipart content into the last user message
    // Point 2 Fix: Use ALL session images for vision injection (Session Persistence)
    if ((visionImages && visionImages.length > 0) || sessionUploadedImages.value.length > 0) {
      const lastUserIdx = messages.map(m => m.role).lastIndexOf('user')
      if (lastUserIdx !== -1) {
        let textContent = messages[lastUserIdx].content || userMessage
        const imageContextLines: string[] = []
        for (const img of sessionUploadedImages.value) {
          const idTag = img.imageId ? ` (imageId: ${img.imageId})` : ''
          imageContextLines.push(`- [${img.filename}]${idTag}`)
        }
        if (imageContextLines.length > 0) {
          textContent += `\n\n<uploaded_images>\nThe following images are available in session memory:\n${imageContextLines.join('\n')}\nTo embed an image in a slide, use insertImageOnSlide with the filename. To extract chart data into Excel, use extract_chart_data with the imageId.\n</uploaded_images>`
        }
        
        const parts: any[] = [{ type: 'text', text: String(textContent) }]
        for (const img of sessionUploadedImages.value) {
          parts.push({ type: 'image_url', image_url: { url: img.dataUri } })
        }
        ;(messages[lastUserIdx] as any).content = parts
      }
    }

    return await runAgentLoop(messages, modelTier)
  }

  async function sendMessage(payload?: string, files?: File[]) {
    // Clear any previous rich context at the start of a new request
    clearLastRichContext()

    let textToSend = ''

    if (payload) {
      textToSend = payload
    } else if (userInput.value && typeof userInput.value === 'string') {
      textToSend = userInput.value
    }

    textToSend = textToSend?.trim() || ''

    if (!textToSend && (!files || files.length === 0)) {
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

    // BUGFIX: Validate credentials are configured before sending request
    const hasCredentials = await areCredentialsConfigured()
    if (!hasCredentials) {
      loading.value = false
      messageUtil.error(t('credentialsRequired'))
      return
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
      let wordCount = selectedText.trim().split(/\s+/).filter(w => w.length > 0).length
      
      if (wordCount < 5 && hostIsPowerPoint) {
        try {
          const { executeOfficeAction } = await import('@/utils/officeAction')
          selectedText = await executeOfficeAction(() => {
            const PPT = (window as any).PowerPoint
            if (!PPT) return Promise.resolve('')
            return PPT.run(async (context: any) => {
              let activeSlideIndex = 0
              try {
                if (typeof context.presentation.getSelectedSlides === 'function') {
                  const selectedSlides = context.presentation.getSelectedSlides()
                  selectedSlides.load('items/id')
                  await context.sync()
                  if (selectedSlides.items.length > 0) {
                    const slides = context.presentation.slides
                    slides.load('items/id')
                    await context.sync()
                    const selectedId = selectedSlides.items[0].id
                    const idx = slides.items.findIndex((s: any) => s.id === selectedId)
                    if (idx !== -1) activeSlideIndex = idx
                  }
                }
              } catch (e) {}

              const slides = context.presentation.slides
              slides.load('items')
              await context.sync()
              if (activeSlideIndex >= slides.items.length) return ''
              const slide = slides.items[activeSlideIndex]

              const shapes = slide.shapes
              shapes.load('items')
              await context.sync()

              for (const shape of shapes.items) {
                try { shape.textFrame.textRange.load('text') } catch {}
              }
              await context.sync()

              const texts = []
              for (const shape of shapes.items) {
                try { texts.push((shape.textFrame.textRange.text || '').trim()) } catch {}
              }
              return texts.filter(Boolean).join('\n')
            })
          })
          wordCount = selectedText.trim().split(/\s+/).filter(w => w.length > 0).length
        } catch (e) {
          console.warn('[AgentLoop] Fallback to PowerPoint slide content failed', e)
        }
      }

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
    const userMsgIdx = history.value.length - 1
    await scrollToVeryBottom() // Scroll to very bottom after user message

    try {
      // Smart reply interception: when user sends after clicking "Reply" quick action
      if (pendingSmartReply.value && hostIsOutlook) {
        await handleSmartReply(userMessage)
        return
      }

      // GEN-L3: Always fetch selected text as Phantom Context (if not already purely image generation)
      if (!isImageFromSelection) {
        selectedText = await fetchSelectionWithTimeout()
      }

      let fullMessage = displayMessageText

      if (files && files.length > 0) {
        currentAction.value = t('agentUploadingFiles') || 'Extraction des fichiers...'
        try {
           const newTextFiles: Array<{ filename: string; content: string; fileId?: string }> = []
           for (const file of files) {
             const result = await uploadFile(file)
             if (result.imageBase64) {
               // Step 4: Store in session images (with imageId for chart extraction)
               sessionUploadedImages.value.push({ filename: result.filename, dataUri: result.imageBase64, imageId: result.imageId })

               // Point 3 Fix: Store in PPT registry for tool access (by filename AND imageId)
               if (hostIsPowerPoint) {
                 const rawBase64 = result.imageBase64.replace(/^data:[^;]+;base64,/, '')
                 powerpointImageRegistry.set(result.filename, rawBase64)
                 if (result.imageId) powerpointImageRegistry.set(result.imageId, rawBase64)
               }
               // Show a preview thumbnail in the user message bubble
               history.value[userMsgIdx].imageSrc = result.imageBase64
             } else {
               // Store extracted text in persistent session memory
               const entry: { filename: string; content: string; fileId?: string } = {
                 filename: result.filename,
                 content: result.extractedText,
               }
               // Tâche 4: Try to upload to LLM provider for file_id referencing (best-effort)
               try {
                 const platformResult = await uploadFileToPlatform(file)
                 if (platformResult.fileId) {
                   entry.fileId = platformResult.fileId
                 }
               } catch {
                 // Provider doesn't support /v1/files or network error — fall back to inline content
                 logService.warn('[AgentLoop] /v1/files upload failed — using inline content fallback', { filename: file.name })
               }
               sessionUploadedFiles.value.push(entry)
               newTextFiles.push({ filename: result.filename, content: result.extractedText, fileId: entry.fileId })
             }
           }
           // Persist file info on the user message for session restore (Tâche 6)
           if (newTextFiles.length > 0) {
             history.value[userMsgIdx].attachedFiles = newTextFiles
           }
        } catch (uploadObjErr: unknown) {
           console.error('[AgentLoop] File upload failed', uploadObjErr)
           return messageUtil.error(t('somethingWentWrong'))
        }
      }

      // Step 4: Pass session uploaded files to processChat (inline or file_id reference)
      const uploadedFilesForChat = sessionUploadedFiles.value.length > 0 ? [...sessionUploadedFiles.value] : undefined

      // Only append context to standard text chats, not pure image generations
      // selectedText is passed separately to processChat so it never pollutes the UI history
      if (isImageFromSelection) {
        if (hostIsPowerPoint) {
          fullMessage = t('pptVisualPrefix') + '\n' + selectedText
        } else {
          fullMessage = t('imageGenerationPrompt').replace('{text}', selectedText)
        }
        await processChat(fullMessage.trim(), undefined, undefined, undefined, uploadedFilesForChat)
      } else {
        // Pass selectedText as selectionContext: injected into LLM payload only, not shown in UI
        await processChat(fullMessage.trim(), undefined, undefined, selectedText || undefined, uploadedFilesForChat)
      }
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

    // BUGFIX: Validate credentials are configured before sending request
    const hasCredentials = await areCredentialsConfigured()
    if (!hasCredentials) {
      messageUtil.error(t('credentialsRequired'))
      return
    }

    // Prevent quick actions from running while another request is in progress
    if (loading.value || abortController.value) {
      return messageUtil.warning(t('requestInProgress') || 'A request is already in progress. Please wait or stop the current request.')
    }
    const selectedQuickAction = hostIsExcel
      ? excelQuickActions.value.find((a: ExcelQuickAction) => a.key === actionKey)
      : hostIsPowerPoint
        ? powerPointQuickActions.value.find((a: PowerPointQuickAction) => a.key === actionKey)
        : hostIsOutlook && outlookQuickActions?.value
          ? outlookQuickActions.value.find((a: OutlookQuickAction) => a.key === actionKey)
          : quickActions.value?.find((a: QuickAction) => a.key === actionKey)

    const selectedExcelQuickAction = hostIsExcel ? selectedQuickAction as ExcelQuickAction | undefined : undefined
    const selectedPowerPointQuickAction = hostIsPowerPoint ? selectedQuickAction as PowerPointQuickAction | undefined : undefined
    const selectedOutlookQuickAction = hostIsOutlook ? selectedQuickAction as OutlookQuickAction | undefined : undefined

    if (actionKey === 'visual' && hostIsPowerPoint) {
      const imageModelTier = Object.entries(availableModels.value).find(([_, info]) => info.type === 'image')?.[0] as ModelTier
      if (!imageModelTier) {
        return messageUtil.error(t('imageError') || 'No image model configured.')
      }

      // Get current slide text selection
      const slideText = await getOfficeSelection({ actionKey })

      // Step 1: call standard LLM to generate a proper image description prompt
      const lang = localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français'
      const visualPrompt = getPowerPointBuiltInPrompt().visual
      const systemMsg = visualPrompt.system(lang)
      const userMsg = visualPrompt.user(slideText || '', lang)

      const actionLabel = selectedQuickAction?.label || t(actionKey)
      history.value.push(createDisplayMessage('user', `[${actionLabel}] ${(slideText || '').substring(0, 100)}...`))
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')))
      await scrollToMessageTop()

      loading.value = true
      abortController.value = new AbortController()
      try {
        let imagePrompt = ''
        await chatStream({
          messages: [{ role: 'system', content: systemMsg }, { role: 'user', content: userMsg }],
          modelTier: resolveChatModelTier(),
          onStream: async (text: string) => { imagePrompt = text },
          abortSignal: abortController.value?.signal,
        })

        if (!imagePrompt.trim()) {
          history.value[history.value.length - 1].content = t('somethingWentWrong')
          return
        }

        // Step 2: use the generated description to produce the image
        history.value[history.value.length - 1].content = t('imageGenerating')
        imageLoading.value = true
        const imageSrc = await generateImage({ prompt: imagePrompt.trim() })
        const message = history.value[history.value.length - 1]
        message.role = 'assistant'; message.content = ''; message.imageSrc = imageSrc
        await scrollToBottom()
      } catch (err: unknown) {
        if (!(err instanceof Error) || err.name !== 'AbortError') {
          console.error('[AgentLoop] visual quick action failed', err)
          const errInfo = categorizeError(err)
          history.value[history.value.length - 1].content = t(errInfo.i18nKey)
        }
      } finally {
        imageLoading.value = false
        loading.value = false
        abortController.value = null
      }
      return
    }

    if (selectedOutlookQuickAction?.mode === 'smart-reply') {
      pendingSmartReply.value = true
      userInput.value = selectedOutlookQuickAction.prefix || ''
      adjustTextareaHeight()
      isDraftFocusGlowing.value = true
      setTimeout(() => { isDraftFocusGlowing.value = false; }, 1500)
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
      isDraftFocusGlowing.value = true
      setTimeout(() => { isDraftFocusGlowing.value = false; }, 1500)
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
      isDraftFocusGlowing.value = true
      setTimeout(() => { isDraftFocusGlowing.value = false; }, 1000)
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
        if (selectedPowerPointQuickAction?.systemPrompt) {
          systemMsg = selectedPowerPointQuickAction.systemPrompt
          userMsg = selectedText || t('applyToCurrentSlide') || 'Apply to the current slide.'
        } else {
          action = getPowerPointBuiltInPrompt()[actionKey as keyof typeof powerPointBuiltInPrompt] || getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt]
        }
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

          // PPT-M2: For speakerNotes action, directly insert the generated text into the slide notes
          if (hostIsPowerPoint && actionKey === 'speakerNotes' && lastMessage?.content?.trim()) {
            const inserted = await setCurrentSlideSpeakerNotes(lastMessage.content.trim())
            if (inserted) {
              lastMessage.content += '\n\n_✓ Notes insérées dans la diapositive._'
            }
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

  /**
   * Rebuilds sessionUploadedFiles from history after a session switch or restore.
   * Call this whenever history is replaced from IndexedDB.
   */
  function rebuildSessionFiles() {
    const seen = new Set<string>()
    sessionUploadedFiles.value = []
    for (const msg of history.value) {
      if (msg.attachedFiles) {
        for (const f of msg.attachedFiles) {
          if (!seen.has(f.filename)) {
            seen.add(f.filename)
            sessionUploadedFiles.value.push(f)
          }
        }
      }
    }
  }

  return { sendMessage, applyQuickAction, runAgentLoop, getOfficeSelection, currentAction, sessionStats, resetSessionStats, rebuildSessionFiles }
}
