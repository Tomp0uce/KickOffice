import { nextTick, ref, type Ref } from 'vue'

import { type ChatMessage, type ChatRequestMessage, chatStream, generateImage } from '@/api/backend'
import { GLOBAL_STYLE_INSTRUCTIONS, buildInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt, getOutlookBuiltInPrompt, getPowerPointBuiltInPrompt, outlookBuiltInPrompt, powerPointBuiltInPrompt } from '@/utils/constant'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { message as messageUtil } from '@/utils/message'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointToolDefinitions } from '@/utils/powerpointTools'
import { prepareMessagesForContext } from '@/utils/tokenManager'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { getEnabledToolNamesFromStorage } from '@/utils/toolStorage'
import { extractTextFromHtml, reassembleWithFragments, getPreservationInstruction, type RichContentContext } from '@/utils/richContentPreserver'
import { applyInheritedStyles, renderOfficeCommonApiHtml } from '@/utils/officeRichText'
import { useAgentPrompts } from '@/composables/useAgentPrompts'
import { useOfficeSelection } from '@/composables/useOfficeSelection'

import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, OutlookQuickAction, QuickAction } from '@/types/chat'


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
  replyLanguage: Ref<string>
  agentMaxIterations: Ref<number>
  useSelectedText: Ref<boolean>
  excelFormulaLanguage: Ref<'en' | 'fr'>
  userGender: Ref<string>
  userFirstName: Ref<string>
  userLastName: Ref<string>
}

interface AgentLoopActions {
  quickActions: Ref<QuickAction[]>
  outlookQuickActions?: OutlookQuickAction[]
  excelQuickActions: Ref<ExcelQuickAction[]>
  powerPointQuickActions: PowerPointQuickAction[]
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
function isCredentialError(error: any): boolean {
  if (!error) return false
  const message = error.message || String(error)
  return (
    message.includes('401') ||
    message.includes('LiteLLM user credentials') ||
    message.includes('X-User-Key') ||
    message.includes('X-User-Email')
  )
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
    isWord: hostIsWord,
  } = host

  // Destructure settings
  const {
    customSystemPrompt,
    replyLanguage,
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
    hostIsWord,
  })

  function buildChatMessages(systemPrompt: string): ChatMessage[] {
    return [{ role: 'system', content: systemPrompt }, ...history.value.filter(m => m.role === 'user' || m.role === 'assistant').map(m => ({ role: m.role, content: m.content }))]
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
    const appToolDefs = hostIsOutlook ? getOutlookToolDefinitions() : hostIsPowerPoint ? getPowerPointToolDefinitions() : hostIsExcel ? getExcelToolDefinitions() : getWordToolDefinitions()
    const generalToolDefs = getGeneralToolDefinitions()
    const allToolDefs = [...generalToolDefs, ...appToolDefs]
    const enabledToolNames = getEnabledToolNamesFromStorage(allToolDefs.map(def => def.name))
    const enabledToolDefs = allToolDefs.filter(def => enabledToolNames.has(def.name))
    const tools = enabledToolDefs.map(def => ({ type: 'function' as const, function: { name: def.name, description: def.description, parameters: def.inputSchema as Record<string, unknown> } }))
    let iteration = 0
    const maxIter = Number(agentMaxIterations.value) || 10
    let currentMessages: ChatRequestMessage[] = [...messages]
    let lastToolSignature: string | null = null
    let toolsWereExecuted = false // Track if any tools were successfully executed
    currentAction.value = t('agentAnalyzing')
    history.value.push(createDisplayMessage('assistant', ''))
    await scrollToMessageTop() // Scroll to show start of assistant response
    const lastIndex = history.value.length - 1
    let abortedByUser = false
    while (iteration < maxIter) {
      if (abortController.value?.signal.aborted) {
        abortedByUser = true
        break
      }

      iteration++
      currentAction.value = t('agentAnalyzing')
      const currentSystemPrompt = messages[0]?.role === 'system' ? messages[0].content : ''
      const contextSafeMessages = prepareMessagesForContext(currentMessages, currentSystemPrompt)
      let response: StreamResponse = { choices: [{ message: { role: 'assistant', content: '', tool_calls: [] } }] }
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
            history.value[lastIndex].content = text
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
          }
        })
        response.choices[0].message.tool_calls = response.choices[0].message.tool_calls.filter(Boolean)
      } catch (err: any) {
        if (err.name === 'AbortError' || abortController.value?.signal.aborted) {
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
        // Display user-friendly message for credential errors
        if (isCredentialError(err)) {
          history.value[lastIndex].content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
        } else {
          history.value[lastIndex].content = `Error: The model or API failed to respond. ${err.message || ''}`
        }
        currentAction.value = ''
        break
      }
      const choice = response.choices?.[0]
      if (!choice) break
      const assistantMsg = choice.message
      currentMessages.push({
        role: 'assistant',
        content: assistantMsg.content || '',
        tool_calls: assistantMsg.tool_calls,
      })
      if (assistantMsg.content) history.value[lastIndex].content = assistantMsg.content
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

        const toolName = toolCall.function.name
        let toolArgs: Record<string, any> = {}
        try {
          toolArgs = JSON.parse(toolCall.function.arguments)
        } catch (parseErr) {
          console.error('[AgentLoop] Failed to parse tool call arguments', { toolName, arguments: toolCall.function.arguments, error: parseErr })
          toolResults.push({ tool_call_id: toolCall.id, content: `Error in ${toolName}: malformed tool arguments — JSON parse failed` })
          continue
        }
        let result = ''
        const toolDef = enabledToolDefs.find(tool => tool.name === toolName)
        if (toolDef) {
          const currentSignature = `${toolName}${JSON.stringify(toolArgs)}`
          if (currentSignature === lastToolSignature) {
            result = 'Error: You just executed this exact tool with the same arguments. It is a loop. Stop or change your arguments.'
          } else {
            currentAction.value = getActionLabelForCategory(toolDef.category)
            await scrollToBottom()
            try {
              result = await toolDef.execute(toolArgs)
              toolsWereExecuted = true // Mark that at least one tool was successfully executed
            } catch (err: any) {
              console.error('[AgentLoop] tool execution failed', { toolName, toolArgs, error: err })
              result = `Error in ${toolName}: ${err.message}`
            }
            currentAction.value = ''
            lastToolSignature = currentSignature
          }
        }

        // Check abort after tool execution
        if (abortController.value?.signal.aborted) {
          toolLoopAborted = true
          break
        }

        let safeContent = ''
        if (result === null || result === undefined) {
          safeContent = ''
        } else if (typeof result === 'object') {
          try {
            safeContent = JSON.stringify(result)
          } catch {
            safeContent = String(result)
          }
        } else {
          safeContent = String(result)
        }

        toolResults.push({ tool_call_id: toolCall.id, content: safeContent })
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

    if (abortedByUser) {
      currentAction.value = ''
      history.value.push(createDisplayMessage('system', t('agentStoppedByUser')))
      return
    }

    const assistantContent = history.value[lastIndex]?.content?.trim() || ''
    if (!assistantContent) {
      // If tools were executed successfully but no text response, that's OK (e.g., proofreading with comments)
      if (toolsWereExecuted) {
        history.value[lastIndex].content = t('toolsExecutedSuccessfully')
      } else {
        history.value[lastIndex].content = t('noModelResponse')
      }
    }

    if (iteration >= maxIter) messageUtil.warning(t('recursionLimitExceeded'))
    currentAction.value = ''
  }

  async function processChat(userMessage: string) {
    const modelConfig = availableModels.value[selectedModelTier.value]
    if (modelConfig?.type === 'image') {
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')))
      await scrollToMessageTop() // Scroll to top of assistant message
      imageLoading.value = true
      try {
        const imageSrc = await generateImage({ prompt: userMessage })
        const message = history.value[history.value.length - 1]
        message.role = 'assistant'; message.content = ''; message.imageSrc = imageSrc
      } catch (err: any) {
        const message = history.value[history.value.length - 1]
        message.role = 'assistant'; message.content = `${t('imageError')}: ${err.message}`; message.imageSrc = undefined
      } finally {
        imageLoading.value = false
      }
      await scrollToBottom() // Final scroll after image loads
      return
    }
    const systemPrompt = customSystemPrompt.value || agentPrompt(replyLanguage.value || 'Français')
    const messages = buildChatMessages(systemPrompt)
    const modelTier = resolveChatModelTier()

    await runAgentLoop(messages, modelTier)
  }

  async function sendMessage(payload?: unknown) {
    let textToSend = ''

    if (typeof payload === 'string') {
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

    if (!backendOnline.value) {
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
        return messageUtil.error(t('imageSelectionTooShort'))
      }
      isImageFromSelection = true
    }

    loading.value = true
    abortController.value = new AbortController()

    // If it's pure selection image, we show the selection as the user message bubble
    const displayMessageText = isImageFromSelection ? selectedText : userMessage
    history.value.push(createDisplayMessage('user', displayMessageText))
    await scrollToVeryBottom() // Scroll to very bottom after user message

    try {
      // Smart reply interception: when user sends after clicking "Reply" quick action
      if (pendingSmartReply.value && hostIsOutlook) {
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
        const lang = replyLanguage.value || 'Français'
        const replyPrompt = getOutlookBuiltInPrompt()['reply']
        const systemMsg = replyPrompt.system(lang) + `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`
        const userMsg = replyPrompt.user(emailBody, lang).replace('[REPLY_INTENT]', replyIntent)
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
            abortSignal: abortController.value?.signal,
          })
          const lastMessage = history.value[history.value.length - 1]
          if (!lastMessage?.content?.trim()) {
            lastMessage.content = t('noModelResponse')
          }
        } catch (err: any) {
          if (err.name === 'AbortError') return
          console.error('[AgentLoop] Smart reply chatStream failed', err)
          const lastMessage = history.value[history.value.length - 1]
          if (isCredentialError(err)) {
            lastMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
          } else {
            lastMessage.content = `Error: ${err.message || t('failedToResponse')}`
          }
        }
        return
      }

      // If we haven't fetched it yet and it's enabled
      if (useSelectedText.value && !isImageFromSelection) {
        let timeoutId: ReturnType<typeof setTimeout> | null = null
        try {
          const timeoutPromise = new Promise<string>((_, reject) => {
            timeoutId = setTimeout(() => reject(new Error('getOfficeSelection timeout')), 3000)
          })
          
          if (!hostIsExcel) {
            // F1: Extract formatted HTML natively and convert to markdown to preserve styling (Word, PPT, Outlook)
            const htmlPromise = new Promise<string>((_, reject) => {
              timeoutId = setTimeout(() => reject(new Error('getOfficeSelectionAsHtml timeout')), 3000)
            })
            
            try {
              const htmlContent = await Promise.race([getOfficeSelectionAsHtml({ includeOutlookSelectedText: true }), htmlPromise])
              if (htmlContent) {
                 const richContext = extractTextFromHtml(htmlContent)
                 selectedText = richContext.cleanText || selectedText
              } else {
                 selectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
              }
            } catch {
              selectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
            }
          } else {
            selectedText = await Promise.race([getOfficeSelection({ includeOutlookSelectedText: true }), timeoutPromise])
          }
        } catch (error) {
          console.warn('[AgentLoop] Failed to fetch selection before sending message', error)
        } finally {
          if (timeoutId) clearTimeout(timeoutId)
        }
      }

      let fullMessage = displayMessageText

      // Only append context to standard text chats, not pure image generations
      if (isImageFromSelection) {
        fullMessage = t('imageGenerationPrompt').replace('{text}', selectedText)
      } else if (selectedText && !isImageFromSelection) {
        const selectionLabel = hostIsOutlook ? 'Selected text' : hostIsPowerPoint ? 'Selected slide text' : hostIsExcel ? 'Selected cells' : 'Selected text'
        fullMessage = `${userMessage}

[${selectionLabel}: "${selectedText}"]`
        history.value[history.value.length - 1].content = fullMessage
      }

      await processChat(fullMessage)
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        console.error('[AgentLoop] sendMessage failed', error)
        if (isCredentialError(error)) {
          messageUtil.warning(t('credentialsRequired'))
        } else {
          messageUtil.error(t('failedToResponse'))
        }
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
        ? powerPointQuickActions.find(a => a.key === actionKey)
        : hostIsOutlook && outlookQuickActions
          ? outlookQuickActions.find(a => a.key === actionKey)
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
      const textForLlm = richContext ? richContext.cleanText : selectedText

      let action: { system: (lang: string) => string, user: (text: string, lang: string) => string } | undefined
      let systemMsg = ''
      let userMsg = ''
      if (hostIsOutlook) action = getOutlookBuiltInPrompt()[actionKey as keyof typeof outlookBuiltInPrompt]
      else if (hostIsPowerPoint) action = getPowerPointBuiltInPrompt()[actionKey as keyof typeof powerPointBuiltInPrompt]
      else if (hostIsExcel) {
        if (selectedExcelQuickAction?.mode === 'immediate' && selectedExcelQuickAction.systemPrompt) {
          systemMsg = selectedExcelQuickAction.systemPrompt
          userMsg = `Selection:\n${selectedText}`
        } else action = getExcelBuiltInPrompt()[actionKey as keyof typeof excelBuiltInPrompt]
      } else action = getBuiltInPrompt()[actionKey as keyof typeof buildInPrompt]
      if (!systemMsg || !userMsg) {
        if (!action) action = getBuiltInPrompt()[actionKey as keyof typeof buildInPrompt]
        if (!action) return
        const lang = replyLanguage.value || 'Français'
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
          if (isCredentialError(err)) {
            lastMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`
            messageUtil.warning(t('credentialsRequired'))
          } else {
            lastMessage.content = `Error: ${err.message || t('failedToResponse')}`
            messageUtil.error(t('failedToResponse'))
          }
        }
      }
    } finally {
      loading.value = false
      abortController.value = null
    }
  }

  return { sendMessage, applyQuickAction, runAgentLoop, getOfficeSelection, currentAction }
}
