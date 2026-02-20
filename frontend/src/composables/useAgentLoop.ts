import { nextTick, ref, type Ref } from 'vue'

import { type ChatMessage, type ChatRequestMessage, chatStream, chatSync, generateImage } from '@/api/backend'
import { GLOBAL_STYLE_INSTRUCTIONS, buildInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt, getOutlookBuiltInPrompt, getPowerPointBuiltInPrompt, outlookBuiltInPrompt, powerPointBuiltInPrompt } from '@/utils/constant'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { message as messageUtil } from '@/utils/message'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointToolDefinitions } from '@/utils/powerpointTools'
import { prepareMessagesForContext } from '@/utils/tokenManager'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { getEnabledToolNamesFromStorage } from '@/utils/toolStorage'
import { useAgentPrompts } from '@/composables/useAgentPrompts'
import { useOfficeSelection } from '@/composables/useOfficeSelection'

import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, OutlookQuickAction, QuickAction } from '@/types/chat'


interface EnabledToolsStorageState {
  version: number
  signature: string
  enabledToolNames: string[]
}

interface UseAgentLoopOptions {
  t: (key: string) => string
  history: Ref<DisplayMessage[]>
  userInput: Ref<string>
  loading: Ref<boolean>
  imageLoading: Ref<boolean>
  backendOnline: Ref<boolean>
  availableModels: Ref<Record<string, ModelInfo>>
  selectedModelTier: Ref<ModelTier>
  selectedModelInfo: Ref<ModelInfo | undefined>
  firstChatModelTier: Ref<ModelTier>
  customSystemPrompt: Ref<string>
  replyLanguage: Ref<string>
  agentMaxIterations: Ref<number>
  useSelectedText: Ref<boolean>
  excelFormulaLanguage: Ref<'en' | 'fr'>
  userGender: Ref<string>
  userFirstName: Ref<string>
  userLastName: Ref<string>
  abortController: Ref<AbortController | null>
  inputTextarea: Ref<HTMLTextAreaElement | undefined>
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
  hostIsWord: boolean
  quickActions: Ref<QuickAction[]>
  outlookQuickActions?: Ref<OutlookQuickAction[]>
  excelQuickActions: Ref<ExcelQuickAction[]>
  powerPointQuickActions: PowerPointQuickAction[]
  draftFocusGlow: Ref<boolean>
  createDisplayMessage: (role: DisplayMessage['role'], content: string, imageSrc?: string) => DisplayMessage
  adjustTextareaHeight: () => void
  scrollToBottom: () => Promise<void>
}

export function useAgentLoop(options: UseAgentLoopOptions) {
  const {
    t,
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    availableModels,
    selectedModelTier,
    selectedModelInfo,
    firstChatModelTier,
    customSystemPrompt,
    replyLanguage,
    agentMaxIterations,
    useSelectedText,
    excelFormulaLanguage,
    userGender,
    userFirstName,
    userLastName,
    abortController,
    inputTextarea,
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
    hostIsWord,
    quickActions,
    outlookQuickActions,
    excelQuickActions,
    powerPointQuickActions,
    draftFocusGlow,
    createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
  } = options

  const currentAction = ref('')

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

  const { getOfficeSelection } = useOfficeSelection({
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
    const maxIter = Math.min(Number(agentMaxIterations.value) || 10, 10)
    let currentMessages: ChatRequestMessage[] = [...messages]
    let lastToolSignature: string | null = null
    currentAction.value = t('agentAnalyzing')
    history.value.push(createDisplayMessage('assistant', ''))
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
      let response: any = { choices: [{ message: { role: 'assistant', content: '', tool_calls: [] } }] }
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
        history.value[lastIndex].content = `Error: The model or API failed to respond. ${err.message || ''}`
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
      for (const toolCall of assistantMsg.tool_calls) {
        const toolName = toolCall.function.name
        let toolArgs: Record<string, any> = {}
        try { toolArgs = JSON.parse(toolCall.function.arguments) } catch {}
        let result = ''
        const toolDef = enabledToolDefs.find(tool => tool.name === toolName)
        if (toolDef) {
          const currentSignature = `${toolName}${JSON.stringify(toolArgs)}`
          if (currentSignature === lastToolSignature) {
            result = 'Error: You just executed this exact tool with the same arguments. It is a loop. Stop or change your arguments.'
          } else {
            currentAction.value = getActionLabelForCategory(toolDef.category)
            await scrollToBottom()
            try { result = await toolDef.execute(toolArgs) } catch (err: any) { console.error('[AgentLoop] tool execution failed', { toolName, toolArgs, error: err }); result = `Error: ${err.message}` }
            currentAction.value = ''
            lastToolSignature = currentSignature
          }
        }
        if (abortController.value?.signal.aborted) {
          abortedByUser = true
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

        currentMessages.push({ role: 'tool', tool_call_id: toolCall.id, content: safeContent })
      }
      if (abortedByUser) {
        break
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
      history.value[lastIndex].content = t('noModelResponse')
    }

    if (iteration >= maxIter) messageUtil.warning(t('recursionLimitExceeded'))
    currentAction.value = ''
  }

  async function processChat(userMessage: string) {
    const modelConfig = availableModels.value[selectedModelTier.value]
    if (modelConfig?.type === 'image') {
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')))
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
      await scrollToBottom()
      return
    }
    const systemPrompt = customSystemPrompt.value || agentPrompt(replyLanguage.value || 'Français')
    const messages = buildChatMessages(systemPrompt)
    const modelTier = resolveChatModelTier()

    try {
      await runAgentLoop(messages, modelTier)
    } catch (error) {
      throw error
    }
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
    await scrollToBottom()

    try {
      // If we haven't fetched it yet and it's enabled
      if (useSelectedText.value && !isImageFromSelection) {
        let timeoutId: ReturnType<typeof setTimeout> | null = null
        try {
          const timeoutPromise = new Promise<string>((_, reject) => {
            timeoutId = setTimeout(() => reject(new Error('getOfficeSelection timeout')), 3000)
          })
          selectedText = await Promise.race([getOfficeSelection(), timeoutPromise])
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
        const selectionLabel = hostIsOutlook ? 'Email body' : hostIsPowerPoint ? 'Selected slide text' : hostIsExcel ? 'Selected cells' : 'Selected text'
        fullMessage = `${userMessage}

[${selectionLabel}: "${selectedText}"]`
        history.value[history.value.length - 1].content = fullMessage
      }

      await processChat(fullMessage)
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        console.error('[AgentLoop] sendMessage failed', error)
        messageUtil.error(t('failedToResponse'))
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
        : hostIsOutlook && outlookQuickActions?.value
          ? outlookQuickActions.value.find(a => a.key === actionKey)
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
    const selectedText = await getOfficeSelection({ includeOutlookSelectedText: true })
    if (!selectedText) return messageUtil.error(t(hostIsOutlook ? 'selectEmailPrompt' : hostIsPowerPoint ? 'selectSlideTextPrompt' : hostIsExcel ? 'selectCellsPrompt' : 'selectTextPrompt'))

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
      if (!action) return
      const lang = replyLanguage.value || 'Français'
      systemMsg = action.system(lang)
      userMsg = action.user(selectedText, lang)
    }
    
    // Enforce global formatting constraints on all Quick Actions
    systemMsg += `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`

    const actionLabel = selectedQuickAction?.label || t(actionKey)
    history.value.push(createDisplayMessage('user', `[${actionLabel}] ${selectedText.substring(0, 100)}...`))

    if (selectedQuickAction?.executeWithAgent) {
      await runAgentLoop([{ role: 'system', content: systemMsg }, { role: 'user', content: userMsg }], resolveChatModelTier())
    } else {
      history.value.push(createDisplayMessage('assistant', ''))
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
    }
  }

  return { sendMessage, applyQuickAction, runAgentLoop, getOfficeSelection, currentAction }
}
