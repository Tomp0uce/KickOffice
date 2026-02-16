import { nextTick, type Ref } from 'vue'

import { type ChatMessage, type ChatRequestMessage, chatStream, chatSync, generateImage } from '@/api/backend'
import { buildInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt, getOutlookBuiltInPrompt, getPowerPointBuiltInPrompt, outlookBuiltInPrompt, powerPointBuiltInPrompt } from '@/utils/constant'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { message as messageUtil } from '@/utils/message'
import { getOfficeTextCoercionType, getOutlookMailbox, isOfficeAsyncSucceeded, type OfficeAsyncResult } from '@/utils/officeOutlook'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointSelection, getPowerPointToolDefinitions } from '@/utils/powerpointTools'
import { getWordToolDefinitions } from '@/utils/wordTools'

import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, QuickAction } from '@/types/chat'

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
  quickActions: Ref<QuickAction[]>
  excelQuickActions: Ref<ExcelQuickAction[]>
  powerPointQuickActions: PowerPointQuickAction[]
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
    quickActions,
    excelQuickActions,
    powerPointQuickActions,
    createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
  } = options

  const excelFormulaLanguageInstruction = () => excelFormulaLanguage.value === 'fr'
    ? 'Excel interface locale: French. Use localized French function names and separators when providing formulas, and prefer localized formula tool behavior.'
    : 'Excel interface locale: English. Use English function names and standard English formula syntax.'

  const userProfilePromptBlock = () => {
    const firstName = userFirstName.value.trim()
    const lastName = userLastName.value.trim()
    const fullName = `${firstName} ${lastName}`.trim() || t('userProfileUnknownName')
    const genderMap: Record<string, string> = {
      female: t('userGenderFemale'), male: t('userGenderMale'), nonbinary: t('userGenderNonBinary'), unspecified: t('userGenderUnspecified'),
    }
    const genderLabel = genderMap[userGender.value] || t('userGenderUnspecified')
    return `\n\nUser profile context for communications (especially emails):\n- First name: ${firstName || t('userProfileUnknownFirstName')}\n- Last name: ${lastName || t('userProfileUnknownLastName')}\n- Full name: ${fullName}\n- Gender: ${genderLabel}\nUse this profile when drafting salutations, signatures, and tone, unless the user asks otherwise.`
  }

  const wordAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Word Expert Agent. Your goal is to assist users in creating, editing, and formatting documents with professional precision.\n\n# Capabilities\n- You can interact with the document directly using provided tools (reading text, applying styles, inserting content, etc.).\n- You understand document structure, typography, and professional writing standards.\n\n# Guidelines\n1. **Tool First**: If a request requires document modification or inspection, prioritize using the available tools.\n2. **Direct Actions**: For Word formatting requests (bold, underline, highlight, size, color, superscript, uppercase, tags like <format>...</format>, etc.), execute the change directly with tools instead of giving manual steps.\n3. **Accuracy**: Ensure formatting and content changes are precise and follow the user's intent.\n4. **Conciseness**: Provide brief, helpful explanations of your actions.\n5. **Language**: You must communicate entirely in ${lang}.\n\n# Safety\nDo not perform destructive actions (like clearing the whole document) unless explicitly instructed.`
  const excelAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Excel Expert Agent. Your goal is to assist users with data analysis, formulas, charts, formatting, and spreadsheet operations with professional precision.\n\n# Guidelines\n1. **Tool First**\n2. **Read First**\n3. **Accuracy**\n4. **Conciseness**\n5. **Language**: You must communicate entirely in ${lang}.\n6. **Formula locale**: ${excelFormulaLanguageInstruction()}\n7. **Formula duplication**: use fillFormulaDown when applying same formula across rows.`
  const powerPointAgentPrompt = (lang: string) => `# Role\nYou are a PowerPoint presentation expert.\n# Guidelines\n5. **Language**: You must communicate entirely in ${lang}.`
  const outlookAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Outlook Email Expert Agent.\n# Guidelines\n4. **Language**: You must communicate entirely in ${lang}.`

  const agentPrompt = (lang: string) => {
    let base = hostIsOutlook ? outlookAgentPrompt(lang) : hostIsPowerPoint ? powerPointAgentPrompt(lang) : hostIsExcel ? excelAgentPrompt(lang) : wordAgentPrompt(lang)
    return `${base}${userProfilePromptBlock()}`
  }

  function buildChatMessages(systemPrompt: string): ChatMessage[] {
    return [{ role: 'system', content: systemPrompt }, ...history.value.filter(m => m.role === 'user' || m.role === 'assistant').map(m => ({ role: m.role, content: m.content }))]
  }

  const getOutlookMailBody = (): Promise<string> => new Promise((resolve) => {
    try {
      const mailbox = getOutlookMailbox()
      if (!mailbox?.item) return resolve('')
      mailbox.item.body.getAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<string>) => resolve(isOfficeAsyncSucceeded(result.status) ? (result.value || '') : ''))
    } catch { resolve('') }
  })

  const getOutlookSelectedText = (): Promise<string> => new Promise((resolve) => {
    try {
      const mailbox = getOutlookMailbox()
      if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function') return resolve('')
      mailbox.item.getSelectedDataAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<{ data?: string }>) => resolve(isOfficeAsyncSucceeded(result.status) && result.value?.data ? result.value.data : ''))
    } catch { resolve('') }
  })

  async function getOfficeSelection(options?: { includeOutlookSelectedText?: boolean }): Promise<string> {
    if (hostIsOutlook) {
      if (options?.includeOutlookSelectedText) {
        const selected = await getOutlookSelectedText()
        if (selected) return selected
      }
      return getOutlookMailBody()
    }
    if (hostIsPowerPoint) return getPowerPointSelection()
    if (hostIsExcel) {
      return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange()
        range.load('values, address')
        await ctx.sync()
        return `[${range.address}]\n${range.values.map((row: any[]) => row.join('\t')).join('\n')}`
      })
    }
    return Word.run(async (ctx) => {
      const range = ctx.document.getSelection()
      range.load('text')
      await ctx.sync()
      return range.text
    })
  }

  const resolveChatModelTier = (): ModelTier => (
    selectedModelInfo.value?.type === 'image' ? firstChatModelTier.value : selectedModelTier.value
  )

  function getEnabledToolNamesFromStorage(allToolNames: string[]): Set<string> {
    const fallback = new Set(allToolNames)
    try {
      const stored = localStorage.getItem('enabledTools')
      if (!stored) return fallback
      const parsed = JSON.parse(stored)
      if (!Array.isArray(parsed)) return fallback
      return new Set(parsed.filter((name): name is string => typeof name === 'string'))
    } catch {
      return fallback
    }
  }



  async function runAgentLoop(messages: ChatMessage[], modelTier: ModelTier) {
    const appToolDefs = hostIsOutlook ? getOutlookToolDefinitions() : hostIsPowerPoint ? getPowerPointToolDefinitions() : hostIsExcel ? getExcelToolDefinitions() : getWordToolDefinitions()
    const generalToolDefs = getGeneralToolDefinitions()
    const allToolDefs = [...generalToolDefs, ...appToolDefs]
    const enabledToolNames = getEnabledToolNamesFromStorage(allToolDefs.map(def => def.name))
    const enabledToolDefs = allToolDefs.filter(def => enabledToolNames.has(def.name))
    const tools = enabledToolDefs.map(def => ({ type: 'function' as const, function: { name: def.name, description: def.description, parameters: def.inputSchema } }))
    let iteration = 0
    const maxIter = Number(agentMaxIterations.value) || 25
    let currentMessages: ChatRequestMessage[] = [...messages]
    const analyzingPlaceholder = t('agentAnalyzing')
    history.value.push(createDisplayMessage('assistant', analyzingPlaceholder))
    const lastIndex = history.value.length - 1
    let abortedByUser = false
    while (iteration < maxIter) {
      if (abortController.value?.signal.aborted) {
        abortedByUser = true
        break
      }

      iteration++
      let response
      try {
        response = await chatSync({ messages: currentMessages, modelTier, tools, abortSignal: abortController.value?.signal })
      } catch (err: any) {
        if (err.name === 'AbortError' || abortController.value?.signal.aborted) {
          abortedByUser = true
          break
        }
        console.error('[AgentLoop] chatSync failed', {
          host: hostIsOutlook ? 'outlook' : hostIsPowerPoint ? 'powerpoint' : hostIsExcel ? 'excel' : 'word',
          modelTier,
          iteration,
          messageCount: currentMessages.length,
          error: err,
        })
        throw err
      }
      const choice = response.choices?.[0]
      if (!choice) break
      const assistantMsg = choice.message
      currentMessages.push({ role: 'assistant', content: assistantMsg.content || '' })
      if (assistantMsg.content) history.value[lastIndex].content = assistantMsg.content
      if (!assistantMsg.tool_calls?.length) break
      for (const toolCall of assistantMsg.tool_calls) {
        const toolName = toolCall.function.name
        let toolArgs: Record<string, any> = {}
        try { toolArgs = JSON.parse(toolCall.function.arguments) } catch {}
        let result = ''
        const toolDef = enabledToolDefs.find(tool => tool.name === toolName)
        if (toolDef) {
          try { result = await toolDef.execute(toolArgs) } catch (err: any) { console.error('[AgentLoop] tool execution failed', { toolName, toolArgs, error: err }); result = `Error: ${err.message}` }
        }
        if (abortController.value?.signal.aborted) {
          abortedByUser = true
          break
        }
        currentMessages.push({ role: 'tool', tool_call_id: toolCall.id, content: result })
      }
      if (abortedByUser) {
        break
      }
    }

    if (abortedByUser) {
      history.value.push(createDisplayMessage('system', t('agentStoppedByUser')))
      return
    }

    const assistantContent = history.value[lastIndex]?.content?.trim() || ''
    if (!assistantContent || assistantContent === analyzingPlaceholder) {
      history.value[lastIndex].content = t('noModelResponse')
    }

    if (iteration >= maxIter) messageUtil.warning(t('recursionLimitExceeded'))
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

    await runAgentLoop(messages, modelTier)
  }

  async function sendMessage() {
    if (!userInput.value.trim() || loading.value) return
    if (!backendOnline.value) return messageUtil.error(t('backendOffline'))
    const userMessage = userInput.value.trim()
    userInput.value = ''
    adjustTextareaHeight()

    let selectedText = ''
    if (useSelectedText.value) {
      try { selectedText = await getOfficeSelection() } catch {}
    }
    const selectionLabel = hostIsOutlook ? 'Email body' : hostIsPowerPoint ? 'Selected slide text' : hostIsExcel ? 'Selected cells' : 'Selected text'
    const fullMessage = selectedText ? `${userMessage}\n\n[${selectionLabel}: "${selectedText}"]` : userMessage
    history.value.push(createDisplayMessage('user', fullMessage))
    await scrollToBottom()

    loading.value = true
    abortController.value = new AbortController()
    try {
      await processChat(fullMessage)
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        console.error('[AgentLoop] sendMessage failed', error)
        messageUtil.error(t('failedToResponse'))
      }
    } finally {
      loading.value = false
      abortController.value = null
    }
  }

  async function applyQuickAction(actionKey: string) {
    if (!backendOnline.value) return messageUtil.error(t('backendOffline'))
    const selectedQuickAction = hostIsExcel ? excelQuickActions.value.find(a => a.key === actionKey) : hostIsPowerPoint ? powerPointQuickActions.find(a => a.key === actionKey) : quickActions.value.find(a => a.key === actionKey)
    if (hostIsExcel && selectedQuickAction?.mode === 'draft') {
      userInput.value = selectedQuickAction.prefix || ''
      adjustTextareaHeight(); await nextTick(); inputTextarea.value?.focus(); return
    }
    if (hostIsPowerPoint && (selectedQuickAction as PowerPointQuickAction)?.mode === 'draft') {
      userInput.value = t('pptVisualPrefix')
      adjustTextareaHeight(); await nextTick(); inputTextarea.value?.focus(); return
    }
    const selectedText = await getOfficeSelection({ includeOutlookSelectedText: true })
    if (!selectedText) return messageUtil.error(t(hostIsOutlook ? 'selectEmailPrompt' : hostIsPowerPoint ? 'selectSlideTextPrompt' : hostIsExcel ? 'selectCellsPrompt' : 'selectTextPrompt'))

    let action: { system: (lang: string) => string, user: (text: string, lang: string) => string } | undefined
    let systemMsg = ''
    let userMsg = ''
    if (hostIsOutlook) action = getOutlookBuiltInPrompt()[actionKey as keyof typeof outlookBuiltInPrompt]
    else if (hostIsPowerPoint) action = getPowerPointBuiltInPrompt()[actionKey as keyof typeof powerPointBuiltInPrompt]
    else if (hostIsExcel) {
      if (selectedQuickAction?.mode === 'immediate' && selectedQuickAction.systemPrompt) {
        systemMsg = selectedQuickAction.systemPrompt
        userMsg = `Selection:\n${selectedText}`
      } else action = getExcelBuiltInPrompt()[actionKey as keyof typeof excelBuiltInPrompt]
    } else action = getBuiltInPrompt()[actionKey as keyof typeof buildInPrompt]
    if (!systemMsg || !userMsg) {
      if (!action) return
      const lang = replyLanguage.value || 'Français'
      systemMsg = action.system(lang)
      userMsg = action.user(selectedText, lang)
    }

    const actionLabel = selectedQuickAction?.label || t(actionKey)
    history.value.push(createDisplayMessage('user', `[${actionLabel}] ${selectedText.substring(0, 100)}...`))
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

  return { sendMessage, applyQuickAction, runAgentLoop, getOfficeSelection }
}
