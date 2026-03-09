/**
 * QC-M2 — Extracts orchestration/business logic out of HomePage.vue.
 *
 * Exposes scroll helpers, textarea sizing, session handlers, chat actions,
 * and prompt management as composable functions so HomePage.vue only handles
 * template binding and high-level composition.
 */
import { nextTick } from 'vue'
import type { Ref } from 'vue'
import { useRouter } from 'vue-router'
import type ChatInput from '@/components/chat/ChatInput.vue'
import type ChatMessageList from '@/components/chat/ChatMessageList.vue'
import type { SavedPrompt } from '@/utils/savedPrompts'
import { loadSavedPromptsFromStorage } from '@/utils/savedPrompts'
import { TEXTAREA_MAX_HEIGHT_PX } from '@/constants/limits'
import type { useSessionManager } from '@/composables/useSessionManager'

type SessionManager = ReturnType<typeof useSessionManager>

type ScrollMode = 'bottom' | 'message-top' | 'auto'

export function useHomePage(deps: {
  chatInputRef: Ref<InstanceType<typeof ChatInput> | undefined>
  messageListRef: Ref<InstanceType<typeof ChatMessageList> | undefined>
  savedPrompts: Ref<SavedPrompt[]>
  userInput: Ref<string>
  customSystemPrompt: Ref<string>
  selectedPromptId: Ref<string>
  loading: Ref<boolean>
  isDeleteConfirmVisible: Ref<boolean>
  isNewChatConfirmVisible: Ref<boolean>
  sessionManager: SessionManager
  resetSessionStats: () => void
  rebuildSessionFiles: () => void
  stopGeneration: () => void
}) {
  const router = useRouter()
  const {
    chatInputRef,
    messageListRef,
    savedPrompts,
    userInput,
    customSystemPrompt,
    selectedPromptId,
    loading,
    isDeleteConfirmVisible,
    isNewChatConfirmVisible,
    sessionManager,
    resetSessionStats,
    rebuildSessionFiles,
    stopGeneration,
  } = deps

  // ─── Textarea ─────────────────────────────────────────────────────────────

  function adjustTextareaHeight() {
    const candidate = chatInputRef.value?.textareaEl
    const textarea =
      candidate && 'style' in candidate
        ? (candidate as HTMLTextAreaElement)
        : (candidate as unknown as HTMLTextAreaElement)

    if (textarea && textarea.style) {
      textarea.style.height = 'auto'
      textarea.style.height = `${Math.min(textarea.scrollHeight, TEXTAREA_MAX_HEIGHT_PX)}px`
    }
  }

  // ─── Scroll helpers ────────────────────────────────────────────────────────

  async function scrollToBottom(mode: ScrollMode = 'auto') {
    await nextTick()
    const rawContainer = messageListRef.value?.containerEl
    const container = ((rawContainer as any)?.value || rawContainer) as HTMLElement | undefined
    if (!container) return

    const messageElements = container.querySelectorAll('[data-message]')
    const lastMessage = messageElements[messageElements.length - 1] as HTMLElement | undefined

    if (!lastMessage) {
      container.scrollTop = container.scrollHeight
      return
    }

    const msgTop = lastMessage.offsetTop
    const padding = 12

    if (mode === 'bottom') {
      container.scrollTo({ top: container.scrollHeight, behavior: 'smooth' })
    } else if (mode === 'message-top') {
      container.scrollTo({ top: msgTop - padding, behavior: 'smooth' })
    } else {
      if (lastMessage.offsetHeight > container.clientHeight) {
        container.scrollTo({ top: msgTop - padding, behavior: 'smooth' })
      } else {
        container.scrollTo({ top: container.scrollHeight, behavior: 'smooth' })
      }
    }
  }

  async function scrollToMessageTop() {
    await scrollToBottom('message-top')
  }

  async function scrollToVeryBottom() {
    await scrollToBottom('bottom')
  }

  async function scrollToConversationTop() {
    await nextTick()
    const rawContainer = messageListRef.value?.containerEl
    const container = ((rawContainer as any)?.value || rawContainer) as HTMLElement | undefined
    if (!container) return
    container.scrollTo({ top: 0, behavior: 'smooth' })
  }

  // ─── Navigation ────────────────────────────────────────────────────────────

  function goToSettings() {
    router.push('/settings')
  }

  // ─── Chat lifecycle ────────────────────────────────────────────────────────

  async function doNewChat() {
    if (loading.value) stopGeneration()
    await sessionManager.newSession()
    resetSessionStats()
    userInput.value = ''
    customSystemPrompt.value = ''
    selectedPromptId.value = ''
    await nextTick()
    const el = chatInputRef.value?.textareaEl as unknown as { focus?: () => void }
    el?.focus?.()
    adjustTextareaHeight()
  }

  async function executeNewChat() {
    if (userInput.value.trim()) {
      isNewChatConfirmVisible.value = true
      return
    }
    await doNewChat()
  }

  async function confirmNewChat() {
    isNewChatConfirmVisible.value = false
    await doNewChat()
  }

  // ─── Session management ────────────────────────────────────────────────────

  async function handleSwitchSession(sessionId: string) {
    if (loading.value) return
    await sessionManager.switchSession(sessionId)
    rebuildSessionFiles()
    resetSessionStats()
    await nextTick()
    scrollToConversationTop()
  }

  function handleDeleteSession() {
    if (loading.value) return
    isDeleteConfirmVisible.value = true
  }

  async function confirmDeleteSession() {
    isDeleteConfirmVisible.value = false
    await sessionManager.deleteCurrentSession()
    await nextTick()
    scrollToConversationTop()
  }

  // ─── Message actions ───────────────────────────────────────────────────────

  function handleRegenerate(
    history: Ref<Array<{ role: string; content?: string }>>,
    sendMessage: (content: string, files?: File[]) => void,
  ) {
    if (loading.value) return
    const lastUserMsg = [...history.value].reverse().find(m => m.role === 'user')
    if (!lastUserMsg?.content) return
    sendMessage(lastUserMsg.content, [])
  }

  async function handleEditMessage(
    message: { content?: string },
  ) {
    userInput.value = message.content ?? ''
    await nextTick()
    const el = chatInputRef.value?.textareaEl as unknown as { focus?: () => void }
    el?.focus?.()
  }

  // ─── Saved prompts ─────────────────────────────────────────────────────────

  function loadSavedPrompts() {
    savedPrompts.value = loadSavedPromptsFromStorage([])
  }

  function loadSelectedPrompt() {
    const prompt = savedPrompts.value.find(p => p.id === selectedPromptId.value)
    if (!prompt) {
      customSystemPrompt.value = ''
      return
    }
    customSystemPrompt.value = prompt.systemPrompt
    userInput.value = prompt.userPrompt
    adjustTextareaHeight()
    const el = chatInputRef.value?.textareaEl as unknown as { focus?: () => void }
    el?.focus?.()
  }

  return {
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop,
    scrollToVeryBottom,
    scrollToConversationTop,
    goToSettings,
    executeNewChat,
    confirmNewChat,
    handleSwitchSession,
    handleDeleteSession,
    confirmDeleteSession,
    handleRegenerate,
    handleEditMessage,
    loadSavedPrompts,
    loadSelectedPrompt,
  }
}
