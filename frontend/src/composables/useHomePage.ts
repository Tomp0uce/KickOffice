/**
 * QC-M2 — Extracts orchestration/business logic out of HomePage.vue.
 *
 * Exposes scroll helpers, textarea sizing, session handlers, chat actions,
 * and prompt management as composable functions so HomePage.vue only handles
 * template binding and high-level composition.
 *
 * UX-H1 — Smart scroll with manual interruption: auto-scroll is disabled
 * when the user manually scrolls away from the bottom, and re-enabled when
 * they scroll back near the bottom.
 */
import { nextTick, ref } from 'vue';
import type { Ref } from 'vue';
import { useRouter } from 'vue-router';
import type ChatInput from '@/components/chat/ChatInput.vue';
import type ChatMessageList from '@/components/chat/ChatMessageList.vue';
import { TEXTAREA_MAX_HEIGHT_PX } from '@/constants/limits';
import type { useSessionManager } from '@/composables/useSessionManager';
import { resetVfs } from '@/utils/vfs';

type SessionManager = ReturnType<typeof useSessionManager>;

type ScrollMode = 'bottom' | 'message-top' | 'auto';

// UX-H1: Threshold in pixels from bottom to consider user "at bottom"
const SCROLL_BOTTOM_THRESHOLD_PX = 100;

export function useHomePage(deps: {
  chatInputRef: Ref<InstanceType<typeof ChatInput> | undefined>;
  messageListRef: Ref<InstanceType<typeof ChatMessageList> | undefined>;
  userInput: Ref<string>;
  loading: Ref<boolean>;
  isDeleteConfirmVisible: Ref<boolean>;
  isNewChatConfirmVisible: Ref<boolean>;
  sessionManager: SessionManager;
  resetSessionStats: () => void;
  rebuildSessionFiles: () => void;
  stopGeneration: () => void;
}) {
  const router = useRouter();
  const {
    chatInputRef,
    messageListRef,
    userInput,
    loading,
    isDeleteConfirmVisible,
    isNewChatConfirmVisible,
    sessionManager,
    resetSessionStats,
    rebuildSessionFiles,
    stopGeneration,
  } = deps;

  // UX-H1 — Smart scroll: auto-scroll enabled by default, disabled when user scrolls up
  const isAutoScrollEnabled = ref(true);
  // Timestamp of the last programmatic scroll — handleScroll ignores events within 300ms of it
  let programmaticScrollTs = 0;

  // ─── Textarea ─────────────────────────────────────────────────────────────

  function adjustTextareaHeight() {
    const candidate = chatInputRef.value?.textareaEl;
    const textarea =
      candidate && 'style' in candidate
        ? (candidate as HTMLTextAreaElement)
        : (candidate as unknown as HTMLTextAreaElement);

    if (textarea && textarea.style) {
      textarea.style.height = 'auto';
      textarea.style.height = `${Math.min(textarea.scrollHeight, TEXTAREA_MAX_HEIGHT_PX)}px`;
    }
  }

  // ─── Scroll helpers ────────────────────────────────────────────────────────

  // UX-H1 — Check if container is scrolled near bottom
  function isNearBottom(container: HTMLElement): boolean {
    const scrollBottom = container.scrollHeight - container.scrollTop - container.clientHeight;
    return scrollBottom <= SCROLL_BOTTOM_THRESHOLD_PX;
  }

  // UX-H1 — Handle scroll event to detect manual user scrolling
  function handleScroll() {
    // Ignore scroll events fired by our own programmatic scrolls
    // 600ms covers smooth scrolls which can take 500ms+ to complete
    if (Date.now() - programmaticScrollTs < 600) return;

    const rawContainer = messageListRef.value?.containerEl;
    const container = ((rawContainer as any)?.value || rawContainer) as HTMLElement | undefined;
    if (!container) return;

    // If user scrolls near the bottom, re-enable auto-scroll
    // If user scrolls away from the bottom, disable auto-scroll
    isAutoScrollEnabled.value = isNearBottom(container);
  }

  async function scrollToBottom(mode: ScrollMode = 'auto', force = false) {
    await nextTick();
    const rawContainer = messageListRef.value?.containerEl;
    const container = ((rawContainer as any)?.value || rawContainer) as HTMLElement | undefined;
    if (!container) return;

    // UX-H1 — If forced, re-enable auto-scroll
    if (force) {
      isAutoScrollEnabled.value = true;
    }

    // UX-H1 — Respect auto-scroll flag unless forced
    if (!force && !isAutoScrollEnabled.value) return;

    // Mark this as a programmatic scroll so handleScroll ignores the resulting event
    programmaticScrollTs = Date.now();

    const messageElements = container.querySelectorAll('[data-message]');
    const lastMessage = messageElements[messageElements.length - 1] as HTMLElement | undefined;

    if (!lastMessage) {
      container.scrollTop = container.scrollHeight;
      return;
    }

    // Compute position relative to the scroll container using getBoundingClientRect,
    // NOT offsetTop (which is relative to offsetParent, not the scroll container).
    const containerRect = container.getBoundingClientRect();
    const msgRect = lastMessage.getBoundingClientRect();
    const msgTopRelative = msgRect.top - containerRect.top + container.scrollTop;
    const padding = 12;

    if (mode === 'bottom') {
      container.scrollTo({ top: container.scrollHeight, behavior: 'smooth' });
    } else if (mode === 'message-top') {
      container.scrollTo({ top: msgTopRelative - padding, behavior: 'smooth' });
    } else {
      if (lastMessage.offsetHeight > container.clientHeight) {
        container.scrollTo({ top: msgTopRelative - padding, behavior: 'smooth' });
      } else {
        container.scrollTo({ top: container.scrollHeight, behavior: 'smooth' });
      }
    }
  }

  async function scrollToMessageTop() {
    // UX-H1 — scrollToMessageTop is always called when new content arrives,
    // so we always force-enable auto-scroll to ensure the user sees new messages
    await scrollToBottom('message-top', true);
  }

  async function scrollToVeryBottom() {
    await scrollToBottom('bottom', false);
  }

  async function scrollToConversationTop() {
    await nextTick();
    const rawContainer = messageListRef.value?.containerEl;
    const container = ((rawContainer as any)?.value || rawContainer) as HTMLElement | undefined;
    if (!container) return;
    // UX-H1 — Disable auto-scroll when explicitly scrolling to top
    isAutoScrollEnabled.value = false;
    container.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ─── Navigation ────────────────────────────────────────────────────────────

  function goToSettings() {
    router.push('/settings');
  }

  // ─── Chat lifecycle ────────────────────────────────────────────────────────

  async function doNewChat() {
    if (loading.value) stopGeneration();
    await sessionManager.newSession();
    resetSessionStats();
    rebuildSessionFiles(); // clear session files — prevents leaking files into the new session
    resetVfs();            // clear VFS — new session starts with a clean filesystem
    userInput.value = '';
    await nextTick();
    const el = chatInputRef.value?.textareaEl as unknown as { focus?: () => void };
    el?.focus?.();
    adjustTextareaHeight();
  }

  async function executeNewChat() {
    if (userInput.value.trim()) {
      isNewChatConfirmVisible.value = true;
      return;
    }
    await doNewChat();
  }

  async function confirmNewChat() {
    isNewChatConfirmVisible.value = false;
    await doNewChat();
  }

  // ─── Session management ────────────────────────────────────────────────────

  async function handleSwitchSession(sessionId: string) {
    if (loading.value) return;
    await sessionManager.switchSession(sessionId);
    rebuildSessionFiles();
    resetVfs(); // each session has its own VFS state
    resetSessionStats();
    await nextTick();
    scrollToConversationTop();
  }

  function handleDeleteSession() {
    if (loading.value) return;
    isDeleteConfirmVisible.value = true;
  }

  async function confirmDeleteSession() {
    isDeleteConfirmVisible.value = false;
    await sessionManager.deleteCurrentSession();
    await nextTick();
    scrollToConversationTop();
  }

  // ─── Message actions ───────────────────────────────────────────────────────

  function handleRegenerate(
    history: Ref<Array<{ role: string; content?: string }>>,
    sendMessage: (content: string, files?: File[]) => void,
  ) {
    if (loading.value) return;
    const lastUserMsg = [...history.value].reverse().find(m => m.role === 'user');
    if (!lastUserMsg?.content) return;
    sendMessage(lastUserMsg.content, []);
  }

  async function handleEditMessage(message: { content?: string }) {
    userInput.value = message.content ?? '';
    await nextTick();
    const el = chatInputRef.value?.textareaEl as unknown as { focus?: () => void };
    el?.focus?.();
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
    // UX-H1 — Smart scroll exports
    handleScroll,
    isAutoScrollEnabled,
  };
}
