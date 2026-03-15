/**
 * ARCH-H2 — HomePage context via provide/inject to reduce prop drilling
 *
 * This composable provides a centralized context for HomePage.vue and its children,
 * eliminating the need to pass ~44 props through multiple component layers.
 */
import { inject, provide, type InjectionKey, type Ref, type ComputedRef } from 'vue';
import type { DisplayMessage, QuickAction, RenderSegment } from '@/types/chat';
import type { ModelInfo, ModelTier } from '@/types';
import type { SavedPrompt } from '@/utils/savedPrompts';
import type { SessionStats } from '@/composables/useAgentLoop';
import type { useSessionManager } from '@/composables/useSessionManager';

type SessionManager = ReturnType<typeof useSessionManager>;

export interface HomePageContext {
  // ─── State ───────────────────────────────────────────────────────────────
  history: Ref<DisplayMessage[]>;
  historyWithSegments: ComputedRef<
    Array<{
      key: string;
      message: DisplayMessage;
      segments: RenderSegment[];
    }>
  >;
  loading: Ref<boolean>;
  imageLoading: Ref<boolean>;
  backendOnline: Ref<boolean>;
  backendChecked: Ref<boolean>;
  currentAction: Ref<string>;
  userInput: Ref<string>;
  customSystemPrompt: Ref<string>;
  selectedPromptId: Ref<string>;
  savedPrompts: Ref<SavedPrompt[]>;
  isDraftFocusGlowing: Ref<boolean>;
  isAutoScrollEnabled: Ref<boolean>; // UX-H1

  // ─── Models ──────────────────────────────────────────────────────────────
  availableModels: Ref<Record<string, ModelInfo>>;
  selectedModelTier: Ref<ModelTier>;
  selectedModelInfo: ComputedRef<ModelInfo | undefined>;

  // ─── Quick Actions ───────────────────────────────────────────────────────
  quickActions: ComputedRef<QuickAction[] | undefined>;

  // ─── Session ─────────────────────────────────────────────────────────────
  sessionManager: SessionManager;
  sessionStats: Ref<SessionStats>;

  // ─── Translations ────────────────────────────────────────────────────────
  t: (key: string, fallback?: string) => string;

  // ─── Handlers ────────────────────────────────────────────────────────────
  sendMessage: (content: string, files?: File[]) => void;
  applyQuickAction: (action: QuickAction) => void;
  stopGeneration: () => void;
  handleScroll: () => void; // UX-H1
  handleRegenerate: () => void;
  handleEditMessage: (message: DisplayMessage) => void;
  insertMessageToDocument: (message: DisplayMessage, type: 'replace' | 'append') => void;
  copyMessageToClipboard: (message: DisplayMessage) => void;
  undoLastInsert: () => Promise<boolean>;
  canUndo: Ref<boolean>;
  goToSettings: () => void;
  executeNewChat: () => Promise<void>;
  handleSwitchSession: (sessionId: string) => Promise<void>;
  handleDeleteSession: () => void;
  loadSelectedPrompt: () => void;
  adjustTextareaHeight: () => void;

  // ─── Computed ────────────────────────────────────────────────────────────
  inputPlaceholder: ComputedRef<string>;
}

const HOME_PAGE_CONTEXT_KEY: InjectionKey<HomePageContext> = Symbol('homePageContext');

export function provideHomePageContext(context: HomePageContext) {
  provide(HOME_PAGE_CONTEXT_KEY, context);
}

export function useHomePageContext(): HomePageContext {
  const context = inject(HOME_PAGE_CONTEXT_KEY);
  if (!context) {
    throw new Error(
      'useHomePageContext must be called within a component that has provideHomePageContext',
    );
  }
  return context;
}
