<template>
  <div class="items-center relative flex h-full w-full flex-col justify-center bg-bg-secondary p-1">
    <div class="relative flex h-full w-full flex-col gap-1 rounded-md">
      <ChatHeader
        :settings-title="t('settings')"
        :loading="loading"
        :sessions="sessionManager.sessions.value"
        :current-session-id="sessionManager.currentSessionId.value"
        @new-chat="executeNewChat"
        @settings="goToSettings"
        @switch-session="handleSwitchSession"
        @delete-session="handleDeleteSession"
      />

      <!-- Persistent Offline Indicator — only shown after first check to avoid false negative flash -->
      <div
        v-if="backendChecked && !backendOnline"
        class="flex items-center justify-center bg-red-500/10 py-1.5 px-3 rounded-md border border-red-500/20 shadow-xs mx-4 mt-2"
      >
        <span class="text-xs text-red-500 font-medium flex items-center gap-2">
          <span class="relative flex h-2 w-2">
            <span
              class="animate-ping absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-75"
            ></span>
            <span class="relative inline-flex rounded-full h-2 w-2 bg-red-500"></span>
          </span>
          {{ t('backendOffline') }}
        </span>
      </div>

      <!-- Auth Error Indicator (UX-M3) -->
      <div
        v-if="backendOnline && Object.keys(availableModels).length === 0 && !loading"
        class="flex flex-col items-center justify-center bg-warning/10 py-2 px-3 rounded-md border border-warning/30 shadow-xs mx-4 mt-2 mb-2 animate-in fade-in"
      >
        <span class="text-xs text-warning-700 dark:text-warning font-medium text-center mb-1">
          {{ t('authErrorBanner', 'Authentication required or invalid API key.') }}
        </span>
        <button
          class="text-[11px] underline text-accent hover:text-accent-hover transition-colors"
          @click="goToSettings"
        >
          {{ t('goToSettings', 'Go to Settings to configure') }}
        </button>
      </div>

      <div
        v-if="isDeleteConfirmVisible"
        class="absolute inset-x-4 top-14 z-50 flex flex-col gap-3 rounded-md border border-border-secondary bg-bg-tertiary p-4 shadow-lg animate-in fade-in slide-in-from-top-4"
      >
        <p class="text-[13px] font-medium leading-tight text-main">
          {{ t('deleteSessionConfirm') }}
        </p>
        <div class="mt-1 flex justify-end gap-2">
          <button
            class="rounded-md border border-border-secondary bg-bg-secondary px-3 py-1.5 text-xs font-medium text-main transition-colors hover:bg-bg-tertiary"
            @click="isDeleteConfirmVisible = false"
          >
            {{ t('cancel') }}
          </button>
          <button
            class="rounded-md bg-red-600 px-3 py-1.5 text-xs font-medium text-white transition-colors hover:bg-red-700 focus:outline-hidden focus:ring-2 focus:ring-red-500 focus:ring-offset-2 focus:ring-offset-bg-tertiary"
            @click="confirmDeleteSession"
          >
            {{ t('confirm') }}
          </button>
        </div>
      </div>

      <div
        v-if="isNewChatConfirmVisible"
        class="absolute inset-x-4 top-14 z-50 flex flex-col gap-3 rounded-md border border-border-secondary bg-bg-tertiary p-4 shadow-lg animate-in fade-in slide-in-from-top-4"
      >
        <p class="text-[13px] font-medium leading-tight text-main">
          {{ t('newChatConfirm') }}
        </p>
        <div class="mt-1 flex justify-end gap-2">
          <button
            class="rounded-md border border-border-secondary bg-bg-secondary px-3 py-1.5 text-xs font-medium text-main transition-colors hover:bg-bg-tertiary"
            @click="isNewChatConfirmVisible = false"
          >
            {{ t('cancel') }}
          </button>
          <button
            class="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-white transition-colors hover:bg-primary/90"
            @click="confirmNewChat"
          >
            {{ t('confirm') }}
          </button>
        </div>
      </div>

      <QuickActionsBar
        v-model:selected-prompt-id="selectedPromptId"
        :quick-actions="quickActions ?? []"
        :loading="loading"
        :saved-prompts="savedPrompts"
        :select-prompt-title="t('selectPrompt')"
        @apply-action="applyQuickAction"
        @load-prompt="loadSelectedPrompt"
      />

      <!-- ARCH-H2 — Props removed, uses context via provide/inject -->
      <ChatMessageList ref="messageListRef" />

      <StatsBar
        v-model:selected-model-tier="selectedModelTier"
        :session-stats="sessionStats"
        :model-name="selectedModelInfo?.id ?? selectedModelTier"
        :current-action="currentAction"
        :context-window-tokens="selectedModelInfo?.contextWindow ?? 400_000"
        :loading="loading"
        :available-models="availableModels"
      />

      <ChatInput
        ref="chatInputRef"
        v-model="userInput"
        :use-word-formatting="true"
        :use-selected-text="true"
        :input-placeholder="inputPlaceholder"
        :loading="loading"
        :backend-online="backendOnline"
        :show-word-formatting="false"
        :send-label="t('send')"
        :stop-label="t('stop')"
        :draft-focus-glow="isDraftFocusGlowing"
        @submit="sendMessage"
        @stop="stopGeneration"
      />
    </div>
  </div>
</template>

<script lang="ts" setup>
import type { InsertType, ModelTier, ModelInfo } from '@/types';
defineOptions({ name: 'Home' });
import {
  ref,
  computed,
  watch,
  nextTick,
  onBeforeMount,
  onMounted,
  onActivated,
  onDeactivated,
} from 'vue';
import { useStorage } from '@vueuse/core';
import {
  BookOpen,
  ChartBarBig,
  CheckCheck,
  FileCheck,
  FunctionSquare,
  Globe,
  Grid3X3,
  Image,
  ListTodo,
  Mail,
  MessageSquare,
  NotebookPen,
  ScanSearch,
  Sparkle,
  Table,
  TrendingUp,
  Zap,
} from 'lucide-vue-next';
import { useI18n } from 'vue-i18n';

import { useHealthCheck } from '@/composables/useHealthCheck';
import { provideHomePageContext } from '@/composables/useHomePageContext'; // ARCH-H2
import ChatHeader from '@/components/chat/ChatHeader.vue';
import ChatInput from '@/components/chat/ChatInput.vue';
import ChatMessageList from '@/components/chat/ChatMessageList.vue';
import QuickActionsBar from '@/components/chat/QuickActionsBar.vue';
import StatsBar from '@/components/chat/StatsBar.vue';
import { useAgentLoop } from '@/composables/useAgentLoop';
import { useImageActions } from '@/composables/useImageActions';
import { useOfficeInsert } from '@/composables/useOfficeInsert';
import { useSessionManager } from '@/composables/useSessionManager';
import type {
  DisplayMessage,
  ExcelQuickAction,
  PowerPointQuickAction,
  OutlookQuickAction,
  QuickAction,
} from '@/types/chat';
import { localStorageKey } from '@/utils/enum';
import { isPowerPoint, isWord, isExcel, isOutlook, forHost } from '@/utils/hostDetection';
import { type SavedPrompt } from '@/utils/savedPrompts';
import { useHomePage } from '@/composables/useHomePage';
import type { ExcelFormulaLanguage } from '@/utils/constant'; // TOOL-M4

const { t } = useI18n();

const savedPrompts = ref<SavedPrompt[]>([]);
const selectedPromptId = ref('');
const customSystemPrompt = ref('');
const isDraftFocusGlowing = ref(false);
const isDeleteConfirmVisible = ref(false);
const isNewChatConfirmVisible = ref(false);
const availableModels = ref<Record<string, ModelInfo>>({});
const selectedModelTier = useStorage<ModelTier>(localStorageKey.modelTier, 'standard');

const { backendOnline, backendChecked } = useHealthCheck(availableModels, selectedModelTier);

const hostIsExcel = isExcel();
const hostIsWord = isWord();
const hostIsPowerPoint = isPowerPoint();
const hostIsOutlook = isOutlook();

const currentHost =
  forHost({
    outlook: 'outlook',
    powerpoint: 'powerpoint',
    excel: 'excel',
    word: 'word',
  }) || 'word';
const history = ref<DisplayMessage[]>([]);

// RACE-C1: pass loading so switchSession is blocked while the agent loop is active
const sessionManager = useSessionManager(currentHost, history, loading);
const userInput = ref('');
const loading = ref(false);
const imageLoading = ref(false);
const abortController = ref<AbortController | null>(null);
// GEN-L3: Format UI options removed, but keeping logic vars true implicitly in prompt logic
const agentMaxIterationsRaw = useStorage(localStorageKey.agentMaxIterations, 25);
const agentMaxIterations = computed(() => {
  const val = Number(agentMaxIterationsRaw.value);
  if (isNaN(val) || val < 1) return 1;
  if (val > 100) return 100;
  return Math.floor(val);
});
const userGender = useStorage(localStorageKey.userGender, 'unspecified');
const userFirstName = useStorage(localStorageKey.userFirstName, '');
const userLastName = useStorage(localStorageKey.userLastName, '');
const excelFormulaLanguage = useStorage<ExcelFormulaLanguage>(
  localStorageKey.excelFormulaLanguage,
  'en',
); // TOOL-M4
const insertType = ref<InsertType>('replace');

const chatInputRef = ref<InstanceType<typeof ChatInput>>();
const messageListRef = ref<InstanceType<typeof ChatMessageList>>();

const wordQuickActions = computed<QuickAction[]>(() => [
  {
    key: 'word-proofread',
    label: t('proofread'),
    icon: CheckCheck,
    executeWithAgent: true,
    tooltipKey: 'proofread_tooltip',
  },
  {
    key: 'word-translate',
    label: t('translate'),
    icon: Globe,
    executeWithAgent: true,
    tooltipKey: 'translate_tooltip',
  },
  {
    key: 'word-review',
    label: t('wordReview', 'Review'),
    icon: BookOpen,
    executeWithAgent: true,
    tooltipKey: 'wordReview_tooltip',
  },
  {
    key: 'polish',
    label: t('polish'),
    icon: Sparkle,
    tooltipKey: 'polish_tooltip',
  },
  {
    key: 'summary',
    label: t('summary'),
    icon: FileCheck,
    tooltipKey: 'summary_tooltip',
  },
]);
const excelQuickActions = computed<ExcelQuickAction[]>(() => [
  {
    key: 'ingest',
    label: t('excelIngest', 'Smart Ingestion'),
    icon: Table,
    mode: 'immediate',
    executeWithAgent: true,
    tooltipKey: 'excelIngest_tooltip',
  },
  {
    key: 'digitizeChart',
    label: t('excelDigitizeChart', 'Digitize Chart'),
    icon: ChartBarBig,
    mode: 'immediate',
    executeWithAgent: true,
    imageUpload: true,
    tooltipKey: 'excelDigitizeChart_tooltip',
  },
  {
    key: 'explain',
    label: t('excelExplain', 'Explain Formula'),
    icon: BookOpen,
    mode: 'immediate',
    executeWithAgent: true,
    systemPrompt:
      'You are an Excel expert. Explain the selected formula or data in simple terms: what it does, how it works, and any edge cases to be aware of.',
    tooltipKey: 'excelExplain_tooltip',
  },
  {
    key: 'formulaGenerator',
    label: t('excelFormulaGenerator', 'Formula Generator'),
    icon: FunctionSquare,
    mode: 'draft',
    prefix: t('excelFormulaGeneratorPrefix', 'Help me build a formula'),
    tooltipKey: 'excelFormulaGenerator_tooltip',
  },
  {
    key: 'dataTrend',
    label: t('excelDataTrend', 'Data Trend'),
    icon: TrendingUp,
    mode: 'immediate',
    executeWithAgent: true,
    systemPrompt:
      'You are a data analyst. Analyze the trends in the selected data: identify patterns, outliers, growth rates, and provide a concise summary with actionable insights.',
    tooltipKey: 'excelDataTrend_tooltip',
  },
  {
    key: 'pixelArt',
    label: t('excelPixelArt', 'Pixel Art'),
    icon: Grid3X3,
    mode: 'immediate',
    executeWithAgent: true,
    imageUpload: true,
    tooltipKey: 'excelPixelArt_tooltip',
  },
]);
const outlookQuickActions = computed<OutlookQuickAction[]>(() => [
  {
    key: 'proofread',
    label: t('outlookProofread'),
    icon: CheckCheck,
    tooltipKey: 'outlookProofread_tooltip',
  },
  {
    key: 'translate',
    label: t('translate'),
    icon: Globe,
    tooltipKey: 'translate_tooltip',
  },
  {
    key: 'reply',
    label: t('outlookReply'),
    icon: Mail,
    mode: 'smart-reply',
    prefix: t('outlookReplyPrePrompt'),
    tooltipKey: 'outlookReply_tooltip',
  },
  {
    key: 'extract',
    label: t('outlookExtract'),
    icon: ListTodo,
    tooltipKey: 'outlookExtract_tooltip',
  },
  {
    key: 'mom',
    label: t('outlookMoM', 'MoM'),
    icon: NotebookPen,
    mode: 'mom',
    prefix: t('outlookMoMPrefix', 'Génère moi un compte rendu de réunion pour ces notes de réunion : '),
    tooltipKey: 'outlookMoM_tooltip',
  },
]);
const powerPointQuickActions = computed<PowerPointQuickAction[]>(() => [
  {
    key: 'ppt-proofread',
    label: t('proofread'),
    icon: CheckCheck,
    mode: 'immediate',
    executeWithAgent: true,
    tooltipKey: 'ppt_proofread_tooltip',
  },
  {
    key: 'ppt-translate',
    label: t('translate'),
    icon: Globe,
    mode: 'immediate',
    executeWithAgent: true,
    tooltipKey: 'translate_tooltip',
  },
  {
    // PPT-H2: replaced speakerNotes with review — no text selection required
    key: 'review',
    label: t('pptReview'),
    icon: ScanSearch,
    mode: 'immediate',
    tooltipKey: 'pptReview_tooltip',
  },
  {
    key: 'punchify',
    label: t('pptPunchify'),
    icon: Zap,
    mode: 'immediate',
    tooltipKey: 'pptPunchify_tooltip',
    executeWithAgent: true,
  },
  {
    key: 'visual',
    label: t('pptVisual'),
    icon: Image,
    mode: 'immediate',
    tooltipKey: 'pptVisual_tooltip',
  },
]);

const quickActions = computed(() =>
  forHost({
    outlook: outlookQuickActions.value,
    powerpoint: powerPointQuickActions.value,
    excel: excelQuickActions.value,
    word: wordQuickActions.value,
  }),
);
const selectedModelInfo = computed(() => availableModels.value[selectedModelTier.value]);
const firstChatModelTier = computed<ModelTier>(
  () =>
    (Object.entries(availableModels.value).find(
      ([, model]) => model.type !== 'image',
    )?.[0] as ModelTier) || 'standard',
);
const inputPlaceholder = computed(() =>
  selectedModelInfo.value?.type === 'image' ? t('describeImage') : t('directTheAgent'),
);

const imageActions = useImageActions(t);
const historyWithSegments = computed(() => imageActions.historyWithSegments(history));

const officeInsert = useOfficeInsert({
  hostIsOutlook,
  hostIsPowerPoint,
  hostIsExcel,
  hostIsWord,
  useWordFormatting: ref(true), // GEN-L3: Always true
  t,
  shouldTreatMessageAsImage: imageActions.shouldTreatMessageAsImage,
  getMessageActionPayload: imageActions.getMessageActionPayload,
  copyImageToClipboard: imageActions.copyImageToClipboard,
  insertImageToWord: imageActions.insertImageToWord,
  insertImageToPowerPoint: imageActions.insertImageToPowerPoint,
});

function stopGeneration() {
  abortController.value?.abort();
  abortController.value = null;
  loading.value = false;
}

const homePage = useHomePage({
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
  resetSessionStats: () => resetSessionStats?.(),
  rebuildSessionFiles: () => rebuildSessionFiles?.(),
  stopGeneration,
});

watch(userInput, () => {
  homePage.adjustTextareaHeight();
});

const {
  adjustTextareaHeight,
  scrollToBottom,
  scrollToMessageTop,
  scrollToVeryBottom,
  goToSettings,
  executeNewChat,
  confirmNewChat,
  handleSwitchSession,
  handleDeleteSession,
  confirmDeleteSession,
  loadSavedPrompts,
  loadSelectedPrompt,
  handleScroll, // UX-H1 — Smart scroll handler
  isAutoScrollEnabled, // UX-H1 — Auto-scroll state
} = homePage;

const {
  sendMessage,
  applyQuickAction,
  currentAction,
  sessionStats,
  resetSessionStats,
  rebuildSessionFiles,
} = useAgentLoop({
  t,
  refs: {
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    abortController,
    inputTextarea: computed(() => chatInputRef.value?.textareaEl),
    isDraftFocusGlowing,
  },
  models: {
    availableModels,
    selectedModelTier,
    selectedModelInfo,
    firstChatModelTier,
  },
  host: {
    isOutlook: hostIsOutlook,
    isPowerPoint: hostIsPowerPoint,
    isExcel: hostIsExcel,
    isWord: hostIsWord,
  },
  settings: {
    customSystemPrompt,
    agentMaxIterations,
    useSelectedText: ref(true), // GEN-L3: Always true
    excelFormulaLanguage,
    userGender,
    userFirstName,
    userLastName,
  },
  actions: {
    quickActions,
    outlookQuickActions,
    excelQuickActions,
    powerPointQuickActions,
  },
  helpers: {
    createDisplayMessage: imageActions.createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop,
    scrollToVeryBottom,
    captureBeforeInsert: officeInsert.captureBeforeInsert,
    saveSnapshot: officeInsert.saveSnapshot,
  },
});

function handleRegenerate() {
  homePage.handleRegenerate(history, sendMessage);
}

function handleEditMessage(message: DisplayMessage) {
  homePage.handleEditMessage(message);
}

const { insertMessageToDocument, copyMessageToClipboard, undoLastInsert, canUndo } = officeInsert;

// ARCH-H2 — Provide context to eliminate prop drilling (~44 bindings → 0)
provideHomePageContext({
  // State
  history,
  historyWithSegments,
  loading,
  imageLoading,
  backendOnline,
  backendChecked,
  currentAction,
  userInput,
  customSystemPrompt,
  selectedPromptId,
  savedPrompts,
  isDraftFocusGlowing,
  isAutoScrollEnabled,
  // Models
  availableModels,
  selectedModelTier,
  selectedModelInfo,
  // Quick Actions
  quickActions,
  // Session
  sessionManager,
  sessionStats,
  // Translations
  t,
  // Handlers
  sendMessage,
  applyQuickAction,
  stopGeneration,
  handleScroll,
  handleRegenerate,
  handleEditMessage,
  insertMessageToDocument,
  copyMessageToClipboard,
  undoLastInsert,
  canUndo,
  goToSettings,
  executeNewChat,
  handleSwitchSession,
  handleDeleteSession,
  loadSelectedPrompt,
  adjustTextareaHeight,
  // Computed
  inputPlaceholder,
});

// Persist session after each agent turn completes
watch(loading, async (isLoading, wasLoading) => {
  if (wasLoading && !isLoading) {
    await sessionManager.persistCurrentSession();
  }
});

onBeforeMount(async () => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as InsertType) || 'replace';
  loadSavedPrompts();
  await sessionManager.init();
  rebuildSessionFiles();
});

onActivated(() => {
  loadSavedPrompts();
});

onDeactivated(() => {
  if (loading.value) stopGeneration();
});

onMounted(() => {
  // On initial load, scroll to the top of the last message so the user
  // can read the most recent exchange from its beginning.
  if (history.value.length > 0) {
    nextTick(() => {
      scrollToMessageTop();
    });
  }
});
</script>
