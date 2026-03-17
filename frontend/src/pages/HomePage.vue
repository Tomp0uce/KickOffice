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

      <!-- UX-H1: Extracted sub-components -->
      <OfflineBanner />
      <AuthErrorBanner />
      <SessionConfirmDialogs
        :is-delete-confirm-visible="isDeleteConfirmVisible"
        :is-new-chat-confirm-visible="isNewChatConfirmVisible"
        @cancel-delete="isDeleteConfirmVisible = false"
        @confirm-delete="confirmDeleteSession"
        @cancel-new-chat="isNewChatConfirmVisible = false"
        @confirm-new-chat="confirmNewChat"
      />

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
        :can-validate-ai-changes="canValidateAiChanges"
        :on-validate-ai-changes="handleValidateAiChanges"
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
import { useI18n } from 'vue-i18n';

import { useHealthCheck } from '@/composables/useHealthCheck';
import { provideHomePageContext } from '@/composables/useHomePageContext'; // ARCH-H2
import ChatHeader from '@/components/chat/ChatHeader.vue';
import ChatInput from '@/components/chat/ChatInput.vue';
import ChatMessageList from '@/components/chat/ChatMessageList.vue';
import QuickActionsBar from '@/components/chat/QuickActionsBar.vue';
import StatsBar from '@/components/chat/StatsBar.vue';
import OfflineBanner from '@/components/chat/OfflineBanner.vue';
import AuthErrorBanner from '@/components/chat/AuthErrorBanner.vue';
import SessionConfirmDialogs from '@/components/chat/SessionConfirmDialogs.vue';
import { useAgentLoop } from '@/composables/useAgentLoop';
import { useImageActions } from '@/composables/useImageActions';
import { useOfficeInsert } from '@/composables/useOfficeInsert';
import { useSessionManager } from '@/composables/useSessionManager';
import type { DisplayMessage } from '@/types/chat';
import { useWordQuickActions } from '@/composables/quickActions/useWordQuickActions';
import { useExcelQuickActions } from '@/composables/quickActions/useExcelQuickActions';
import { useOutlookQuickActions } from '@/composables/quickActions/useOutlookQuickActions';
import { usePowerPointQuickActions } from '@/composables/quickActions/usePowerPointQuickActions';
import { localStorageKey } from '@/utils/enum';
import { isPowerPoint, isWord, isExcel, isOutlook, forHost } from '@/utils/hostDetection';
import { acceptAiChangesInDocument, hasAiTrackedChanges } from '@/utils/wordTools';
import { clearAllAgentHighlightsInWorkbook } from '@/utils/excelTools';
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

// "Valider les modifications IA" — shown only when there are actual AI changes to validate.
// For Word: checks trackedChanges attributed to the KickOffice AI author (WordApi 1.6).
// For Excel: shown after each agent turn since Excel highlights are always present post-run.
const canValidateAiChanges = ref(false);

async function refreshCanValidateAiChanges(): Promise<void> {
  if (hostIsWord) {
    canValidateAiChanges.value = await hasAiTrackedChanges();
  } else if (hostIsExcel) {
    canValidateAiChanges.value = true;
  } else {
    canValidateAiChanges.value = false;
  }
}

async function handleValidateAiChanges(): Promise<string> {
  let result = '';
  if (hostIsWord) result = await acceptAiChangesInDocument();
  else if (hostIsExcel) result = await clearAllAgentHighlightsInWorkbook();
  // After validation, re-check — if accepted, button should disappear
  await refreshCanValidateAiChanges();
  if (hostIsExcel) canValidateAiChanges.value = false;
  return result;
}

const currentHost =
  forHost({
    outlook: 'outlook',
    powerpoint: 'powerpoint',
    excel: 'excel',
    word: 'word',
  }) || 'word';
const history = ref<DisplayMessage[]>([]);
const loading = ref(false);

// RACE-C1: pass loading so switchSession is blocked while the agent loop is active
const sessionManager = useSessionManager(currentHost, history, loading);
const userInput = ref('');
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

// UX-H1 / QUAL-H2: Quick action definitions extracted to per-host composables
const { wordQuickActions } = useWordQuickActions();
const { excelQuickActions } = useExcelQuickActions();
const { outlookQuickActions } = useOutlookQuickActions();
const { powerPointQuickActions } = usePowerPointQuickActions();

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
    captureDocumentState: officeInsert.captureDocumentState,
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

// Persist session and refresh "Valider" button visibility after each agent turn
watch(loading, async (isLoading, wasLoading) => {
  if (wasLoading && !isLoading) {
    await sessionManager.persistCurrentSession();
    await refreshCanValidateAiChanges();
  }
});

onBeforeMount(async () => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as InsertType) || 'replace';
  loadSavedPrompts();
  await sessionManager.init();
  rebuildSessionFiles();
  await refreshCanValidateAiChanges();
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
