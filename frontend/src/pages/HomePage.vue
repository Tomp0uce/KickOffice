<template>
  <div
    class="itemse-center relative flex h-full w-full flex-col justify-center bg-bg-secondary p-1"
  >
    <div class="relative flex h-full w-full flex-col gap-1 rounded-md">
      <ChatHeader
        :new-chat-title="t('newChat')"
        :settings-title="t('settings')"
        @new-chat="startNewChat"
        @settings="goToSettings"
      />

      <QuickActionsBar
        v-model:selected-prompt-id="selectedPromptId"
        :quick-actions="quickActions"
        :loading="loading"
        :saved-prompts="savedPrompts"
        :select-prompt-title="t('selectPrompt')"
        @apply-action="applyQuickAction"
        @load-prompt="loadSelectedPrompt"
      />

      <ChatMessageList
        ref="messageListRef"
        :history="history"
        :history-with-segments="historyWithSegments"
        :current-action="currentAction"
        :backend-online="backendOnline"
        :empty-title="$t('emptyTitle')"
        :empty-subtitle="
          $t(
            hostIsOutlook
              ? 'emptySubtitleOutlook'
              : hostIsPowerPoint
                ? 'emptySubtitlePowerPoint'
                : hostIsExcel
                  ? 'emptySubtitleExcel'
                  : 'emptySubtitle',
          )
        "
        :backend-online-label="t('backendOnline')"
        :backend-offline-label="t('backendOffline')"
        :replace-selected-text="t('replaceSelectedText')"
        :append-to-selection="t('appendToSelection')"
        :copy-to-clipboard="t('copyToClipboard')"
        :thought-process-label="t('thoughtProcess')"
        @insert-message="insertMessageToDocument"
        @copy-message="copyMessageToClipboard"
      />

      <ChatInput
        ref="chatInputRef"
        v-model:selected-model-tier="selectedModelTier"
        v-model="userInput"
        v-model:use-word-formatting="useWordFormatting"
        v-model:use-selected-text="useSelectedText"
        :available-models="availableModels"
        :input-placeholder="inputPlaceholder"
        :loading="loading"
        :backend-online="backendOnline"
        :show-word-formatting="
          !hostIsExcel && !hostIsPowerPoint && !hostIsOutlook
        "
        :use-word-formatting-label="$t('useWordFormattingLabel')"
        :include-selection-label="
          $t(
            hostIsOutlook
              ? 'includeSelectionLabelOutlook'
              : hostIsPowerPoint
                ? 'includeSelectionLabelPowerPoint'
                : hostIsExcel
                  ? 'includeSelectionLabelExcel'
                  : 'includeSelectionLabel',
          )
        "
        :task-type-label="t('taskTypeLabel')"
        :send-label="t('send')"
        :stop-label="t('stop')"
        :draft-focus-glow="draftFocusGlow"
        @submit="sendMessage"
        @stop="stopGeneration"
      />
    </div>
  </div>
</template>

<script lang="ts" setup>
import { useStorage } from "@vueuse/core";
import {
  BookOpen,
  Brush,
  Briefcase,
  CheckCheck,
  CheckCircle,
  Eraser,
  Eye,
  FileCheck,
  FunctionSquare,
  Globe,
  Image,
  ListTodo,
  Mail,
  MessageSquare,
  Scissors,
  Sparkle,
  Wand2,
  Zap,
} from "lucide-vue-next";
import {
  computed,
  nextTick,
  onBeforeMount,
  onMounted,
  onUnmounted,
  ref,
  watch,
} from "vue";
import { useI18n } from "vue-i18n";
import { useRouter } from "vue-router";

import { fetchModels, healthCheck } from "@/api/backend";
import ChatHeader from "@/components/chat/ChatHeader.vue";
import ChatInput from "@/components/chat/ChatInput.vue";
import ChatMessageList from "@/components/chat/ChatMessageList.vue";
import QuickActionsBar from "@/components/chat/QuickActionsBar.vue";
import { useAgentLoop } from "@/composables/useAgentLoop";
import { useImageActions } from "@/composables/useImageActions";
import { useOfficeInsert } from "@/composables/useOfficeInsert";
import type {
  DisplayMessage,
  ExcelQuickAction,
  PowerPointQuickAction,
  OutlookQuickAction,
  QuickAction,
} from "@/types/chat";
import { localStorageKey } from "@/utils/enum";
import {
  isExcel,
  isOutlook,
  isPowerPoint,
  isWord,
} from "@/utils/hostDetection";
import {
  loadSavedPromptsFromStorage,
  type SavedPrompt,
} from "@/utils/savedPrompts";

const router = useRouter();
const { t } = useI18n();

const savedPrompts = ref<SavedPrompt[]>([]);
const selectedPromptId = ref("");
const customSystemPrompt = ref("");
const draftFocusGlow = ref(false);
const backendOnline = ref(false);
const availableModels = ref<Record<string, ModelInfo>>({});
const selectedModelTier = useStorage<ModelTier>(
  localStorageKey.modelTier,
  "standard",
);
const hostIsExcel = isExcel();
const hostIsWord = isWord();
const hostIsPowerPoint = isPowerPoint();
const hostIsOutlook = isOutlook();

const currentHost = hostIsWord
  ? "word"
  : hostIsExcel
    ? "excel"
    : hostIsPowerPoint
      ? "powerpoint"
      : hostIsOutlook
        ? "outlook"
        : "unknown";
const MAX_HISTORY_MESSAGES = 100;
const history = useStorage<DisplayMessage[]>(`chatHistory_${currentHost}`, []);
watch(
  () => history.value.length,
  (len) => {
    if (len > MAX_HISTORY_MESSAGES) {
      history.value = history.value.slice(len - MAX_HISTORY_MESSAGES);
    }
  },
);
const userInput = ref("");
const loading = ref(false);
const imageLoading = ref(false);
const abortController = ref<AbortController | null>(null);
const backendCheckInterval = ref<number | null>(null);
const useWordFormatting = useStorage(localStorageKey.useWordFormatting, true);
const useSelectedText = useStorage(localStorageKey.useSelectedText, true);
const replyLanguage = useStorage(localStorageKey.replyLanguage, "Français");
const agentMaxIterations = useStorage(localStorageKey.agentMaxIterations, 25);
const userGender = useStorage(localStorageKey.userGender, "unspecified");
const userFirstName = useStorage(localStorageKey.userFirstName, "");
const userLastName = useStorage(localStorageKey.userLastName, "");
const excelFormulaLanguage = useStorage<"en" | "fr">(
  localStorageKey.excelFormulaLanguage,
  "en",
);
const insertType = ref<insertTypes>("replace");

const chatInputRef = ref<InstanceType<typeof ChatInput>>();
const messageListRef = ref<InstanceType<typeof ChatMessageList>>();

const wordQuickActions: QuickAction[] = [
  {
    key: "proofread",
    label: t("proofread"),
    icon: CheckCheck,
    executeWithAgent: true,
    tooltipKey: "proofread_tooltip",
  },
  {
    key: "translate",
    label: t("translate"),
    icon: Globe,
    tooltipKey: "translate_tooltip",
  },
  {
    key: "polish",
    label: t("polish"),
    icon: Sparkle,
    tooltipKey: "polish_tooltip",
  },
  {
    key: "academic",
    label: t("academic"),
    icon: BookOpen,
    tooltipKey: "academic_tooltip",
  },
  {
    key: "summary",
    label: t("summary"),
    icon: FileCheck,
    tooltipKey: "summary_tooltip",
  },
];
const excelQuickActions = computed<ExcelQuickAction[]>(() => [
  {
    key: "clean",
    label: t("clean"),
    icon: Eraser,
    mode: "immediate",
    systemPrompt: "You are a data cleaning expert.",
    tooltipKey: "excelClean_tooltip",
  },
  {
    key: "beautify",
    label: t("beautify"),
    icon: Brush,
    mode: "immediate",
    systemPrompt: "You are an Excel formatting expert.",
    tooltipKey: "excelBeautify_tooltip",
  },
  {
    key: "formula",
    label: t("excelFormula"),
    icon: FunctionSquare,
    mode: "draft",
    prefix: t("excelFormulaPrefix"),
    tooltipKey: "excelFormula_tooltip",
  },
  {
    key: "transform",
    label: t("transform"),
    icon: Wand2,
    mode: "draft",
    prefix: t("excelTransformPrefix"),
    tooltipKey: "excelTransform_tooltip",
  },
  {
    key: "highlight",
    label: t("highlight"),
    icon: Eye,
    mode: "draft",
    prefix: t("excelHighlightPrefix"),
    tooltipKey: "excelHighlight_tooltip",
  },
]);
const outlookQuickActions: OutlookQuickAction[] = [
  {
    key: "proofread",
    label: t("outlookProofread"),
    icon: CheckCheck,
    tooltipKey: "outlookProofread_tooltip",
  },
  {
    key: "translate",
    label: t("translate"),
    icon: Globe,
    tooltipKey: "translate_tooltip",
  },
  {
    key: "concise",
    label: t("outlookConcise"),
    icon: Scissors,
    tooltipKey: "outlookConcise_tooltip",
  },
  {
    key: "extract",
    label: t("outlookExtract"),
    icon: ListTodo,
    tooltipKey: "outlookExtract_tooltip",
  },
  {
    key: "reply",
    label: t("outlookReply"),
    icon: Mail,
    mode: "smart-reply",
    prefix: t("outlookReplyPrePrompt"),
    tooltipKey: "outlookReply_tooltip",
  },
];
const powerPointQuickActions: PowerPointQuickAction[] = [
  {
    key: "proofread",
    label: t("proofread"),
    icon: CheckCheck,
    mode: "immediate",
    executeWithAgent: true,
    tooltipKey: "proofread_tooltip",
  },
  {
    key: "translate",
    label: t("translate"),
    icon: Globe,
    mode: "immediate",
    tooltipKey: "translate_tooltip",
  },
  {
    key: "speakerNotes",
    label: t("pptSpeakerNotes"),
    icon: MessageSquare,
    mode: "immediate",
    tooltipKey: "pptSpeakerNotes_tooltip",
  },
  {
    key: "punchify",
    label: t("pptPunchify"),
    icon: Zap,
    mode: "immediate",
    tooltipKey: "pptPunchify_tooltip",
  },
  {
    key: "visual",
    label: t("pptVisual"),
    icon: Image,
    mode: "immediate",
    tooltipKey: "pptVisual_tooltip",
  },
];

const quickActions = computed(() =>
  hostIsOutlook
    ? outlookQuickActions
    : hostIsPowerPoint
      ? powerPointQuickActions
      : hostIsExcel
        ? excelQuickActions.value
        : wordQuickActions,
);
const selectedModelInfo = computed(
  () => availableModels.value[selectedModelTier.value],
);
const firstChatModelTier = computed<ModelTier>(
  () =>
    (Object.entries(availableModels.value).find(
      ([, model]) => model.type !== "image",
    )?.[0] as ModelTier) || "standard",
);
const inputPlaceholder = computed(() =>
  selectedModelInfo.value?.type === "image"
    ? t("describeImage")
    : t("directTheAgent"),
);

function adjustTextareaHeight() {
  const candidate = chatInputRef.value?.textareaEl;
  const textarea =
    candidate && "style" in candidate
      ? (candidate as HTMLTextAreaElement)
      : candidate?.value;

  if (textarea && textarea.style) {
    textarea.style.height = "auto";
    textarea.style.height = `${Math.min(textarea.scrollHeight, 120)}px`;
  }
}

watch(userInput, () => {
  adjustTextareaHeight();
});

type ScrollMode = 'bottom' | 'message-top' | 'auto';

/**
 * Scroll the chat container
 * @param mode - 'bottom': scroll to very bottom, 'message-top': scroll to top of last message, 'auto': smart behavior
 */
async function scrollToBottom(mode: ScrollMode = 'auto') {
  await nextTick();
  const rawContainer = messageListRef.value?.containerEl;
  const container = ((rawContainer as any)?.value || rawContainer) as
    | HTMLElement
    | undefined;
  if (!container) return;

  const messageElements = container.querySelectorAll(".group");
  const lastMessage = messageElements[messageElements.length - 1] as
    | HTMLElement
    | undefined;

  if (!lastMessage) {
    container.scrollTop = container.scrollHeight;
    return;
  }

  const msgTop = lastMessage.offsetTop;
  const padding = 12;

  if (mode === 'bottom') {
    // Always scroll to the very bottom
    container.scrollTo({ top: container.scrollHeight, behavior: "smooth" });
  } else if (mode === 'message-top') {
    // Scroll so the top of the last message is visible at the top of the container
    container.scrollTo({ top: msgTop - padding, behavior: "smooth" });
  } else {
    // Auto mode: smart behavior based on message height
    // If the message is taller than the container, keep its start visible
    // Otherwise, scroll to bottom as content grows
    if (lastMessage.offsetHeight > container.clientHeight) {
      container.scrollTo({ top: msgTop - padding, behavior: "smooth" });
    } else {
      container.scrollTo({ top: container.scrollHeight, behavior: "smooth" });
    }
  }
}

/**
 * Scroll to show the top of the last message (for receiving new assistant messages)
 */
async function scrollToMessageTop() {
  await scrollToBottom('message-top');
}

/**
 * Scroll to very bottom (for user sending messages and on startup)
 */
async function scrollToVeryBottom() {
  await scrollToBottom('bottom');
}

const imageActions = useImageActions(t);
const historyWithSegments = computed(() =>
  imageActions.historyWithSegments(history),
);

const officeInsert = useOfficeInsert({
  hostIsOutlook,
  hostIsPowerPoint,
  hostIsExcel,
  hostIsWord,
  useWordFormatting,
  insertType,
  t,
  shouldTreatMessageAsImage: imageActions.shouldTreatMessageAsImage,
  getMessageActionPayload: imageActions.getMessageActionPayload,
  copyImageToClipboard: imageActions.copyImageToClipboard,
  insertImageToWord: imageActions.insertImageToWord,
  insertImageToPowerPoint: imageActions.insertImageToPowerPoint,
});

const { sendMessage, applyQuickAction, currentAction } = useAgentLoop({
  t,
  refs: {
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    abortController,
    inputTextarea: computed(() => chatInputRef.value?.textareaEl),
    draftFocusGlow,
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
    replyLanguage,
    agentMaxIterations,
    useSelectedText,
    excelFormulaLanguage,
    userGender,
    userFirstName,
    userLastName,
  },
  actions: {
    quickActions,
    outlookQuickActions: computed(() => outlookQuickActions),
    excelQuickActions,
    powerPointQuickActions,
  },
  helpers: {
    createDisplayMessage: imageActions.createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop,
    scrollToVeryBottom,
  },
});

function stopGeneration() {
  abortController.value?.abort();
  abortController.value = null;
  loading.value = false;
}

function goToSettings() {
  router.push("/settings");
}

function startNewChat() {
  if (loading.value) stopGeneration();
  userInput.value = "";
  history.value = [];
  customSystemPrompt.value = "";
  selectedPromptId.value = "";
  adjustTextareaHeight();
}

function loadSavedPrompts() {
  savedPrompts.value = loadSavedPromptsFromStorage([]);
}

function loadSelectedPrompt() {
  const prompt = savedPrompts.value.find(
    (p) => p.id === selectedPromptId.value,
  );
  if (!prompt) {
    customSystemPrompt.value = "";
    return;
  }
  customSystemPrompt.value = prompt.systemPrompt;
  userInput.value = prompt.userPrompt;
  adjustTextareaHeight();
  chatInputRef.value?.textareaEl?.value?.focus();
}

async function checkBackend() {
  backendOnline.value = await healthCheck();
  if (!backendOnline.value) return;
  try {
    availableModels.value = await fetchModels();
    if (!availableModels.value[selectedModelTier.value]) {
      const [firstTier] = Object.keys(availableModels.value);
      if (firstTier) selectedModelTier.value = firstTier as ModelTier;
    }
  } catch {
    console.error("Failed to fetch models");
  }
}

const { insertMessageToDocument, copyMessageToClipboard } = officeInsert;

onBeforeMount(() => {
  insertType.value =
    (localStorage.getItem(localStorageKey.insertType) as insertTypes) ||
    "replace";
  loadSavedPrompts();
  checkBackend();
  backendCheckInterval.value = window.setInterval(checkBackend, 30000);
});

onMounted(() => {
  // Scroll to bottom of history on initial load
  if (history.value.length > 0) {
    nextTick(() => {
      scrollToVeryBottom();
    });
  }
});

onUnmounted(() => {
  if (backendCheckInterval.value !== null)
    window.clearInterval(backendCheckInterval.value);
});
</script>
