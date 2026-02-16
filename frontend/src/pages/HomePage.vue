<template>
  <div class="itemse-center relative flex h-full w-full flex-col justify-center bg-bg-secondary p-1">
    <div class="relative flex h-full w-full flex-col gap-1 rounded-md">
      <ChatHeader :new-chat-title="t('newChat')" :settings-title="t('settings')" @new-chat="startNewChat" @settings="goToSettings" />

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
        :backend-online="backendOnline"
        :empty-title="$t('emptyTitle')"
        :empty-subtitle="$t(hostIsOutlook ? 'emptySubtitleOutlook' : hostIsPowerPoint ? 'emptySubtitlePowerPoint' : hostIsExcel ? 'emptySubtitleExcel' : 'emptySubtitle')"
        :backend-online-label="t('backendOnline')"
        :backend-offline-label="t('backendOffline')"
        :replace-selected-text="t('replaceSelectedText')"
        :append-to-selection="t('appendToSelection')"
        :copy-to-clipboard="t('copyToClipboard')"
        @insert-message="insertMessageToDocument"
        @copy-message="copyMessageToClipboard"
      />

      <ChatInput
        ref="chatInputRef"
        v-model:selected-model-tier="selectedModelTier"
        v-model:user-input="userInput"
        v-model:use-word-formatting="useWordFormatting"
        v-model:use-selected-text="useSelectedText"
        :available-models="availableModels"
        :input-placeholder="inputPlaceholder"
        :loading="loading"
        :backend-online="backendOnline"
        :show-word-formatting="!hostIsExcel && !hostIsPowerPoint && !hostIsOutlook"
        :use-word-formatting-label="$t('useWordFormattingLabel')"
        :include-selection-label="$t(hostIsOutlook ? 'includeSelectionLabelOutlook' : hostIsPowerPoint ? 'includeSelectionLabelPowerPoint' : hostIsExcel ? 'includeSelectionLabelExcel' : 'includeSelectionLabel')"
        :task-type-label="t('taskTypeLabel')"
        :send-label="t('send')"
        :stop-label="t('stop')"
        @send="sendMessage"
        @stop="stopGeneration"
        @input="adjustTextareaHeight"
      />
    </div>
  </div>
</template>

<script lang="ts" setup>
import { useStorage } from '@vueuse/core'
import { BookOpen, Brush, Briefcase, CheckCheck, CheckCircle, Eraser, Eye, FileCheck, FunctionSquare, Globe, Image, ListTodo, Mail, MessageSquare, Minus, Scissors, Sparkle, Wand2, Zap } from 'lucide-vue-next'
import { computed, nextTick, onBeforeMount, onUnmounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { useRouter } from 'vue-router'

import { fetchModels, healthCheck } from '@/api/backend'
import ChatHeader from '@/components/chat/ChatHeader.vue'
import ChatInput from '@/components/chat/ChatInput.vue'
import ChatMessageList from '@/components/chat/ChatMessageList.vue'
import QuickActionsBar from '@/components/chat/QuickActionsBar.vue'
import { useAgentLoop } from '@/composables/useAgentLoop'
import { useImageActions } from '@/composables/useImageActions'
import { useOfficeInsert } from '@/composables/useOfficeInsert'
import type { DisplayMessage, ExcelQuickAction, PowerPointQuickAction, QuickAction } from '@/types/chat'
import { localStorageKey } from '@/utils/enum'
import { isExcel, isOutlook, isPowerPoint, isWord } from '@/utils/hostDetection'
import { loadSavedPromptsFromStorage, type SavedPrompt } from '@/utils/savedPrompts'

const router = useRouter()
const { t } = useI18n()

const savedPrompts = ref<SavedPrompt[]>([])
const selectedPromptId = ref('')
const customSystemPrompt = ref('')
const backendOnline = ref(false)
const availableModels = ref<Record<string, ModelInfo>>({})
const selectedModelTier = useStorage<ModelTier>(localStorageKey.modelTier, 'standard')
const history = ref<DisplayMessage[]>([])
const userInput = ref('')
const loading = ref(false)
const imageLoading = ref(false)
const abortController = ref<AbortController | null>(null)
const backendCheckInterval = ref<number | null>(null)
const useWordFormatting = useStorage(localStorageKey.useWordFormatting, true)
const useSelectedText = useStorage(localStorageKey.useSelectedText, true)
const replyLanguage = useStorage(localStorageKey.replyLanguage, 'Français')
const agentMaxIterations = useStorage(localStorageKey.agentMaxIterations, 25)
const userGender = useStorage(localStorageKey.userGender, 'unspecified')
const userFirstName = useStorage(localStorageKey.userFirstName, '')
const userLastName = useStorage(localStorageKey.userLastName, '')
const excelFormulaLanguage = useStorage<'en' | 'fr'>(localStorageKey.excelFormulaLanguage, 'en')
const insertType = ref<insertTypes>('replace')

const chatInputRef = ref<InstanceType<typeof ChatInput>>()
const messageListRef = ref<InstanceType<typeof ChatMessageList>>()

const hostIsExcel = isExcel()
const hostIsWord = isWord()
const hostIsPowerPoint = isPowerPoint()
const hostIsOutlook = isOutlook()

const wordQuickActions: QuickAction[] = [
  { key: 'translate', label: t('translate'), icon: Globe },
  { key: 'polish', label: t('polish'), icon: Sparkle },
  { key: 'academic', label: t('academic'), icon: BookOpen },
  { key: 'summary', label: t('summary'), icon: FileCheck },
  { key: 'grammar', label: t('grammar'), icon: CheckCircle },
]
const excelQuickActions = computed<ExcelQuickAction[]>(() => [
  { key: 'clean', label: t('clean'), icon: Eraser, mode: 'immediate', systemPrompt: 'You are a data cleaning expert.' },
  { key: 'beautify', label: t('beautify'), icon: Brush, mode: 'immediate', systemPrompt: 'You are an Excel formatting expert.' },
  { key: 'formula', label: t('excelFormula'), icon: FunctionSquare, mode: 'draft', prefix: 'Génère une formule Excel pour : ' },
  { key: 'transform', label: t('transform'), icon: Wand2, mode: 'draft', prefix: 'Transforme la sélection pour : ' },
  { key: 'highlight', label: t('highlight'), icon: Eye, mode: 'draft', prefix: 'Mets en évidence (couleur) les cellules qui : ' },
])
const outlookQuickActions: QuickAction[] = [
  { key: 'reply', label: t('outlookReply'), icon: Mail },
  { key: 'formalize', label: t('outlookFormalize'), icon: Briefcase },
  { key: 'concise', label: t('outlookConcise'), icon: Scissors },
  { key: 'proofread', label: t('outlookProofread'), icon: CheckCheck },
  { key: 'extract', label: t('outlookExtract'), icon: ListTodo },
]
const powerPointQuickActions: PowerPointQuickAction[] = [
  { key: 'bullets', label: t('pptBullets'), icon: ListTodo, mode: 'immediate' },
  { key: 'speakerNotes', label: t('pptSpeakerNotes'), icon: MessageSquare, mode: 'immediate' },
  { key: 'punchify', label: t('pptPunchify'), icon: Zap, mode: 'immediate' },
  { key: 'shrink', label: t('pptShrink'), icon: Minus, mode: 'immediate' },
  { key: 'visual', label: t('pptVisual'), icon: Image, mode: 'draft' },
]

const quickActions = computed(() => hostIsOutlook ? outlookQuickActions : hostIsPowerPoint ? powerPointQuickActions : hostIsExcel ? excelQuickActions.value : wordQuickActions)
const selectedModelInfo = computed(() => availableModels.value[selectedModelTier.value])
const firstChatModelTier = computed<ModelTier>(() => Object.entries(availableModels.value).find(([, model]) => model.type !== 'image')?.[0] as ModelTier || 'standard')
const inputPlaceholder = computed(() => selectedModelInfo.value?.type === 'image' ? t('describeImage') : t('directTheAgent'))

function adjustTextareaHeight() {
  const textarea = chatInputRef.value?.textareaEl?.value
  if (textarea) {
    textarea.style.height = 'auto'
    textarea.style.height = `${Math.min(textarea.scrollHeight, 120)}px`
  }
}

async function scrollToBottom() {
  await nextTick()
  const container = messageListRef.value?.containerEl?.value
  if (container) container.scrollTop = container.scrollHeight
}

const imageActions = useImageActions(t)
const historyWithSegments = computed(() => imageActions.historyWithSegments(history))

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
})

const { sendMessage, applyQuickAction } = useAgentLoop({
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
  inputTextarea: computed(() => chatInputRef.value?.textareaEl?.value),
  hostIsOutlook,
  hostIsPowerPoint,
  hostIsExcel,
  quickActions,
  excelQuickActions,
  powerPointQuickActions,
  createDisplayMessage: imageActions.createDisplayMessage,
  adjustTextareaHeight,
  scrollToBottom,
})

function stopGeneration() {
  abortController.value?.abort()
  abortController.value = null
  loading.value = false
}

function goToSettings() {
  router.push('/settings')
}

function startNewChat() {
  if (loading.value) stopGeneration()
  userInput.value = ''
  history.value = []
  customSystemPrompt.value = ''
  selectedPromptId.value = ''
  adjustTextareaHeight()
}

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
  chatInputRef.value?.textareaEl?.value?.focus()
}

async function checkBackend() {
  backendOnline.value = await healthCheck()
  if (!backendOnline.value) return
  try {
    availableModels.value = await fetchModels()
    if (!availableModels.value[selectedModelTier.value]) {
      const [firstTier] = Object.keys(availableModels.value)
      if (firstTier) selectedModelTier.value = firstTier as ModelTier
    }
  } catch {
    console.error('Failed to fetch models')
  }
}

const { insertMessageToDocument, copyMessageToClipboard } = officeInsert

onBeforeMount(() => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as insertTypes) || 'replace'
  loadSavedPrompts()
  checkBackend()
  backendCheckInterval.value = window.setInterval(checkBackend, 30000)
})

onUnmounted(() => {
  if (backendCheckInterval.value !== null) window.clearInterval(backendCheckInterval.value)
})
</script>
