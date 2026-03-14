<template>
  <div
    ref="containerEl"
    role="log"
    aria-live="polite"
    aria-relevant="additions text"
    class="card flex flex-1 flex-col gap-4 overflow-y-auto min-h-0"
    @scroll="handleScrollEvent"
  >
    <div class="sr-only" role="status" aria-live="polite" aria-atomic="true">
      {{ liveAnnouncement }}
    </div>
    <div
      v-if="history.length === 0"
      class="flex h-full flex-col items-center justify-center gap-4 p-8 text-center text-accent"
    >
      <Sparkles :size="32" />
      <p class="font-semibold text-main">{{ emptyTitle }}</p>
      <p class="text-xs font-semibold text-secondary">{{ emptySubtitle }}</p>
      <div
        role="status"
        class="flex items-center gap-1 rounded-md px-2 py-1 text-xs"
        :class="backendOnline ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'"
      >
        <div class="h-2 w-2 rounded-full" :class="backendOnline ? 'bg-green-500' : 'bg-red-500'" />
        {{ backendOnline ? backendOnlineLabel : backendOfflineLabel }}
      </div>
    </div>

    <div
      v-for="item in historyWithSegments"
      :key="item.key"
      v-show="hasContent(item)"
      data-message
      class="group flex items-end gap-4 [.user]:flex-row-reverse"
      :class="item.message.role === 'assistant' ? 'assistant' : 'user'"
    >
      <div
        class="flex min-w-0 flex-1 flex-col gap-1 group-[.assistant]:items-start group-[.assistant]:text-left group-[.user]:items-end group-[.user]:text-left"
      >
        <div class="flex flex-col gap-1">
          <div
            v-show="
              item.segments.some(s => s.type !== 'text' || s.text.trim() !== '') ||
              (item.message.toolCalls && item.message.toolCalls.length > 0) ||
              item.message.imageSrc
            "
            class="group max-w-[95%] rounded-md border border-border-secondary p-1 text-sm leading-[1.4] wrap-break-word text-main/90 shadow-sm group-[.assistant]:bg-bg-tertiary group-[.assistant]:text-left group-[.user]:bg-accent/10"
          >
            <template v-for="(segment, idx) in item.segments" :key="`${item.key}-segment-${idx}`">
              <MarkdownRenderer
                v-if="segment.type === 'text' && segment.text.trim() !== ''"
                :content="segment.text"
              />
              <!-- Enhanced thinking block with brain icon + streaming dots -->
              <div
                v-else
                class="mb-1 rounded-sm border border-border-secondary bg-bg-secondary overflow-hidden"
              >
                <button
                  type="button"
                  class="w-full flex items-center gap-1.5 px-2 py-1 text-[10px] uppercase tracking-wider text-accent hover:bg-bg-tertiary transition-colors"
                  :aria-expanded="isThoughtOpen(item.key, idx)"
                  :aria-label="thoughtProcessLabel"
                  @click="toggleThought(item.key, idx)"
                >
                  <component
                    :is="isThoughtOpen(item.key, idx) ? ChevronDown : ChevronRight"
                    :size="10"
                  />
                  <Brain :size="10" />
                  <span>{{ thoughtProcessLabel }}</span>
                  <!-- Streaming dots: show when loading and this is last segment -->
                  <span
                    v-if="
                      loading &&
                      idx === item.segments.length - 1 &&
                      item.key === historyWithSegments[historyWithSegments.length - 1]?.key
                    "
                    class="animate-pulse ml-1"
                    >...</span
                  >
                </button>
                <pre
                  v-if="isThoughtOpen(item.key, idx)"
                  class="m-0 px-2 py-1.5 text-xs wrap-break-word whitespace-pre-wrap text-secondary border-t border-border-secondary max-h-20 overflow-y-auto"
                  >{{ segment.text.trim() }}</pre
                >
              </div>
            </template>
            <ToolCallBlock v-for="tc in item.message.toolCalls" :key="tc.id" :tool-call="tc" />
            <img
              v-if="item.message.imageSrc"
              :src="item.message.imageSrc"
              class="mt-2 max-w-full rounded-md"
              alt="Generated image"
            />
            <!-- File attachment badges (Tâche 5) -->
            <div
              v-if="item.message.attachedFiles && item.message.attachedFiles.length > 0"
              class="mt-1.5 flex flex-wrap gap-1"
            >
              <div
                v-for="(f, i) in item.message.attachedFiles"
                :key="i"
                class="flex items-center gap-1 rounded-sm bg-accent/15 px-1.5 py-0.5 text-[10px] text-accent font-medium"
                :title="f.fileId ? `file_id: ${f.fileId}` : f.filename"
              >
                <Paperclip :size="9" />
                <span class="max-w-[120px] truncate">{{ f.filename }}</span>
              </div>
            </div>
          </div>
          <div
            v-if="
              item.message.timestamp &&
              (item.segments.some(s => s.type !== 'text' || s.text.trim() !== '') ||
                (item.message.toolCalls && item.message.toolCalls.length > 0) ||
                item.message.imageSrc)
            "
            class="text-[10px] text-secondary/60 px-1"
          >
            {{ formatTime(item.message.timestamp) }}
          </div>
        </div>
        <!-- Assistant action buttons: hidden until hover (U-L1) -->
        <div
          v-if="item.message.role === 'assistant' && hasContent(item)"
          class="flex gap-1 opacity-0 group-hover:opacity-100 focus-within:opacity-100 transition-opacity duration-150"
        >
          <CustomButton
            :title="replaceSelectedText"
            text=""
            :icon="FileText"
            type="secondary"
            class="bg-surface! p-1.5! text-secondary!"
            :icon-size="12"
            @click="context.insertMessageToDocument(item.message, 'replace')"
          />
          <CustomButton
            :title="appendToSelection"
            text=""
            :icon="Plus"
            type="secondary"
            class="bg-surface! p-1.5! text-secondary!"
            :icon-size="12"
            @click="context.insertMessageToDocument(item.message, 'append')"
          />
          <CustomButton
            :title="copyToClipboard"
            text=""
            :icon="Copy"
            type="secondary"
            class="bg-surface! p-1.5! text-secondary!"
            :icon-size="12"
            @click="context.copyMessageToClipboard(item.message)"
          />
          <!-- Regenerate: only on the last assistant message (U-L2) -->
          <CustomButton
            v-if="item.key === lastAssistantKey && !loading"
            :title="regenerateLabel"
            text=""
            :icon="RotateCcw"
            type="secondary"
            class="bg-surface! p-1.5! text-secondary!"
            :icon-size="12"
            @click="context.handleRegenerate()"
          />
        </div>
        <!-- User message edit button: hidden until hover (U-L2) -->
        <div
          v-if="item.message.role === 'user' && hasContent(item)"
          class="flex gap-1 opacity-0 group-hover:opacity-100 focus-within:opacity-100 transition-opacity duration-150"
        >
          <CustomButton
            :title="editMessageLabel"
            text=""
            :icon="Pencil"
            type="secondary"
            class="bg-surface! p-1.5! text-secondary!"
            :icon-size="12"
            @click="context.handleEditMessage(item.message)"
          />
        </div>
      </div>
    </div>

    <!-- Agent Action Indicator (Transferred from StatsBar) -->
    <div
      v-if="
        currentAction ||
        (loading && history.length > 0 && history[history.length - 1].role !== 'assistant')
      "
      class="flex items-end gap-4 assistant mt-2"
    >
      <div class="flex min-w-0 flex-1 flex-col gap-1 items-start text-left">
        <div
          class="max-w-[95%] rounded-md border border-border-secondary px-3 py-2 text-[10px] leading-[1.4] wrap-break-word text-main/90 shadow-sm bg-bg-tertiary"
        >
          <div class="flex items-start gap-2 text-accent" role="status" aria-live="polite">
            <span class="inline-flex mt-0.5 h-2 w-2 shrink-0 animate-pulse rounded-full bg-accent" />
            <Terminal class="shrink-0 mt-0.5" :size="ICON_SIZE_SM" v-if="currentAction" />
            <span class="line-clamp-2 break-words" v-if="currentAction">{{ currentAction }}</span>
            <span v-else class="animate-pulse">▊</span>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import {
  Brain,
  ChevronDown,
  ChevronRight,
  Copy,
  FileText,
  Paperclip,
  Pencil,
  Plus,
  RotateCcw,
  Sparkles,
  Terminal,
} from 'lucide-vue-next'
import { computed, ref, watch } from 'vue'

import CustomButton from '@/components/CustomButton.vue'
import MarkdownRenderer from '@/components/chat/MarkdownRenderer.vue'
import ToolCallBlock from '@/components/chat/ToolCallBlock.vue'
import type { DisplayMessage, RenderSegment } from '@/types/chat'
import { ICON_SIZE_SM } from '@/constants/limits'
import { useHomePageContext } from '@/composables/useHomePageContext' // ARCH-H2
import { forHost } from '@/utils/hostDetection' // ARCH-H2

// ARCH-H2 — Use context instead of props to eliminate prop drilling
const context = useHomePageContext()

// Keep props optional for backward compatibility during migration
const props = withDefaults(defineProps<{
  history?: DisplayMessage[]
  historyWithSegments?: Array<{
    key: string
    message: DisplayMessage
    segments: RenderSegment[]
  }>
  currentAction?: string
  loading?: boolean
  backendOnline?: boolean
  emptyTitle?: string
  emptySubtitle?: string
  backendOnlineLabel?: string
  backendOfflineLabel?: string
  replaceSelectedText?: string
  appendToSelection?: string
  copyToClipboard?: string
  thoughtProcessLabel?: string
  regenerateLabel?: string
  editMessageLabel?: string
  onScroll?: () => void // UX-H1 — Optional scroll handler
}>(), {})

const emit = defineEmits<{
  (e: 'insert-message', message: DisplayMessage, type: 'replace' | 'append'): void
  (e: 'copy-message', message: DisplayMessage): void
  (e: 'regenerate'): void
  (e: 'edit-message', message: DisplayMessage): void
}>()

// ARCH-H2 — Use context values with props as fallback
const history = computed(() => props.history ?? context.history.value)
const historyWithSegments = computed(() => props.historyWithSegments ?? context.historyWithSegments.value)
const currentAction = computed(() => props.currentAction ?? context.currentAction.value)
const loading = computed(() => props.loading ?? context.loading.value)
const backendOnline = computed(() => props.backendOnline ?? context.backendOnline.value)
const emptyTitle = computed(() => props.emptyTitle ?? context.t('emptyTitle'))
const emptySubtitle = computed(() => {
  if (props.emptySubtitle) return props.emptySubtitle
  // ARCH-H2 — Use forHost to determine the correct subtitle key
  const subtitleKey = forHost({
    outlook: 'emptySubtitleOutlook',
    powerpoint: 'emptySubtitlePowerPoint',
    excel: 'emptySubtitleExcel',
    word: 'emptySubtitle',
  }) || 'emptySubtitle'
  return context.t(subtitleKey)
})
const backendOnlineLabel = computed(() => props.backendOnlineLabel ?? context.t('backendOnline'))
const backendOfflineLabel = computed(() => props.backendOfflineLabel ?? context.t('backendOffline'))
const replaceSelectedText = computed(() => props.replaceSelectedText ?? context.t('replaceSelectedText'))
const appendToSelection = computed(() => props.appendToSelection ?? context.t('appendToSelection'))
const copyToClipboard = computed(() => props.copyToClipboard ?? context.t('copyToClipboard'))
const thoughtProcessLabel = computed(() => props.thoughtProcessLabel ?? context.t('thoughtProcess'))
const regenerateLabel = computed(() => props.regenerateLabel ?? context.t('regenerate'))
const editMessageLabel = computed(() => props.editMessageLabel ?? context.t('editMessage'))

// UX-H1 — Delegate scroll event to context or prop handler
function handleScrollEvent() {
  if (props.onScroll) {
    props.onScroll()
  } else {
    context.handleScroll()
  }
}

const containerEl = ref<HTMLElement>()

function hasContent(item: { message: DisplayMessage; segments: RenderSegment[] }): boolean {
  return (
    item.segments.some(s => s.type !== 'text' || s.text.trim() !== '') ||
    (item.message.toolCalls != null && item.message.toolCalls.length > 0) ||
    !!item.message.imageSrc
  )
}

const lastAssistantKey = computed(() => {
  const items = props.historyWithSegments.filter(item => item.message.role === 'assistant')
  return items[items.length - 1]?.key ?? null
})

const expandedThoughts = ref<Record<string, boolean>>({})

function thoughtKey(itemKey: string, segmentIndex: number): string {
  return `${itemKey}-${segmentIndex}`
}

function isThoughtOpen(itemKey: string, segmentIndex: number): boolean {
  return expandedThoughts.value[thoughtKey(itemKey, segmentIndex)] || false
}

function toggleThought(itemKey: string, segmentIndex: number): void {
  const key = thoughtKey(itemKey, segmentIndex)
  expandedThoughts.value[key] = !expandedThoughts.value[key]
}

const liveAnnouncement = ref('')

function formatTime(timestamp: number): string {
  const date = new Date(timestamp)
  const hours = date.getHours().toString().padStart(2, '0')
  const minutes = date.getMinutes().toString().padStart(2, '0')
  return `${hours}:${minutes}`
}

watch(
  () => props.history.length,
  (nextLength, previousLength = 0) => {
    // Clean up expandedThoughts for removed messages to prevent memory leaks (PM11)
    if (nextLength < previousLength) {
      const currentKeys = new Set(props.historyWithSegments.map(item => item.key))
      for (const key of Object.keys(expandedThoughts.value)) {
        const itemKey = key.split('-')[0]
        if (!currentKeys.has(itemKey)) {
          delete expandedThoughts.value[key]
        }
      }
    }

    if (nextLength <= previousLength || nextLength === 0) return
    const latestMessage = props.history[nextLength - 1]
    if (!latestMessage) return

    if (latestMessage.role === 'assistant') {
      liveAnnouncement.value = latestMessage.content
        ? `Assistant: ${latestMessage.content}`
        : 'Assistant is generating a response.'
      return
    }

    if (latestMessage.role === 'user') {
      liveAnnouncement.value = `User: ${latestMessage.content}`
      return
    }

    liveAnnouncement.value = latestMessage.content
  },
)

defineExpose({ containerEl })
</script>
