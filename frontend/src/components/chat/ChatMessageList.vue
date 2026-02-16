<template>
  <div ref="containerEl" class="flex flex-1 flex-col gap-4 overflow-y-auto rounded-md border border-border-secondary bg-surface p-2 shadow-sm">
    <div v-if="history.length === 0" class="flex h-full flex-col items-center justify-center gap-4 p-8 text-center text-accent">
      <Sparkles :size="32" />
      <p class="font-semibold text-main">{{ emptyTitle }}</p>
      <p class="text-xs font-semibold text-secondary">{{ emptySubtitle }}</p>
      <div class="flex items-center gap-1 rounded-md px-2 py-1 text-xs" :class="backendOnline ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'">
        <div class="h-2 w-2 rounded-full" :class="backendOnline ? 'bg-green-500' : 'bg-red-500'" />
        {{ backendOnline ? backendOnlineLabel : backendOfflineLabel }}
      </div>
    </div>

    <div v-for="item in historyWithSegments" :key="item.key" class="group flex items-end gap-4 [.user]:flex-row-reverse" :class="item.message.role === 'assistant' ? 'assistant' : 'user'">
      <div class="flex min-w-0 flex-1 flex-col gap-1 group-[.assistant]:items-start group-[.assistant]:text-left group-[.user]:items-end group-[.user]:text-left">
        <div class="group max-w-[95%] rounded-md border border-border-secondary p-1 text-sm leading-[1.4] wrap-break-word whitespace-pre-wrap text-main/90 shadow-sm group-[.assistant]:bg-bg-tertiary group-[.assistant]:text-left group-[.user]:bg-accent/10">
          <template v-for="(segment, idx) in item.segments" :key="`${item.key}-segment-${idx}`">
            <span v-if="segment.type === 'text'">{{ segment.text.trim() }}</span>
            <details v-else class="mb-1 rounded-sm border border-border-secondary bg-bg-secondary">
              <summary class="cursor-pointer list-none p-1 text-sm font-semibold text-secondary">Thought process</summary>
              <pre class="m-0 p-1 text-xs wrap-break-word whitespace-pre-wrap text-secondary">{{ segment.text.trim() }}</pre>
            </details>
          </template>
          <img v-if="item.message.imageSrc" :src="item.message.imageSrc" class="mt-2 max-w-full rounded-md" alt="Generated image" />
        </div>
        <div v-if="item.message.role === 'assistant'" class="flex gap-1">
          <CustomButton :title="replaceSelectedText" text="" :icon="FileText" type="secondary" class="bg-surface! p-1.5! text-secondary!" :icon-size="12" @click="$emit('insert-message', item.message, 'replace')" />
          <CustomButton :title="appendToSelection" text="" :icon="Plus" type="secondary" class="bg-surface! p-1.5! text-secondary!" :icon-size="12" @click="$emit('insert-message', item.message, 'append')" />
          <CustomButton :title="copyToClipboard" text="" :icon="Copy" type="secondary" class="bg-surface! p-1.5! text-secondary!" :icon-size="12" @click="$emit('copy-message', item.message)" />
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { Copy, FileText, Plus, Sparkles } from 'lucide-vue-next'
import { ref } from 'vue'

import CustomButton from '@/components/CustomButton.vue'
import type { DisplayMessage, RenderSegment } from '@/types/chat'

defineProps<{
  history: DisplayMessage[]
  historyWithSegments: Array<{ key: string, message: DisplayMessage, segments: RenderSegment[] }>
  backendOnline: boolean
  emptyTitle: string
  emptySubtitle: string
  backendOnlineLabel: string
  backendOfflineLabel: string
  replaceSelectedText: string
  appendToSelection: string
  copyToClipboard: string
}>()

defineEmits<{
  (e: 'insert-message', message: DisplayMessage, type: 'replace' | 'append'): void
  (e: 'copy-message', message: DisplayMessage): void
}>()

const containerEl = ref<HTMLElement>()
defineExpose({ containerEl })
</script>
