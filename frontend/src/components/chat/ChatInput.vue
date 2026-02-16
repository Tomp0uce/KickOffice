<template>
  <div class="flex flex-col gap-1 rounded-md">
    <div class="flex items-center justify-between gap-2 overflow-hidden">
      <div class="flex min-w-0 flex-1 items-center gap-2 overflow-hidden">
        <span class="shrink-0 text-xs font-medium text-secondary">Type de tâche :</span>
        <select :value="selectedModelTier" class="h-7 max-w-full min-w-0 cursor-pointer rounded-md border border-border bg-surface p-1 text-xs text-secondary hover:border-accent focus:outline-none" @change="$emit('update:selectedModelTier', ($event.target as HTMLSelectElement).value)">
          <option v-for="(info, tier) in availableModels" :key="tier" :value="tier">
            {{ info.label }}
          </option>
        </select>
      </div>
    </div>
    <div class="flex min-w-12 items-center gap-2 rounded-md border border-border bg-surface p-2 focus-within:border-accent">
      <textarea
        ref="textareaEl"
        :value="userInput"
        class="placeholder:text-secondary block max-h-30 flex-1 resize-none overflow-y-auto border-none bg-transparent py-2 text-xs leading-normal text-main outline-none placeholder:text-xs"
        :placeholder="inputPlaceholder"
        rows="1"
        @keydown.enter.exact.prevent="$emit('send')"
        @input="$emit('update:userInput', ($event.target as HTMLTextAreaElement).value); $emit('input')"
      />
      <button
        v-if="loading"
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-danger text-white"
        title="Stop"
        aria-label="Stop"
        @click="$emit('stop')"
      >
        <Square :size="18" />
      </button>
      <button
        v-else
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-accent text-white disabled:cursor-not-allowed disabled:bg-accent/50"
        title="Send"
        :disabled="!userInput.trim() || !backendOnline"
        aria-label="Send"
        @click="$emit('send')"
      >
        <Send :size="18" />
      </button>
    </div>
    <div class="flex justify-center gap-3 px-1">
      <label v-if="showWordFormatting" class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
        <input :checked="useWordFormatting" type="checkbox" @change="$emit('update:useWordFormatting', ($event.target as HTMLInputElement).checked)" />
        <span>{{ useWordFormattingLabel }}</span>
      </label>
      <label class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
        <input :checked="useSelectedText" type="checkbox" @change="$emit('update:useSelectedText', ($event.target as HTMLInputElement).checked)" />
        <span>{{ includeSelectionLabel }}</span>
      </label>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { Send, Square } from 'lucide-vue-next'
import { ref } from 'vue'

defineProps<{
  availableModels: Record<string, ModelInfo>
  selectedModelTier: string
  userInput: string
  inputPlaceholder: string
  loading: boolean
  backendOnline: boolean
  showWordFormatting: boolean
  useWordFormatting: boolean
  useSelectedText: boolean
  useWordFormattingLabel: string
  includeSelectionLabel: string
}>()

defineEmits<{
  (e: 'update:selectedModelTier', value: string): void
  (e: 'update:userInput', value: string): void
  (e: 'update:useWordFormatting', value: boolean): void
  (e: 'update:useSelectedText', value: boolean): void
  (e: 'send'): void
  (e: 'stop'): void
  (e: 'input'): void
}>()

const textareaEl = ref<HTMLTextAreaElement>()
defineExpose({ textareaEl })
</script>
