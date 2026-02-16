<template>
  <div class="flex flex-col gap-1 rounded-md">
    <div class="flex items-center justify-between gap-2 overflow-hidden">
      <div class="flex min-w-0 flex-1 items-center gap-2 overflow-hidden">
        <label :id="modelTierLabelId" :for="modelTierSelectId" class="shrink-0 text-xs font-medium text-secondary">{{ taskTypeLabel }}</label>
        <select :id="modelTierSelectId" :value="selectedModelTier" :aria-labelledby="modelTierLabelId" class="h-7 max-w-full min-w-0 cursor-pointer rounded-md border border-border bg-surface p-1 text-xs text-secondary hover:border-accent focus:outline-none" @change="$emit('update:selectedModelTier', ($event.target as HTMLSelectElement).value)">
          <option v-for="(info, tier) in availableModels" :key="tier" :value="tier">
            {{ info.label }}
          </option>
        </select>
      </div>
    </div>
    <div class="flex min-w-12 items-center gap-2 rounded-md border border-border bg-surface p-2 focus-within:border-accent">
      <textarea
        ref="textareaEl"
        :value="modelValue"
        class="placeholder:text-secondary block max-h-30 flex-1 resize-none overflow-y-auto border-none bg-transparent py-2 text-xs leading-normal text-main outline-none placeholder:text-xs"
        :placeholder="inputPlaceholder"
        rows="1"
        @keydown.enter.exact.prevent="$emit('submit')"
        @input="$emit('update:modelValue', ($event.target as HTMLTextAreaElement).value); $emit('input')"
      />
      <button
        v-if="loading"
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-danger text-white"
        :title="stopLabel"
        :aria-label="stopLabel"
        @click="$emit('stop')"
      >
        <Square :size="18" />
      </button>
      <button
        v-else
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-accent text-white disabled:cursor-not-allowed disabled:bg-accent/50"
        :title="sendLabel"
        :disabled="!modelValue.trim() || !backendOnline"
        :aria-label="sendLabel"
        @click="$emit('submit')"
      >
        <Send :size="18" />
      </button>
    </div>
    <div class="flex justify-center gap-3 px-1">
      <label v-if="showWordFormatting" :for="wordFormattingCheckboxId" class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
        <input :id="wordFormattingCheckboxId" :checked="useWordFormatting" :aria-label="useWordFormattingLabel" type="checkbox" @change="$emit('update:useWordFormatting', ($event.target as HTMLInputElement).checked)" />
        <span>{{ useWordFormattingLabel }}</span>
      </label>
      <label :for="selectedTextCheckboxId" class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
        <input :id="selectedTextCheckboxId" :checked="useSelectedText" :aria-label="includeSelectionLabel" type="checkbox" @change="$emit('update:useSelectedText', ($event.target as HTMLInputElement).checked)" />
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
  modelValue: string
  inputPlaceholder: string
  loading: boolean
  backendOnline: boolean
  showWordFormatting: boolean
  useWordFormatting: boolean
  useSelectedText: boolean
  useWordFormattingLabel: string
  includeSelectionLabel: string
  taskTypeLabel: string
  sendLabel: string
  stopLabel: string
}>()

defineEmits<{
  (e: 'update:selectedModelTier', value: string): void
  (e: 'update:modelValue', value: string): void
  (e: 'update:useWordFormatting', value: boolean): void
  (e: 'update:useSelectedText', value: boolean): void
  (e: 'submit'): void
  (e: 'stop'): void
  (e: 'input'): void
}>()

const textareaEl = ref<HTMLTextAreaElement>()
const modelTierSelectId = 'chat-model-tier-select'
const modelTierLabelId = 'chat-model-tier-label'
const wordFormattingCheckboxId = 'chat-word-formatting-checkbox'
const selectedTextCheckboxId = 'chat-selected-text-checkbox'
defineExpose({ textareaEl })
</script>
