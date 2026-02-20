<template>
  <div
    class="flex w-full items-center justify-center gap-2 overflow-hidden rounded-md"
  >
    <CustomButton
      v-for="action in quickActions"
      :key="action.key"
      :title="$t(action.key + '_tooltip')"
      text=""
      :icon="action.icon"
      type="secondary"
      :icon-size="16"
      class="shrink-0! bg-surface! p-1.5!"
      :disabled="loading"
      :aria-label="action.label"
      @click="$emit('apply-action', action.key)"
    />
    <SingleSelect
      :model-value="selectedPromptId"
      :key-list="savedPrompts.map((prompt) => prompt.id)"
      :placeholder="selectPromptTitle"
      title=""
      :fronticon="false"
      class="max-w-xs! flex-1! bg-surface! text-xs!"
      @update:model-value="$emit('update:selectedPromptId', $event as string)"
      @change="$emit('load-prompt')"
    >
      <template #item="{ item }">
        {{ savedPrompts.find((prompt) => prompt.id === item)?.name || item }}
      </template>
    </SingleSelect>
  </div>
</template>

<script lang="ts" setup>
import CustomButton from "@/components/CustomButton.vue";
import SingleSelect from "@/components/SingleSelect.vue";
import type { QuickAction } from "@/types/chat";
import type { SavedPrompt } from "@/utils/savedPrompts";

defineProps<{
  quickActions: QuickAction[];
  loading: boolean;
  savedPrompts: SavedPrompt[];
  selectedPromptId: string;
  selectPromptTitle: string;
}>();

defineEmits<{
  (e: "apply-action", key: string): void;
  (e: "update:selectedPromptId", value: string): void;
  (e: "load-prompt"): void;
}>();
</script>
