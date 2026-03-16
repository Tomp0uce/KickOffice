<template>
  <div class="mt-2 mb-1 rounded-sm border border-border-secondary bg-bg-secondary overflow-hidden">
    <button
      type="button"
      class="w-full flex items-center gap-1.5 px-2 py-1 text-[10px] tracking-wider text-secondary hover:bg-bg-tertiary transition-colors text-left"
      @click="expanded = !expanded"
    >
      <component :is="expanded ? ChevronDown : ChevronRight" :size="10" />
      <Wrench :size="10" />
      <span class="flex-1 font-medium truncate">{{ props.toolCall.name }}</span>
      <component :is="statusIcon" :size="10" :class="statusIconClass" class="shrink-0" />
    </button>
    <div v-if="expanded" class="border-t border-border-secondary">
      <div class="px-2 py-1.5 text-xs">
        <div class="text-[10px] uppercase text-secondary mb-1">{{ t('toolCall.args') }}</div>
        <pre
          class="text-[10px] max-h-28 overflow-y-auto whitespace-pre-wrap break-words text-main/70 bg-bg-secondary rounded p-1"
          >{{ argsText }}</pre
        >
      </div>
      <div
        v-if="props.toolCall.result"
        class="px-2 py-1.5 text-xs border-t border-border-secondary"
      >
        <div
          class="text-[10px] uppercase mb-1"
          :class="props.toolCall.status === 'error' ? 'text-red-400' : 'text-secondary'"
        >
          {{ props.toolCall.status === 'error' ? t('toolCall.error') : t('toolCall.result') }}
        </div>
        <pre
          class="text-[10px] max-h-32 overflow-y-auto whitespace-pre-wrap break-words rounded p-1"
          :class="
            props.toolCall.status === 'error'
              ? 'text-red-400 bg-red-50 dark:bg-red-950/20'
              : 'text-main/70 bg-bg-secondary'
          "
          >{{ props.toolCall.result }}</pre
        >
      </div>
      <div v-if="props.toolCall.screenshotSrc" class="px-2 py-1.5 border-t border-border-secondary">
        <div class="text-[10px] uppercase text-secondary mb-1">screenshot</div>
        <img :src="props.toolCall.screenshotSrc" alt="screenshot" class="max-w-full rounded-sm" />
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { computed, ref } from 'vue';
import { ChevronDown, ChevronRight, CheckCircle2, Loader2, Wrench, XCircle } from 'lucide-vue-next';
import { useI18n } from 'vue-i18n';
import type { ToolCallPart } from '@/types/chat';

const { t } = useI18n();
const props = defineProps<{ toolCall: ToolCallPart }>();

const expanded = ref(false);

const argsText = computed(() => {
  try {
    return JSON.stringify(props.toolCall.args, null, 2);
  } catch {
    return String(props.toolCall.args);
  }
});

const statusIcon = computed(() => {
  switch (props.toolCall.status) {
    case 'running':
      return Loader2;
    case 'complete':
      return CheckCircle2;
    case 'error':
      return XCircle;
    default:
      return Loader2;
  }
});

const statusIconClass = computed(() => {
  switch (props.toolCall.status) {
    case 'running':
      return 'animate-spin text-accent';
    case 'complete':
      return 'text-green-500';
    case 'error':
      return 'text-red-500';
    default:
      return 'animate-spin text-secondary';
  }
});
</script>
