<template>
  <div
    v-if="hasStats || modelName"
    class="flex items-center justify-between border-t border-border-secondary bg-bg-secondary px-2 py-1 text-[10px] text-secondary"
    style="font-family: monospace"
  >
    <!-- Token counts -->
    <div class="flex items-center gap-2">
      <span v-if="hasStats" :title="t('stats.inputTokens', { count: sessionStats.inputTokens })">
        ↑{{ formatTokens(sessionStats.inputTokens) }}
      </span>
      <span v-if="hasStats" :title="t('stats.outputTokens', { count: sessionStats.outputTokens })">
        ↓{{ formatTokens(sessionStats.outputTokens) }}
      </span>
      <div
        v-if="(contextWindowTokens ?? 0) > 0 && sessionStats.inputTokens > 0"
        class="ml-1 flex items-center gap-1 w-16"
        :title="
          contextPctNum >= 80
            ? t('stats.contextWarning', { pct: contextPct })
            : t('stats.contextUsage', {
                used: sessionStats.inputTokens,
                total: contextWindowTokens,
                pct: contextPct,
              })
        "
      >
        <div class="h-1.5 flex-1 bg-border rounded-full overflow-hidden">
          <div
            class="h-full transition-all duration-medium"
            :class="contextBarColor"
            :style="{ width: contextPctClamped + '%' }"
          ></div>
        </div>
        <span class="text-[10px]" :class="contextTextColor">{{ contextPct }}%</span>
      </div>
    </div>

    <!-- Model selector (custom, opens upward) + model name -->
    <div
      v-if="availableModels && selectedModelTier !== undefined"
      class="relative flex items-center gap-1.5 ml-2 min-w-0"
      @click.stop
    >
      <!-- Trigger button -->
      <button
        type="button"
        class="flex items-center gap-0.5 h-5 cursor-pointer rounded border border-border bg-surface px-1 text-[10px] text-secondary hover:border-accent focus:outline-none focus:ring-1 focus:ring-primary/50"
        :aria-label="t('stats.modelSelector', 'Select model')"
        :aria-haspopup="true"
        :aria-expanded="dropdownOpen"
        style="font-family: inherit"
        @click="dropdownOpen = !dropdownOpen"
      >
        {{ currentLabel }}
        <svg
          class="w-2.5 h-2.5 opacity-60 transition-transform"
          :class="{ 'rotate-180': dropdownOpen }"
          viewBox="0 0 10 6"
          fill="none"
          stroke="currentColor"
          stroke-width="1.5"
        >
          <path d="M1 1l4 4 4-4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>

      <!-- Dropdown list — opens UPWARD -->
      <Transition
        enter-active-class="transition-all duration-fast ease-apple"
        leave-active-class="transition-all duration-fast ease-apple"
        enter-from-class="opacity-0 translate-y-1"
        leave-to-class="opacity-0 translate-y-1"
      >
        <ul
          v-if="dropdownOpen"
          class="absolute right-0 bottom-full mb-1 z-50 min-w-full rounded border border-border bg-surface shadow-md text-[10px] text-secondary overflow-hidden"
          style="font-family: inherit"
        >
          <li
            v-for="(info, tier) in availableModels"
            :key="tier"
            class="cursor-pointer px-2 py-1 hover:bg-accent/10"
            :class="{ 'font-semibold text-accent': tier === selectedModelTier }"
            @click="selectTier(String(tier))"
          >
            {{ info.label }}
          </li>
        </ul>
      </Transition>

      <span v-if="modelName" class="truncate text-secondary/80">{{ modelName }}</span>
    </div>
    <div v-else-if="modelName" class="truncate ml-2 text-secondary/80">
      {{ modelName }}
    </div>
  </div>

  <!-- Click-outside overlay to close dropdown -->
  <div v-if="dropdownOpen" class="fixed inset-0 z-40" @click="dropdownOpen = false" />
</template>

<script lang="ts" setup>
import type { ModelInfo } from '@/types';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

const { t } = useI18n();

interface TokenStats {
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  lastCallInputTokens?: number;
}

const props = defineProps<{
  sessionStats: TokenStats;
  modelName?: string;
  contextWindowTokens?: number;
  currentAction?: string;
  loading?: boolean;
  availableModels?: Record<string, ModelInfo>;
  selectedModelTier?: string;
}>();

const emit = defineEmits<{
  (e: 'update:selectedModelTier', value: string): void;
}>();

const dropdownOpen = ref(false);

const currentLabel = computed(() => {
  if (!props.availableModels || !props.selectedModelTier) return '';
  return props.availableModels[props.selectedModelTier]?.label ?? props.selectedModelTier;
});

const selectTier = (tier: string) => {
  emit('update:selectedModelTier', tier);
  dropdownOpen.value = false;
};

const hasStats = computed(
  () => props.sessionStats.inputTokens > 0 || props.sessionStats.outputTokens > 0,
);

const contextPctNum = computed(() => {
  if (!props.contextWindowTokens || props.contextWindowTokens === 0) return 0;
  // Use the most recent call's input tokens (not cumulative) so the bar stays meaningful
  const tokensForBar = props.sessionStats.lastCallInputTokens ?? props.sessionStats.inputTokens;
  return (tokensForBar / props.contextWindowTokens) * 100;
});

const contextPct = computed(() => contextPctNum.value.toFixed(1));
const contextPctClamped = computed(() => Math.min(contextPctNum.value, 100).toFixed(1));

const contextBarColor = computed(() => {
  const pct = contextPctNum.value;
  if (pct >= 90) return 'bg-danger';
  if (pct >= 70) return 'bg-warning';
  return 'bg-success';
});

const contextTextColor = computed(() => {
  const pct = contextPctNum.value;
  if (pct >= 90) return 'text-danger';
  if (pct >= 70) return 'text-warning';
  return '';
});

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}k`;
  return n.toString();
}
</script>
