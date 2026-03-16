<template>
  <!-- Auth Error Indicator (UX-M3) -->
  <div
    v-if="backendOnline && Object.keys(availableModels).length === 0 && !loading"
    class="flex flex-col items-center justify-center bg-warning/10 py-2 px-3 rounded-md border border-warning/30 shadow-xs mx-4 mt-2 mb-2 animate-in fade-in"
  >
    <span class="text-xs text-warning-700 dark:text-warning font-medium text-center mb-1">
      {{ t('authErrorBanner') }}
    </span>
    <button
      class="text-[11px] underline text-accent hover:text-accent-hover transition-colors"
      @click="goToSettings"
    >
      {{ t('goToSettings') }}
    </button>
  </div>
</template>

<script lang="ts" setup>
import { useHomePageContext } from '@/composables/useHomePageContext';

const { backendOnline, availableModels, loading, goToSettings, t } = useHomePageContext();
</script>
