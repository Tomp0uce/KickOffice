<template>
  <!-- Persistent Offline Indicator — only shown after first check to avoid false negative flash -->
  <div
    v-if="backendChecked && !backendOnline"
    class="flex items-center justify-center bg-danger/10 py-1.5 px-3 rounded-md border border-danger/20 shadow-xs mx-4 mt-2"
  >
    <span class="text-xs text-danger font-medium flex items-center gap-2">
      <span class="relative flex h-2 w-2">
        <span
          class="animate-ping absolute inline-flex h-full w-full rounded-full bg-danger/75 opacity-75"
        ></span>
        <span class="relative inline-flex rounded-full h-2 w-2 bg-danger"></span>
      </span>
      {{ t('backendOffline') }}
    </span>
  </div>
</template>

<script lang="ts" setup>
import { useHomePageContext } from '@/composables/useHomePageContext';

const { backendChecked, backendOnline, t } = useHomePageContext();
</script>
