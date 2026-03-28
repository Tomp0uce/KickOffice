<template>
  <!-- Delete session confirm dialog -->
  <div
    v-if="isDeleteConfirmVisible"
    class="absolute inset-x-4 top-14 z-50 flex flex-col gap-3 rounded-md border border-border-secondary bg-bg-tertiary p-4 shadow-lg animate-in fade-in slide-in-from-top-4"
  >
    <p class="text-[13px] font-medium leading-tight text-main">
      {{ t('deleteSessionConfirm') }}
    </p>
    <div class="mt-1 flex justify-end gap-2">
      <button
        class="rounded-md border border-border-secondary bg-bg-secondary px-3 py-1.5 text-xs font-medium text-main transition-colors duration-fast cursor-pointer hover:bg-bg-tertiary"
        @click="$emit('cancel-delete')"
      >
        {{ t('cancel') }}
      </button>
      <button
        class="rounded-md bg-danger px-3 py-1.5 text-xs font-medium text-white transition-colors cursor-pointer hover:bg-danger/80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-danger/50 focus-visible:ring-offset-2 focus-visible:ring-offset-bg-tertiary"
        @click="$emit('confirm-delete')"
      >
        {{ t('confirm') }}
      </button>
    </div>
  </div>

  <!-- New chat confirm dialog -->
  <div
    v-if="isNewChatConfirmVisible"
    class="absolute inset-x-4 top-14 z-50 flex flex-col gap-3 rounded-md border border-border-secondary bg-bg-tertiary p-4 shadow-lg animate-in fade-in slide-in-from-top-4"
  >
    <p class="text-[13px] font-medium leading-tight text-main">
      {{ t('newChatConfirm') }}
    </p>
    <div class="mt-1 flex justify-end gap-2">
      <button
        class="rounded-md border border-border-secondary bg-bg-secondary px-3 py-1.5 text-xs font-medium text-main transition-colors duration-fast cursor-pointer hover:bg-bg-tertiary"
        @click="$emit('cancel-new-chat')"
      >
        {{ t('cancel') }}
      </button>
      <button
        class="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-white transition-colors duration-fast cursor-pointer hover:bg-primary/90"
        @click="$emit('confirm-new-chat')"
      >
        {{ t('confirm') }}
      </button>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { useHomePageContext } from '@/composables/useHomePageContext';

defineProps<{
  isDeleteConfirmVisible: boolean;
  isNewChatConfirmVisible: boolean;
}>();

defineEmits<{
  (e: 'cancel-delete'): void;
  (e: 'confirm-delete'): void;
  (e: 'cancel-new-chat'): void;
  (e: 'confirm-new-chat'): void;
}>();

const { t } = useHomePageContext();
</script>
