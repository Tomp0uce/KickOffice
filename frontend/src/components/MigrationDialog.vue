<template>
  <div
    class="fixed inset-0 z-50 flex items-end justify-center bg-black/50 p-2 sm:items-center sm:p-4"
    aria-modal="true"
    role="dialog"
  >
    <div class="card-base flex w-full max-w-sm flex-col gap-4 bg-surface p-4">
      <div class="flex items-start gap-3">
        <div class="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-accent/10">
          <Zap class="text-accent" :size="16" />
        </div>
        <div>
          <h3 class="text-sm font-semibold text-main">
            {{ t('migrationTitle') || 'Vos prompts deviennent des Skills' }}
          </h3>
          <p class="mt-1 text-xs leading-normal text-secondary">
            {{
              t('migrationSubtitle', { count: promptCount }) ||
              `${promptCount} prompt(s) personnalisé(s) trouvé(s). Voulez-vous les convertir en Skills pour continuer à les utiliser ?`
            }}
          </p>
        </div>
      </div>

      <div class="flex gap-2">
        <CustomButton
          type="secondary"
          class="flex-1"
          :text="t('ignore') || 'Ignorer'"
          @click="$emit('dismiss')"
        />
        <CustomButton
          type="primary"
          class="flex-1"
          :text="t('convertToSkills') || 'Convertir en Skills'"
          @click="$emit('convert')"
        />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { Zap } from 'lucide-vue-next';
import { useI18n } from 'vue-i18n';
import CustomButton from '@/components/CustomButton.vue';

const { t } = useI18n();

defineProps<{ promptCount: number }>();
defineEmits<{
  (e: 'convert'): void;
  (e: 'dismiss'): void;
}>();
</script>
