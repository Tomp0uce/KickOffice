<template>
  <div class="flex w-full flex-wrap items-center justify-center gap-2 rounded-md">
    <!-- User Skills dropdown + create button -->
    <div class="flex flex-1 items-center gap-1 min-w-0 max-w-full">
      <SingleSelect
        :key-list="userSkillsForHost.map(s => s.id)"
        :placeholder="t('mySkills') || 'Mes skills...'"
        :title="t('mySkills') || 'My skills'"
        :fronticon="false"
        class="flex-1! bg-surface! text-xs!"
        @update:model-value="id => $emit('execute-user-skill', String(id))"
      >
        <template #item="{ item }">
          {{ userSkillsForHost.find(s => s.id === item)?.name || item }}
        </template>
      </SingleSelect>
      <CustomButton
        :icon="Plus"
        text=""
        :title="t('createSkill') || 'Créer un skill'"
        type="secondary"
        :icon-size="14"
        class="shrink-0! bg-surface! p-1.5!"
        @click="$emit('open-skill-creator')"
      />
    </div>

    <!-- Built-in quick action buttons (unchanged) -->
    <CustomButton
      v-for="action in quickActions"
      :key="action.key"
      :title="$t(action.tooltipKey || action.key + '_tooltip')"
      text=""
      :icon="action.icon"
      type="secondary"
      :icon-size="16"
      class="shrink-0! bg-surface! p-1.5!"
      :disabled="loading"
      :aria-label="action.label"
      @click="$emit('apply-action', action.key)"
    />
  </div>
</template>

<script lang="ts" setup>
import { Plus } from 'lucide-vue-next';
import { useI18n } from 'vue-i18n';
import CustomButton from '@/components/CustomButton.vue';
import SingleSelect from '@/components/SingleSelect.vue';
import type { QuickAction } from '@/types/chat';
import type { UserSkill } from '@/types/userSkill';

const { t } = useI18n();

defineProps<{
  quickActions: QuickAction[];
  loading: boolean;
  userSkillsForHost: UserSkill[];
}>();

defineEmits<{
  (e: 'apply-action', key: string): void;
  (e: 'execute-user-skill', id: string): void;
  (e: 'open-skill-creator'): void;
}>();
</script>
