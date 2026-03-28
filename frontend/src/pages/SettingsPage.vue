<template>
  <div class="relative flex h-full w-full items-center justify-center bg-bg-secondary">
    <div
      class="relative z-1 flex h-full w-full flex-col items-center justify-start gap-2 rounded-xl border-none p-2"
    >
      <div
        class="flex w-full items-center justify-between gap-1 overflow-visible rounded-2xl border border-border-secondary p-0 shadow-sm"
      >
        <div class="flex flex-wrap items-center gap-2 p-1">
          <CustomButton
            :icon="ArrowLeft"
            type="secondary"
            class="border-none p-1.5!"
            text=""
            :title="t('back')"
            @click="backToHome"
          />
        </div>
        <div class="flex-1">
          <h2 class="text-sm font-semibold text-main">
            {{ t('settings') }}
          </h2>
        </div>
      </div>

      <!-- Tab Navigation -->
      <div
        role="tablist"
        class="flex w-full justify-between rounded-2xl border border-border-secondary p-0"
      >
        <CustomButton
          v-for="tab in tabs"
          :key="tab.id"
          role="tab"
          :aria-selected="currentTab === tab.id"
          text=""
          :type="currentTab === tab.id ? 'primary' : 'secondary'"
          :title="t(tab.label) || tab.defaultLabel"
          :icon="tab.icon"
          :icon-size="16"
          class="flex-1 rounded-sm border-none! p-1.5!"
          @click="currentTab = tab.id"
        />
      </div>

      <!-- Main Content -->
      <div class="w-full flex-1 overflow-hidden">
        <div class="no-scrollbar h-full w-full overflow-auto rounded-md shadow-md">
          <div v-if="resolvingConfig" class="flex flex-col gap-2 p-1">
            <div
              v-for="i in 3"
              :key="i"
              class="card-base h-16 w-full animate-pulse bg-surface/50"
            ></div>
          </div>
          <template v-else>
            <AccountTab v-if="currentTab === 'account'" ref="accountTabRef" />

            <GeneralTab
              v-if="currentTab === 'general'"
              :backend-online="backendOnline"
              :available-models="availableModels"
              :app-version="appVersion"
              :loading-models="loadingModels"
              @open-feedback="showFeedbackDialog = true"
            />

            <SkillLibraryTab v-if="currentTab === 'skills'" @open-creator="() => {}" />

            <ToolsTab v-if="currentTab === 'tools'" />
          </template>
        </div>
      </div>
    </div>
    <FeedbackDialog v-if="showFeedbackDialog" @close="showFeedbackDialog = false" />
  </div>
</template>

<script lang="ts" setup>
import type { ModelInfo } from '@/types';
import { ArrowLeft, Globe, KeyRound, Wrench, Zap } from 'lucide-vue-next';
import { onBeforeMount, onMounted, ref } from 'vue';
import { useI18n } from 'vue-i18n';
import { useRouter } from 'vue-router';

import { fetchModels, healthCheck } from '@/api/backend';
import { logService } from '@/utils/logger';
import CustomButton from '@/components/CustomButton.vue';
import AccountTab from '@/components/settings/AccountTab.vue';
import FeedbackDialog from '@/components/settings/FeedbackDialog.vue';
import GeneralTab from '@/components/settings/GeneralTab.vue';
import SkillLibraryTab from '@/components/settings/SkillLibraryTab.vue';
import ToolsTab from '@/components/settings/ToolsTab.vue';
import { migrateFromPlaintext, getUserKey, getUserEmail } from '@/utils/credentialStorage';

const { t } = useI18n();
const router = useRouter();
const appVersion = __APP_VERSION__;

const currentTab = ref('account');
const showFeedbackDialog = ref(false);

// Backend state shared with GeneralTab
const backendOnline = ref(false);
const availableModels = ref<Record<string, ModelInfo>>({});
const loadingModels = ref(true);
const resolvingConfig = ref(true);

// Ref to AccountTab to set credentials after migration
const accountTabRef = ref<InstanceType<typeof AccountTab> | null>(null);

const tabs = [
  { id: 'account', label: 'account', defaultLabel: 'Account', icon: KeyRound },
  { id: 'general', label: 'general', defaultLabel: 'General', icon: Globe },
  { id: 'skills', label: 'skills', defaultLabel: 'Skills', icon: Zap },
  { id: 'tools', label: 'tools', defaultLabel: 'Tools', icon: Wrench },
];

async function checkBackend() {
  backendOnline.value = await healthCheck();
  if (backendOnline.value) {
    try {
      loadingModels.value = true;
      availableModels.value = await fetchModels();
    } catch {
      logService.error('Failed to fetch models');
    } finally {
      loadingModels.value = false;
    }
  } else {
    loadingModels.value = false;
  }
}

function backToHome() {
  router.push('/');
}

onBeforeMount(() => {
  checkBackend();
});

onMounted(async () => {
  // Simulate config resolution time for smooth UX
  await new Promise(resolve => setTimeout(resolve, 600));
  // Migrate old plaintext credentials if needed
  await migrateFromPlaintext();
  // Load credentials and push them into AccountTab
  const key = await getUserKey();
  const email = await getUserEmail();
  accountTabRef.value?.setCredentials(key, email);
  resolvingConfig.value = false;
});
</script>
