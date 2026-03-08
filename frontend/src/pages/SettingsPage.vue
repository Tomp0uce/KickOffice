<template>
  <div class="relative flex h-full w-full items-center justify-center bg-bg-secondary">
    <div
      class="relative z-1 flex h-full w-full flex-col items-center justify-start gap-2 rounded-xl border-none p-2"
    >
      <div
        class="flex w-full items-center justify-between gap-1 overflow-visible rounded-2xl border border-border-secondary p-0 shadow-sm"
      >
        <div class="flex flex-wrap items-center gap-4 p-1">
          <CustomButton
            :icon="ArrowLeft"
            type="secondary"
            class="border-none p-1!"
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
          class="flex-1 rounded-sm border-none! p-1!"
          @click="currentTab = tab.id"
        />
      </div>

      <!-- Main Content -->
      <div class="w-full flex-1 overflow-hidden">
        <div class="no-scrollbar h-full w-full overflow-auto rounded-md shadow-md">
          <AccountTab v-if="currentTab === 'account'" ref="accountTabRef" />

          <GeneralTab
            v-if="currentTab === 'general'"
            :backend-online="backendOnline"
            :available-models="availableModels"
            :app-version="appVersion"
            @open-feedback="showFeedbackDialog = true"
          />

          <PromptsTab v-if="currentTab === 'prompts'" />

          <BuiltinPromptsTab v-if="currentTab === 'builtinPrompts'" />

          <ToolsTab v-if="currentTab === 'tools'" />
        </div>
      </div>
    </div>
    <FeedbackDialog v-if="showFeedbackDialog" @close="showFeedbackDialog = false" />
  </div>
</template>

<script lang="ts" setup>
import type { ModelInfo } from '@/types'
import { ArrowLeft, Globe, KeyRound, MessageSquare, Settings, Wrench } from 'lucide-vue-next'
import { onBeforeMount, onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { useRouter } from 'vue-router'

import { fetchModels, healthCheck } from '@/api/backend'
import CustomButton from '@/components/CustomButton.vue'
import AccountTab from '@/components/settings/AccountTab.vue'
import BuiltinPromptsTab from '@/components/settings/BuiltinPromptsTab.vue'
import FeedbackDialog from '@/components/settings/FeedbackDialog.vue'
import GeneralTab from '@/components/settings/GeneralTab.vue'
import PromptsTab from '@/components/settings/PromptsTab.vue'
import ToolsTab from '@/components/settings/ToolsTab.vue'
import { migrateFromPlaintext, getUserKey, getUserEmail } from '@/utils/credentialStorage'

const { t } = useI18n()
const router = useRouter()
const appVersion = __APP_VERSION__

const currentTab = ref('account')
const showFeedbackDialog = ref(false)

// Backend state shared with GeneralTab
const backendOnline = ref(false)
const availableModels = ref<Record<string, ModelInfo>>({})

// Ref to AccountTab to set credentials after migration
const accountTabRef = ref<InstanceType<typeof AccountTab> | null>(null)

const tabs = [
  { id: 'account', label: 'account', defaultLabel: 'Account', icon: KeyRound },
  { id: 'general', label: 'general', defaultLabel: 'General', icon: Globe },
  { id: 'prompts', label: 'prompts', defaultLabel: 'Prompts', icon: MessageSquare },
  {
    id: 'builtinPrompts',
    label: 'builtinPrompts',
    defaultLabel: 'Built-in Prompts',
    icon: Settings,
  },
  { id: 'tools', label: 'tools', defaultLabel: 'Tools', icon: Wrench },
]

async function checkBackend() {
  backendOnline.value = await healthCheck()
  if (backendOnline.value) {
    try {
      availableModels.value = await fetchModels()
    } catch {
      console.error('Failed to fetch models')
    }
  }
}

function backToHome() {
  router.push('/')
}

onBeforeMount(() => {
  checkBackend()
})

onMounted(async () => {
  // Migrate old plaintext credentials if needed
  await migrateFromPlaintext()
  // Load credentials and push them into AccountTab
  const key = await getUserKey()
  const email = await getUserEmail()
  accountTabRef.value?.setCredentials(key, email)
})
</script>
