<template>
  <div role="tabpanel" class="flex h-full w-full flex-col items-center gap-2 bg-bg-secondary p-1">
    <SettingCard>
      <SingleSelect
        v-model="localLanguage"
        :tight="false"
        :key-list="localLanguageOptions.map(item => item.value)"
        :title="t('localLanguageLabel')"
        :fronticon="false"
        :placeholder="
          localLanguageOptions.find(o => o.value === localLanguage)?.label || localLanguage
        "
      >
        <template #item="{ item }">
          {{ localLanguageOptions.find(o => o.value === item)?.label || item }}
        </template>
      </SingleSelect>
    </SettingCard>

    <SettingCard>
      <label class="flex cursor-pointer items-center justify-between gap-2">
        <div class="flex flex-col">
          <span class="text-sm font-semibold text-main">{{ t('darkModeLabel') }}</span>
          <span class="text-xs text-secondary">{{ t('darkModeDescription') }}</span>
        </div>
        <input
          v-model="darkMode"
          type="checkbox"
          class="h-4 w-4 cursor-pointer accent-accent"
          :aria-label="t('darkModeLabel')"
        />
      </label>
    </SettingCard>

    <SettingCard v-if="hostIsExcel">
      <SingleSelect
        v-model="excelFormulaLanguage"
        :tight="false"
        :key-list="excelFormulaLanguageOptions.map(item => item.value)"
        :title="t('excelFormulaLanguageLabel')"
        :fronticon="false"
        :placeholder="
          excelFormulaLanguageOptions.find(o => o.value === excelFormulaLanguage)?.label ||
          excelFormulaLanguage
        "
      >
        <template #item="{ item }">
          {{ excelFormulaLanguageOptions.find(o => o.value === item)?.label || item }}
        </template>
      </SingleSelect>
    </SettingCard>

    <SettingCard>
      <div class="grid grid-cols-1 gap-2 md:grid-cols-2">
        <CustomInput
          v-model="userFirstName"
          :title="t('userFirstNameLabel')"
          :placeholder="t('userFirstNamePlaceholder')"
        />
        <CustomInput
          v-model="userLastName"
          :title="t('userLastNameLabel')"
          :placeholder="t('userLastNamePlaceholder')"
        />
      </div>
    </SettingCard>

    <SettingCard>
      <SingleSelect
        v-model="userGender"
        :tight="false"
        :key-list="genderOptions.map(item => item.value)"
        :title="t('userGenderLabel')"
        :fronticon="false"
        :placeholder="
          genderOptions.find(o => o.value === userGender)?.label || t('userGenderUnspecified')
        "
      >
        <template #item="{ item }">
          {{ genderOptions.find(o => o.value === item)?.label || item }}
        </template>
      </SingleSelect>
    </SettingCard>

    <SettingCard>
      <CustomInput
        v-model.number="agentMaxIterations"
        :title="t('agentMaxIterationsLabel')"
        placeholder="25"
        input-type="number"
      />
    </SettingCard>

    <!-- Backend status -->
    <SettingCard>
      <div class="flex items-center justify-between">
        <span class="text-sm font-semibold text-secondary">{{ t('backendStatus') }}</span>
        <div
          class="flex items-center gap-1 rounded-md px-2 py-1 text-xs"
          :class="backendOnline ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'"
        >
          <div
            class="h-2 w-2 rounded-full"
            :class="backendOnline ? 'bg-green-500' : 'bg-red-500'"
          />
          {{ backendOnline ? t('backendOnline') : t('backendOffline') }}
        </div>
      </div>
    </SettingCard>

    <SettingCard>
      <div class="flex items-center justify-between gap-2">
        <span class="text-sm font-semibold text-secondary">{{ t('appVersion') }}</span>
        <span class="rounded-sm bg-bg-secondary px-1.5 py-0.5 text-xs text-secondary">{{
          appVersion
        }}</span>
      </div>
    </SettingCard>

    <SettingCard>
      <div class="flex items-center justify-between gap-2">
        <span class="text-sm font-semibold text-secondary">{{
          t('reportBugOrFeedback') || 'Report a Bug / Feedback'
        }}</span>
        <CustomButton
          type="secondary"
          class="max-w-[160px] shrink-0"
          :text="t('feedbackButtonText')"
          @click="emit('open-feedback')"
        />
      </div>
    </SettingCard>

    <!-- Available models (read-only) -->
    <SettingCard v-if="loadingModels || Object.keys(availableModels).length > 0">
      <div class="flex flex-col gap-2">
        <span class="text-sm font-semibold text-secondary">{{ t('configuredModels') }}</span>
        <template v-if="loadingModels">
          <div
            v-for="i in 3"
            :key="i"
            class="card-base flex items-center justify-between animate-pulse"
          >
            <div class="h-4 w-24 bg-surface/80 rounded"></div>
            <div class="flex items-center gap-1.5">
              <div class="h-4 w-16 bg-surface/80 rounded"></div>
              <div class="h-4 w-12 bg-accent/20 rounded"></div>
            </div>
          </div>
        </template>
        <template v-else>
          <div
            v-for="(info, tier) in availableModels"
            :key="tier"
            class="card-base flex items-center justify-between"
          >
            <span class="text-xs font-medium text-main">{{ info.label }}</span>
            <div class="flex items-center gap-1.5">
              <span class="rounded-sm bg-bg-secondary px-1.5 py-0.5 text-xs text-secondary">{{
                info.id
              }}</span>
              <span class="rounded-sm bg-accent/10 px-2 py-0.5 text-xs text-accent">{{
                tier
              }}</span>
            </div>
          </div>
        </template>
      </div>
    </SettingCard>
  </div>
</template>

<script setup lang="ts">
import { useStorage } from '@vueuse/core'
import { watch } from 'vue'
import { useI18n } from 'vue-i18n'

import type { ModelInfo } from '@/types'
import CustomButton from '@/components/CustomButton.vue'
import CustomInput from '@/components/CustomInput.vue'
import SettingCard from '@/components/SettingCard.vue'
import SingleSelect from '@/components/SingleSelect.vue'
import { i18n } from '@/i18n'
import { optionLists } from '@/utils/common'
import { localStorageKey } from '@/utils/enum'
import { isExcel } from '@/utils/hostDetection'

const props = defineProps<{
  backendOnline: boolean
  availableModels: Record<string, ModelInfo>
  appVersion: string
  loadingModels?: boolean
}>()

const emit = defineEmits<{
  (e: 'open-feedback'): void
}>()

const { t } = useI18n()

const hostIsExcel = isExcel()

const localLanguage = useStorage(localStorageKey.localLanguage, 'fr')
const darkMode = useStorage(localStorageKey.darkMode, false)
const excelFormulaLanguage = useStorage(localStorageKey.excelFormulaLanguage, 'en')
const userGender = useStorage(localStorageKey.userGender, 'unspecified')
const userFirstName = useStorage(localStorageKey.userFirstName, '')
const userLastName = useStorage(localStorageKey.userLastName, '')
const agentMaxIterations = useStorage(localStorageKey.agentMaxIterations, 25)

const AGENT_MAX_ITERATIONS_MIN = 1
const AGENT_MAX_ITERATIONS_MAX = 100

function sanitizeAgentMaxIterations(value: unknown): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed)) return 25
  const normalized = Math.trunc(parsed)
  return Math.min(AGENT_MAX_ITERATIONS_MAX, Math.max(AGENT_MAX_ITERATIONS_MIN, normalized))
}

watch(
  agentMaxIterations,
  value => {
    const sanitized = sanitizeAgentMaxIterations(value)
    if (sanitized !== value) {
      agentMaxIterations.value = sanitized
    }
  },
  { immediate: true },
)

watch(localLanguage, val => {
  i18n.global.locale.value = val as 'en' | 'fr'
})

const localLanguageOptions = optionLists.localLanguageList

const excelFormulaLanguageOptions = [
  { label: t('excelFormulaLanguageEnglish'), value: 'en' },
  { label: t('excelFormulaLanguageFrench'), value: 'fr' },
]
const genderOptions = [
  { label: t('userGenderUnspecified'), value: 'unspecified' },
  { label: t('userGenderFemale'), value: 'female' },
  { label: t('userGenderMale'), value: 'male' },
  { label: t('userGenderNonBinary'), value: 'nonbinary' },
]
</script>
