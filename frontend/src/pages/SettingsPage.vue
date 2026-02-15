<template>
  <div class="relative flex h-full w-full items-center justify-center bg-bg-secondary">
    <div class="relative z-1 flex h-full w-full flex-col items-center justify-start gap-2 rounded-xl border-none p-2">
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
            {{ $t('settings') || 'Settings' }}
          </h2>
        </div>
      </div>

      <!-- Tab Navigation -->
      <div class="flex w-full justify-between rounded-2xl border border-border-secondary p-0">
        <CustomButton
          v-for="tab in tabs"
          :key="tab.id"
          text=""
          :type="currentTab === tab.id ? 'primary' : 'secondary'"
          :title="$t(tab.label) || tab.defaultLabel"
          :icon="tab.icon"
          :icon-size="16"
          class="flex-1 rounded-sm border-none! p-1!"
          @click="currentTab = tab.id"
        />
      </div>

      <!-- Main Content -->
      <div class="w-full flex-1 overflow-hidden">
        <div class="no-scrollbar h-full w-full overflow-auto rounded-md shadow-md">
          <!-- General Settings -->
          <div
            v-show="currentTab === 'general'"
            class="flex h-full w-full flex-col items-center gap-2 bg-bg-secondary p-1"
          >
            <SettingCard>
              <SingleSelect
                v-model="localLanguage"
                :tight="false"
                :key-list="localLanguageOptions.map(item => item.value)"
                :title="$t('localLanguageLabel')"
                :fronticon="false"
                :placeholder="localLanguageOptions.find(o => o.value === localLanguage)?.label || localLanguage"
              >
                <template #item="{ item }">
                  {{ localLanguageOptions.find(o => o.value === item)?.label || item }}
                </template>
              </SingleSelect>
            </SettingCard>

            <SettingCard>
              <SingleSelect
                v-model="replyLanguage"
                :tight="false"
                :key-list="replyLanguageOptions.map(item => item.value)"
                :title="$t('replyLanguageLabel')"
                :fronticon="false"
                :placeholder="replyLanguageOptions.find(o => o.value === replyLanguage)?.label || replyLanguage"
              >
                <template #item="{ item }">
                  {{ replyLanguageOptions.find(o => o.value === item)?.label || item }}
                </template>
              </SingleSelect>
            </SettingCard>

            <SettingCard>
              <div class="grid grid-cols-1 gap-2 md:grid-cols-2">
                <CustomInput
                  v-model="userFirstName"
                  :title="$t('userFirstNameLabel')"
                  :placeholder="$t('userFirstNamePlaceholder')"
                />
                <CustomInput
                  v-model="userLastName"
                  :title="$t('userLastNameLabel')"
                  :placeholder="$t('userLastNamePlaceholder')"
                />
              </div>
            </SettingCard>

            <SettingCard>
              <SingleSelect
                v-model="userGender"
                :tight="false"
                :key-list="genderOptions.map(item => item.value)"
                :title="$t('userGenderLabel')"
                :fronticon="false"
                :placeholder="genderOptions.find(o => o.value === userGender)?.label || t('userGenderUnspecified')"
              >
                <template #item="{ item }">
                  {{ genderOptions.find(o => o.value === item)?.label || item }}
                </template>
              </SingleSelect>
            </SettingCard>

            <SettingCard>
              <CustomInput
                v-model.number="agentMaxIterations"
                :title="$t('agentMaxIterationsLabel')"
                placeholder="25"
                input-type="number"
              />
            </SettingCard>

            <!-- Backend status -->
            <SettingCard>
              <div class="flex items-center justify-between">
                <span class="text-sm font-semibold text-secondary">{{ $t('backendStatus') }}</span>
                <div
                  class="flex items-center gap-1 rounded-md px-2 py-1 text-xs"
                  :class="backendOnline ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'"
                >
                  <div
                    class="h-2 w-2 rounded-full"
                    :class="backendOnline ? 'bg-green-500' : 'bg-red-500'"
                  />
                  {{ backendOnline ? $t('backendOnline') : $t('backendOffline') }}
                </div>
              </div>
            </SettingCard>

            <!-- Available models (read-only) -->
            <SettingCard v-if="Object.keys(availableModels).length > 0">
              <div class="flex flex-col gap-2">
                <span class="text-sm font-semibold text-secondary">{{ $t('configuredModels') }}</span>
                <div
                  v-for="(info, tier) in availableModels"
                  :key="tier"
                  class="flex items-center justify-between rounded-md border border-border bg-surface p-2"
                >
                  <span class="text-xs font-medium text-main">{{ info.label }}</span>
                  <div class="flex items-center gap-1.5">
                    <span class="rounded-sm bg-bg-secondary px-1.5 py-0.5 text-xs text-secondary">{{ info.id }}</span>
                    <span class="rounded-sm bg-accent/10 px-2 py-0.5 text-xs text-accent">{{ tier }}</span>
                  </div>
                </div>
              </div>
            </SettingCard>
          </div>

          <!-- Prompts Settings -->
          <div
            v-show="currentTab === 'prompts'"
            class="flex w-full flex-1 flex-col items-center gap-2 bg-bg-secondary p-1"
          >
            <div
              class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
            >
              <div class="flex items-center justify-between">
                <h3 class="text-center text-sm font-semibold text-main">
                  {{ $t('savedPrompts') }}
                </h3>
                <CustomButton
                  :icon="Plus"
                  text=""
                  :title="t('addPrompt')"
                  class="p-1!"
                  type="secondary"
                  @click="addNewPrompt"
                />
              </div>

              <div
                v-for="prompt in savedPrompts"
                :key="prompt.id"
                class="rounded-md border border-border bg-surface p-3"
              >
                <div class="flex items-start justify-between">
                  <div class="flex flex-1 flex-wrap items-center gap-2">
                    <input
                      v-if="editingPromptId === prompt.id"
                      v-model="editingPrompt.name"
                      class="max-w-37.5 min-w-25 flex-1 rounded-md border border-border px-2 py-1 text-sm font-semibold text-secondary focus:border-accent focus:outline-none"
                      @blur="savePromptEdit"
                      @keyup.enter="savePromptEdit"
                    />
                    <span v-else class="text-sm font-semibold text-main">{{ prompt.name }}</span>
                  </div>
                  <div class="flex shrink-0 gap-1">
                    <CustomButton
                      type="secondary"
                      :title="t('edit')"
                      :icon="Edit2"
                      class="border-none! bg-surface! p-1.5!"
                      :icon-size="14"
                      text=""
                      @click="startEditPrompt(prompt)"
                    />
                    <CustomButton
                      v-if="savedPrompts.length > 1"
                      class="border-none! bg-surface! p-1.5!"
                      :title="t('delete')"
                      type="secondary"
                      :icon="Trash2"
                      text=""
                      :icon-size="14"
                      @click="deletePrompt(prompt.id)"
                    />
                  </div>
                </div>

                <div v-if="editingPromptId === prompt.id" class="mt-3 border-t border-t-border pt-3">
                  <label class="mb-1 block text-xs font-semibold text-secondary">{{ $t('systemPrompt') }}</label>
                  <textarea
                    v-model="editingPrompt.systemPrompt"
                    class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm leading-normal text-main transition-all duration-200 ease-apple focus:border-accent focus:outline-none"
                    rows="3"
                    :placeholder="$t('systemPromptPlaceholder')"
                  />

                  <label class="mb-1 block text-xs font-semibold text-secondary">{{ $t('userPrompt') }}</label>
                  <textarea
                    v-model="editingPrompt.userPrompt"
                    class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm leading-normal text-main transition-all duration-200 ease-apple focus:border-accent focus:outline-none"
                    rows="3"
                    :placeholder="$t('userPromptPlaceholder')"
                  />

                  <div class="mt-3 flex gap-2">
                    <CustomButton type="primary" class="flex-1" :text="t('save')" @click="savePromptEdit" />
                    <CustomButton type="secondary" class="flex-1" :text="t('cancel')" @click="cancelEdit" />
                  </div>
                </div>

                <div v-else class="mt-2">
                  <p class="overflow-hidden text-xs font-semibold text-ellipsis text-secondary">
                    {{ prompt.systemPrompt.substring(0, 100) }}{{ prompt.systemPrompt.length > 100 ? '...' : '' }}
                  </p>
                </div>
              </div>
            </div>
          </div>

          <!-- Built-in Prompts Settings -->
          <div
            v-show="currentTab === 'builtinPrompts'"
            class="flex w-full flex-1 items-center gap-2 overflow-hidden bg-bg-secondary p-1"
          >
            <div
              class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
            >
              <div class="rounded-md border border-border-secondary p-1 shadow-sm">
                <h3 class="text-center text-sm font-semibold text-accent/70">
                  {{ t('builtinPrompts') || 'Built-in Prompts' }}
                </h3>
              </div>
              <div class="rounded-md border border-border-secondary p-1 shadow-sm">
                <p class="text-xs leading-normal font-medium wrap-break-word text-secondary">
                  {{
                    t('builtinPromptsDescription', {
                      language: '${language}',
                      text: '${text}',
                    })
                  }}
                </p>
              </div>

              <div
                v-for="(promptConfig, key) in builtInPromptsData"
                :key="key"
                class="flex flex-col gap-2 rounded-md border border-border bg-surface p-2 hover:border-2 hover:border-accent"
              >
                <div class="flex flex-row items-start justify-between">
                  <div class="flex items-center gap-2">
                    <span class="text-sm font-semibold text-secondary">{{ t(hostIsExcel ? `excel${key.charAt(0).toUpperCase() + key.slice(1)}` : key) || key }}</span>
                  </div>
                  <div class="flex gap-1">
                    <CustomButton
                      :icon="editingBuiltinPromptKey === key ? Save : Edit2"
                      text=""
                      :title="editingBuiltinPromptKey === key ? t('save') : t('edit')"
                      class="border-none bg-surface! p-1.5!"
                      type="secondary"
                      :icon-size="14"
                      @click="toggleEditBuiltinPrompt(key)"
                    />
                    <CustomButton
                      v-if="isBuiltinPromptModified(key)"
                      :icon="RotateCcwIcon"
                      text=""
                      :title="t('reset')"
                      class="border-none bg-surface! p-1.5!"
                      type="secondary"
                      :icon-size="14"
                      @click="resetBuiltinPrompt(key)"
                    />
                  </div>
                </div>

                <div v-if="editingBuiltinPromptKey === key">
                  <label class="mt-2 block text-xs font-semibold text-secondary">{{ $t('systemPrompt') }}</label>
                  <textarea
                    v-model="editingBuiltinPrompt.system"
                    class="min-h-20 w-full rounded-md border border-border bg-bg-secondary p-2 text-xs text-main focus:border-accent focus:outline-none"
                    rows="3"
                  />
                  <label class="mt-2 block text-xs font-semibold text-secondary">{{ $t('userPrompt') }}</label>
                  <textarea
                    v-model="editingBuiltinPrompt.user"
                    class="min-h-20 w-full rounded-md border border-border bg-bg-secondary p-2 text-xs text-main focus:border-accent focus:outline-none"
                    rows="4"
                  />
                </div>

                <div v-else class="mt-2">
                  <p class="mb-2 text-xs font-semibold text-secondary">{{ $t('systemPrompt') }}:</p>
                  <p class="text-xs leading-normal wrap-break-word text-secondary">
                    {{ getSystemPromptPreview(promptConfig.system) }}
                  </p>
                  <p class="mt-2 mb-2 text-xs font-semibold text-secondary">{{ $t('userPrompt') }}:</p>
                  <p class="text-xs leading-normal wrap-break-word text-secondary">
                    {{ getUserPromptPreview(promptConfig.user) }}
                  </p>
                </div>
              </div>
            </div>
          </div>

          <!-- Tools Settings -->
          <div
            v-show="currentTab === 'tools'"
            class="w-full flex-1 items-center gap-2 overflow-hidden bg-bg-secondary p-1"
          >
            <div
              class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
            >
              <div class="rounded-md border border-border-secondary p-1 shadow-sm">
                <h3 class="text-center text-sm font-semibold text-accent/70">
                  {{ t('tools') }}
                </h3>
              </div>
              <div class="rounded-md border border-border-secondary p-1 shadow-sm">
                <p class="text-xs leading-normal font-medium wrap-break-word text-secondary">
                  {{ t(hostIsExcel ? 'excelToolsDescription' : 'wordToolsDescription') }}
                </p>
              </div>
              <div class="flex flex-col gap-2">
                <div
                  v-for="tool in allToolsList"
                  :key="tool.name"
                  class="flex items-center gap-2 rounded-md border border-border bg-surface p-2 hover:border-accent"
                >
                  <input
                    :id="'tool-' + tool.name"
                    type="checkbox"
                    :checked="enabledTools.has(tool.name)"
                    class="h-4 w-4 cursor-pointer"
                    @change="toggleTool(tool.name)"
                  />
                  <div class="flex flex-col" @click="toggleTool(tool.name)">
                    <label :for="'tool-' + tool.name" class="text-xs font-semibold text-secondary">
                      {{ $t(`${hostIsExcel ? 'excelTool' : 'wordTool'}_${tool.name}`) }}
                    </label>
                    <span class="text-xs text-secondary/90">
                      {{ $t(`${hostIsExcel ? 'excelTool' : 'wordTool'}_${tool.name}_desc`) }}
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { useStorage } from '@vueuse/core'
import {
  ArrowLeft,
  Edit2,
  Globe,
  MessageSquare,
  Plus,
  RotateCcwIcon,
  Save,
  Settings,
  Trash2,
  Wrench,
} from 'lucide-vue-next'
import { onBeforeMount, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { useRouter } from 'vue-router'

import { fetchModels, healthCheck } from '@/api/backend'
import CustomButton from '@/components/CustomButton.vue'
import CustomInput from '@/components/CustomInput.vue'
import SettingCard from '@/components/SettingCard.vue'
import SingleSelect from '@/components/SingleSelect.vue'
import { buildInPrompt, excelBuiltInPrompt } from '@/utils/constant'
import { optionLists } from '@/utils/common'
import { localStorageKey } from '@/utils/enum'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { isExcel } from '@/utils/hostDetection'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { i18n } from '@/i18n'

const { t } = useI18n()
const router = useRouter()

const currentTab = ref('general')

// Settings
const localLanguage = useStorage(localStorageKey.localLanguage, 'fr')
const replyLanguage = useStorage(localStorageKey.replyLanguage, 'Fran\u00e7ais')
const agentMaxIterations = useStorage(localStorageKey.agentMaxIterations, 25)
const userGender = useStorage(localStorageKey.userGender, 'unspecified')
const userFirstName = useStorage(localStorageKey.userFirstName, '')
const userLastName = useStorage(localStorageKey.userLastName, '')

const localLanguageOptions = optionLists.localLanguageList
const replyLanguageOptions = optionLists.replyLanguageList
const genderOptions = [
  { label: t('userGenderUnspecified'), value: 'unspecified' },
  { label: t('userGenderFemale'), value: 'female' },
  { label: t('userGenderMale'), value: 'male' },
  { label: t('userGenderNonBinary'), value: 'nonbinary' },
]

// Backend
const backendOnline = ref(false)
const availableModels = ref<Record<string, ModelInfo>>({})

// Host detection
const hostIsExcel = isExcel()

// Tools - switch based on host
const appToolsList = hostIsExcel ? getExcelToolDefinitions() : getWordToolDefinitions()
const allToolsList = [...getGeneralToolDefinitions(), ...appToolsList]
const enabledTools = ref<Set<string>>(new Set())

// Prompt management
interface Prompt {
  id: string
  name: string
  systemPrompt: string
  userPrompt: string
}

const savedPrompts = ref<Prompt[]>([])
const editingPromptId = ref<string>('')
const editingPrompt = ref<Prompt>({ id: '', name: '', systemPrompt: '', userPrompt: '' })

// Built-in prompts - switch between Word and Excel
type WordBuiltinPromptKey = 'translate' | 'polish' | 'academic' | 'summary' | 'grammar'
type ExcelBuiltinPromptKey = 'analyze' | 'chart' | 'formula' | 'format' | 'explain'
type BuiltinPromptKey = WordBuiltinPromptKey | ExcelBuiltinPromptKey

interface BuiltinPromptConfig {
  system: (language: string) => string
  user: (text: string, language: string) => string
}

const wordBuiltInPromptsData: Record<WordBuiltinPromptKey, BuiltinPromptConfig> = {
  translate: { ...buildInPrompt.translate },
  polish: { ...buildInPrompt.polish },
  academic: { ...buildInPrompt.academic },
  summary: { ...buildInPrompt.summary },
  grammar: { ...buildInPrompt.grammar },
}

const excelBuiltInPromptsData: Record<ExcelBuiltinPromptKey, BuiltinPromptConfig> = {
  analyze: { ...excelBuiltInPrompt.analyze },
  chart: { ...excelBuiltInPrompt.chart },
  formula: { ...excelBuiltInPrompt.formula },
  format: { ...excelBuiltInPrompt.format },
  explain: { ...excelBuiltInPrompt.explain },
}

const builtInPromptsData = ref<Record<string, BuiltinPromptConfig>>(
  hostIsExcel ? { ...excelBuiltInPromptsData } : { ...wordBuiltInPromptsData },
)

const editingBuiltinPromptKey = ref<BuiltinPromptKey | ''>('')
const editingBuiltinPrompt = ref<{ system: string; user: string }>({ system: '', user: '' })
const originalBuiltInPrompts: Record<string, BuiltinPromptConfig> = hostIsExcel
  ? { ...excelBuiltInPrompt }
  : { ...buildInPrompt }
const builtInPromptsStorageKey = hostIsExcel ? 'customExcelBuiltInPrompts' : 'customBuiltInPrompts'

const tabs = [
  { id: 'general', label: 'general', defaultLabel: 'General', icon: Globe },
  { id: 'prompts', label: 'prompts', defaultLabel: 'Prompts', icon: MessageSquare },
  { id: 'builtinPrompts', label: 'builtinPrompts', defaultLabel: 'Built-in Prompts', icon: Settings },
  { id: 'tools', label: 'tools', defaultLabel: 'Tools', icon: Wrench },
]

// Watchers
watch(localLanguage, (val) => {
  i18n.global.locale.value = val as 'en' | 'fr'
  localStorage.setItem(localStorageKey.localLanguage, val)
})

watch(replyLanguage, (val) => {
  localStorage.setItem(localStorageKey.replyLanguage, val)
})

watch(agentMaxIterations, (val) => {
  localStorage.setItem(localStorageKey.agentMaxIterations, String(val))
})

watch(userGender, (val) => {
  localStorage.setItem(localStorageKey.userGender, val)
})

watch(userFirstName, (val) => {
  localStorage.setItem(localStorageKey.userFirstName, val.trim())
})

watch(userLastName, (val) => {
  localStorage.setItem(localStorageKey.userLastName, val.trim())
})

// Prompt management
function loadPrompts() {
  const stored = localStorage.getItem('savedPrompts')
  if (stored) {
    try {
      savedPrompts.value = JSON.parse(stored)
      return
    } catch {
      localStorage.removeItem('savedPrompts')
    }
  }
  savedPrompts.value = [
    { id: 'default', name: 'Default', systemPrompt: '', userPrompt: '' },
  ]
  savePromptsToStorage()
}

function savePromptsToStorage() {
  localStorage.setItem('savedPrompts', JSON.stringify(savedPrompts.value))
}

function addNewPrompt() {
  const newPrompt: Prompt = {
    id: `prompt_${Date.now()}`,
    name: `Prompt ${savedPrompts.value.length + 1}`,
    systemPrompt: '',
    userPrompt: '',
  }
  savedPrompts.value.push(newPrompt)
  savePromptsToStorage()
  startEditPrompt(newPrompt)
}

function startEditPrompt(prompt: Prompt) {
  editingPromptId.value = prompt.id
  editingPrompt.value = { ...prompt }
}

function savePromptEdit() {
  const index = savedPrompts.value.findIndex(p => p.id === editingPromptId.value)
  if (index !== -1) {
    savedPrompts.value[index] = { ...editingPrompt.value }
    savePromptsToStorage()
  }
  editingPromptId.value = ''
}

function cancelEdit() {
  editingPromptId.value = ''
}

function deletePrompt(id: string) {
  if (savedPrompts.value.length <= 1) return
  const index = savedPrompts.value.findIndex(p => p.id === id)
  if (index !== -1) {
    savedPrompts.value.splice(index, 1)
    savePromptsToStorage()
  }
}

// Built-in prompts
function loadBuiltInPrompts() {
  const stored = localStorage.getItem(builtInPromptsStorageKey)
  if (stored) {
    try {
      const customPrompts = JSON.parse(stored)
      Object.keys(customPrompts).forEach(key => {
        if (builtInPromptsData.value[key]) {
          builtInPromptsData.value[key] = {
            system: (language: string) => customPrompts[key].system.replace(/\$\{language\}/g, language),
            user: (text: string, language: string) =>
              customPrompts[key].user.replace(/\$\{text\}/g, text).replace(/\$\{language\}/g, language),
          }
        }
      })
    } catch (error) {
      console.error('Error loading custom built-in prompts:', error)
    }
  }
}

function saveBuiltInPrompts() {
  const customPrompts: Record<string, { system: string; user: string }> = {}
  Object.keys(builtInPromptsData.value).forEach(key => {
    customPrompts[key] = {
      system: builtInPromptsData.value[key].system('${language}'),
      user: builtInPromptsData.value[key].user('${text}', '${language}'),
    }
  })
  localStorage.setItem(builtInPromptsStorageKey, JSON.stringify(customPrompts))
}

function toggleEditBuiltinPrompt(key: string) {
  if (editingBuiltinPromptKey.value === key) {
    builtInPromptsData.value[key] = {
      system: (language: string) => editingBuiltinPrompt.value.system.replace(/\$\{language\}/g, language),
      user: (text: string, language: string) =>
        editingBuiltinPrompt.value.user.replace(/\$\{text\}/g, text).replace(/\$\{language\}/g, language),
    }
    saveBuiltInPrompts()
    editingBuiltinPromptKey.value = ''
  } else {
    editingBuiltinPromptKey.value = key as BuiltinPromptKey
    editingBuiltinPrompt.value = {
      system: builtInPromptsData.value[key].system('${language}'),
      user: builtInPromptsData.value[key].user('${text}', '${language}'),
    }
  }
}

function isBuiltinPromptModified(key: string): boolean {
  if (!originalBuiltInPrompts[key]) return false
  const current = {
    system: builtInPromptsData.value[key].system('English'),
    user: builtInPromptsData.value[key].user('sample text', 'English'),
  }
  const original = {
    system: originalBuiltInPrompts[key].system('English'),
    user: originalBuiltInPrompts[key].user('sample text', 'English'),
  }
  return current.system !== original.system || current.user !== original.user
}

function resetBuiltinPrompt(key: string) {
  if (!originalBuiltInPrompts[key]) return
  builtInPromptsData.value[key] = { ...originalBuiltInPrompts[key] }
  saveBuiltInPrompts()
  if (editingBuiltinPromptKey.value === key) {
    editingBuiltinPromptKey.value = ''
  }
}

function getSystemPromptPreview(systemFunc: (language: string) => string): string {
  const full = systemFunc('English')
  return full.length > 100 ? full.substring(0, 100) + '...' : full
}

function getUserPromptPreview(userFunc: (text: string, language: string) => string): string {
  const full = userFunc('[selected text]', 'English')
  return full.length > 100 ? full.substring(0, 100) + '...' : full
}

// Tools
function loadToolPreferences() {
  const stored = localStorage.getItem('enabledTools')
  if (stored) {
    try {
      enabledTools.value = new Set(JSON.parse(stored))
    } catch {
      enabledTools.value = new Set(allToolsList.map(t => t.name))
    }
  } else {
    enabledTools.value = new Set(allToolsList.map(t => t.name))
  }
}

function toggleTool(toolName: string) {
  if (enabledTools.value.has(toolName)) {
    enabledTools.value.delete(toolName)
  } else {
    enabledTools.value.add(toolName)
  }
  localStorage.setItem('enabledTools', JSON.stringify([...enabledTools.value]))
}

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
  loadPrompts()
  loadBuiltInPrompts()
  loadToolPreferences()
  checkBackend()
})
</script>
