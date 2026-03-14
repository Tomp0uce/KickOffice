<template>
  <div
    role="tabpanel"
    class="flex w-full flex-1 items-center gap-2 overflow-hidden bg-bg-secondary p-1"
  >
    <div
      class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
    >
      <div class="rounded-md border border-border-secondary p-1 shadow-sm">
        <h3 class="text-center text-sm font-semibold text-accent/70">
          {{ t('builtinPrompts') }}
        </h3>
      </div>
      <div class="rounded-md border border-border-secondary p-1 shadow-sm">
        <p class="text-xs leading-normal font-medium wrap-break-word text-secondary">
          {{
            t('builtinPromptsDescription', {
              text: '[TEXT]',
            })
          }}
        </p>
      </div>

      <div
        v-for="(promptConfig, key) in builtInPromptsData"
        :key="key"
        class="card-base flex flex-col gap-2 hover:border-2 hover:border-accent"
      >
        <div class="flex flex-row items-start justify-between">
          <div class="flex items-center gap-2">
            <span class="text-sm font-semibold text-secondary">{{
              t(
                hostIsExcel
                  ? `excel${(key as string).charAt(0).toUpperCase() + (key as string).slice(1)}`
                  : (key as string),
              ) || key
            }}</span>
          </div>
          <div class="flex gap-1">
            <CustomButton
              :icon="editingBuiltinPromptKey === key ? Save : Edit2"
              text=""
              :title="editingBuiltinPromptKey === key ? t('save') : t('edit')"
              class="border-none bg-surface! p-1.5!"
              type="secondary"
              :icon-size="14"
              @click="toggleEditBuiltinPrompt(key as string)"
            />
            <CustomButton
              v-if="isBuiltinPromptModified(key as string)"
              :icon="RotateCcwIcon"
              text=""
              :title="t('reset')"
              class="border-none bg-surface! p-1.5!"
              type="secondary"
              :icon-size="14"
              @click="resetBuiltinPrompt(key as string)"
            />
          </div>
        </div>

        <div v-if="editingBuiltinPromptKey === key">
          <label class="mt-2 block text-xs font-semibold text-secondary">{{
            t('systemPrompt')
          }}</label>
          <textarea
            v-model="editingBuiltinPrompt.system"
            class="min-h-20 w-full rounded-md border border-border bg-bg-secondary p-2 text-xs text-main focus:border-accent focus:outline-none"
            rows="3"
          />
          <label class="mt-2 block text-xs font-semibold text-secondary">{{
            t('userPrompt')
          }}</label>
          <textarea
            v-model="editingBuiltinPrompt.user"
            class="min-h-20 w-full rounded-md border border-border bg-bg-secondary p-2 text-xs text-main focus:border-accent focus:outline-none"
            rows="4"
          />
        </div>

        <div v-else class="mt-2">
          <p class="mb-2 text-xs font-semibold text-secondary">{{ t('systemPrompt') }}:</p>
          <p class="text-xs leading-normal wrap-break-word text-secondary">
            {{ getSystemPromptPreview(promptConfig.system) }}
          </p>
          <p class="mt-2 mb-2 text-xs font-semibold text-secondary">{{ t('userPrompt') }}:</p>
          <p class="text-xs leading-normal wrap-break-word text-secondary">
            {{ getUserPromptPreview(promptConfig.user) }}
          </p>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import { useI18n } from 'vue-i18n';
import { Edit2, RotateCcwIcon, Save } from 'lucide-vue-next';

import CustomButton from '@/components/CustomButton.vue';
import {
  builtInPrompt,
  excelBuiltInPrompt,
  outlookBuiltInPrompt,
  powerPointBuiltInPrompt,
} from '@/utils/constant';
import { isExcel, forHost } from '@/utils/hostDetection';

const { t } = useI18n();

const hostIsExcel = isExcel();

// Types
type WordBuiltinPromptKey = 'translate' | 'polish' | 'academic' | 'summary' | 'proofread';
type ExcelBuiltinPromptKey = 'analyze' | 'chart' | 'formula' | 'format' | 'explain';
type PowerPointBuiltinPromptKey = 'bullets' | 'speakerNotes' | 'punchify' | 'proofread' | 'visual';
type OutlookBuiltinPromptKey =
  | 'reply'
  | 'translate_formalize'
  | 'concise'
  | 'proofread'
  | 'extract';
type BuiltinPromptKey =
  | WordBuiltinPromptKey
  | ExcelBuiltinPromptKey
  | PowerPointBuiltinPromptKey
  | OutlookBuiltinPromptKey;

interface BuiltinPromptConfig {
  system: (language: string) => string;
  user: (text: string, language: string) => string;
}

const wordBuiltInPromptsData: Record<WordBuiltinPromptKey, BuiltinPromptConfig> = {
  translate: { ...builtInPrompt.translate },
  polish: { ...builtInPrompt.polish },
  academic: { ...builtInPrompt.academic },
  summary: { ...builtInPrompt.summary },
  proofread: { ...builtInPrompt.proofread },
};

const excelBuiltInPromptsData: Record<ExcelBuiltinPromptKey, BuiltinPromptConfig> = {
  analyze: { ...excelBuiltInPrompt.analyze },
  chart: { ...excelBuiltInPrompt.chart },
  formula: { ...excelBuiltInPrompt.formula },
  format: { ...excelBuiltInPrompt.format },
  explain: { ...excelBuiltInPrompt.explain },
};

const powerPointBuiltInPromptsData: Record<PowerPointBuiltinPromptKey, BuiltinPromptConfig> = {
  bullets: { ...powerPointBuiltInPrompt.bullets },
  speakerNotes: { ...powerPointBuiltInPrompt.speakerNotes },
  punchify: { ...powerPointBuiltInPrompt.punchify },
  proofread: { ...powerPointBuiltInPrompt.proofread },
  visual: { ...powerPointBuiltInPrompt.visual },
};

const outlookBuiltInPromptsData: Record<OutlookBuiltinPromptKey, BuiltinPromptConfig> = {
  reply: { ...outlookBuiltInPrompt.reply },
  translate_formalize: { ...outlookBuiltInPrompt.translate_formalize },
  concise: { ...outlookBuiltInPrompt.concise },
  proofread: { ...outlookBuiltInPrompt.proofread },
  extract: { ...outlookBuiltInPrompt.extract },
};

const selectedBuiltInPromptsData = forHost({
  outlook: { ...outlookBuiltInPromptsData },
  excel: { ...excelBuiltInPromptsData },
  powerpoint: { ...powerPointBuiltInPromptsData },
  word: { ...wordBuiltInPromptsData },
}) as Record<string, BuiltinPromptConfig>;

const selectedOriginalBuiltInPrompts = forHost({
  outlook: { ...outlookBuiltInPrompt },
  excel: { ...excelBuiltInPrompt },
  powerpoint: { ...powerPointBuiltInPrompt },
  word: { ...builtInPrompt },
}) as Record<string, BuiltinPromptConfig>;

const builtInPromptsData = ref<Record<string, BuiltinPromptConfig>>(selectedBuiltInPromptsData);
const originalBuiltInPrompts: Record<string, BuiltinPromptConfig> = selectedOriginalBuiltInPrompts;

const editingBuiltinPromptKey = ref<BuiltinPromptKey | ''>('');
const editingBuiltinPrompt = ref<{ system: string; user: string }>({
  system: '',
  user: '',
});

const builtInPromptsStorageKey = forHost({
  default: 'ki_Settings_BuiltInPrompts_v5',
  powerpoint: 'ki_Settings_BuiltInPrompts_ppt_v5',
  outlook: 'ki_Settings_BuiltInPrompts_outlook_v5',
  word: 'ki_Settings_BuiltInPrompts_word_v5',
  excel: 'ki_Settings_BuiltInPrompts_excel_v5',
}) as string;

function loadBuiltInPrompts() {
  const stored = localStorage.getItem(builtInPromptsStorageKey);
  if (stored) {
    try {
      const customPrompts = JSON.parse(stored);
      Object.keys(customPrompts).forEach(key => {
        if (builtInPromptsData.value[key]) {
          builtInPromptsData.value[key] = {
            system: (language: string) =>
              customPrompts[key].system.replace(/\[LANGUAGE\]/g, language),
            user: (text: string, language: string) =>
              customPrompts[key].user.replace(/\[TEXT\]/g, text).replace(/\[LANGUAGE\]/g, language),
          };
        }
      });
    } catch (error) {
      console.error('Error loading custom built-in prompts:', error);
    }
  }
}

function saveBuiltInPrompts() {
  const customPrompts: Record<string, { system: string; user: string }> = {};
  Object.keys(builtInPromptsData.value).forEach(key => {
    customPrompts[key] = {
      system: builtInPromptsData.value[key].system('[LANGUAGE]'),
      user: builtInPromptsData.value[key].user('[TEXT]', '[LANGUAGE]'),
    };
  });
  try {
    localStorage.setItem(builtInPromptsStorageKey, JSON.stringify(customPrompts));
  } catch (e) {
    if (e instanceof DOMException && e.name === 'QuotaExceededError') {
      console.warn(
        '[BuiltinPromptsTab] localStorage quota exceeded — built-in prompts not persisted',
      );
    } else {
      throw e;
    }
  }
}

function toggleEditBuiltinPrompt(key: string) {
  if (editingBuiltinPromptKey.value === key) {
    builtInPromptsData.value[key] = {
      system: (language: string) =>
        editingBuiltinPrompt.value.system.replace(/\[LANGUAGE\]/g, language),
      user: (text: string, language: string) =>
        editingBuiltinPrompt.value.user
          .replace(/\[TEXT\]/g, text)
          .replace(/\[LANGUAGE\]/g, language),
    };
    saveBuiltInPrompts();
    editingBuiltinPromptKey.value = '';
  } else {
    editingBuiltinPromptKey.value = key as BuiltinPromptKey;
    editingBuiltinPrompt.value = {
      system: builtInPromptsData.value[key].system('[LANGUAGE]'),
      user: builtInPromptsData.value[key].user('[TEXT]', '[LANGUAGE]'),
    };
  }
}

function isBuiltinPromptModified(key: string): boolean {
  if (!originalBuiltInPrompts[key]) return false;
  const current = {
    system: builtInPromptsData.value[key].system('English'),
    user: builtInPromptsData.value[key].user('sample text', 'English'),
  };
  const original = {
    system: originalBuiltInPrompts[key].system('English'),
    user: originalBuiltInPrompts[key].user('sample text', 'English'),
  };
  return current.system !== original.system || current.user !== original.user;
}

function resetBuiltinPrompt(key: string) {
  if (!originalBuiltInPrompts[key]) return;
  builtInPromptsData.value[key] = { ...originalBuiltInPrompts[key] };
  saveBuiltInPrompts();
  if (editingBuiltinPromptKey.value === key) {
    editingBuiltinPromptKey.value = '';
  }
}

function getSystemPromptPreview(systemFunc: (language: string) => string): string {
  const full = systemFunc('English');
  return full.length > 100 ? full.substring(0, 100) + '...' : full;
}

function getUserPromptPreview(userFunc: (text: string, language: string) => string): string {
  const full = userFunc('[selected text]', 'English');
  return full.length > 100 ? full.substring(0, 100) + '...' : full;
}

// Initial load
loadBuiltInPrompts();
</script>
