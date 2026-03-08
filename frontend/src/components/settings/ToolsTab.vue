<template>
  <div
    role="tabpanel"
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
          {{ t(toolDescriptionKey) }}
        </p>
      </div>
      <div class="flex flex-col gap-2">
        <div
          v-for="tool in allToolsList"
          :key="tool.name"
          class="card-base flex items-center gap-2 hover:border-accent"
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
              {{ t(`${toolTranslationPrefix}_${tool.name}`) }}
            </label>
            <span class="text-xs text-secondary/90">
              {{ t(`${toolTranslationPrefix}_${tool.name}_desc`) }}
            </span>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'

import { forHost } from '@/utils/hostDetection'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { getOutlookToolDefinitions } from '@/utils/outlookTools'
import { getPowerPointToolDefinitions } from '@/utils/powerpointTools'
import { getWordToolDefinitions } from '@/utils/wordTools'
import { getEnabledToolNamesFromStorage, persistEnabledTools } from '@/utils/toolStorage'

const { t } = useI18n()

const appToolsList =
  forHost({
    outlook: getOutlookToolDefinitions(),
    excel: getExcelToolDefinitions(),
    powerpoint: getPowerPointToolDefinitions(),
    word: getWordToolDefinitions(),
  }) || []
const allToolsList = [...getGeneralToolDefinitions(), ...appToolsList]
const enabledTools = ref<Set<string>>(new Set())

const toolDescriptionKey = forHost({
  outlook: 'outlookToolsDescription',
  excel: 'excelToolsDescription',
  powerpoint: 'powerpointToolsDescription',
  word: 'wordToolsDescription',
}) as string

const toolTranslationPrefix = forHost({
  outlook: 'outlookTool',
  excel: 'excelTool',
  powerpoint: 'powerpointTool',
  word: 'wordTool',
}) as string

function loadToolPreferences() {
  enabledTools.value = getEnabledToolNamesFromStorage(allToolsList.map(tool => tool.name))
}

function toggleTool(toolName: string) {
  if (enabledTools.value.has(toolName)) {
    enabledTools.value.delete(toolName)
  } else {
    enabledTools.value.add(toolName)
  }
  persistEnabledTools(
    allToolsList.map(tool => tool.name),
    enabledTools.value,
  )
}

// Initial load
loadToolPreferences()
</script>
