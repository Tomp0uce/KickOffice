<template>
  <div role="tabpanel" class="flex w-full flex-1 flex-col items-center gap-2 bg-bg-secondary p-1">
    <div
      class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
    >
      <div class="flex items-center justify-between">
        <h3 class="text-center text-sm font-semibold text-main">
          {{ t('savedPrompts') }}
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
          <label class="mb-1 block text-xs font-semibold text-secondary">{{
            t('systemPrompt')
          }}</label>
          <textarea
            v-model="editingPrompt.systemPrompt"
            class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm leading-normal text-main transition-all duration-200 ease-apple focus:border-accent focus:outline-none"
            rows="3"
            :placeholder="t('systemPromptPlaceholder')"
          />

          <label class="mb-1 block text-xs font-semibold text-secondary">{{
            t('userPrompt')
          }}</label>
          <textarea
            v-model="editingPrompt.userPrompt"
            class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm leading-normal text-main transition-all duration-200 ease-apple focus:border-accent focus:outline-none"
            rows="3"
            :placeholder="t('userPromptPlaceholder')"
          />

          <div class="mt-3 flex gap-2">
            <CustomButton type="primary" class="flex-1" :text="t('save')" @click="savePromptEdit" />
            <CustomButton type="secondary" class="flex-1" :text="t('cancel')" @click="cancelEdit" />
          </div>
        </div>

        <div v-else class="mt-2">
          <p class="overflow-hidden text-xs font-semibold text-ellipsis text-secondary">
            {{ prompt.systemPrompt.substring(0, 100)
            }}{{ prompt.systemPrompt.length > 100 ? '...' : '' }}
          </p>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import { useI18n } from 'vue-i18n';
import { Edit2, Plus, Trash2 } from 'lucide-vue-next';

import CustomButton from '@/components/CustomButton.vue';
import { loadSavedPromptsFromStorage, type SavedPrompt } from '@/utils/savedPrompts';
import { logService } from '@/utils/logger';

const { t } = useI18n();

const savedPrompts = ref<SavedPrompt[]>([]);
const editingPromptId = ref<string>('');
const editingPrompt = ref<SavedPrompt>({
  id: '',
  name: '',
  systemPrompt: '',
  userPrompt: '',
});

function loadPrompts() {
  const defaultPrompts: SavedPrompt[] = [
    { id: 'default', name: 'Default', systemPrompt: '', userPrompt: '' },
  ];
  savedPrompts.value = loadSavedPromptsFromStorage(defaultPrompts);
  if (savedPrompts.value.length === 0) {
    savedPrompts.value = defaultPrompts;
  }
  savePromptsToStorage();
}

function savePromptsToStorage() {
  try {
    localStorage.setItem('savedPrompts', JSON.stringify(savedPrompts.value));
  } catch (e) {
    if (e instanceof DOMException && e.name === 'QuotaExceededError') {
      logService.warn('[PromptsTab] localStorage quota exceeded — saved prompts not persisted');
    } else {
      throw e;
    }
  }
}

function addNewPrompt() {
  const newPrompt: SavedPrompt = {
    id: `prompt_${Date.now()}`,
    name: `Prompt ${savedPrompts.value.length + 1}`,
    systemPrompt: '',
    userPrompt: '',
  };
  savedPrompts.value.push(newPrompt);
  savePromptsToStorage();
  startEditPrompt(newPrompt);
}

function startEditPrompt(prompt: SavedPrompt) {
  editingPromptId.value = prompt.id;
  editingPrompt.value = { ...prompt };
}

function savePromptEdit() {
  const index = savedPrompts.value.findIndex(p => p.id === editingPromptId.value);
  if (index !== -1) {
    savedPrompts.value[index] = { ...editingPrompt.value };
    savePromptsToStorage();
  }
  editingPromptId.value = '';
}

function cancelEdit() {
  editingPromptId.value = '';
  editingPrompt.value = { id: '', name: '', systemPrompt: '', userPrompt: '' };
}

function deletePrompt(id: string) {
  if (savedPrompts.value.length <= 1) return;
  const index = savedPrompts.value.findIndex(p => p.id === id);
  if (index !== -1) {
    savedPrompts.value.splice(index, 1);
    savePromptsToStorage();
  }
}

// Allow parent to trigger initial load
loadPrompts();
</script>
