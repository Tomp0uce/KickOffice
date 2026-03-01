<template>
  <div
    class="flex flex-col gap-1 rounded-md transition-colors"
    :class="{ 'ring-2 ring-accent border-accent': isDragOver }"
    @dragover.prevent="handleDragOver"
    @dragleave.prevent="handleDragLeave"
    @drop.prevent="handleDrop"
  >
    <div class="flex items-center justify-between gap-2 overflow-hidden">
      <div class="flex min-w-0 flex-1 items-center gap-2 overflow-hidden">
        <label
          :id="modelTierLabelId"
          :for="modelTierSelectId"
          class="shrink-0 text-xs font-medium text-secondary"
          >{{ taskTypeLabel }}</label
        >
        <select
          :id="modelTierSelectId"
          :value="selectedModelTier"
          :aria-labelledby="modelTierLabelId"
          class="h-7 max-w-full min-w-0 cursor-pointer rounded-md border border-border bg-surface p-1 text-xs text-secondary hover:border-accent focus:outline-none"
          @change="handleModelTierChange"
        >
          <option
            v-for="(info, tier) in availableModels"
            :key="tier"
            :value="tier"
          >
            {{ info.label }}
          </option>
        </select>
      </div>
    </div>

    <!-- Zone d'affichage des fichiers attachés -->
    <div v-if="attachedFiles.length > 0" class="flex flex-wrap gap-2 px-1">
      <div
        v-for="(file, index) in attachedFiles"
        :key="index"
        class="flex items-center gap-1 rounded-sm bg-accent/20 px-2 py-1 text-[10px] text-accent font-medium truncate max-w-full"
      >
        <span class="truncate">{{ file.name }}</span>
        <button
          class="cursor-pointer hover:text-danger ml-1 opacity-70 hover:opacity-100"
          @click="removeFile(index)"
          title="Retirer le fichier"
        >
          &times;
        </button>
      </div>
    </div>

    <div
      class="card-base flex min-w-12 items-center gap-2 focus-within:border-accent"
      :class="{
        'ring-2 ring-accent animate-pulse transition-all duration-300':
          draftFocusGlow,
      }"
      :style="
        draftFocusGlow
          ? 'animation-iteration-count: 3; animation-duration: 0.5s;'
          : ''
      "
    >
      <textarea
        ref="textareaEl"
        :value="modelValue"
        :aria-label="inputPlaceholder"
        class="placeholder:text-secondary block max-h-30 flex-1 resize-none overflow-y-auto border-none bg-transparent py-2 text-xs leading-normal text-main outline-none placeholder:text-xs"
        :placeholder="inputPlaceholder"
        rows="1"
        @keydown.enter.exact.prevent="triggerSubmit()"
        @input="handleInput"
      />

      <!-- Bouton trombone pour ajouter un fichier -->
      <button
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-transparent hover:bg-surface text-secondary hover:text-accent disabled:cursor-not-allowed disabled:opacity-50"
        title="Attacher un document (PDF, DOCX, XLSX)"
        :disabled="loading || attachedFiles.length >= 3"
        @click="triggerFileInput"
      >
        <Paperclip :size="16" />
      </button>
      <input
        type="file"
        ref="fileInputEl"
        class="hidden"
        accept=".pdf,.docx,.xlsx,.xls,.csv,.txt,.md"
        multiple
        @change="onFileSelected"
      />

      <button
        v-if="loading"
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-danger text-white"
        :title="stopLabel"
        :aria-label="stopLabel"
        @click="handleStop"
      >
        <Square :size="18" />
      </button>
      <button
        v-else
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-accent text-white disabled:cursor-not-allowed disabled:bg-accent/50"
        :title="sendLabel"
        :disabled="
          (!modelValue.trim() && attachedFiles.length === 0) || !backendOnline
        "
        :aria-label="sendLabel"
        @click="triggerSubmit()"
      >
        <Send :size="18" />
      </button>
    </div>
    <div class="flex justify-center gap-3 px-1">
      <label
        v-if="showWordFormatting"
        :for="wordFormattingCheckboxId"
        class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary"
      >
        <input
          :id="wordFormattingCheckboxId"
          :checked="useWordFormatting"
          :aria-label="useWordFormattingLabel"
          type="checkbox"
          @change="handleWordFormattingChange"
        />
        <span>{{ useWordFormattingLabel }}</span>
      </label>
      <label
        :for="selectedTextCheckboxId"
        class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary"
      >
        <input
          :id="selectedTextCheckboxId"
          :checked="useSelectedText"
          :aria-label="includeSelectionLabel"
          type="checkbox"
          @change="handleSelectedTextChange"
        />
        <span>{{ includeSelectionLabel }}</span>
      </label>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { Send, Square, Paperclip } from "lucide-vue-next";
import { ref } from "vue";
import { messageUtil } from "@/utils/message";

const props = defineProps<{
  availableModels: Record<string, ModelInfo>;
  selectedModelTier: string;
  modelValue: string;
  inputPlaceholder: string;
  loading: boolean;
  backendOnline: boolean;
  showWordFormatting: boolean;
  useWordFormatting: boolean;
  useSelectedText: boolean;
  useWordFormattingLabel: string;
  includeSelectionLabel: string;
  taskTypeLabel: string;
  sendLabel: string;
  stopLabel: string;
  draftFocusGlow?: boolean;
}>();

const emit = defineEmits<{
  (e: "update:selectedModelTier", value: string): void;
  (e: "update:modelValue", value: string): void;
  (e: "update:useWordFormatting", value: boolean): void;
  (e: "update:useSelectedText", value: boolean): void;
  (e: "submit", value: string, files?: File[]): void;
  (e: "stop"): void;
  (e: "input"): void;
}>();

const isDragOver = ref(false);
const attachedFiles = ref<File[]>([]);
const fileInputEl = ref<HTMLInputElement>();

const handleModelTierChange = (event: Event) => {
  emit("update:selectedModelTier", (event.target as HTMLSelectElement).value);
};

const handleInput = (event: Event) => {
  const val = (event.target as HTMLTextAreaElement).value;
  emit("update:modelValue", val);
  emit("input");
};

// --- FILE UPLOAD LOGIC ---
const allowedTypes = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "text/csv",
  "text/plain",
  "text/markdown",
];

const handleDragOver = (e: DragEvent) => {
  if (e.dataTransfer?.types.includes("Files")) {
    isDragOver.value = true;
  }
};

const handleDragLeave = (e: DragEvent) => {
  isDragOver.value = false;
};

const handleDrop = (e: DragEvent) => {
  isDragOver.value = false;
  const files = e.dataTransfer?.files;
  if (files) processFiles(files);
};

const triggerFileInput = () => {
  fileInputEl.value?.click();
};

const onFileSelected = (e: Event) => {
  const target = e.target as HTMLInputElement;
  if (target.files) processFiles(target.files);
  // Reset input so the same file can be selected again if removed
  if (fileInputEl.value) fileInputEl.value.value = "";
};

const processFiles = (fileList: FileList) => {
  // Check for allowed extensions if mime type misses
  const allowedExtensions = [
    ".pdf",
    ".docx",
    ".xlsx",
    ".xls",
    ".csv",
    ".txt",
    ".md",
  ];
  let rejectedCount = 0;
  let oversizedCount = 0;

  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];

    // Limits: Max 3 files
    if (attachedFiles.value.length >= 3) {
      messageUtil.warning("Maximum 3 fichiers autorisés.");
      break;
    }

    const isExtensionOk = allowedExtensions.some((ext) =>
      file.name.toLowerCase().endsWith(ext),
    );

    if (file.size > 10 * 1024 * 1024) {
      oversizedCount++;
      continue;
    }

    if (allowedTypes.includes(file.type) || isExtensionOk) {
      // Avoid duplicate by name
      if (!attachedFiles.value.some((f) => f.name === file.name)) {
        attachedFiles.value.push(file);
      }
    } else {
      rejectedCount++;
    }
  }

  if (oversizedCount > 0) {
    messageUtil.error(
      `${oversizedCount} fichier(s) ignoré(s) : taille > 10MB.`,
    );
  }
  if (rejectedCount > 0) {
    messageUtil.error(
      `${rejectedCount} fichier(s) ignoré(s) : format non supporté.`,
    );
  }
};

const removeFile = (index: number) => {
  attachedFiles.value.splice(index, 1);
};
// -------------------------

const triggerSubmit = () => {
  if (!props.modelValue.trim() && attachedFiles.value.length === 0) {
    return;
  }

  // Pass a copy of the files to prevent issues, then clear immediately
  emit("submit", props.modelValue, [...attachedFiles.value]);
  attachedFiles.value = [];
};

const handleStop = () => {
  emit("stop");
};

const handleWordFormattingChange = (event: Event) => {
  emit("update:useWordFormatting", (event.target as HTMLInputElement).checked);
};

const handleSelectedTextChange = (event: Event) => {
  emit("update:useSelectedText", (event.target as HTMLInputElement).checked);
};

const textareaEl = ref<HTMLTextAreaElement>();
const modelTierSelectId = "chat-model-tier-select";
const modelTierLabelId = "chat-model-tier-label";
const wordFormattingCheckboxId = "chat-word-formatting-checkbox";
const selectedTextCheckboxId = "chat-selected-text-checkbox";
defineExpose({ textareaEl });
</script>
