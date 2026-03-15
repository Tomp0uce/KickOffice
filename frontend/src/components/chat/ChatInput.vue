<template>
  <div
    class="flex flex-col gap-1 rounded-md transition-colors"
    :class="{ 'ring-2 ring-accent border-accent': isDragOver }"
    @dragover.prevent="handleDragOver"
    @dragleave.prevent="handleDragLeave"
    @drop.prevent="handleDrop"
  >
    <!-- Zone d'affichage des fichiers attachés -->
    <div v-if="attachedFiles.length > 0" class="flex flex-wrap gap-2 px-1">
      <div
        v-for="(file, index) in attachedFiles"
        :key="index"
        class="flex items-center gap-1 rounded-sm bg-accent/20 px-2 py-1 text-[10px] text-accent font-medium truncate max-w-full"
      >
        <span class="truncate">{{ file.name }}</span>
        <button
          class="cursor-pointer hover:text-danger ml-1 opacity-70 hover:opacity-100 focus:outline-none focus:ring-2 focus:ring-primary/50 rounded-sm"
          @click="removeFile(index)"
          :title="$t('removeFile')"
        >
          &times;
        </button>
      </div>
    </div>

    <div
      class="card-base flex min-w-12 items-center gap-2 focus-within:border-accent"
      :class="{
        'ring-2 ring-accent draft-focus-glow': isDraftFocusGlowing,
      }"
    >
      <textarea
        ref="textareaEl"
        :value="modelValue"
        :aria-label="inputPlaceholder"
        class="placeholder:text-secondary block max-h-36 flex-1 resize-none overflow-y-auto border-none bg-transparent py-2 text-xs leading-normal text-main outline-none placeholder:text-xs focus:ring-2 focus:ring-primary/50"
        :placeholder="inputPlaceholder"
        rows="2"
        @keydown.enter.exact.prevent="triggerSubmit()"
        @input="handleInput"
        @focus="handleFocus"
        @blur="handleBlur"
        @paste="handlePaste"
      />

      <button
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-transparent hover:bg-surface text-secondary hover:text-accent disabled:cursor-not-allowed disabled:opacity-50 focus:outline-none focus:ring-2 focus:ring-primary/50"
        :title="$t('attachDocument')"
        :aria-label="$t('attachDocument')"
        :disabled="loading || processingFiles || attachedFiles.length >= 3"
        @click="triggerFileInput"
      >
        <Loader2 v-if="processingFiles" :size="ICON_SIZE_MD" class="animate-spin text-accent" />
        <Paperclip v-else :size="ICON_SIZE_MD" />
      </button>
      <input
        type="file"
        ref="fileInputEl"
        class="hidden"
        accept=".pdf,.docx,.xlsx,.xls,.csv,.txt,.md,.png,.jpeg,.jpg"
        multiple
        @change="onFileSelected"
      />

      <button
        v-if="loading"
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-danger text-white focus:outline-none focus:ring-2 focus:ring-primary/50"
        :title="stopLabel"
        :aria-label="stopLabel"
        @click="handleStop"
      >
        <Square :size="18" />
      </button>
      <button
        v-else
        class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-accent text-white disabled:cursor-not-allowed disabled:bg-accent/50 focus:outline-none focus:ring-2 focus:ring-primary/50"
        :title="sendLabel"
        :disabled="(!modelValue.trim() && attachedFiles.length === 0) || !backendOnline"
        :aria-label="sendLabel"
        @click="triggerSubmit()"
      >
        <Send :size="18" />
      </button>
    </div>
    <!-- Formatting checkboxes removed (GEN-L3) -->

    <!-- UX-L2 Keyboard Shortcuts Hint -->
    <div
      class="flex justify-start px-2 mt-0.5 opacity-0 transition-opacity duration-300"
      :class="{ 'opacity-100': isFocused }"
    >
      <span class="text-[10px] text-secondary/60 font-medium">
        {{ t('shiftEnterHint', 'Shift + Enter for new line') }}
      </span>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { Send, Square, Paperclip, Loader2 } from 'lucide-vue-next';
import { ref } from 'vue';
import { useI18n } from 'vue-i18n';
import { message as messageUtil } from '@/utils/message';
import { ICON_SIZE_MD, MAX_UPLOAD_BYTES } from '@/constants/limits';

const { t } = useI18n();

const props = defineProps<{
  modelValue: string;
  inputPlaceholder: string;
  loading: boolean;
  backendOnline: boolean;
  showWordFormatting: boolean;
  useWordFormatting: boolean;
  useSelectedText: boolean;
  useWordFormattingLabel: string;
  includeSelectionLabel: string;
  sendLabel: string;
  stopLabel: string;
  isDraftFocusGlowing?: boolean;
}>();

const emit = defineEmits<{
  (e: 'update:modelValue', value: string): void;
  (e: 'submit', value: string, files?: File[]): void;
  (e: 'stop'): void;
}>();

const isDragOver = ref(false);
const dragCounter = ref(0);
const attachedFiles = ref<File[]>([]);
const fileInputEl = ref<HTMLInputElement>();
const isFocused = ref(false);

const handleInput = (event: Event) => {
  const val = (event.target as HTMLTextAreaElement).value;
  emit('update:modelValue', val);
};

const handleFocus = () => {
  isFocused.value = true;
};

const handleBlur = () => {
  isFocused.value = false;
};

// --- CLIPBOARD PASTE LOGIC (CLIP-M1) ---
const handlePaste = async (event: ClipboardEvent) => {
  const items = event.clipboardData?.items;
  if (!items) return;

  const imageFiles: File[] = [];

  // Check for image items in clipboard
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (item.type.startsWith('image/')) {
      const blob = item.getAsFile();
      if (blob) {
        // Create a File object with a descriptive name
        const timestamp = new Date().getTime();
        const extension = item.type.split('/')[1] || 'png';
        const file = new File([blob], `pasted-image-${timestamp}.${extension}`, {
          type: item.type,
        });
        imageFiles.push(file);
      }
    }
  }

  // Process pasted images through the same pipeline as uploaded files
  if (imageFiles.length > 0) {
    event.preventDefault(); // Prevent default paste behavior for images
    const fileList = createFileList(imageFiles);
    await processFiles(fileList);
  }
};

// Helper to create a FileList-like object from File array
const createFileList = (files: File[]): FileList => {
  const dataTransfer = new DataTransfer();
  files.forEach(file => dataTransfer.items.add(file));
  return dataTransfer.files;
};

// --- FILE UPLOAD LOGIC ---
const allowedTypes = [
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'text/csv',
  'text/plain',
  'text/markdown',
  'image/png',
  'image/jpeg',
];

const handleDragOver = (e: DragEvent) => {
  e.preventDefault(); // Ensure default behavior is prevented
  if (e.dataTransfer?.types.includes('Files')) {
    dragCounter.value++;
    isDragOver.value = true;
  }
};

const handleDragLeave = () => {
  dragCounter.value--;
  if (dragCounter.value === 0) {
    isDragOver.value = false;
  }
};

const handleDrop = (e: DragEvent) => {
  isDragOver.value = false;
  const files = e.dataTransfer?.files;
  if (files) processFiles(files);
};

const triggerFileInput = () => {
  fileInputEl.value?.click();
};

const onFileSelected = async (e: Event) => {
  const target = e.target as HTMLInputElement;
  if (target.files) {
    await processFiles(target.files);
  }
  // Reset input seulement une fois que les fichiers ont été extraits en mémoire
  if (fileInputEl.value) fileInputEl.value.value = '';
};

const processingFiles = ref(false);

const processFiles = async (fileList: FileList) => {
  processingFiles.value = true;
  // Small artificial delay to show UI feedback for large files
  await new Promise(resolve => setTimeout(resolve, 300));

  // Check for allowed extensions if mime type misses
  const allowedExtensions = [
    '.pdf',
    '.docx',
    '.xlsx',
    '.xls',
    '.csv',
    '.txt',
    '.md',
    '.png',
    '.jpeg',
    '.jpg',
  ];
  let rejectedCount = 0;
  let oversizedCount = 0;

  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];

    // Limits: Max 3 files
    if (attachedFiles.value.length >= 3) {
      messageUtil.warning(t('maxFilesWarning'));
      break;
    }

    const isExtensionOk = allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));

    if (file.size > MAX_UPLOAD_BYTES) {
      oversizedCount++;
      continue;
    }

    if (allowedTypes.includes(file.type) || isExtensionOk) {
      // Avoid duplicate by name
      if (!attachedFiles.value.some(f => f.name === file.name)) {
        attachedFiles.value.push(file);
      }
    } else {
      rejectedCount++;
    }
  }

  if (oversizedCount > 0) {
    messageUtil.error(t('filesOversized', { count: oversizedCount }));
  }
  if (rejectedCount > 0) {
    messageUtil.error(t('filesRejected', { count: rejectedCount }));
  }
  processingFiles.value = false;
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
  emit('submit', props.modelValue, [...attachedFiles.value]);
  attachedFiles.value = [];
};

const handleStop = () => {
  emit('stop');
};

const textareaEl = ref<HTMLTextAreaElement>();
defineExpose({ textareaEl });
</script>

<style scoped>
/* UX-L1: Animation for draft focus glow */
.draft-focus-glow {
  animation: pulse 0.5s ease-in-out;
  animation-iteration-count: 3;
}
</style>
