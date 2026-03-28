<template>
  <div role="tabpanel" class="flex w-full flex-1 flex-col items-center gap-2 bg-bg-secondary p-1">
    <div
      class="flex h-full w-full flex-col gap-2 overflow-auto rounded-md border border-border-secondary p-2 shadow-sm"
    >
      <!-- Header -->
      <div class="flex items-center justify-between">
        <h3 class="text-sm font-semibold text-main">
          {{ t('skills') || 'Skills' }}
        </h3>
        <div class="flex gap-1">
          <CustomButton
            :icon="Upload"
            text=""
            type="secondary"
            :title="t('importSkill') || 'Importer un skill'"
            class="border-none! bg-surface! p-1.5!"
            :icon-size="14"
            @click="triggerImport"
          />
          <CustomButton
            :icon="Plus"
            text=""
            type="secondary"
            :title="t('createSkill') || 'Créer un skill'"
            class="border-none! bg-surface! p-1.5!"
            :icon-size="14"
            @click="$emit('open-creator')"
          />
        </div>
      </div>

      <!-- Host filter chips -->
      <div class="flex flex-wrap gap-1">
        <button
          v-for="f in hostFilters"
          :key="f.value"
          class="rounded-full px-2 py-0.5 text-xs font-medium cursor-pointer transition-colors duration-fast"
          :aria-pressed="activeFilter === f.value"
          :class="
            activeFilter === f.value
              ? 'bg-accent text-white'
              : 'bg-surface text-secondary hover:bg-border'
          "
          @click="activeFilter = f.value"
        >
          {{ f.label }}
        </button>
      </div>

      <!-- Empty state -->
      <div
        v-if="filteredSkills.length === 0"
        class="flex flex-1 flex-col items-center justify-center gap-2 py-6 text-center"
      >
        <Zap class="text-border" :size="28" />
        <p class="text-xs text-secondary">
          {{ t('noSkillsYet') || 'Aucun skill. Créez-en un !' }}
        </p>
      </div>

      <!-- Skill list -->
      <div
        v-for="skill in filteredSkills"
        :key="skill.id"
        class="rounded-md border border-border bg-surface"
      >
        <!-- Skill header row -->
        <div class="flex items-start justify-between p-2.5">
          <div class="flex min-w-0 flex-1 items-center gap-1.5">
            <component :is="resolveIcon(skill.icon)" class="shrink-0 text-accent" :size="14" />
            <span class="truncate text-sm font-semibold text-main">{{ skill.name }}</span>
            <span class="shrink-0 rounded bg-bg-secondary px-1 py-0.5 text-xs text-secondary">{{
              skill.host
            }}</span>
            <span class="shrink-0 rounded bg-bg-secondary px-1 py-0.5 text-xs text-secondary">{{
              skill.executionMode
            }}</span>
          </div>
          <div class="ml-1 flex shrink-0 gap-0.5">
            <CustomButton
              type="secondary"
              :title="t('export') || 'Exporter'"
              :icon="Download"
              class="border-none! bg-surface! p-1.5!"
              :icon-size="13"
              text=""
              @click="exportSkillToFile(skill)"
            />
            <CustomButton
              type="secondary"
              :title="t('edit') || 'Modifier'"
              :icon="Edit2"
              class="border-none! bg-surface! p-1.5!"
              :icon-size="13"
              text=""
              @click="startEdit(skill)"
            />
            <CustomButton
              type="secondary"
              :title="t('delete') || 'Supprimer'"
              :icon="Trash2"
              class="border-none! bg-surface! p-1.5!"
              :icon-size="13"
              text=""
              @click="handleDelete(skill.id)"
            />
          </div>
        </div>

        <!-- Description preview (collapsed) -->
        <p
          v-if="editingId !== skill.id"
          class="px-2.5 pb-2.5 text-xs leading-normal text-secondary"
        >
          {{ skill.description.substring(0, 110) }}{{ skill.description.length > 110 ? '…' : '' }}
        </p>

        <!-- Inline edit form -->
        <div v-if="editingId === skill.id" class="border-t border-border px-2.5 pb-2.5 pt-2">
          <!-- Name + Icon -->
          <div class="mb-2 flex gap-2">
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">{{
                t('name') || 'Nom'
              }}</label>
              <input
                v-model="editForm.name"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
              />
            </div>
            <div class="w-20">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">{{
                t('icon') || 'Icône'
              }}</label>
              <input
                v-model="editForm.icon"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-sm text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
                placeholder="Zap"
              />
            </div>
          </div>

          <!-- Description -->
          <label class="mb-0.5 block text-xs font-semibold text-secondary">{{
            t('description') || 'Description'
          }}</label>
          <textarea
            v-model="editForm.description"
            class="mb-2 w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-xs leading-normal text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
            rows="2"
          />

          <!-- Host + Mode -->
          <div class="mb-2 flex gap-2">
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">Host</label>
              <select
                v-model="editForm.host"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-xs text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
              >
                <option value="word">Word</option>
                <option value="excel">Excel</option>
                <option value="powerpoint">PowerPoint</option>
                <option value="outlook">Outlook</option>
                <option value="all">Tous</option>
              </select>
            </div>
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">Mode</label>
              <select
                v-model="editForm.executionMode"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 text-xs text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
              >
                <option value="immediate">Direct (chat)</option>
                <option value="draft">Brouillon</option>
                <option value="agent">Agent (modifie le doc)</option>
              </select>
            </div>
          </div>

          <!-- Skill content (markdown) -->
          <label class="mb-0.5 block text-xs font-semibold text-secondary">{{
            t('skillContent') || 'Contenu (markdown)'
          }}</label>
          <textarea
            v-model="editForm.skillContent"
            class="mb-3 w-full rounded-sm border border-border bg-bg-secondary px-2 py-1 font-mono text-xs leading-normal text-main focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30"
            rows="8"
          />

          <div class="flex gap-2">
            <CustomButton
              type="primary"
              class="flex-1"
              :text="t('save') || 'Enregistrer'"
              @click="saveEdit"
            />
            <CustomButton
              type="secondary"
              class="flex-1"
              :text="t('cancel') || 'Annuler'"
              @click="cancelEdit"
            />
          </div>
        </div>
      </div>
    </div>

    <!-- Hidden file input for import -->
    <input
      ref="fileInputRef"
      type="file"
      accept=".md,.skill.md"
      class="hidden"
      @change="handleImport"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue';
import { useI18n } from 'vue-i18n';
import {
  AlignLeft,
  BarChart,
  Briefcase,
  CheckCircle,
  CheckSquare,
  Database,
  Download,
  Edit2,
  Eye,
  Globe,
  GraduationCap,
  Grid3X3,
  HelpCircle,
  Languages,
  List,
  ListChecks,
  Mail,
  Palette,
  Plus,
  Reply,
  Scissors,
  Sparkles,
  Table,
  TrendingUp,
  Trash2,
  Upload,
  Wand2,
  Zap,
  type LucideIcon,
} from 'lucide-vue-next';
import CustomButton from '@/components/CustomButton.vue';
import { useUserSkills } from '@/composables/useUserSkills';
import type { UserSkill } from '@/types/userSkill';
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser';

const { t } = useI18n();

defineEmits<{
  (e: 'open-creator'): void;
}>();

const { skills, updateSkill, deleteSkill, exportSkillToFile, importSkillFromFile } =
  useUserSkills();

// ── Icon resolver ─────────────────────────────────────────────────────────────

const iconMap: Record<string, LucideIcon> = {
  AlignLeft,
  BarChart,
  Briefcase,
  CheckCircle,
  CheckSquare,
  Database,
  Eye,
  Globe,
  GraduationCap,
  Grid3X3,
  HelpCircle,
  Languages,
  List,
  ListChecks,
  Mail,
  Palette,
  Reply,
  Scissors,
  Sparkles,
  Table,
  TrendingUp,
  Wand2,
  Zap,
};

function resolveIcon(name: string): LucideIcon {
  return iconMap[name] ?? Zap;
}

// ── Filter ────────────────────────────────────────────────────────────────────

const hostFilters = [
  { value: 'all-filter', label: 'Tous' },
  { value: 'word', label: 'Word' },
  { value: 'excel', label: 'Excel' },
  { value: 'powerpoint', label: 'PowerPoint' },
  { value: 'outlook', label: 'Outlook' },
] as const;

type FilterValue = (typeof hostFilters)[number]['value'];

const activeFilter = ref<FilterValue>('all-filter');

const filteredSkills = computed(() => {
  if (activeFilter.value === 'all-filter') return skills.value;
  return skills.value.filter(s => s.host === activeFilter.value || s.host === 'all');
});

// ── Edit ──────────────────────────────────────────────────────────────────────

const editingId = ref<string>('');
const editForm = ref({
  name: '',
  description: '',
  host: 'all' as SkillHost,
  executionMode: 'immediate' as SkillExecutionMode,
  icon: 'Zap',
  skillContent: '',
});

function startEdit(skill: UserSkill): void {
  if (editingId.value === skill.id) {
    cancelEdit();
    return;
  }
  editingId.value = skill.id;
  editForm.value = {
    name: skill.name,
    description: skill.description,
    host: skill.host,
    executionMode: skill.executionMode,
    icon: skill.icon,
    skillContent: skill.skillContent,
  };
}

function saveEdit(): void {
  if (!editingId.value) return;
  updateSkill(editingId.value, { ...editForm.value });
  cancelEdit();
}

function cancelEdit(): void {
  editingId.value = '';
}

function handleDelete(id: string): void {
  if (editingId.value === id) cancelEdit();
  deleteSkill(id);
}

// ── Import ────────────────────────────────────────────────────────────────────

const fileInputRef = ref<HTMLInputElement | null>(null);

function triggerImport(): void {
  fileInputRef.value?.click();
}

async function handleImport(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;
  await importSkillFromFile(file);
  input.value = ''; // reset so same file can be re-imported
}
</script>
