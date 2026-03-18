<template>
  <div
    class="fixed inset-0 z-50 flex items-end justify-center bg-black/50 p-2 sm:items-center sm:p-4"
    aria-modal="true"
    role="dialog"
    @keydown.esc="handleClose"
  >
    <div
      ref="dialogRef"
      class="card-base flex w-full max-w-lg flex-col gap-0 bg-surface"
      style="max-height: 90vh"
    >
      <!-- Header -->
      <div class="flex items-center justify-between border-b border-border px-4 py-3">
        <h3 class="text-sm font-semibold text-main">
          {{ t('createSkill') || 'Créer un skill' }}
        </h3>
        <CustomButton
          :icon="X"
          text=""
          type="secondary"
          :icon-size="14"
          class="border-none! p-1!"
          @click="handleClose"
        />
      </div>

      <!-- Step indicators -->
      <div class="flex border-b border-border">
        <div
          v-for="(s, i) in stepLabels"
          :key="i"
          class="flex-1 py-1.5 text-center text-xs font-medium transition-colors"
          :class="
            i === currentStepIndex
              ? 'border-b-2 border-accent text-accent'
              : i < currentStepIndex
                ? 'text-secondary'
                : 'text-border'
          "
        >
          {{ s }}
        </div>
      </div>

      <!-- Content (scrollable) -->
      <div class="flex-1 overflow-y-auto p-4">

        <!-- Step 1: Describe -->
        <template v-if="step === 'describe'">
          <label class="mb-1 block text-xs font-semibold text-secondary">
            {{ t('describeSkill') || 'Décrivez ce que vous voulez faire...' }}
          </label>
          <textarea
            ref="descriptionRef"
            v-model="descriptionInput"
            class="mb-3 w-full rounded-sm border border-border bg-bg-secondary px-3 py-2 text-sm leading-normal text-main focus:border-accent focus:outline-none"
            rows="4"
            :placeholder="t('skillDescriptionPlaceholder') || 'Ex: Transformer le texte sélectionné en liste de bullet points concis pour PowerPoint'"
            @keydown.ctrl.enter="handleGenerate"
            @keydown.meta.enter="handleGenerate"
          />
          <label class="mb-1 block text-xs font-semibold text-secondary">
            {{ t('targetHost') || 'Application cible' }}
          </label>
          <select
            v-model="selectedHost"
            class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-sm text-main focus:border-accent focus:outline-none"
          >
            <option value="all">{{ t('allHosts') || 'Tous (Word, Excel, PPT, Outlook)' }}</option>
            <option value="word">Word</option>
            <option value="excel">Excel</option>
            <option value="powerpoint">PowerPoint</option>
            <option value="outlook">Outlook</option>
          </select>
          <p v-if="generationError" class="mt-2 text-xs text-red-500">{{ generationError }}</p>
        </template>

        <!-- Step 2: Generating -->
        <template v-if="step === 'generating'">
          <div class="flex flex-col items-center gap-3 py-8">
            <div class="h-8 w-8 animate-spin rounded-full border-2 border-accent border-t-transparent" />
            <p class="text-sm text-secondary">{{ t('generatingSkill') || 'Génération en cours...' }}</p>
            <p class="text-xs text-secondary/60">{{ t('generatingSkillHint') || 'Le LLM analyse votre demande et sélectionne les bons outils Office.' }}</p>
          </div>
        </template>

        <!-- Step 3: Review -->
        <template v-if="step === 'review'">
          <!-- Name + Icon -->
          <div class="mb-3 flex gap-2">
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">{{ t('name') || 'Nom' }}</label>
              <input
                v-model="editName"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-sm text-main focus:border-accent focus:outline-none"
              />
            </div>
            <div class="w-24">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">{{ t('icon') || 'Icône' }}</label>
              <input
                v-model="editIcon"
                class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-sm text-main focus:border-accent focus:outline-none"
                placeholder="Zap"
              />
            </div>
          </div>

          <!-- Description -->
          <label class="mb-0.5 block text-xs font-semibold text-secondary">{{ t('description') || 'Description' }}</label>
          <textarea
            v-model="editDescription"
            class="mb-3 w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-xs leading-normal text-main focus:border-accent focus:outline-none"
            rows="2"
          />

          <!-- Host + Mode -->
          <div class="mb-3 flex gap-2">
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">Host</label>
              <select v-model="editHost" class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-xs text-main focus:border-accent focus:outline-none">
                <option value="word">Word</option>
                <option value="excel">Excel</option>
                <option value="powerpoint">PowerPoint</option>
                <option value="outlook">Outlook</option>
                <option value="all">Tous</option>
              </select>
            </div>
            <div class="flex-1">
              <label class="mb-0.5 block text-xs font-semibold text-secondary">Mode</label>
              <select v-model="editExecutionMode" class="w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 text-xs text-main focus:border-accent focus:outline-none">
                <option value="immediate">Direct (chat)</option>
                <option value="draft">Brouillon</option>
                <option value="agent">Agent (modifie le doc)</option>
              </select>
            </div>
          </div>

          <!-- Skill content -->
          <label class="mb-0.5 block text-xs font-semibold text-secondary">{{ t('skillContent') || 'Contenu (markdown)' }}</label>
          <textarea
            v-model="editContent"
            class="mb-1 w-full rounded-sm border border-border bg-bg-secondary px-2 py-1.5 font-mono text-xs leading-normal text-main focus:border-accent focus:outline-none"
            rows="10"
          />
        </template>

        <!-- Step 4: Testing -->
        <template v-if="step === 'testing'">
          <div class="flex flex-col items-center gap-3 py-6 text-center">
            <div class="rounded-full bg-accent/10 p-3">
              <PlayCircle class="text-accent" :size="28" />
            </div>
            <p class="text-sm font-semibold text-main">{{ t('skillTestRunning') || 'Skill en cours de test' }}</p>
            <p class="text-xs text-secondary">{{ t('skillTestHint') || 'Le résultat apparaît dans le chat. Revenez ici pour modifier ou sauvegarder.' }}</p>
            <CustomButton
              type="secondary"
              :icon="ArrowLeft"
              :text="t('backToReview') || '← Retour à la review'"
              @click="step = 'review'"
            />
          </div>
        </template>

      </div>

      <!-- Footer actions -->
      <div class="flex gap-2 border-t border-border px-4 py-3">
        <!-- Step 1 -->
        <template v-if="step === 'describe'">
          <CustomButton type="secondary" class="flex-1" :text="t('cancel') || 'Annuler'" @click="handleClose" />
          <CustomButton
            type="primary"
            class="flex-1"
            :text="t('generate') || 'Générer →'"
            :disabled="descriptionInput.trim().length < 5"
            @click="handleGenerate"
          />
        </template>

        <!-- Step 3 -->
        <template v-if="step === 'review'">
          <CustomButton type="secondary" :icon="RefreshCw" text="" :title="t('regenerate') || 'Régénérer'" class="shrink-0 p-2!" @click="handleRegenerate" />
          <CustomButton
            type="secondary"
            class="flex-1"
            :text="t('testSkill') || 'Tester ▶'"
            @click="handleTest"
          />
          <CustomButton
            type="primary"
            class="flex-1"
            :text="t('saveSkill') || 'Sauvegarder'"
            @click="handleSave"
          />
        </template>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onMounted } from 'vue';
import { useI18n } from 'vue-i18n';
import { ArrowLeft, PlayCircle, RefreshCw, X } from 'lucide-vue-next';
import CustomButton from '@/components/CustomButton.vue';
import { useSkillCreator } from '@/composables/useSkillCreator';
import { useUserSkills } from '@/composables/useUserSkills';
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser';
import type { UserSkill } from '@/types/userSkill';

const { t } = useI18n();

const emit = defineEmits<{
  (e: 'close'): void;
  (e: 'skill-created', skill: UserSkill): void;
  (e: 'test-skill', skill: Omit<UserSkill, 'id' | 'createdAt' | 'updatedAt'>): void;
}>();

// ── State ─────────────────────────────────────────────────────────────────────

type Step = 'describe' | 'generating' | 'review' | 'testing';
const step = ref<Step>('describe');

const currentStepIndex = computed(() =>
  ['describe', 'generating', 'review', 'testing'].indexOf(step.value),
);
const stepLabels = ['Décrire', 'Générer', 'Réviser', 'Tester'];

// Step 1
const descriptionInput = ref('');
const selectedHost = ref<SkillHost>('all');
const generationError = ref('');

// Step 3 editable fields
const editName = ref('');
const editDescription = ref('');
const editHost = ref<SkillHost>('all');
const editExecutionMode = ref<SkillExecutionMode>('immediate');
const editIcon = ref('Zap');
const editContent = ref('');

const { generating, error, generateSkill } = useSkillCreator();
const { addSkill } = useUserSkills();

const dialogRef = ref<HTMLElement | null>(null);
const descriptionRef = ref<HTMLTextAreaElement | null>(null);

onMounted(() => {
  nextTick(() => descriptionRef.value?.focus());
});

// ── Actions ───────────────────────────────────────────────────────────────────

async function handleGenerate(): Promise<void> {
  if (descriptionInput.value.trim().length < 5) return;
  generationError.value = '';
  step.value = 'generating';

  const result = await generateSkill(descriptionInput.value.trim(), selectedHost.value);

  if (!result) {
    generationError.value = error.value || 'Génération échouée. Réessayez.';
    step.value = 'describe';
    return;
  }

  editName.value = result.name;
  editDescription.value = result.description;
  editHost.value = result.host;
  editExecutionMode.value = result.executionMode;
  editIcon.value = result.icon;
  editContent.value = result.skillContent;
  step.value = 'review';
}

function handleRegenerate(): void {
  step.value = 'describe';
}

function handleTest(): void {
  step.value = 'testing';
  emit('test-skill', {
    name: editName.value,
    description: editDescription.value,
    host: editHost.value,
    executionMode: editExecutionMode.value,
    icon: editIcon.value,
    skillContent: editContent.value,
  });
}

function handleSave(): void {
  const skill = addSkill({
    name: editName.value,
    description: editDescription.value,
    host: editHost.value,
    executionMode: editExecutionMode.value,
    icon: editIcon.value,
    skillContent: editContent.value,
  });
  emit('skill-created', skill);
  emit('close');
  resetState();
}

function handleClose(): void {
  emit('close');
  resetState();
}

function resetState(): void {
  step.value = 'describe';
  descriptionInput.value = '';
  selectedHost.value = 'all';
  generationError.value = '';
}
</script>
