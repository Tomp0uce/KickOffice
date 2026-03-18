/**
 * useUserSkills.ts
 *
 * Manages user-created skills: CRUD, localStorage persistence,
 * export/import as .skill.md, and migration from legacy custom prompts.
 *
 * Uses module-level shared state so all callers see the same skills list
 * (QuickActionsBar, SkillLibraryTab, SkillCreatorModal).
 */

import { ref, computed } from 'vue';
import { logService } from '@/utils/logger';
import { parseSkill, serializeSkillToMd, VALID_HOSTS, VALID_EXECUTION_MODES } from '@/utils/skillParser';
import type { SkillHost } from '@/utils/skillParser';
import type { UserSkill } from '@/types/userSkill';
import { SKILL_STORAGE_KEY, SKILL_MIGRATION_KEY } from '@/types/userSkill';

// ── Module-level shared state ─────────────────────────────────────────────────

const skills = ref<UserSkill[]>([]);
let _loaded = false;

function loadFromStorage(): void {
  if (_loaded) return;
  _loaded = true;

  const stored = localStorage.getItem(SKILL_STORAGE_KEY);
  if (!stored) return;

  try {
    const parsed = JSON.parse(stored);
    skills.value = Array.isArray(parsed) ? parsed.filter(isValidUserSkill) : [];
  } catch (err) {
    logService.warn('[UserSkills] Failed to parse stored skills', err);
    skills.value = [];
  }
}

function saveToStorage(): void {
  try {
    localStorage.setItem(SKILL_STORAGE_KEY, JSON.stringify(skills.value));
  } catch (e) {
    if (e instanceof DOMException && e.name === 'QuotaExceededError') {
      logService.warn('[UserSkills] localStorage quota exceeded — skills not persisted');
    } else {
      throw e;
    }
  }
}

// ── Composable ────────────────────────────────────────────────────────────────

export function useUserSkills() {
  loadFromStorage();

  // ── CRUD ────────────────────────────────────────────────────────────────────

  function addSkill(skill: Omit<UserSkill, 'id' | 'createdAt' | 'updatedAt'>): UserSkill {
    const newSkill: UserSkill = {
      ...skill,
      id: `skill_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
    skills.value = [...skills.value, newSkill];
    saveToStorage();
    return newSkill;
  }

  function updateSkill(id: string, updates: Partial<Omit<UserSkill, 'id' | 'createdAt'>>): void {
    const idx = skills.value.findIndex((s) => s.id === id);
    if (idx === -1) return;
    const updated = [...skills.value];
    updated[idx] = { ...updated[idx], ...updates, updatedAt: Date.now() };
    skills.value = updated;
    saveToStorage();
  }

  function deleteSkill(id: string): void {
    skills.value = skills.value.filter((s) => s.id !== id);
    saveToStorage();
  }

  // ── Filtered view ────────────────────────────────────────────────────────────

  function skillsForHost(host: SkillHost) {
    return computed(() => skills.value.filter((s) => s.host === host || s.host === 'all'));
  }

  // ── Export / Import ──────────────────────────────────────────────────────────

  function exportSkillToFile(skill: UserSkill): void {
    const content = serializeSkillToMd({
      name: skill.name,
      description: skill.description,
      host: skill.host,
      executionMode: skill.executionMode,
      icon: skill.icon,
      skillContent: skill.skillContent,
    });
    const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${skill.name.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '')}.skill.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  async function importSkillFromFile(file: File): Promise<UserSkill | null> {
    try {
      const text = await file.text();
      const parsed = parseSkill(text);
      return addSkill({
        name: parsed.metadata.name,
        description: parsed.metadata.description,
        host: parsed.metadata.host,
        executionMode: parsed.metadata.executionMode,
        icon: parsed.metadata.icon ?? 'Zap',
        skillContent: parsed.body || text, // fallback: use full content if no frontmatter
      });
    } catch (err) {
      logService.warn('[UserSkills] Failed to import skill file', err);
      return null;
    }
  }

  // ── Migration from Custom Prompts ────────────────────────────────────────────

  /**
   * Returns true if there are legacy custom prompts to migrate.
   * Caller is responsible for showing the migration dialog.
   */
  function checkAndMigrateOldPrompts(): boolean {
    if (localStorage.getItem(SKILL_MIGRATION_KEY)) return false;

    const stored = localStorage.getItem('savedPrompts');
    if (!stored) {
      localStorage.setItem(SKILL_MIGRATION_KEY, 'done');
      return false;
    }

    try {
      const prompts = JSON.parse(stored);
      if (!Array.isArray(prompts) || prompts.length === 0) {
        localStorage.setItem(SKILL_MIGRATION_KEY, 'done');
        return false;
      }
      // Skip the default empty placeholder prompt
      const real = prompts.filter(
        (p: any) => p.name !== 'Default' || p.systemPrompt || p.userPrompt,
      );
      return real.length > 0;
    } catch {
      return false;
    }
  }

  /** Convert legacy custom prompts to user skills and persist. */
  function migrateOldPrompts(): void {
    const stored = localStorage.getItem('savedPrompts');
    if (!stored) return;

    try {
      const prompts = JSON.parse(stored) as Array<{
        id: string;
        name: string;
        systemPrompt: string;
        userPrompt: string;
      }>;

      const newSkills: UserSkill[] = [];
      for (const p of prompts) {
        if (p.name === 'Default' && !p.systemPrompt && !p.userPrompt) continue;
        const lines: string[] = [];
        if (p.systemPrompt) {
          lines.push('## Instructions système', '', p.systemPrompt);
        }
        if (p.userPrompt) {
          lines.push('', '## Message type', '', p.userPrompt);
        }
        newSkills.push({
          id: `skill_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
          name: p.name,
          description: p.name,
          host: 'all',
          executionMode: 'immediate',
          icon: 'Zap',
          skillContent: lines.join('\n').trim(),
          createdAt: Date.now(),
          updatedAt: Date.now(),
        });
      }

      if (newSkills.length > 0) {
        skills.value = [...skills.value, ...newSkills];
        saveToStorage(); // single write instead of N writes
      }
    } catch (err) {
      logService.warn('[UserSkills] Migration failed', err);
    }
  }

  /** Call after user confirms migration — cleans up legacy storage. */
  function confirmMigrationDone(): void {
    localStorage.removeItem('savedPrompts');
    localStorage.setItem(SKILL_MIGRATION_KEY, 'done');
  }

  /** Call to skip migration without converting. */
  function dismissMigration(): void {
    localStorage.setItem(SKILL_MIGRATION_KEY, 'done');
  }

  return {
    skills,
    skillsForHost,
    addSkill,
    updateSkill,
    deleteSkill,
    exportSkillToFile,
    importSkillFromFile,
    checkAndMigrateOldPrompts,
    migrateOldPrompts,
    confirmMigrationDone,
    dismissMigration,
  };
}

// ── Validation ────────────────────────────────────────────────────────────────

export function isValidUserSkill(item: unknown): item is UserSkill {
  if (!item || typeof item !== 'object') return false;
  const o = item as Record<string, unknown>;
  return (
    typeof o['id'] === 'string' &&
    typeof o['name'] === 'string' &&
    typeof o['skillContent'] === 'string' &&
    VALID_HOSTS.includes(o['host'] as SkillHost) &&
    VALID_EXECUTION_MODES.includes(o['executionMode'] as 'immediate' | 'draft' | 'agent')
  );
}

/** Reset module state — for testing only. */
export function _resetUserSkillsState(): void {
  skills.value = [];
  _loaded = false;
}
