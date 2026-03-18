import { describe, it, expect, beforeEach, vi } from 'vitest';
import { useUserSkills, isValidUserSkill, _resetUserSkillsState } from '../useUserSkills';
import { SKILL_STORAGE_KEY, SKILL_MIGRATION_KEY } from '@/types/userSkill';

// ── localStorage mock ─────────────────────────────────────────────────────────

const localStorageMock = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: vi.fn((key: string) => store[key] ?? null),
    setItem: vi.fn((key: string, value: string) => { store[key] = value; }),
    removeItem: vi.fn((key: string) => { delete store[key]; }),
    clear: vi.fn(() => { store = {}; }),
    _store: () => store,
  };
})();
vi.stubGlobal('localStorage', localStorageMock);

// ── Helpers ───────────────────────────────────────────────────────────────────

function freshComposable() {
  _resetUserSkillsState();
  return useUserSkills();
}

function makeSkill(overrides = {}) {
  return {
    name: 'Test Skill',
    description: 'Does something useful',
    host: 'word' as const,
    executionMode: 'immediate' as const,
    icon: 'Zap',
    skillContent: '## Instructions\n\nDo something.',
    ...overrides,
  };
}

// ── Setup ─────────────────────────────────────────────────────────────────────

beforeEach(() => {
  localStorageMock.clear();
  vi.clearAllMocks();
  _resetUserSkillsState();
});

// ── addSkill ──────────────────────────────────────────────────────────────────

describe('addSkill', () => {
  it('adds a skill with generated id and timestamps', () => {
    const { skills, addSkill } = freshComposable();
    const skill = addSkill(makeSkill());
    expect(skills.value).toHaveLength(1);
    expect(skill.id).toMatch(/^skill_\d+_[a-z0-9]+$/);
    expect(skill.createdAt).toBeGreaterThan(0);
    expect(skill.updatedAt).toBeGreaterThan(0);
  });

  it('persists to localStorage', () => {
    const { addSkill } = freshComposable();
    addSkill(makeSkill());
    expect(localStorageMock.setItem).toHaveBeenCalledWith(
      SKILL_STORAGE_KEY,
      expect.stringContaining('Test Skill'),
    );
  });

  it('adding multiple skills accumulates them', () => {
    const { skills, addSkill } = freshComposable();
    addSkill(makeSkill({ name: 'A' }));
    addSkill(makeSkill({ name: 'B' }));
    expect(skills.value).toHaveLength(2);
  });

  it('returns the created skill', () => {
    const { addSkill } = freshComposable();
    const result = addSkill(makeSkill({ name: 'Returned' }));
    expect(result.name).toBe('Returned');
    expect(result.host).toBe('word');
  });
});

// ── updateSkill ───────────────────────────────────────────────────────────────

describe('updateSkill', () => {
  it('updates an existing skill by id', () => {
    const { skills, addSkill, updateSkill } = freshComposable();
    const skill = addSkill(makeSkill({ name: 'Original' }));
    updateSkill(skill.id, { name: 'Updated' });
    expect(skills.value[0].name).toBe('Updated');
  });

  it('updates updatedAt but not createdAt', () => {
    const { skills, addSkill, updateSkill } = freshComposable();
    const skill = addSkill(makeSkill());
    const originalCreatedAt = skills.value[0].createdAt;
    updateSkill(skill.id, { name: 'Changed' });
    expect(skills.value[0].createdAt).toBe(originalCreatedAt);
    expect(skills.value[0].updatedAt).toBeGreaterThanOrEqual(originalCreatedAt);
  });

  it('is a no-op for unknown id', () => {
    const { skills, addSkill, updateSkill } = freshComposable();
    addSkill(makeSkill());
    updateSkill('nonexistent-id', { name: 'Ghost' });
    expect(skills.value[0].name).toBe('Test Skill');
  });

  it('persists changes to localStorage', () => {
    const { addSkill, updateSkill } = freshComposable();
    const skill = addSkill(makeSkill());
    vi.clearAllMocks();
    updateSkill(skill.id, { name: 'New Name' });
    expect(localStorageMock.setItem).toHaveBeenCalledWith(
      SKILL_STORAGE_KEY,
      expect.stringContaining('New Name'),
    );
  });
});

// ── deleteSkill ───────────────────────────────────────────────────────────────

describe('deleteSkill', () => {
  it('removes the skill with matching id', () => {
    const { skills, addSkill, deleteSkill } = freshComposable();
    const skill = addSkill(makeSkill());
    deleteSkill(skill.id);
    expect(skills.value).toHaveLength(0);
  });

  it('leaves other skills intact', () => {
    const { skills, addSkill, deleteSkill } = freshComposable();
    const a = addSkill(makeSkill({ name: 'A' }));
    addSkill(makeSkill({ name: 'B' }));
    deleteSkill(a.id);
    expect(skills.value).toHaveLength(1);
    expect(skills.value[0].name).toBe('B');
  });

  it('is a no-op for unknown id', () => {
    const { skills, addSkill, deleteSkill } = freshComposable();
    addSkill(makeSkill());
    deleteSkill('ghost');
    expect(skills.value).toHaveLength(1);
  });
});

// ── skillsForHost ─────────────────────────────────────────────────────────────

describe('skillsForHost', () => {
  it('returns skills matching the host', () => {
    const { addSkill, skillsForHost } = freshComposable();
    addSkill(makeSkill({ host: 'word' }));
    addSkill(makeSkill({ host: 'excel' }));
    const wordSkills = skillsForHost('word');
    expect(wordSkills.value).toHaveLength(1);
    expect(wordSkills.value[0].host).toBe('word');
  });

  it('always includes "all" skills regardless of host filter', () => {
    const { addSkill, skillsForHost } = freshComposable();
    addSkill(makeSkill({ host: 'all', name: 'Universal' }));
    addSkill(makeSkill({ host: 'word', name: 'Word Only' }));
    const excelSkills = skillsForHost('excel');
    expect(excelSkills.value).toHaveLength(1);
    expect(excelSkills.value[0].name).toBe('Universal');
  });

  it('returns empty array when no skills match', () => {
    const { addSkill, skillsForHost } = freshComposable();
    addSkill(makeSkill({ host: 'word' }));
    expect(skillsForHost('outlook').value).toHaveLength(0);
  });

  it('is reactive — updates when a skill is added', () => {
    const { addSkill, skillsForHost } = freshComposable();
    const pptSkills = skillsForHost('powerpoint');
    expect(pptSkills.value).toHaveLength(0);
    addSkill(makeSkill({ host: 'powerpoint' }));
    expect(pptSkills.value).toHaveLength(1);
  });
});

// ── localStorage persistence ──────────────────────────────────────────────────

describe('persistence', () => {
  it('loads skills from localStorage on init', () => {
    const stored = JSON.stringify([
      { id: 'skill_1', name: 'Saved', description: '', host: 'word',
        executionMode: 'immediate', icon: 'Zap', skillContent: '', createdAt: 1, updatedAt: 1 },
    ]);
    localStorageMock.setItem(SKILL_STORAGE_KEY, stored);

    const { skills } = freshComposable();
    expect(skills.value).toHaveLength(1);
    expect(skills.value[0].name).toBe('Saved');
  });

  it('ignores invalid items in storage', () => {
    const stored = JSON.stringify([
      { id: 'skill_1', name: 'Valid', description: '', host: 'word',
        executionMode: 'immediate', icon: 'Zap', skillContent: '', createdAt: 1, updatedAt: 1 },
      { notASkill: true },
      null,
    ]);
    localStorageMock.setItem(SKILL_STORAGE_KEY, stored);

    const { skills } = freshComposable();
    expect(skills.value).toHaveLength(1);
  });

  it('recovers from corrupted localStorage gracefully', () => {
    localStorageMock.setItem(SKILL_STORAGE_KEY, '{not valid json}');
    const { skills } = freshComposable();
    expect(skills.value).toHaveLength(0);
  });
});

// ── exportSkillToFile ─────────────────────────────────────────────────────────

describe('exportSkillToFile', () => {
  it('creates a blob and triggers download with .skill.md extension', () => {
    const createObjectURL = vi.fn(() => 'blob:mock-url');
    const revokeObjectURL = vi.fn();
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL });

    const mockAnchor = { href: '', download: '', click: vi.fn() };
    vi.spyOn(document, 'createElement').mockReturnValueOnce(mockAnchor as any);
    vi.spyOn(document.body, 'appendChild').mockImplementation(() => mockAnchor as any);
    vi.spyOn(document.body, 'removeChild').mockImplementation(() => mockAnchor as any);

    const { addSkill, exportSkillToFile } = freshComposable();
    const skill = addSkill(makeSkill({ name: 'My Export Skill' }));
    exportSkillToFile(skill);

    expect(mockAnchor.download).toMatch(/\.skill\.md$/);
    expect(mockAnchor.click).toHaveBeenCalled();
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
  });
});

// ── importSkillFromFile ───────────────────────────────────────────────────────

describe('importSkillFromFile', () => {
  it('imports a valid .skill.md file and adds it to the list', async () => {
    const content = `---
name: Imported Skill
description: "Imported from file."
host: excel
executionMode: agent
icon: Database
---

## Instructions
Do the thing.`;
    const file = new File([content], 'imported.skill.md', { type: 'text/markdown' });

    const { skills, importSkillFromFile } = freshComposable();
    const result = await importSkillFromFile(file);

    expect(result).not.toBeNull();
    expect(result!.name).toBe('Imported Skill');
    expect(result!.host).toBe('excel');
    expect(result!.executionMode).toBe('agent');
    expect(skills.value).toHaveLength(1);
  });

  it('generates a new id on import (no collision with original id)', async () => {
    const content = `---
name: Skill With Old Id
description: ""
host: word
executionMode: immediate
icon: Zap
---

Body.`;
    const file = new File([content], 'skill.md', { type: 'text/markdown' });
    const { importSkillFromFile } = freshComposable();
    const result = await importSkillFromFile(file);
    expect(result!.id).toMatch(/^skill_\d+_[a-z0-9]+$/);
  });

  it('returns null for a file that throws during read', async () => {
    const file = { text: vi.fn().mockRejectedValue(new Error('Read error')) } as unknown as File;
    const { importSkillFromFile } = freshComposable();
    const result = await importSkillFromFile(file);
    expect(result).toBeNull();
  });
});

// ── migration ─────────────────────────────────────────────────────────────────

describe('migration from custom prompts', () => {
  it('checkAndMigrateOldPrompts returns false when already migrated', () => {
    localStorageMock.setItem(SKILL_MIGRATION_KEY, 'done');
    const { checkAndMigrateOldPrompts } = freshComposable();
    expect(checkAndMigrateOldPrompts()).toBe(false);
  });

  it('returns false when no savedPrompts in storage', () => {
    const { checkAndMigrateOldPrompts } = freshComposable();
    expect(checkAndMigrateOldPrompts()).toBe(false);
    expect(localStorageMock.setItem).toHaveBeenCalledWith(SKILL_MIGRATION_KEY, 'done');
  });

  it('returns true when real prompts exist', () => {
    const prompts = [{ id: '1', name: 'My Prompt', systemPrompt: 'Be helpful', userPrompt: '' }];
    localStorageMock.setItem('savedPrompts', JSON.stringify(prompts));
    const { checkAndMigrateOldPrompts } = freshComposable();
    expect(checkAndMigrateOldPrompts()).toBe(true);
  });

  it('returns false for only the empty Default prompt', () => {
    const prompts = [{ id: 'default', name: 'Default', systemPrompt: '', userPrompt: '' }];
    localStorageMock.setItem('savedPrompts', JSON.stringify(prompts));
    const { checkAndMigrateOldPrompts } = freshComposable();
    expect(checkAndMigrateOldPrompts()).toBe(false);
  });

  it('migrateOldPrompts converts prompts to skills', () => {
    const prompts = [
      { id: '1', name: 'Translator', systemPrompt: 'You are a translator', userPrompt: 'Translate: [TEXT]' },
      { id: '2', name: 'Summarizer', systemPrompt: 'Summarize', userPrompt: '' },
    ];
    localStorageMock.setItem('savedPrompts', JSON.stringify(prompts));
    const { skills, migrateOldPrompts } = freshComposable();
    migrateOldPrompts();
    expect(skills.value).toHaveLength(2);
    expect(skills.value[0].name).toBe('Translator');
    expect(skills.value[0].host).toBe('all');
    expect(skills.value[0].executionMode).toBe('immediate');
    expect(skills.value[0].skillContent).toContain('You are a translator');
    expect(skills.value[0].skillContent).toContain('Translate: [TEXT]');
  });

  it('migrateOldPrompts skips empty Default prompt', () => {
    const prompts = [
      { id: 'default', name: 'Default', systemPrompt: '', userPrompt: '' },
      { id: '1', name: 'Real', systemPrompt: 'Instructions', userPrompt: '' },
    ];
    localStorageMock.setItem('savedPrompts', JSON.stringify(prompts));
    const { skills, migrateOldPrompts } = freshComposable();
    migrateOldPrompts();
    expect(skills.value).toHaveLength(1);
    expect(skills.value[0].name).toBe('Real');
  });

  it('confirmMigrationDone removes savedPrompts and sets migration flag', () => {
    localStorageMock.setItem('savedPrompts', '[]');
    const { confirmMigrationDone } = freshComposable();
    confirmMigrationDone();
    expect(localStorageMock.removeItem).toHaveBeenCalledWith('savedPrompts');
    expect(localStorageMock.setItem).toHaveBeenCalledWith(SKILL_MIGRATION_KEY, 'done');
  });

  it('dismissMigration sets migration flag without removing savedPrompts', () => {
    localStorageMock.setItem('savedPrompts', '[]');
    const { dismissMigration } = freshComposable();
    dismissMigration();
    expect(localStorageMock.removeItem).not.toHaveBeenCalledWith('savedPrompts');
    expect(localStorageMock.setItem).toHaveBeenCalledWith(SKILL_MIGRATION_KEY, 'done');
  });
});

// ── isValidUserSkill ──────────────────────────────────────────────────────────

describe('isValidUserSkill', () => {
  it('accepts a valid UserSkill', () => {
    expect(isValidUserSkill({
      id: 'skill_1', name: 'Test', description: '', host: 'word',
      executionMode: 'agent', icon: 'Zap', skillContent: '', createdAt: 1, updatedAt: 1,
    })).toBe(true);
  });

  it('rejects missing required fields', () => {
    expect(isValidUserSkill({ name: 'No id' })).toBe(false);
    expect(isValidUserSkill({ id: 'x', host: 'word' })).toBe(false);
  });

  it('rejects invalid host value', () => {
    expect(isValidUserSkill({
      id: 'x', name: 'T', description: '', host: 'browser',
      executionMode: 'immediate', icon: 'Zap', skillContent: '', createdAt: 1, updatedAt: 1,
    })).toBe(false);
  });

  it('rejects invalid executionMode', () => {
    expect(isValidUserSkill({
      id: 'x', name: 'T', description: '', host: 'word',
      executionMode: 'stream', icon: 'Zap', skillContent: '', createdAt: 1, updatedAt: 1,
    })).toBe(false);
  });

  it('rejects null and non-objects', () => {
    expect(isValidUserSkill(null)).toBe(false);
    expect(isValidUserSkill('string')).toBe(false);
    expect(isValidUserSkill(42)).toBe(false);
  });
});
