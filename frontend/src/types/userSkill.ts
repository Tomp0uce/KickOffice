import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser';

export type { SkillHost, SkillExecutionMode };

export interface UserSkill {
  id: string; // "skill_1710766200000"
  name: string; // "Reformuler en bullets"
  description: string; // Displayed in dropdown and library
  host: SkillHost;
  executionMode: SkillExecutionMode;
  icon: string; // Lucide icon name, default: "Zap"
  skillContent: string; // Markdown body (without frontmatter)
  createdAt: number; // timestamp ms
  updatedAt: number; // timestamp ms
}

export const SKILL_STORAGE_KEY = 'ki_UserSkills_v1';
export const SKILL_MIGRATION_KEY = 'ki_UserSkillsMigrated_v1';
