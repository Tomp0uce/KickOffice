/**
 * Skills Loader — metadata-driven.
 *
 * Built-in skills are imported via Vite ?raw at build-time.
 * Frontmatter is parsed at module load (one-time cost, not per-request).
 * getQuickActionSkill() signature is unchanged for backwards compatibility.
 */

import { logService } from '@/utils/logger';
import { parseSkill } from '@/utils/skillParser';
import type { ParsedSkill } from '@/utils/skillParser';

// Host skills (no frontmatter — always injected by host, not user-selectable)
import commonSkill from './common.skill.md?raw';
import wordSkill from './word.skill.md?raw';
import excelSkill from './excel.skill.md?raw';
import powerpointSkill from './powerpoint.skill.md?raw';
import outlookSkill from './outlook.skill.md?raw';

// Quick Action skills — raw imports
import bulletsSkillRaw from './quickactions/bullets.skill.md?raw';
import punchifySkillRaw from './quickactions/punchify.skill.md?raw';
import reviewSkillRaw from './quickactions/review.skill.md?raw';
import translateSkillRaw from './quickactions/translate.skill.md?raw';
import formalizeSkillRaw from './quickactions/formalize.skill.md?raw';
import conciseSkillRaw from './quickactions/concise.skill.md?raw';
import proofreadSkillRaw from './quickactions/proofread.skill.md?raw';
import pptProofreadSkillRaw from './quickactions/ppt-proofread.skill.md?raw';
import pptTranslateSkillRaw from './quickactions/ppt-translate.skill.md?raw';
import wordTranslateSkillRaw from './quickactions/word-translate.skill.md?raw';
import wordProofreadSkillRaw from './quickactions/word-proofread.skill.md?raw';
import wordReviewSkillRaw from './quickactions/word-review.skill.md?raw';
import polishSkillRaw from './quickactions/polish.skill.md?raw';
import academicSkillRaw from './quickactions/academic.skill.md?raw';
import summarySkillRaw from './quickactions/summary.skill.md?raw';
import extractSkillRaw from './quickactions/extract.skill.md?raw';
import replySkillRaw from './quickactions/reply.skill.md?raw';
import ingestSkillRaw from './quickactions/ingest.skill.md?raw';
import autographSkillRaw from './quickactions/autograph.skill.md?raw';
import chartDigitizerSkillRaw from './quickactions/chart-digitizer.skill.md?raw';
import pixelArtSkillRaw from './quickactions/pixel-art.skill.md?raw';
import explainExcelSkillRaw from './quickactions/explain-excel.skill.md?raw';
import formulaGeneratorSkillRaw from './quickactions/formula-generator.skill.md?raw';
import dataTrendSkillRaw from './quickactions/data-trend.skill.md?raw';

export type { SkillHost, SkillExecutionMode, ParsedSkill, SkillMetadata } from '@/utils/skillParser';

/** Backwards-compatible alias — capitalized host names used by useAgentPrompts and host skills. */
export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook';

// ── Internal registry ─────────────────────────────────────────────────────────

/** Raw skill files keyed by actionKey — order matches original quickActionSkillMap. */
const rawSkillFiles: Record<string, string> = {
  // PowerPoint
  bullets: bulletsSkillRaw,
  punchify: punchifySkillRaw,
  review: reviewSkillRaw,
  'ppt-proofread': pptProofreadSkillRaw,
  'ppt-translate': pptTranslateSkillRaw,
  // Word (agent-based, surgical)
  'word-translate': wordTranslateSkillRaw,
  'word-proofread': wordProofreadSkillRaw,
  'word-review': wordReviewSkillRaw,
  // Word / Outlook (non-agent)
  translate: translateSkillRaw,
  formalize: formalizeSkillRaw,
  concise: conciseSkillRaw,
  proofread: proofreadSkillRaw,
  polish: polishSkillRaw,
  academic: academicSkillRaw,
  summary: summarySkillRaw,
  // Outlook
  extract: extractSkillRaw,
  reply: replySkillRaw,
  // Excel
  ingest: ingestSkillRaw,
  autograph: autographSkillRaw,
  digitizeChart: chartDigitizerSkillRaw,
  pixelArt: pixelArtSkillRaw,
  explain: explainExcelSkillRaw,
  formulaGenerator: formulaGeneratorSkillRaw,
  dataTrend: dataTrendSkillRaw,
};

/** Parsed skills registry — populated once at module load. */
const parsedSkills: Map<string, ParsedSkill> = new Map(
  Object.entries(rawSkillFiles).map(([key, raw]) => [key, parseSkill(raw, key)]),
);

const hostSkillMap: Record<string, string> = {
  Word: wordSkill,
  Excel: excelSkill,
  PowerPoint: powerpointSkill,
  Outlook: outlookSkill,
};

// ── Host skills ───────────────────────────────────────────────────────────────

/**
 * Get the combined skill document for a specific Office host.
 * Returns common rules + host-specific rules concatenated.
 */
export function getSkillForHost(host: string): string {
  const hostSkill = hostSkillMap[host];
  if (!hostSkill) {
    logService.warn(`[Skills] Unknown host: ${host}, using common skills only`);
    return commonSkill;
  }
  return `${commonSkill}\n\n---\n\n${hostSkill}`;
}

export function getCommonSkill(): string {
  return commonSkill;
}

export function getHostSpecificSkill(host: string): string {
  return hostSkillMap[host] || '';
}

export function getAvailableHosts(): string[] {
  return ['Word', 'Excel', 'PowerPoint', 'Outlook'];
}

// ── Quick Action skills ───────────────────────────────────────────────────────

/**
 * Get the full raw .skill.md content for a Quick Action (for system prompt injection).
 * Signature unchanged — backwards compatible with useQuickActions.ts.
 */
export function getQuickActionSkill(actionKey: string): string | undefined {
  return parsedSkills.get(actionKey)?.raw;
}

/**
 * Get the parsed metadata for a Quick Action skill (for UI display).
 * Returns undefined if the action key has no corresponding skill.
 */
export function getQuickActionSkillMetadata(actionKey: string) {
  return parsedSkills.get(actionKey)?.metadata;
}

/**
 * Get metadata for all built-in Quick Action skills.
 * Useful for displaying a catalog of available built-in skills.
 */
export function getAllBuiltInSkillsMetadata() {
  return Array.from(parsedSkills.values()).map((s) => s.metadata);
}

export function hasQuickActionSkill(actionKey: string): boolean {
  return parsedSkills.has(actionKey);
}

export function getAvailableQuickActionSkills(): string[] {
  return Array.from(parsedSkills.keys());
}
