/**
 * Skills Loader
 *
 * Loads and combines skill documents for injection into agent prompts.
 * Skills are defensive prompting guidelines that prevent common Office.js errors.
 */

import { logService } from '@/utils/logger'

// Import skill documents as raw strings
// Note: Vite supports ?raw suffix for importing file contents
import commonSkill from './common.skill.md?raw'
import wordSkill from './word.skill.md?raw'
import excelSkill from './excel.skill.md?raw'
import powerpointSkill from './powerpoint.skill.md?raw'
import outlookSkill from './outlook.skill.md?raw'

// Import Quick Action skills
import bulletsSkill from './quickactions/bullets.skill.md?raw'
import punchifySkill from './quickactions/punchify.skill.md?raw'
import reviewSkill from './quickactions/review.skill.md?raw'
import translateSkill from './quickactions/translate.skill.md?raw'
import formalizeSkill from './quickactions/formalize.skill.md?raw'
import conciseSkill from './quickactions/concise.skill.md?raw'
import proofreadSkill from './quickactions/proofread.skill.md?raw'
import polishSkill from './quickactions/polish.skill.md?raw'
import academicSkill from './quickactions/academic.skill.md?raw'
import summarySkill from './quickactions/summary.skill.md?raw'
import extractSkill from './quickactions/extract.skill.md?raw'
import replySkill from './quickactions/reply.skill.md?raw'
import ingestSkill from './quickactions/ingest.skill.md?raw'
import autographSkill from './quickactions/autograph.skill.md?raw'
import explainExcelSkill from './quickactions/explain-excel.skill.md?raw'
import formulaGeneratorSkill from './quickactions/formula-generator.skill.md?raw'
import dataTrendSkill from './quickactions/data-trend.skill.md?raw'

export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'
export type QuickActionKey = 'bullets' | 'punchify' | 'review' | 'visual' | 'translate' | 'formalize' | 'concise' | 'proofread'

const hostSkillMap: Record<OfficeHost, string> = {
  Word: wordSkill,
  Excel: excelSkill,
  PowerPoint: powerpointSkill,
  Outlook: outlookSkill,
}

/**
 * Get the combined skill document for a specific Office host.
 *
 * @param host - The Office application (Word, Excel, PowerPoint, Outlook)
 * @returns Combined skill markdown (common rules + host-specific rules)
 */
export function getSkillForHost(host: OfficeHost): string {
  const hostSkill = hostSkillMap[host]

  if (!hostSkill) {
    logService.warn(`[Skills] Unknown host: ${host}, using common skills only`)
    return commonSkill
  }

  return `${commonSkill}\n\n---\n\n${hostSkill}`
}

/**
 * Get just the common skill document (shared rules).
 */
export function getCommonSkill(): string {
  return commonSkill
}

/**
 * Get just the host-specific skill document (without common rules).
 */
export function getHostSpecificSkill(host: OfficeHost): string {
  return hostSkillMap[host] || ''
}

/**
 * List all available hosts.
 */
export function getAvailableHosts(): OfficeHost[] {
  return ['Word', 'Excel', 'PowerPoint', 'Outlook']
}

/**
 * Quick Action skills map.
 */
const quickActionSkillMap: Record<string, string> = {
  // PowerPoint
  bullets: bulletsSkill,
  punchify: punchifySkill,
  review: reviewSkill,

  // Word
  translate: translateSkill,
  formalize: formalizeSkill,
  concise: conciseSkill,
  proofread: proofreadSkill,
  polish: polishSkill,
  academic: academicSkill,
  summary: summarySkill,

  // Outlook
  extract: extractSkill,
  reply: replySkill,

  // Excel
  ingest: ingestSkill,
  autograph: autographSkill,
  explain: explainExcelSkill,
  formulaGenerator: formulaGeneratorSkill,
  dataTrend: dataTrendSkill,
}

/**
 * Get the skill document for a specific Quick Action.
 *
 * @param actionKey - The Quick Action key (e.g., 'bullets', 'translate', 'review')
 * @returns Skill markdown for the Quick Action, or undefined if not found
 *
 * @example
 * const skill = getQuickActionSkill('bullets')
 * if (skill) {
 *   // Inject as system message: { role: 'system', content: skill }
 * }
 */
export function getQuickActionSkill(actionKey: string): string | undefined {
  return quickActionSkillMap[actionKey]
}

/**
 * Check if a Quick Action has a corresponding skill file.
 *
 * @param actionKey - The Quick Action key
 * @returns true if a skill exists for this action
 */
export function hasQuickActionSkill(actionKey: string): boolean {
  return actionKey in quickActionSkillMap
}

/**
 * List all available Quick Action skills.
 */
export function getAvailableQuickActionSkills(): string[] {
  return Object.keys(quickActionSkillMap)
}
