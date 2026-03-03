/**
 * Skills Loader
 *
 * Loads and combines skill documents for injection into agent prompts.
 * Skills are defensive prompting guidelines that prevent common Office.js errors.
 */

// Import skill documents as raw strings
// Note: Vite supports ?raw suffix for importing file contents
import commonSkill from './common.skill.md?raw'
import wordSkill from './word.skill.md?raw'
import excelSkill from './excel.skill.md?raw'
import powerpointSkill from './powerpoint.skill.md?raw'
import outlookSkill from './outlook.skill.md?raw'

export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'

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
    console.warn(`[Skills] Unknown host: ${host}, using common skills only`)
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
