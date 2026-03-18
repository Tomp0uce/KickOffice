/**
 * skillParser.ts
 *
 * Parses .skill.md files with YAML frontmatter.
 * Lightweight parser — no external dependency.
 * Supports flat key: value pairs (quoted or unquoted), no nested objects.
 */

export type SkillHost = 'word' | 'excel' | 'powerpoint' | 'outlook' | 'all';
export type SkillExecutionMode = 'immediate' | 'draft' | 'agent';

export interface SkillMetadata {
  name: string;
  description: string;
  host: SkillHost;
  executionMode: SkillExecutionMode;
  icon: string;
  actionKey?: string; // built-in skills only
}

export interface ParsedSkill {
  metadata: SkillMetadata;
  body: string; // markdown body without frontmatter
  raw: string; // full file content (for system prompt injection)
}

export const VALID_HOSTS: SkillHost[] = ['word', 'excel', 'powerpoint', 'outlook', 'all'];
export const VALID_EXECUTION_MODES: SkillExecutionMode[] = ['immediate', 'draft', 'agent'];

/**
 * Parse a raw .skill.md string into metadata + body.
 * Falls back to default metadata if frontmatter is missing or invalid.
 */
export function parseSkill(raw: string, fallbackActionKey?: string): ParsedSkill {
  const match = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/);

  if (!match) {
    return {
      metadata: {
        name: fallbackActionKey ?? 'Unnamed Skill',
        description: '',
        host: 'all',
        executionMode: 'immediate',
        icon: 'Zap',
        actionKey: fallbackActionKey,
      },
      body: raw,
      raw,
    };
  }

  const frontmatterStr = match[1];
  const body = match[2].trim();
  const metadata = parseFrontmatter(frontmatterStr, fallbackActionKey);

  return { metadata, body, raw };
}

/**
 * Minimal YAML parser for flat key: value pairs.
 * Handles quoted strings (single or double).
 */
function parseFrontmatter(yaml: string, fallbackActionKey?: string): SkillMetadata {
  const obj: Record<string, string> = {};

  for (const line of yaml.split('\n')) {
    const colonIdx = line.indexOf(':');
    if (colonIdx === -1) continue;
    const key = line.slice(0, colonIdx).trim();
    if (!key) continue;
    let val = line.slice(colonIdx + 1).trim();
    // Strip surrounding quotes (single or double)
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    // Unescape escaped double quotes inside the value
    val = val.replace(/\\"/g, '"');
    obj[key] = val;
  }

  return {
    name: obj['name'] ?? fallbackActionKey ?? 'Unnamed',
    description: obj['description'] ?? '',
    host: VALID_HOSTS.includes(obj['host'] as SkillHost) ? (obj['host'] as SkillHost) : 'all',
    executionMode: VALID_EXECUTION_MODES.includes(obj['executionMode'] as SkillExecutionMode)
      ? (obj['executionMode'] as SkillExecutionMode)
      : 'immediate',
    icon: obj['icon'] ?? 'Zap',
    actionKey: obj['actionKey'] || fallbackActionKey,
  };
}

/**
 * Serialize a skill back to .skill.md format (for export).
 */
export function serializeSkillToMd(skill: {
  name: string;
  description: string;
  host: SkillHost;
  executionMode: SkillExecutionMode;
  icon: string;
  skillContent: string;
  actionKey?: string;
}): string {
  const lines = [
    '---',
    `name: ${skill.name}`,
    `description: "${skill.description.replace(/"/g, '\\"')}"`,
    `host: ${skill.host}`,
    `executionMode: ${skill.executionMode}`,
    `icon: ${skill.icon}`,
  ];
  if (skill.actionKey) lines.push(`actionKey: ${skill.actionKey}`);
  lines.push('---', '', skill.skillContent);
  return lines.join('\n');
}
