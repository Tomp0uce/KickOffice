import { describe, it, expect } from 'vitest';
import { parseSkill, serializeSkillToMd } from '@/utils/skillParser';

// ── Fixtures ──────────────────────────────────────────────────────────────────

const FULL_SKILL_MD = `---
name: Traduire le texte
description: "Traduit le texte sélectionné entre français et anglais."
host: all
executionMode: immediate
icon: Languages
actionKey: translate
---

Traduis le texte reçu vers l'autre langue (FR ↔ EN).

## Règles
- Préserve le formatage
- Détecte la langue source`;

const SKILL_WITHOUT_ACTION_KEY = `---
name: Mon Skill
description: "Un skill utilisateur custom."
host: word
executionMode: agent
icon: Wand2
---

Instructions de mon skill.`;

const SKILL_NO_FRONTMATTER = `# Old Style Skill

## Purpose
Do something useful.

## Rules
- Rule 1`;

const SKILL_INVALID_ENUM_VALUES = `---
name: Skill Weird
description: "Test avec valeurs invalides."
host: invalidhost
executionMode: invalidmode
icon: Zap
---

Body content.`;

const SKILL_QUOTED_SINGLE = `---
name: 'Single Quoted'
description: 'Description avec des guillemets simples.'
host: excel
executionMode: draft
icon: Table
---

Corps.`;

const SKILL_DESCRIPTION_WITH_ESCAPED_QUOTES = `---
name: Test Quotes
description: "Il dit \\"bonjour\\" et repart."
host: outlook
executionMode: immediate
icon: Mail
---

Corps.`;

const SKILL_WINDOWS_LINE_ENDINGS = `---\r\nname: Windows Skill\r\ndescription: "Test CRLF."\r\nhost: powerpoint\r\nexecutionMode: immediate\r\nicon: List\r\n---\r\n\r\nCorps avec CRLF.`;

// ── parseSkill ─────────────────────────────────────────────────────────────────

describe('parseSkill', () => {
  describe('with valid frontmatter', () => {
    it('parses all metadata fields correctly', () => {
      const result = parseSkill(FULL_SKILL_MD);
      expect(result.metadata.name).toBe('Traduire le texte');
      expect(result.metadata.description).toBe(
        'Traduit le texte sélectionné entre français et anglais.',
      );
      expect(result.metadata.host).toBe('all');
      expect(result.metadata.executionMode).toBe('immediate');
      expect(result.metadata.icon).toBe('Languages');
      expect(result.metadata.actionKey).toBe('translate');
    });

    it('returns raw = original file content', () => {
      const result = parseSkill(FULL_SKILL_MD);
      expect(result.raw).toBe(FULL_SKILL_MD);
    });

    it('returns body without frontmatter delimiters', () => {
      const result = parseSkill(FULL_SKILL_MD);
      expect(result.body).not.toContain('---');
      expect(result.body).toContain('Traduis le texte reçu');
      expect(result.body).toContain('## Règles');
    });

    it('handles skill without actionKey', () => {
      const result = parseSkill(SKILL_WITHOUT_ACTION_KEY);
      expect(result.metadata.actionKey).toBeUndefined();
      expect(result.metadata.host).toBe('word');
      expect(result.metadata.executionMode).toBe('agent');
    });

    it('uses fallbackActionKey when actionKey is absent from frontmatter', () => {
      const result = parseSkill(SKILL_WITHOUT_ACTION_KEY, 'my-fallback');
      expect(result.metadata.actionKey).toBe('my-fallback');
    });

    it('frontmatter actionKey takes precedence over fallbackActionKey', () => {
      const result = parseSkill(FULL_SKILL_MD, 'should-be-ignored');
      expect(result.metadata.actionKey).toBe('translate');
    });

    it('parses single-quoted description', () => {
      const result = parseSkill(SKILL_QUOTED_SINGLE);
      expect(result.metadata.name).toBe('Single Quoted');
      expect(result.metadata.description).toBe('Description avec des guillemets simples.');
      expect(result.metadata.host).toBe('excel');
      expect(result.metadata.executionMode).toBe('draft');
    });

    it('unescapes double quotes inside description', () => {
      const result = parseSkill(SKILL_DESCRIPTION_WITH_ESCAPED_QUOTES);
      expect(result.metadata.description).toBe('Il dit "bonjour" et repart.');
    });

    it('handles Windows line endings (CRLF)', () => {
      const result = parseSkill(SKILL_WINDOWS_LINE_ENDINGS);
      expect(result.metadata.name).toBe('Windows Skill');
      expect(result.metadata.host).toBe('powerpoint');
      expect(result.body).toContain('Corps avec CRLF.');
    });
  });

  describe('with invalid or missing frontmatter', () => {
    it('falls back gracefully when no frontmatter', () => {
      const result = parseSkill(SKILL_NO_FRONTMATTER);
      expect(result.metadata.name).toBe('Unnamed Skill');
      expect(result.metadata.host).toBe('all');
      expect(result.metadata.executionMode).toBe('immediate');
      expect(result.metadata.icon).toBe('Zap');
      expect(result.raw).toBe(SKILL_NO_FRONTMATTER);
      expect(result.body).toBe(SKILL_NO_FRONTMATTER);
    });

    it('uses fallbackActionKey when no frontmatter', () => {
      const result = parseSkill(SKILL_NO_FRONTMATTER, 'legacy-key');
      expect(result.metadata.name).toBe('legacy-key');
      expect(result.metadata.actionKey).toBe('legacy-key');
    });

    it('falls back to "all" for invalid host value', () => {
      const result = parseSkill(SKILL_INVALID_ENUM_VALUES);
      expect(result.metadata.host).toBe('all');
    });

    it('falls back to "immediate" for invalid executionMode value', () => {
      const result = parseSkill(SKILL_INVALID_ENUM_VALUES);
      expect(result.metadata.executionMode).toBe('immediate');
    });

    it('handles empty string gracefully', () => {
      const result = parseSkill('');
      expect(result.metadata.name).toBe('Unnamed Skill');
      expect(result.body).toBe('');
    });
  });

  describe('all valid host values', () => {
    const hosts = ['word', 'excel', 'powerpoint', 'outlook', 'all'] as const;
    for (const host of hosts) {
      it(`accepts host "${host}"`, () => {
        const raw = `---\nname: Test\ndescription: ""\nhost: ${host}\nexecutionMode: immediate\nicon: Zap\n---\n\nBody.`;
        expect(parseSkill(raw).metadata.host).toBe(host);
      });
    }
  });

  describe('all valid executionMode values', () => {
    const modes = ['immediate', 'draft', 'agent'] as const;
    for (const mode of modes) {
      it(`accepts executionMode "${mode}"`, () => {
        const raw = `---\nname: Test\ndescription: ""\nhost: all\nexecutionMode: ${mode}\nicon: Zap\n---\n\nBody.`;
        expect(parseSkill(raw).metadata.executionMode).toBe(mode);
      });
    }
  });
});

// ── serializeSkillToMd ─────────────────────────────────────────────────────────

describe('serializeSkillToMd', () => {
  it('produces valid frontmatter that can be re-parsed', () => {
    const original = {
      name: 'Mon Skill Test',
      description: 'Fait quelque chose d\'utile.',
      host: 'word' as const,
      executionMode: 'agent' as const,
      icon: 'Wand2',
      skillContent: '## Instructions\n\nFais ceci.',
    };
    const serialized = serializeSkillToMd(original);
    const reparsed = parseSkill(serialized);

    expect(reparsed.metadata.name).toBe(original.name);
    expect(reparsed.metadata.description).toBe(original.description);
    expect(reparsed.metadata.host).toBe(original.host);
    expect(reparsed.metadata.executionMode).toBe(original.executionMode);
    expect(reparsed.metadata.icon).toBe(original.icon);
    expect(reparsed.body).toBe(original.skillContent);
  });

  it('escapes double quotes in description', () => {
    const serialized = serializeSkillToMd({
      name: 'Test',
      description: 'Il dit "bonjour".',
      host: 'all',
      executionMode: 'immediate',
      icon: 'Zap',
      skillContent: 'Body.',
    });
    expect(serialized).toContain('description: "Il dit \\"bonjour\\"."');
    // Re-parse should recover the original description
    expect(parseSkill(serialized).metadata.description).toBe('Il dit "bonjour".');
  });

  it('includes actionKey in frontmatter when provided', () => {
    const serialized = serializeSkillToMd({
      name: 'Built-in',
      description: 'Test.',
      host: 'word',
      executionMode: 'agent',
      icon: 'Zap',
      skillContent: 'Body.',
      actionKey: 'word-translate',
    });
    expect(serialized).toContain('actionKey: word-translate');
    expect(parseSkill(serialized).metadata.actionKey).toBe('word-translate');
  });

  it('omits actionKey line when not provided', () => {
    const serialized = serializeSkillToMd({
      name: 'User Skill',
      description: 'Test.',
      host: 'all',
      executionMode: 'immediate',
      icon: 'Zap',
      skillContent: 'Body.',
    });
    expect(serialized).not.toContain('actionKey');
  });

  it('round-trips correctly for all host values', () => {
    const hosts = ['word', 'excel', 'powerpoint', 'outlook', 'all'] as const;
    for (const host of hosts) {
      const serialized = serializeSkillToMd({
        name: 'Test',
        description: 'Desc.',
        host,
        executionMode: 'immediate',
        icon: 'Zap',
        skillContent: 'Body.',
      });
      expect(parseSkill(serialized).metadata.host).toBe(host);
    }
  });
});
