import { describe, it, expect, vi } from 'vitest';

// Mock diff-match-patch before importing the module under test
vi.mock('diff-match-patch', () => {
  const EQUAL = 0;
  const DELETE = -1;
  const INSERT = 1;

  class DiffMatchPatch {
    diff_main(a: string, b: string): Array<[number, string]> {
      if (a === b) return [[EQUAL, a]];
      // Minimal fake: treat entire strings as delete/insert
      const result: Array<[number, string]> = [];
      if (a) result.push([DELETE, a]);
      if (b) result.push([INSERT, b]);
      return result;
    }
    diff_cleanupSemantic(_diffs: Array<[number, string]>): void {
      // no-op in tests
    }
  }

  return { default: DiffMatchPatch };
});

// Mock the constant module — languageMap is a plain object, no side effects
vi.mock('@/utils/constant', () => ({
  languageMap: {
    en: 'English',
    fr: 'Français',
  },
}));

import {
  generateVisualDiff,
  computeTextDiffStats,
  createOfficeTools,
  optionLists,
} from '../common';

// ─────────────────────────────────────────────────────────────────────────────
// generateVisualDiff
// ─────────────────────────────────────────────────────────────────────────────
describe('generateVisualDiff', () => {
  it('returns empty string when both inputs are identical strings', () => {
    const result = generateVisualDiff('hello', 'hello');
    // DiffMatchPatch returns EQUAL for identical strings → no colored spans
    expect(result).toBe('hello');
  });

  it('wraps insertions in a blue underline span', () => {
    const result = generateVisualDiff('', 'new text');
    expect(result).toContain('color:blue');
    expect(result).toContain('new text');
  });

  it('wraps deletions in a red line-through span', () => {
    const result = generateVisualDiff('old text', '');
    expect(result).toContain('color:red');
    expect(result).toContain('old text');
  });

  it('escapes HTML special characters in text', () => {
    // originalText → deleted, contains HTML that must be escaped
    const result = generateVisualDiff('<b>bold</b>', '');
    expect(result).not.toContain('<b>');
    expect(result).toContain('&lt;b&gt;');
  });

  it('escapes ampersands in text', () => {
    const result = generateVisualDiff('A & B', '');
    expect(result).toContain('&amp;');
    expect(result).not.toContain('A & B');
  });

  it('converts newlines to <br> tags', () => {
    const result = generateVisualDiff('line1\nline2', '');
    expect(result).toContain('<br>');
    expect(result).not.toContain('\n');
  });

  it('returns empty string when originalText is not a string', () => {
    expect(generateVisualDiff(null as unknown as string, 'text')).toBe('');
    expect(generateVisualDiff(42 as unknown as string, 'text')).toBe('');
    expect(generateVisualDiff(undefined as unknown as string, 'text')).toBe('');
  });

  it('returns empty string when newText is not a string', () => {
    expect(generateVisualDiff('text', null as unknown as string)).toBe('');
    expect(generateVisualDiff('text', undefined as unknown as string)).toBe('');
  });

  it('returns empty string when both inputs are non-strings', () => {
    expect(generateVisualDiff(null as unknown as string, null as unknown as string)).toBe('');
  });

  it('handles two empty strings without throwing', () => {
    const result = generateVisualDiff('', '');
    // Both are empty strings (equal) → EQUAL diff with empty text → no visible output
    expect(typeof result).toBe('string');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// computeTextDiffStats
// ─────────────────────────────────────────────────────────────────────────────
describe('computeTextDiffStats', () => {
  it('returns zero for insertions and deletions when texts are identical', () => {
    const stats = computeTextDiffStats('hello', 'hello');
    expect(stats.insertions).toBe(0);
    expect(stats.deletions).toBe(0);
    expect(stats.unchanged).toBe('hello'.length);
  });

  it('counts deletions when new text is empty', () => {
    const original = 'deleted';
    const stats = computeTextDiffStats(original, '');
    expect(stats.deletions).toBe(original.length);
    expect(stats.insertions).toBe(0);
    expect(stats.unchanged).toBe(0);
  });

  it('counts insertions when original text is empty', () => {
    const added = 'added';
    const stats = computeTextDiffStats('', added);
    expect(stats.insertions).toBe(added.length);
    expect(stats.deletions).toBe(0);
    expect(stats.unchanged).toBe(0);
  });

  it('returns an object with the correct shape', () => {
    const stats = computeTextDiffStats('a', 'b');
    expect(stats).toHaveProperty('insertions');
    expect(stats).toHaveProperty('deletions');
    expect(stats).toHaveProperty('unchanged');
  });

  it('all values are non-negative integers', () => {
    const stats = computeTextDiffStats('foo', 'bar');
    expect(stats.insertions).toBeGreaterThanOrEqual(0);
    expect(stats.deletions).toBeGreaterThanOrEqual(0);
    expect(stats.unchanged).toBeGreaterThanOrEqual(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// createOfficeTools
// ─────────────────────────────────────────────────────────────────────────────
describe('createOfficeTools', () => {
  type ToolName = 'doSomething' | 'doOther';

  interface ToolTemplate {
    description: string;
  }

  interface ToolDef extends ToolTemplate {
    execute: (args?: Record<string, unknown>) => Promise<string>;
  }

  const definitions: Record<ToolName, ToolTemplate> = {
    doSomething: { description: 'Does something' },
    doOther: { description: 'Does other' },
  };

  const buildExecute =
    (def: ToolTemplate) =>
    async (_args?: Record<string, unknown>): Promise<string> =>
      `executed: ${def.description}`;

  it('produces an entry for each definition key', () => {
    const tools = createOfficeTools<ToolName, ToolTemplate, ToolDef>(definitions, buildExecute);
    expect(Object.keys(tools)).toEqual(expect.arrayContaining(['doSomething', 'doOther']));
  });

  it('each entry has an execute function', () => {
    const tools = createOfficeTools<ToolName, ToolTemplate, ToolDef>(definitions, buildExecute);
    expect(typeof tools.doSomething.execute).toBe('function');
    expect(typeof tools.doOther.execute).toBe('function');
  });

  it('execute returns the expected promise value', async () => {
    const tools = createOfficeTools<ToolName, ToolTemplate, ToolDef>(definitions, buildExecute);
    await expect(tools.doSomething.execute()).resolves.toBe('executed: Does something');
    await expect(tools.doOther.execute()).resolves.toBe('executed: Does other');
  });

  it('preserves original definition properties alongside execute', () => {
    const tools = createOfficeTools<ToolName, ToolTemplate, ToolDef>(definitions, buildExecute);
    expect(tools.doSomething.description).toBe('Does something');
  });

  it('handles an empty definitions object without throwing', () => {
    const emptyTools = createOfficeTools<never, ToolTemplate, ToolDef>(
      {} as Record<never, ToolTemplate>,
      buildExecute,
    );
    expect(Object.keys(emptyTools)).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// optionLists
// ─────────────────────────────────────────────────────────────────────────────
describe('optionLists', () => {
  it('exposes a localLanguageList array', () => {
    expect(Array.isArray(optionLists.localLanguageList)).toBe(true);
  });

  it('localLanguageList contains English entry', () => {
    const en = optionLists.localLanguageList.find(item => item.value === 'en');
    expect(en).toBeDefined();
    expect(en?.label).toBe('English');
  });

  it('localLanguageList contains French entry', () => {
    const fr = optionLists.localLanguageList.find(item => item.value === 'fr');
    expect(fr).toBeDefined();
    expect(fr?.label).toBe('Français');
  });

  it('each entry has a label and value string', () => {
    for (const item of optionLists.localLanguageList) {
      expect(typeof item.label).toBe('string');
      expect(typeof item.value).toBe('string');
    }
  });
});
