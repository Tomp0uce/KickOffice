import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock officeCodeValidator before importing the module under test
vi.mock('@/utils/officeCodeValidator', () => ({
  validateOfficeCode: vi.fn(),
}));

// Mock sandboxedEval before importing the module under test
vi.mock('@/utils/sandbox', () => ({
  sandboxedEval: vi.fn(),
}));

// Mock logService before importing the module under test
vi.mock('@/utils/logger', () => ({
  logService: {
    warn: vi.fn(),
    error: vi.fn(),
    info: vi.fn(),
    debug: vi.fn(),
  },
}));

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
  getDisplayLanguage,
  truncateString,
  getErrorMessage,
  getDetailedOfficeError,
  normalizeLineEndings,
  buildScreenshotResult,
  createEvalExecutor,
  buildExecuteWrapper,
} from '../common';
import { validateOfficeCode } from '@/utils/officeCodeValidator';
import { sandboxedEval } from '@/utils/sandbox';
import { logService } from '@/utils/logger';

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

// ─────────────────────────────────────────────────────────────────────────────
// getDisplayLanguage
// ─────────────────────────────────────────────────────────────────────────────
describe('getDisplayLanguage', () => {
  afterEach(() => {
    localStorage.clear();
  });

  it('returns English when localStorage has "en"', () => {
    localStorage.setItem('localLanguage', 'en');
    expect(getDisplayLanguage()).toBe('English');
  });

  it('returns Français when localStorage has "fr"', () => {
    localStorage.setItem('localLanguage', 'fr');
    expect(getDisplayLanguage()).toBe('Français');
  });

  it('returns Français when localStorage key is absent', () => {
    expect(getDisplayLanguage()).toBe('Français');
  });

  it('returns Français when localStorage has an invalid value', () => {
    localStorage.setItem('localLanguage', 'de');
    expect(getDisplayLanguage()).toBe('Français');
  });

  it('returns Français when localStorage has an empty string', () => {
    localStorage.setItem('localLanguage', '');
    expect(getDisplayLanguage()).toBe('Français');
  });

  it('returns Français when localStorage throws a SecurityError', () => {
    const originalGetItem = localStorage.getItem.bind(localStorage);
    const spy = vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('SecurityError');
    });
    expect(getDisplayLanguage()).toBe('Français');
    spy.mockRestore();
    // restore normal behaviour for other tests
    void originalGetItem;
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// truncateString
// ─────────────────────────────────────────────────────────────────────────────
describe('truncateString', () => {
  it('returns the string unchanged when shorter than maxLen', () => {
    expect(truncateString('hello', 10)).toBe('hello');
  });

  it('returns the string unchanged when exactly maxLen', () => {
    expect(truncateString('hello', 5)).toBe('hello');
  });

  it('truncates and appends ... when longer than maxLen', () => {
    expect(truncateString('hello world', 5)).toBe('hello...');
  });

  it('returns ... for a 0-length max on a non-empty string', () => {
    expect(truncateString('abc', 0)).toBe('...');
  });

  it('handles an empty string without appending ...', () => {
    expect(truncateString('', 5)).toBe('');
  });

  it('handles unicode characters correctly', () => {
    const result = truncateString('caf\u00e9 au lait', 4);
    expect(result).toBe('caf\u00e9...');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getErrorMessage
// ─────────────────────────────────────────────────────────────────────────────
describe('getErrorMessage', () => {
  it('returns error.message for an Error instance', () => {
    expect(getErrorMessage(new Error('boom'))).toBe('boom');
  });

  it('returns error.message for a plain object with a string message property', () => {
    expect(getErrorMessage({ message: 'plain object error' })).toBe('plain object error');
  });

  it('returns String(value) for a plain string', () => {
    expect(getErrorMessage('raw string')).toBe('raw string');
  });

  it('returns String(value) for a number', () => {
    expect(getErrorMessage(42)).toBe('42');
  });

  it('returns "null" for null input', () => {
    expect(getErrorMessage(null)).toBe('null');
  });

  it('returns "undefined" for undefined input', () => {
    expect(getErrorMessage(undefined)).toBe('undefined');
  });

  it('ignores non-string message property on object', () => {
    // message is present but not a string — falls through to String()
    const obj = { message: 123 };
    expect(getErrorMessage(obj)).toBe('[object Object]');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getDetailedOfficeError
// ─────────────────────────────────────────────────────────────────────────────
describe('getDetailedOfficeError', () => {
  it('extracts message from a minimal Office error', () => {
    const err = { message: 'Office error', debugInfo: {} };
    expect(getDetailedOfficeError(err)).toBe('Office error');
  });

  it('includes code when present', () => {
    const err = { message: 'Office error', code: 'ItemNotFound', debugInfo: {} };
    const result = getDetailedOfficeError(err);
    expect(result).toContain('Code: ItemNotFound');
  });

  it('includes errorLocation from debugInfo', () => {
    const err = {
      message: 'fail',
      debugInfo: { errorLocation: 'context.body' },
    };
    const result = getDetailedOfficeError(err);
    expect(result).toContain('Location: context.body');
  });

  it('includes failing statement from debugInfo', () => {
    const err = {
      message: 'fail',
      debugInfo: { statement: 'context.body.load("text")' },
    };
    const result = getDetailedOfficeError(err);
    expect(result).toContain('Failing statement: context.body.load("text")');
  });

  it('includes surrounding statements when present', () => {
    const err = {
      message: 'fail',
      debugInfo: { surroundingStatements: ['stmt1', 'stmt2'] },
    };
    const result = getDetailedOfficeError(err);
    expect(result).toContain('Surrounding context: stmt1; stmt2');
  });

  it('omits surrounding context when surroundingStatements is empty', () => {
    const err = {
      message: 'fail',
      debugInfo: { surroundingStatements: [] },
    };
    const result = getDetailedOfficeError(err);
    expect(result).not.toContain('Surrounding context');
  });

  it('falls back to getErrorMessage for a regular Error', () => {
    const err = new Error('regular error');
    expect(getDetailedOfficeError(err)).toBe('regular error');
  });

  it('falls back to String() for null', () => {
    expect(getDetailedOfficeError(null)).toBe('null');
  });

  it('falls back to String() for a plain string', () => {
    expect(getDetailedOfficeError('raw')).toBe('raw');
  });

  it('handles an Office error without debugInfo key', () => {
    // Has message but no debugInfo — should fall through to getErrorMessage
    const err = { message: 'no debug info' };
    // No debugInfo key, so falls back to getErrorMessage branch
    expect(getDetailedOfficeError(err)).toBe('no debug info');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// normalizeLineEndings
// ─────────────────────────────────────────────────────────────────────────────
describe('normalizeLineEndings', () => {
  it('returns empty string for empty input', () => {
    expect(normalizeLineEndings('')).toBe('');
  });

  it('replaces \\r\\n with \\n', () => {
    expect(normalizeLineEndings('line1\r\nline2')).toBe('line1\nline2');
  });

  it('replaces lone \\r with \\n', () => {
    expect(normalizeLineEndings('line1\rline2')).toBe('line1\nline2');
  });

  it('handles mixed \\r\\n and \\r endings', () => {
    expect(normalizeLineEndings('a\r\nb\rc')).toBe('a\nb\nc');
  });

  it('leaves text with only \\n unchanged', () => {
    expect(normalizeLineEndings('a\nb\nc')).toBe('a\nb\nc');
  });

  it('handles string with no line endings', () => {
    expect(normalizeLineEndings('hello world')).toBe('hello world');
  });

  it('handles multiple consecutive \\r\\n sequences', () => {
    expect(normalizeLineEndings('a\r\n\r\nb')).toBe('a\n\nb');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// buildScreenshotResult
// ─────────────────────────────────────────────────────────────────────────────
describe('buildScreenshotResult', () => {
  it('returns valid JSON', () => {
    const result = buildScreenshotResult('base64data', 'A chart');
    expect(() => JSON.parse(result)).not.toThrow();
  });

  it('sets __screenshot__ to true', () => {
    const parsed = JSON.parse(buildScreenshotResult('data', 'desc'));
    expect(parsed.__screenshot__).toBe(true);
  });

  it('sets mimeType to image/png', () => {
    const parsed = JSON.parse(buildScreenshotResult('data', 'desc'));
    expect(parsed.mimeType).toBe('image/png');
  });

  it('includes the provided base64 value', () => {
    const parsed = JSON.parse(buildScreenshotResult('abc123', 'chart'));
    expect(parsed.base64).toBe('abc123');
  });

  it('includes the provided description', () => {
    const parsed = JSON.parse(buildScreenshotResult('data', 'my description'));
    expect(parsed.description).toBe('my description');
  });

  it('handles empty strings without throwing', () => {
    const result = buildScreenshotResult('', '');
    const parsed = JSON.parse(result);
    expect(parsed.base64).toBe('');
    expect(parsed.description).toBe('');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// buildExecuteWrapper
// ─────────────────────────────────────────────────────────────────────────────
describe('buildExecuteWrapper', () => {
  const mockRunner = vi.fn();

  beforeEach(() => {
    mockRunner.mockReset();
  });

  it('returns error JSON when the executeKey function is missing', async () => {
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const def = { name: 'myTool' } as Record<string, unknown>;
    const result = JSON.parse(await buildExecute(def)());
    expect(result.success).toBe(false);
    expect(result.error).toContain('executeWord');
    expect(result.tool).toBe('myTool');
  });

  it('returns error JSON when the executeKey value is not a function', async () => {
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const def = { name: 'myTool', executeWord: 'not-a-function' } as Record<string, unknown>;
    const result = JSON.parse(await buildExecute(def)());
    expect(result.success).toBe(false);
    expect(result.error).toContain('executeWord');
  });

  it('uses "unknown" as tool name when name is absent', async () => {
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const def = {} as Record<string, unknown>;
    const result = JSON.parse(await buildExecute(def)());
    expect(result.tool).toBe('unknown');
  });

  it('calls runner and returns its result on success', async () => {
    mockRunner.mockResolvedValue('{"success":true}');
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const hostExecute = vi.fn();
    const def = { name: 'myTool', executeWord: hostExecute } as Record<string, unknown>;
    const result = await buildExecute(def)({ arg1: 'val' });
    expect(result).toBe('{"success":true}');
    expect(mockRunner).toHaveBeenCalledOnce();
  });

  it('passes args to the host execute function via runner', async () => {
    let capturedAction: ((ctx: unknown) => Promise<unknown>) | undefined;
    mockRunner.mockImplementation(async (action: (ctx: unknown) => Promise<unknown>) => {
      capturedAction = action;
      return 'result';
    });
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const hostExecute = vi.fn().mockResolvedValue('ok');
    const def = { name: 'myTool', executeWord: hostExecute } as Record<string, unknown>;
    await buildExecute(def)({ key: 'value' });
    expect(capturedAction).toBeDefined();
    await capturedAction!('fakeContext');
    expect(hostExecute).toHaveBeenCalledWith('fakeContext', { key: 'value' });
  });

  it('returns error JSON when runner throws', async () => {
    mockRunner.mockRejectedValue(new Error('Office failure'));
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const def = { name: 'myTool', executeWord: vi.fn() } as Record<string, unknown>;
    const result = JSON.parse(await buildExecute(def)());
    expect(result.success).toBe(false);
    expect(result.error).toBe('Office failure');
    expect(result.suggestion).toBeDefined();
  });

  it('defaults args to {} when called with no arguments', async () => {
    mockRunner.mockResolvedValue('done');
    const buildExecute = buildExecuteWrapper('executeWord', mockRunner);
    const def = { name: 'myTool', executeWord: vi.fn() } as Record<string, unknown>;
    await expect(buildExecute(def)()).resolves.toBe('done');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// createEvalExecutor
// ─────────────────────────────────────────────────────────────────────────────
describe('createEvalExecutor', () => {
  const mockValidateOfficeCode = vi.mocked(validateOfficeCode);
  const mockSandboxedEval = vi.mocked(sandboxedEval);
  const mockLogWarn = vi.mocked(logService.warn);

  const baseConfig = {
    host: 'Word' as const,
    toolName: 'eval_wordjs',
    suggestion: 'Fix your code',
    buildSandboxContext: (_ctx: unknown) => ({ context: _ctx }),
  };

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns validation error JSON when code is invalid', async () => {
    mockValidateOfficeCode.mockReturnValue({
      valid: false,
      errors: ['Missing context.sync()'],
      warnings: [],
    });
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('fakeCtx', { code: 'bad code', explanation: 'test' }));
    expect(result.success).toBe(false);
    expect(result.error).toContain('Code validation failed');
    expect(result.validationErrors).toEqual(['Missing context.sync()']);
    expect(result.suggestion).toBe('Fix your code');
    expect(result.codeReceived).toBe('bad code');
  });

  it('truncates code in validation error when longer than validationCodePreviewLength', async () => {
    mockValidateOfficeCode.mockReturnValue({
      valid: false,
      errors: ['error'],
      warnings: [],
    });
    const executor = createEvalExecutor({ ...baseConfig, validationCodePreviewLength: 5 });
    const longCode = 'a'.repeat(20);
    const result = JSON.parse(await executor('ctx', { code: longCode, explanation: '' }));
    expect(result.codeReceived).toBe('aaaaa...');
  });

  it('logs warnings when validation passes with warnings', async () => {
    mockValidateOfficeCode.mockReturnValue({
      valid: true,
      errors: [],
      warnings: ['Use .load()'],
    });
    mockSandboxedEval.mockResolvedValue('result data');
    const executor = createEvalExecutor(baseConfig);
    await executor('ctx', { code: 'valid code', explanation: 'x' });
    expect(mockLogWarn).toHaveBeenCalledWith('[eval_wordjs] Validation warnings:', ['Use .load()']);
  });

  it('includes warnings array in success response when warnings exist', async () => {
    mockValidateOfficeCode.mockReturnValue({
      valid: true,
      errors: [],
      warnings: ['warn1'],
    });
    mockSandboxedEval.mockResolvedValue('ok');
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: 'e' }));
    expect(result.success).toBe(true);
    expect(result.warnings).toEqual(['warn1']);
  });

  it('omits warnings field in success response when there are no warnings', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue('ok');
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: 'e' }));
    expect(result.success).toBe(true);
    expect(result.warnings).toBeUndefined();
  });

  it('returns success JSON with result and explanation', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue({ data: 42 });
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'valid', explanation: 'gets data' }));
    expect(result.success).toBe(true);
    expect(result.result).toEqual({ data: 42 });
    expect(result.explanation).toBe('gets data');
  });

  it('sets result to null when sandboxedEval returns undefined', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue(undefined);
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: '' }));
    expect(result.result).toBeNull();
  });

  it('includes hasMutated in response when mutationDetector is provided', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue('done');
    const executor = createEvalExecutor({
      ...baseConfig,
      mutationDetector: (_code: string) => true,
    });
    const result = JSON.parse(await executor('ctx', { code: 'mutating code', explanation: '' }));
    expect(result.hasMutated).toBe(true);
  });

  it('omits hasMutated when no mutationDetector is provided', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue('done');
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: '' }));
    expect('hasMutated' in result).toBe(false);
  });

  it('calls preExecuteHook before sandboxedEval', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    const callOrder: string[] = [];
    mockSandboxedEval.mockImplementation(async () => {
      callOrder.push('sandbox');
      return 'ok';
    });
    const preExecuteHook = vi.fn().mockImplementation(() => {
      callOrder.push('hook');
    });
    const executor = createEvalExecutor({ ...baseConfig, preExecuteHook });
    await executor('ctx', { code: 'code', explanation: '' });
    expect(callOrder).toEqual(['hook', 'sandbox']);
    expect(preExecuteHook).toHaveBeenCalledWith('ctx');
  });

  it('returns catch error JSON when sandboxedEval throws', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockRejectedValue(new Error('sandbox boom'));
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'bad exec', explanation: 'x' }));
    expect(result.success).toBe(false);
    expect(result.error).toContain('sandbox boom');
    expect(result.codeExecuted).toBe('bad exec');
    expect(result.hint).toBeDefined();
  });

  it('truncates codeExecuted in catch error when longer than catchCodePreviewLength', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockRejectedValue(new Error('fail'));
    const executor = createEvalExecutor({ ...baseConfig, catchCodePreviewLength: 5 });
    const longCode = 'b'.repeat(30);
    const result = JSON.parse(await executor('ctx', { code: longCode, explanation: '' }));
    expect(result.codeExecuted).toBe('bbbbb...');
  });

  it('uses custom catchHint when provided', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockRejectedValue(new Error('fail'));
    const executor = createEvalExecutor({ ...baseConfig, catchHint: 'Custom hint here' });
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: '' }));
    expect(result.hint).toBe('Custom hint here');
  });

  it('uses default hint when catchHint is not provided', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockRejectedValue(new Error('fail'));
    const executor = createEvalExecutor(baseConfig);
    const result = JSON.parse(await executor('ctx', { code: 'code', explanation: '' }));
    expect(result.hint).toContain('context.sync()');
  });

  it('passes the buildSandboxContext output to sandboxedEval', async () => {
    mockValidateOfficeCode.mockReturnValue({ valid: true, errors: [], warnings: [] });
    mockSandboxedEval.mockResolvedValue(null);
    const buildSandboxContext = vi.fn().mockReturnValue({ ctx: 'built' });
    const executor = createEvalExecutor({ ...baseConfig, buildSandboxContext });
    await executor('myCtx', { code: 'code', explanation: '' });
    expect(buildSandboxContext).toHaveBeenCalledWith('myCtx');
    expect(mockSandboxedEval).toHaveBeenCalledWith('code', { ctx: 'built' }, 'Word');
  });
});
