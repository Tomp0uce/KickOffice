import DiffMatchPatch from 'diff-match-patch';
import type { ToolDefinition } from '@/types';

// R17/CH5 — Generate a visual diff HTML string (insertions in blue/underline, deletions in red/strikethrough)
export function generateVisualDiff(originalText: unknown, newText: unknown): string {
  if (typeof originalText !== 'string' || typeof newText !== 'string') {
    return '';
  }
  const dmp = new DiffMatchPatch();
  const diffs = dmp.diff_main(originalText, newText);
  dmp.diff_cleanupSemantic(diffs);

  return diffs
    .map(([op, text]) => {
      const escaped = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');
      if (op === 1) return `<span style="color:blue;text-decoration:underline;">${escaped}</span>`;
      if (op === -1)
        return `<span style="color:red;text-decoration:line-through;">${escaped}</span>`;
      return escaped;
    })
    .join('');
}

export interface TextDiffStats {
  insertions: number;
  deletions: number;
  unchanged: number;
}

/** Compute word-level diff stats between two strings (used by hosts that report diff info without a visual HTML diff). */
export function computeTextDiffStats(originalText: string, revisedText: string): TextDiffStats {
  const dmp = new DiffMatchPatch();
  const diffs = dmp.diff_main(originalText, revisedText);
  dmp.diff_cleanupSemantic(diffs);
  let insertions = 0,
    deletions = 0,
    unchanged = 0;
  for (const [op, text] of diffs) {
    if (op === 0) unchanged += text.length;
    else if (op === -1) deletions += text.length;
    else if (op === 1) insertions += text.length;
  }
  return { insertions, deletions, unchanged };
}

/**
 * Truncate a string to maxLen characters, appending '...' if truncated.
 * Used by wordTools and outlookTools for error message truncation.
 */
export function truncateString(str: string, maxLen: number): string {
  if (str.length <= maxLen) return str;
  return str.slice(0, maxLen) + '...';
}

/**
 * Extract error message from unknown error type.
 * Safely extracts error.message or converts to string for TypeScript strict mode.
 */
export function getErrorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  if (
    error &&
    typeof error === 'object' &&
    'message' in error &&
    typeof error.message === 'string'
  ) {
    return error.message;
  }
  return String(error);
}

/**
 * Extract detailed error info from Office.js OfficeExtension.Error objects.
 * Provides errorLocation, failing statement, and surrounding context for LLM auto-correction.
 * Ported from Office Agents error handling pattern.
 * Falls back to getErrorMessage() for non-Office errors.
 */
export function getDetailedOfficeError(error: unknown): string {
  if (
    error &&
    typeof error === 'object' &&
    'debugInfo' in error &&
    'message' in error
  ) {
    const officeError = error as {
      message: string;
      code?: string;
      debugInfo?: {
        errorLocation?: string;
        statement?: string;
        surroundingStatements?: string[];
      };
    };

    const parts = [officeError.message];
    if (officeError.code) parts.push(`Code: ${officeError.code}`);

    if (officeError.debugInfo) {
      const { errorLocation, statement, surroundingStatements } = officeError.debugInfo;
      if (errorLocation) parts.push(`Location: ${errorLocation}`);
      if (statement) parts.push(`Failing statement: ${statement}`);
      if (surroundingStatements?.length) {
        parts.push(`Surrounding context: ${surroundingStatements.join('; ')}`);
      }
    }

    return parts.join('\n');
  }

  return getErrorMessage(error);
}

/**
 * Generic Office Tool Template.
 * Defines the structure for host-specific tool definitions before wrapping with execute().
 *
 * @template TContext - The Office.js context type (Word.RequestContext, Excel.RequestContext, etc.)
 *
 * Each host-specific tool file should extend this with their own executeXxx property:
 * - Word: { executeWord: (context: Word.RequestContext, args: Record<string, unknown>) => Promise<string> }
 * - Excel: { executeExcel: (context: Excel.RequestContext, args: Record<string, unknown>) => Promise<string> }
 * - PowerPoint: { executePowerPoint: (context: PowerPoint.RequestContext, args: Record<string, unknown>) => Promise<string> }
 * - Outlook: { executeOutlook: (item: Office.MessageCompose, args: Record<string, unknown>) => Promise<string> }
 */
export type OfficeToolTemplate<TContext = any> = Omit<ToolDefinition, 'execute'>;

/**
 * Generic factory that wraps host-specific tool templates with a uniform `execute` method.
 * Each tool file passes a `buildExecute` callback that closes over its host runner
 * (runWord, runExcel, runPowerPoint, runOutlook).
 */
export function createOfficeTools<TName extends string, TTemplate extends object, TDef>(
  definitions: Record<TName, TTemplate>,
  buildExecute: (definition: TTemplate) => (args?: Record<string, any>) => Promise<string>,
): Record<TName, TDef> {
  return Object.fromEntries(
    Object.entries(definitions).map(([name, def]) => [
      name,
      { ...(def as object), execute: buildExecute(def as TTemplate) },
    ]),
  ) as unknown as Record<TName, TDef>;
}

/**
 * Generic wrapper builder for Office.js tool execution.
 * Creates the execute wrapper that bridges the generic Tool interface with host-specific execution.
 *
 * Wraps host-specific executeXxx methods with:
 * - Office.js context runner (runWord, runExcel, etc.)
 * - Standard error handling
 * - JSON stringified error responses
 *
 * @template TTemplate - The tool template type
 * @param executeKey - The property name for the host-specific executor (e.g., 'executeWord', 'executeExcel')
 * @param runner - The Office.js runner function that provides the context
 * @returns A function that builds execute wrappers for tool definitions
 *
 * @example
 * const buildExecute = buildExecuteWrapper<WordToolTemplate>(
 *   'executeWord',
 *   <T>(action: (ctx: Word.RequestContext) => Promise<T>) => executeOfficeAction(() => Word.run(action))
 * )
 */
export function buildExecuteWrapper<TTemplate extends Record<string, any>>(
  executeKey: string,
  runner: <T>(action: (context: any) => Promise<T>) => Promise<T>,
): (definition: TTemplate) => (args?: Record<string, any>) => Promise<string> {
  return (def: TTemplate) =>
    async (args: Record<string, any> = {}): Promise<string> => {
      const hostExecute = def[executeKey];
      if (!hostExecute || typeof hostExecute !== 'function') {
        return JSON.stringify(
          {
            success: false,
            error: `Tool definition missing ${executeKey} function`,
            tool: def.name || 'unknown',
          },
          null,
          2,
        );
      }

      try {
        return await runner((context: any) => hostExecute(context, args));
      } catch (error: unknown) {
        return JSON.stringify(
          {
            success: false,
            error: getErrorMessage(error),
            tool: def.name || 'unknown',
            suggestion: 'Fix the error parameters or context and try again.',
          },
          null,
          2,
        );
      }
    };
}

export const optionLists = {
  localLanguageList: [
    { label: 'English', value: 'en' },
    { label: 'Fran\u00e7ais', value: 'fr' },
  ],
};

/**
 * Normalizes line endings to \n (removes \r).
 * This standardizes line endings from various inputs (like LLMs or copy/paste)
 * across Word and PowerPoint.
 */
export function normalizeLineEndings(text: string): string {
  if (!text) return '';
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}
