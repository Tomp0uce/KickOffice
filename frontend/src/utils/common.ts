import DiffMatchPatch from 'diff-match-patch'
import { languageMap } from './constant'

// R17/CH5 — Generate a visual diff HTML string (insertions in blue/underline, deletions in red/strikethrough)
export function generateVisualDiff(originalText: unknown, newText: unknown): string {
  if (typeof originalText !== 'string' || typeof newText !== 'string') {
    return ''
  }
  const dmp = new DiffMatchPatch()
  const diffs = dmp.diff_main(originalText, newText)
  dmp.diff_cleanupSemantic(diffs)

  return diffs
    .map(([op, text]) => {
      const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')
      if (op === 1) return `<span style="color:blue;text-decoration:underline;">${escaped}</span>`
      if (op === -1) return `<span style="color:red;text-decoration:line-through;">${escaped}</span>`
      return escaped
    })
    .join('')
}

export interface TextDiffStats {
  insertions: number
  deletions: number
  unchanged: number
}

/** Compute word-level diff stats between two strings (used by hosts that report diff info without a visual HTML diff). */
export function computeTextDiffStats(originalText: string, revisedText: string): TextDiffStats {
  const dmp = new DiffMatchPatch()
  const diffs = dmp.diff_main(originalText, revisedText)
  dmp.diff_cleanupSemantic(diffs)
  let insertions = 0, deletions = 0, unchanged = 0
  for (const [op, text] of diffs) {
    if (op === 0) unchanged += text.length
    else if (op === -1) deletions += text.length
    else if (op === 1) insertions += text.length
  }
  return { insertions, deletions, unchanged }
}

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
  ) as unknown as Record<TName, TDef>
}

export const optionLists = {
  localLanguageList: [
    { label: 'English', value: 'en' },
    { label: 'Fran\u00e7ais', value: 'fr' },
  ],
}
