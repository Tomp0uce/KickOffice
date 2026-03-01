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

export const optionLists = {
  localLanguageList: [
    { label: 'English', value: 'en' },
    { label: 'Fran\u00e7ais', value: 'fr' },
  ],
}
