/**
 * Word Diff Utilities
 *
 * Wrapper around office-word-diff for surgical text editing.
 * Preserves formatting by computing word-level diffs and applying
 * only the changes, not replacing entire ranges.
 */

import { OfficeWordDiff, getDiffStats, computeDiff } from 'office-word-diff'
import type { DiffResult, DiffStats } from 'office-word-diff'

export interface RevisionResult {
  success: boolean
  strategy: 'token' | 'sentence' | 'block'
  insertions: number
  deletions: number
  unchanged: number
  message: string
}

/**
 * Apply a revision to selected text using word-level diffing.
 *
 * IMPORTANT: Must be called within Word.run() context.
 *
 * @param context - Word.RequestContext from Word.run()
 * @param revisedText - The new version of the text
 * @param enableTrackChanges - Show changes in Word's Track Changes (default: true)
 * @returns RevisionResult with operation details
 *
 * @example
 * await Word.run(async (context) => {
 *   const result = await applyRevisionToSelection(context, "New text here", true);
 *   console.log(`Applied ${result.insertions} insertions using ${result.strategy} strategy`);
 * });
 */
export async function applyRevisionToSelection(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true
): Promise<RevisionResult> {
  // 1. Get selection and load text
  const range = context.document.getSelection()
  range.load('text')
  await context.sync()

  const originalText = range.text

  // 2. Handle edge cases
  if (!originalText || !originalText.trim()) {
    return {
      success: false,
      strategy: 'block',
      insertions: 0,
      deletions: 0,
      unchanged: 0,
      message: 'Error: No text selected. Please select text before using proposeRevision.',
    }
  }

  if (originalText === revisedText) {
    return {
      success: true,
      strategy: 'token',
      insertions: 0,
      deletions: 0,
      unchanged: originalText.length,
      message: 'Text is identical, no changes needed.',
    }
  }

  // 3. Preview stats before applying
  const stats = getDiffStats(originalText, revisedText)

  // 4. Apply diff with cascading fallback
  const differ = new OfficeWordDiff({
    enableTracking: enableTrackChanges,
    logLevel: 'info',
    onLog: (msg, level) => {
      if (level === 'error') console.error('[WordDiff]', msg)
      else if (level === 'warn') console.warn('[WordDiff]', msg)
      else console.log('[WordDiff]', msg)
    },
  })

  try {
    const result = await differ.applyDiff(context, range, originalText, revisedText)

    return {
      success: result.success,
      strategy: result.strategyUsed,
      insertions: result.insertions,
      deletions: result.deletions,
      unchanged: stats.unchanged,
      message: result.success
        ? `Successfully applied ${result.insertions} insertions and ${result.deletions} deletions using ${result.strategyUsed} strategy.`
        : `Diff application failed. Check logs for details.`,
    }
  } catch (error: any) {
    console.error('[WordDiff] Unexpected error:', error)
    return {
      success: false,
      strategy: 'block',
      insertions: 0,
      deletions: 0,
      unchanged: 0,
      message: `Error applying revision: ${error.message || String(error)}`,
    }
  }
}

/**
 * Preview diff statistics without applying changes.
 * Does NOT require Word context - can be used for UI preview.
 */
export function previewDiffStats(originalText: string, revisedText: string): DiffStats {
  return getDiffStats(originalText, revisedText)
}

/**
 * Compute raw diff operations for debugging/display.
 * Does NOT require Word context.
 *
 * @returns Array of [operation, text] tuples:
 *   - [0, "text"] = unchanged
 *   - [-1, "text"] = deletion
 *   - [1, "text"] = insertion
 */
export function computeRawDiff(originalText: string, revisedText: string): Array<[number, string]> {
  return computeDiff(originalText, revisedText)
}

/**
 * Check if text has complex content that may not diff well.
 * Warns about tables, images, and other non-text content.
 */
export function hasComplexContent(text: string): boolean {
  // Check for table markers, image placeholders, or other special content
  const complexPatterns = [
    /\t.*\t/,           // Tab-separated (likely table)
    /\[Image\]/i,       // Image placeholder
    /\[Figure\]/i,      // Figure placeholder
    /^\s*\|.*\|/m,      // Markdown table row
  ]
  return complexPatterns.some(pattern => pattern.test(text))
}
