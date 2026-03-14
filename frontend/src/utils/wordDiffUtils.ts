/**
 * Word Diff Utilities — v2 (docx-redline-js)
 *
 * Generates native Word Track Changes (w:ins / w:del) by injecting
 * revision markup into paragraph OOXML via docx-redline-js.
 *
 * Integration pattern from Gemini AI for Office (MIT License):
 * https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
 * OOXML engine: https://github.com/AnsonLai/docx-redline-js (MIT License)
 */

import { applyRedlineToOxml, setDefaultAuthor } from '@ansonlai/docx-redline-js';

import {
  setChangeTrackingForAi,
  restoreChangeTracking,
  loadRedlineAuthor,
} from './wordTrackChanges';

export interface RevisionResult {
  success: boolean;
  strategy: 'redline' | 'direct-replace' | 'none';
  author?: string;
  message: string;
}

/**
 * Apply a revision to the current selection using docx-redline-js.
 *
 * Follows the Gemini AI for Office pattern:
 * 1. Extract selection text + OOXML via getOoxml()
 * 2. Generate revision markup via applyRedlineToOxml() (w:ins / w:del)
 * 3. Disable Track Changes via setChangeTrackingForAi() (prevent double-tracking)
 * 4. Insert modified OOXML via insertOoxml() (revision markup survives)
 * 5. Restore Track Changes via restoreChangeTracking()
 *
 * IMPORTANT: Must be called within Word.run() context.
 */
export async function applyRevisionToSelection(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true,
): Promise<RevisionResult> {
  const redlineAuthor = loadRedlineAuthor();

  // 1. Get selection text + OOXML in a single sync batch
  const selection = context.document.getSelection();
  selection.load('text');
  const ooxmlResult = selection.getOoxml();
  await context.sync();

  const originalText = selection.text;
  const ooxml = ooxmlResult.value;

  // 2. Edge cases
  if (!originalText || !originalText.trim()) {
    return {
      success: false,
      strategy: 'none',
      message: 'Error: No text selected. Please select text before using proposeRevision.',
    };
  }

  if (originalText === revisedText) {
    return {
      success: true,
      strategy: 'none',
      message: 'Text is identical, no changes needed.',
    };
  }

  // 3. Generate revision markup via docx-redline-js
  setDefaultAuthor(redlineAuthor);

  let resultOoxml: string;
  try {
    const redlineResult = await applyRedlineToOxml(ooxml, originalText, revisedText, {
      author: enableTrackChanges ? redlineAuthor : undefined,
      generateRedlines: enableTrackChanges,
    });
    resultOoxml = redlineResult.oxml;
  } catch (error: any) {
    console.error('[WordDiff] docx-redline-js error:', error);
    return {
      success: false,
      strategy: 'none',
      message: `Error generating revision markup: ${error.message || String(error)}`,
    };
  }

  // 4. Disable Track Changes, insert, restore — pattern from Gemini AI for Office
  const trackingState = await setChangeTrackingForAi(
    context,
    enableTrackChanges,
    'proposeRevision',
  );

  try {
    // Insert the modified OOXML
    // w:ins/w:del survive because native tracking is OFF
    selection.insertOoxml(resultOoxml, 'Replace');
    await context.sync();
  } catch (insertError: any) {
    // Fallback: if insertOoxml fails (Word Online), use direct text replacement
    console.warn('[WordDiff] insertOoxml failed, falling back to insertText:', insertError);
    try {
      selection.insertText(revisedText, 'Replace');
      await context.sync();
    } catch (fallbackError: any) {
      return {
        success: false,
        strategy: 'none',
        message: `Error applying revision: ${fallbackError.message || String(fallbackError)}`,
      };
    }
    return {
      success: true,
      strategy: 'direct-replace',
      message: 'Revision applied with direct replacement (insertOoxml unavailable).',
    };
  } finally {
    // 5. ALWAYS restore the original tracking mode
    await restoreChangeTracking(context, trackingState, 'proposeRevision');
  }

  return {
    success: true,
    strategy: enableTrackChanges ? 'redline' : 'direct-replace',
    author: enableTrackChanges ? redlineAuthor : undefined,
    message: enableTrackChanges
      ? `Revision applied with Track Changes (author: "${redlineAuthor}").`
      : 'Revision applied with direct replacement (no Track Changes).',
  };
}
