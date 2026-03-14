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

import { logService } from './logger';
import { getErrorMessage } from './common';
import {
  setChangeTrackingForAi,
  restoreChangeTracking,
  loadRedlineAuthor,
} from './wordTrackChanges';

export interface DocumentRevisionEntry {
  /** Text of the paragraph to revise (used to locate it in the document). */
  originalText: string;
  /** Full revised version of the paragraph text. */
  revisedText: string;
}

export interface DocumentRevisionResult {
  success: boolean;
  applied: number;
  failed: number;
  skipped: number;
  author?: string;
  details: string[];
}

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
  } catch (error: unknown) {
    logService.error('[WordDiff] docx-redline-js error:', error instanceof Error ? error : new Error(String(error)));
    return {
      success: false,
      strategy: 'none',
      message: `Error generating revision markup: ${getErrorMessage(error)}`,
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
  } catch (insertError: unknown) {
    // Fallback: if insertOoxml fails (Word Online), use direct text replacement
    logService.warn('[WordDiff] insertOoxml failed, falling back to insertText:', insertError instanceof Error ? insertError : new Error(String(insertError)));
    try {
      selection.insertText(revisedText, 'Replace');
      await context.sync();
    } catch (fallbackError: unknown) {
      return {
        success: false,
        strategy: 'none',
        message: `Error applying revision: ${getErrorMessage(fallbackError)}`,
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

/**
 * Apply multiple paragraph-level revisions to the entire document using docx-redline-js.
 *
 * Each entry identifies a paragraph by its current text (originalText) and provides
 * the full revised version (revisedText). The function locates the first matching
 * paragraph, generates Track Changes OOXML via applyRedlineToOxml(), and inserts it.
 *
 * Track Changes is disabled once for the entire batch (same pattern as applyRevisionToSelection)
 * so native tracking doesn't double-record the injected w:ins/w:del markup.
 *
 * IMPORTANT: Must be called within Word.run() context.
 */
export async function applyRevisionToDocument(
  context: Word.RequestContext,
  revisions: DocumentRevisionEntry[],
  enableTrackChanges: boolean = true,
): Promise<DocumentRevisionResult> {
  const redlineAuthor = loadRedlineAuthor();

  // Load all paragraphs once
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items');
  await context.sync();

  const items = paragraphs.items;
  items.forEach(p => p.load('text'));
  await context.sync();

  const details: string[] = [];
  let applied = 0;
  let failed = 0;
  let skipped = 0;

  const trackingState = await setChangeTrackingForAi(
    context,
    enableTrackChanges,
    'proposeDocumentRevision',
  );

  try {
    for (const { originalText, revisedText } of revisions) {
      // Find the first paragraph whose trimmed text matches
      const paraIndex = items.findIndex(p => p.text.trim() === originalText.trim());
      if (paraIndex === -1) {
        details.push(
          `[NOT FOUND] "${originalText.slice(0, 60)}${originalText.length > 60 ? '…' : ''}"`,
        );
        failed++;
        continue;
      }

      if (originalText.trim() === revisedText.trim()) {
        details.push(`[${paraIndex}] SKIPPED: text identical`);
        skipped++;
        continue;
      }

      const para = items[paraIndex];
      const ooxmlResult = para.getOoxml();
      await context.sync();
      const ooxml = ooxmlResult.value;

      setDefaultAuthor(redlineAuthor);
      let resultOoxml: string;
      try {
        const redlineResult = await applyRedlineToOxml(ooxml, para.text, revisedText, {
          author: enableTrackChanges ? redlineAuthor : undefined,
          generateRedlines: enableTrackChanges,
        });
        resultOoxml = redlineResult.oxml;
      } catch (err: any) {
        details.push(
          `[${paraIndex}] ERROR (redline): ${err.message || String(err)}`,
        );
        failed++;
        continue;
      }

      try {
        para.insertOoxml(resultOoxml, 'Replace');
        await context.sync();
        details.push(`[${paraIndex}] OK`);
        applied++;
      } catch (insertError: unknown) {
        // Fallback: direct text replacement (Word Online may not support insertOoxml on paragraphs)
        try {
          para.insertText(revisedText, 'Replace');
          await context.sync();
          details.push(`[${paraIndex}] OK (direct-replace fallback)`);
          applied++;
        } catch (fallbackError: unknown) {
          details.push(`[${paraIndex}] ERROR (insert): ${getErrorMessage(fallbackError)}`);
          failed++;
        }
      }
    }
  } finally {
    await restoreChangeTracking(context, trackingState, 'proposeDocumentRevision');
  }

  return {
    success: failed === 0,
    applied,
    failed,
    skipped,
    author: enableTrackChanges ? redlineAuthor : undefined,
    details,
  };
}
