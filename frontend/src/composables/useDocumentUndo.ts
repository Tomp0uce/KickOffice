/**
 * useDocumentUndo — Save and restore document state before/after LLM interventions.
 *
 * Captures the current selection content before an insert operation
 * and allows the user to revert to the previous state with a single click.
 *
 * Supported hosts:
 * - Word: captures selection HTML, wraps inserted content in a Content Control
 *   for reliable undo targeting, then restores original HTML on undo.
 * - Outlook: captures and restores selected HTML via body.setSelectedDataAsync.
 * - Excel: captures selected range address + cell values, restores them on undo.
 * - PowerPoint: captures selected text, restores via setSelectedDataAsync on undo
 *   (relies on the inserted text remaining selected immediately after insertion).
 */

import { ref } from 'vue';
import { logService } from '@/utils/logger';
import {
  getOfficeHtmlCoercionType,
  getOutlookMailbox,
  isOfficeAsyncSucceeded,
  type OfficeAsyncResult,
} from '@/utils/officeOutlook';

declare const Word: any;
declare const Excel: any;
declare const Office: any;

interface UndoSnapshot {
  /** Which host the snapshot came from */
  host: 'word' | 'outlook' | 'excel' | 'powerpoint';
  /** Timestamp of the snapshot */
  timestamp: number;

  // Word / Outlook
  /** The original HTML content before the LLM intervention */
  html?: string;
  /** Tag of the content control wrapping the inserted content (Word only) */
  contentControlTag?: string;

  // Excel
  /** Address of the captured range (e.g. "Sheet1!A1:C3") */
  excelRangeAddress?: string;
  /** Original cell values of the captured range */
  excelValues?: any[][];

  // PowerPoint
  /** Original plain-text content of the captured selection */
  pptText?: string;
}

const UNDO_CC_TAG_PREFIX = 'ko-undo-';

export function useDocumentUndo(options: {
  hostIsWord: boolean;
  hostIsOutlook: boolean;
  hostIsExcel?: boolean;
  hostIsPowerPoint?: boolean;
}) {
  const { hostIsWord, hostIsOutlook, hostIsExcel = false, hostIsPowerPoint = false } = options;

  /** The saved state that can be restored */
  const undoSnapshot = ref<UndoSnapshot | null>(null);

  /** Whether an undo operation is available */
  const canUndo = ref(false);

  /**
   * Capture current selection state before an insert/quick action.
   * Must be called BEFORE the insert operation modifies the document.
   * Returns a partial snapshot object (or null if capture failed).
   */
  async function captureBeforeInsert(): Promise<Partial<UndoSnapshot> | null> {
    try {
      if (hostIsWord) {
        return await captureWordSelectionAndWrap();
      }
      if (hostIsOutlook) {
        const html = await captureOutlookSelection();
        return html != null ? { host: 'outlook', html } : null;
      }
      if (hostIsExcel) {
        return await captureExcelSelection();
      }
      if (hostIsPowerPoint) {
        return await capturePowerPointSelection();
      }
    } catch (err) {
      logService.warn('[useDocumentUndo] Failed to capture selection for undo', err);
    }
    return null;
  }

  /**
   * Capture the current Word selection AND wrap it in a Content Control.
   * The CC acts as a persistent anchor so the subsequent insert (which targets the CC)
   * and the eventual undo (which finds the CC by tag) are always in sync.
   */
  async function captureWordSelectionAndWrap(): Promise<Partial<UndoSnapshot> | null> {
    const tag = `${UNDO_CC_TAG_PREFIX}${Date.now()}`;
    return new Promise<Partial<UndoSnapshot> | null>((resolve) => {
      Word.run(async (context: any) => {
        const selection = context.document.getSelection();
        const htmlResult = selection.getHtml();
        // Wrap the selection in a hidden CC — this becomes the undo anchor.
        // Even for empty selections (cursor only), an empty CC is created at the cursor.
        const cc = selection.insertContentControl();
        cc.tag = tag;
        cc.appearance = 'Hidden';
        cc.title = 'KickOffice Undo';
        await context.sync();
        resolve({
          host: 'word',
          html: htmlResult.value ?? '',
          contentControlTag: tag,
        });
      }).catch(() => resolve(null));
    });
  }

  async function captureOutlookSelection(): Promise<string | null> {
    return new Promise<string | null>((resolve) => {
      const mailbox = getOutlookMailbox();
      const item = mailbox?.item;
      if (!(item?.body as any)?.getSelectedDataAsync) {
        resolve(null);
        return;
      }
      // @ts-expect-error Office.js async callback
      item.body.getSelectedDataAsync(
        getOfficeHtmlCoercionType(),
        (result: OfficeAsyncResult) => {
          if (isOfficeAsyncSucceeded(result.status) && result.value) {
            resolve(result.value as string);
          } else {
            resolve(null);
          }
        },
      );
    });
  }

  async function captureExcelSelection(): Promise<Partial<UndoSnapshot> | null> {
    return new Promise<Partial<UndoSnapshot> | null>((resolve) => {
      Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'values']);
        await context.sync();
        resolve({
          host: 'excel',
          excelRangeAddress: range.address as string,
          excelValues: range.values as any[][],
        });
      }).catch(() => resolve(null));
    });
  }

  async function capturePowerPointSelection(): Promise<Partial<UndoSnapshot> | null> {
    return new Promise<Partial<UndoSnapshot> | null>((resolve) => {
      if (typeof Office === 'undefined' || !Office?.context?.document?.getSelectedDataAsync) {
        resolve(null);
        return;
      }
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        (result: OfficeAsyncResult) => {
          if (isOfficeAsyncSucceeded(result.status)) {
            resolve({ host: 'powerpoint', pptText: (result.value as string) ?? '' });
          } else {
            resolve(null);
          }
        },
      );
    });
  }

  /**
   * Save the snapshot and mark undo as available.
   * Called after captureBeforeInsert returns.
   */
  function saveSnapshot(partial: Partial<UndoSnapshot>, tag?: string) {
    undoSnapshot.value = {
      ...partial,
      timestamp: Date.now(),
      // Only override contentControlTag when an explicit tag is passed (backward compat).
      // The new Word flow embeds the tag directly in the partial from captureWordSelectionAndWrap.
      ...(tag !== undefined ? { contentControlTag: tag } : {}),
    } as UndoSnapshot;
    canUndo.value = true;
  }

  /**
   * Undo the last insert operation — restore the original content.
   * Snapshot is always cleared after one attempt (success or failure) so the
   * undo button never gets stuck showing the error on repeated clicks.
   */
  async function undoLastInsert(): Promise<boolean> {
    const snapshot = undoSnapshot.value;
    if (!snapshot) return false;

    // Clear immediately so the button disappears regardless of outcome.
    // If the CC was removed by an OOXML tool the undo is unrecoverable anyway.
    undoSnapshot.value = null;
    canUndo.value = false;

    try {
      if (snapshot.host === 'word' && snapshot.contentControlTag) {
        return await undoWordInsert(snapshot);
      }
      if (snapshot.host === 'outlook') {
        return await undoOutlookInsert(snapshot);
      }
      if (snapshot.host === 'excel') {
        return await undoExcelInsert(snapshot);
      }
      if (snapshot.host === 'powerpoint') {
        return await undoPowerPointInsert(snapshot);
      }
    } catch (err) {
      logService.error('[useDocumentUndo] Undo failed', err);
    }
    return false;
  }

  async function undoWordInsert(snapshot: UndoSnapshot): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      Word.run(async (context: any) => {
        const ccs = context.document.contentControls.getByTag(snapshot.contentControlTag!);
        ccs.load('items');
        await context.sync();

        if (ccs.items.length === 0) {
          resolve(false);
          return;
        }

        const cc = ccs.items[0];

        if (snapshot.html) {
          // Restore original HTML content
          cc.insertHtml(snapshot.html, 'Replace');
        } else {
          // Original selection was empty — delete the inserted content
          cc.delete(false);
        }
        await context.sync();

        // Clean up: remove the content control wrapper
        try {
          ccs.load('items');
          await context.sync();
          if (ccs.items.length > 0) {
            ccs.items[0].delete(true); // keep content, remove CC wrapper
            await context.sync();
          }
        } catch {
          // Content control may already be gone after delete — OK
        }

        undoSnapshot.value = null;
        canUndo.value = false;
        resolve(true);
      }).catch(() => resolve(false));
    });
  }

  async function undoOutlookInsert(snapshot: UndoSnapshot): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      const mailbox = getOutlookMailbox();
      const item = mailbox?.item;
      if (!item?.body?.setSelectedDataAsync) {
        resolve(false);
        return;
      }
      item.body.setSelectedDataAsync!(
        snapshot.html ?? '',
        { coercionType: getOfficeHtmlCoercionType() },
        (result: OfficeAsyncResult) => {
          if (isOfficeAsyncSucceeded(result.status)) {
            undoSnapshot.value = null;
            canUndo.value = false;
            resolve(true);
          } else {
            resolve(false);
          }
        },
      );
    });
  }

  async function undoExcelInsert(snapshot: UndoSnapshot): Promise<boolean> {
    if (!snapshot.excelRangeAddress || !snapshot.excelValues) return false;
    return new Promise<boolean>((resolve) => {
      Excel.run(async (context: any) => {
        // range.address includes sheet name (e.g. "Sheet1!A1:B2") — use workbook-level getRange
        const range = context.workbook.getRange(snapshot.excelRangeAddress!);
        range.values = snapshot.excelValues!;
        await context.sync();
        undoSnapshot.value = null;
        canUndo.value = false;
        resolve(true);
      }).catch(() => resolve(false));
    });
  }

  async function undoPowerPointInsert(snapshot: UndoSnapshot): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      if (typeof Office === 'undefined' || !Office?.context?.document?.setSelectedDataAsync) {
        resolve(false);
        return;
      }
      // Restore: replace current selection (which should still be the just-inserted text)
      // with the original text that was there before the insert.
      Office.context.document.setSelectedDataAsync(
        snapshot.pptText ?? '',
        { coercionType: Office.CoercionType.Text },
        (result: OfficeAsyncResult) => {
          if (isOfficeAsyncSucceeded(result.status)) {
            undoSnapshot.value = null;
            canUndo.value = false;
            resolve(true);
          } else {
            resolve(false);
          }
        },
      );
    });
  }

  /** Clear the undo state (e.g., when starting a new chat) */
  function clearUndo() {
    undoSnapshot.value = null;
    canUndo.value = false;
  }

  return {
    canUndo,
    captureBeforeInsert,
    saveSnapshot,
    undoLastInsert,
    clearUndo,
  };
}
