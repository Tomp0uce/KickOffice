/**
 * useDocumentUndo — Save and restore document state before/after LLM interventions.
 *
 * Captures the current selection content (HTML) before an insert operation
 * and allows the user to revert to the previous state with a single click.
 *
 * Supported hosts:
 * - Word: captures selection HTML, wraps inserted content in a Content Control
 *   for reliable undo targeting, then restores original HTML on undo.
 * - Outlook: captures and restores selected HTML via body.setSelectedDataAsync.
 *
 * For Excel/PowerPoint, undo is not yet implemented (agent-based actions
 * use Track Changes or surgical replacements which have their own accept/reject).
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

interface UndoSnapshot {
  /** The original HTML content before the LLM intervention */
  html: string;
  /** Which host the snapshot came from */
  host: 'word' | 'outlook';
  /** Timestamp of the snapshot */
  timestamp: number;
  /** Tag of the content control wrapping the inserted content (Word only) */
  contentControlTag?: string;
}

const UNDO_CC_TAG_PREFIX = 'ko-undo-';

export function useDocumentUndo(options: {
  hostIsWord: boolean;
  hostIsOutlook: boolean;
}) {
  const { hostIsWord, hostIsOutlook } = options;

  /** The saved state that can be restored */
  const undoSnapshot = ref<UndoSnapshot | null>(null);

  /** Whether an undo operation is available */
  const canUndo = ref(false);

  /**
   * Capture current selection state before an insert/quick action.
   * Must be called BEFORE the insert operation modifies the document.
   */
  async function captureBeforeInsert(): Promise<string | null> {
    try {
      if (hostIsWord) {
        return await captureWordSelection();
      }
      if (hostIsOutlook) {
        return await captureOutlookSelection();
      }
    } catch (err) {
      logService.warn('[useDocumentUndo] Failed to capture selection for undo', err);
    }
    return null;
  }

  async function captureWordSelection(): Promise<string | null> {
    return new Promise<string | null>((resolve) => {
      Word.run(async (context: any) => {
        const selection = context.document.getSelection();
        const htmlResult = selection.getHtml();
        await context.sync();
        resolve(htmlResult.value || null);
      }).catch(() => resolve(null));
    });
  }

  async function captureOutlookSelection(): Promise<string | null> {
    return new Promise<string | null>((resolve) => {
      const mailbox = getOutlookMailbox();
      const item = mailbox?.item;
      if (!item?.body?.getSelectedDataAsync) {
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

  /**
   * Save the snapshot and mark undo as available.
   * Called after captureBeforeInsert and before the actual insert.
   */
  function saveSnapshot(html: string, host: 'word' | 'outlook', tag?: string) {
    undoSnapshot.value = {
      html,
      host,
      timestamp: Date.now(),
      contentControlTag: tag,
    };
    canUndo.value = true;
  }

  /**
   * Wrap the just-inserted content in a Word Content Control for undo targeting.
   * Returns the tag assigned to the content control.
   */
  async function wrapInsertedContentInWord(): Promise<string | null> {
    const tag = `${UNDO_CC_TAG_PREFIX}${Date.now()}`;
    try {
      await Word.run(async (context: any) => {
        const selection = context.document.getSelection();
        const cc = selection.insertContentControl();
        cc.tag = tag;
        cc.appearance = 'BoundingBox';
        cc.color = '#4A90D9';
        cc.title = 'KickOffice — Ctrl+Z to undo';
        await context.sync();
      });
      return tag;
    } catch (err) {
      logService.warn('[useDocumentUndo] Failed to wrap inserted content in CC', err);
      return null;
    }
  }

  /**
   * Undo the last insert operation — restore the original content.
   */
  async function undoLastInsert(): Promise<boolean> {
    const snapshot = undoSnapshot.value;
    if (!snapshot) return false;

    try {
      if (snapshot.host === 'word' && snapshot.contentControlTag) {
        return await undoWordInsert(snapshot);
      }
      if (snapshot.host === 'outlook') {
        return await undoOutlookInsert(snapshot);
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
        snapshot.html,
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

  /** Clear the undo state (e.g., when starting a new chat) */
  function clearUndo() {
    undoSnapshot.value = null;
    canUndo.value = false;
  }

  return {
    canUndo,
    captureBeforeInsert,
    saveSnapshot,
    wrapInsertedContentInWord,
    undoLastInsert,
    clearUndo,
  };
}
