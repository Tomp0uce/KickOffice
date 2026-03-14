import {
  getPowerPointSelection,
  getPowerPointSelectionAsHtml,
  getCurrentSlideNumber,
  getSlideContentStandalone,
} from '@/utils/powerpointTools';
import { executeOfficeAction as runOfficeAction } from '@/utils/officeAction';
import {
  getOfficeTextCoercionType,
  getOfficeHtmlCoercionType,
  getOutlookMailbox,
  isOfficeAsyncSucceeded,
  type OfficeAsyncResult,
} from '@/utils/officeOutlook';

declare const Office: any;
declare const Word: any;
declare const Excel: any;
declare const PowerPoint: any;

export interface UseOfficeSelectionOptions {
  hostIsOutlook: boolean;
  hostIsPowerPoint: boolean;
  hostIsExcel: boolean;
}

export function useOfficeSelection(options: UseOfficeSelectionOptions) {
  const { hostIsOutlook, hostIsPowerPoint, hostIsExcel } = options;

  function withTimeout<T>(promise: Promise<T>, ms: number, fallbackValue: T): Promise<T> {
    return new Promise(resolve => {
      let resolved = false;
      const timeoutId = setTimeout(() => {
        if (!resolved) {
          resolved = true;
          resolve(fallbackValue);
        }
      }, ms);

      promise
        .then(val => {
          if (!resolved) {
            resolved = true;
            clearTimeout(timeoutId);
            resolve(val);
          }
        })
        .catch(() => {
          if (!resolved) {
            resolved = true;
            clearTimeout(timeoutId);
            resolve(fallbackValue);
          }
        });
    });
  }

  const getOutlookMailBody = (): Promise<string> => {
    return withTimeout(
      new Promise<string>(resolve => {
        try {
          const mailbox = getOutlookMailbox();
          if (!mailbox?.item) return resolve('');
          mailbox.item.body.getAsync(
            getOfficeTextCoercionType(),
            (result: OfficeAsyncResult<string>) =>
              resolve(isOfficeAsyncSucceeded(result.status) ? result.value || '' : ''),
          );
        } catch {
          resolve('');
        }
      }),
      3000,
      '',
    );
  };

  const getOutlookMailBodyAsHtml = (): Promise<string> => {
    return withTimeout(
      new Promise<string>(resolve => {
        try {
          const mailbox = getOutlookMailbox();
          if (!mailbox?.item) return resolve('');
          const htmlType = getOfficeHtmlCoercionType();
          if (!htmlType) return resolve('');
          mailbox.item.body.getAsync(htmlType, (result: OfficeAsyncResult<string>) =>
            resolve(isOfficeAsyncSucceeded(result.status) ? result.value || '' : ''),
          );
        } catch {
          resolve('');
        }
      }),
      3000,
      '',
    );
  };

  const getOutlookSelectedText = (): Promise<string> => {
    return withTimeout(
      new Promise<string>(resolve => {
        try {
          const mailbox = getOutlookMailbox();
          if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function')
            return resolve('');
          mailbox.item.getSelectedDataAsync(
            getOfficeTextCoercionType(),
            (result: OfficeAsyncResult<{ data?: string }>) =>
              resolve(
                isOfficeAsyncSucceeded(result.status) && result.value?.data
                  ? result.value.data
                  : '',
              ),
          );
        } catch {
          resolve('');
        }
      }),
      3000,
      '',
    );
  };

  const getOutlookSelectedHtml = (): Promise<string> => {
    return withTimeout(
      new Promise<string>(resolve => {
        try {
          const mailbox = getOutlookMailbox();
          if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function')
            return resolve('');
          const htmlType = getOfficeHtmlCoercionType();
          if (!htmlType) return resolve('');
          mailbox.item.getSelectedDataAsync(
            htmlType,
            (result: OfficeAsyncResult<{ data?: string }>) =>
              resolve(
                isOfficeAsyncSucceeded(result.status) && result.value?.data
                  ? result.value.data
                  : '',
              ),
          );
        } catch {
          resolve('');
        }
      }),
      3000,
      '',
    );
  };

  async function getOfficeSelection(selectionOptions?: {
    includeOutlookSelectedText?: boolean;
    actionKey?: string;
  }): Promise<string> {
    if (hostIsOutlook) {
      if (selectionOptions?.includeOutlookSelectedText) {
        const selected = await getOutlookSelectedText();
        if (selected) return selected;
      }
      return getOutlookMailBody();
    }

    if (hostIsPowerPoint) {
      const selection = await getPowerPointSelection();
      if (selection) return selection;

      // PowerPoint fallback: get current slide content
      try {
        const slideNum = await getCurrentSlideNumber();
        return await runOfficeAction(() =>
          PowerPoint.run((ctx: any) => getSlideContentStandalone(ctx, slideNum)),
        );
      } catch {
        return '';
      }
    }

    if (hostIsExcel) {
      return Excel.run(async (ctx: any) => {
        const range = ctx.workbook.getSelectedRange();
        range.load('values, address, rowCount, columnCount');
        await ctx.sync();

        let targetRange = range;
        // If selection is just one cell, check if we should fallback to used range
        if (range.rowCount === 1 && range.columnCount === 1) {
          const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
          const usedRange = activeSheet.getUsedRangeOrNullObject();
          usedRange.load('values, address, rowCount, columnCount, isNullObject');
          await ctx.sync();
          if (!usedRange.isNullObject) {
            targetRange = usedRange;
          }
        }

        const escapeCell = (val: any) => {
          if (val === null || val === undefined) return '';
          const str = String(val);
          if (str.includes('\t') || str.includes('\n') || str.includes('\r')) {
            return `"${str.replace(/"/g, '""')}"`;
          }
          return str;
        };

        return `[${targetRange.address}]\n${targetRange.values.map((row: any[]) => row.map(escapeCell).join('\t')).join('\n')}`;
      });
    }

    // Word Fallback
    return Word.run(async (ctx: any) => {
      const range = ctx.document.getSelection();
      range.load('text');
      await ctx.sync();

      if (range.text) return range.text;

      // Fallback to full document body
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return body.text || '';
    });
  }

  /**
   * Get the selection as HTML to preserve rich content (images, formatting, etc.).
   * Falls back to plain text if HTML is not available.
   * Used by quick actions to preserve non-text elements during LLM processing.
   */
  async function getOfficeSelectionAsHtml(selectionOptions?: {
    includeOutlookSelectedText?: boolean;
    actionKey?: string;
  }): Promise<string> {
    if (hostIsOutlook) {
      if (selectionOptions?.includeOutlookSelectedText) {
        const selectedHtml = await getOutlookSelectedHtml();
        if (selectedHtml) return selectedHtml;
      }

      const html = await getOutlookMailBodyAsHtml();
      return html || getOutlookMailBody();
    }

    if (hostIsExcel) {
      return ''; // Still no HTML for Excel
    }

    if (hostIsPowerPoint) {
      const selectionHtml = await getPowerPointSelectionAsHtml();
      if (selectionHtml) return selectionHtml;

      // PPT Fallback: get text of current slide (no real HTML fallback since it's shape-based)
      try {
        const slideNum = await getCurrentSlideNumber();
        return await runOfficeAction(() =>
          PowerPoint.run((ctx: any) => getSlideContentStandalone(ctx, slideNum)),
        );
      } catch {
        return '';
      }
    }

    // Word Fallback
    try {
      return await Word.run(async (ctx: any) => {
        const range = ctx.document.getSelection();
        const htmlResult = range.getHtml();
        range.load('text');
        await ctx.sync();

        if (range.text) return htmlResult.value || '';

        // Fallback to full document HTML
        const bodyHtml = ctx.document.body.getHtml();
        await ctx.sync();
        return bodyHtml.value || '';
      });
    } catch (err) {
      console.warn('[useOfficeSelection] Word getHtml failed', err);
      return '';
    }
  }

  return { getOfficeSelection, getOfficeSelectionAsHtml };
}
