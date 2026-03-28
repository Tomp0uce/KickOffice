/**
 * officeDocumentContext.ts
 *
 * Fetches lightweight document metadata for each Office host and returns it as a
 * JSON string to be injected as <doc_context> into every agent request.
 *
 * This mirrors Open_Excel's `getWorkbookMetadata()` pattern (Issue #1 of AGENT_MODE_ANALYSIS.md):
 * the model receives the workbook/document structure automatically without needing to call
 * a discovery tool first.
 */

import { executeOfficeAction } from './officeAction';

declare const Excel: any;
declare const PowerPoint: any;

/**
 * Excel — workbook metadata: all sheet names + usedRange dimensions, active sheet, selected range.
 */
export async function getExcelDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() =>
      Excel.run(async (context: any) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;

        worksheets.load('items/name');

        const activeSheet = worksheets.getActiveWorksheet();
        activeSheet.load('name');

        const selectedRange = workbook.getSelectedRange();
        selectedRange.load('address');

        await context.sync();

        // Batch-load usedRange for every sheet in a single sync
        const usedRanges = worksheets.items.map((sheet: any) => {
          const ur = sheet.getUsedRangeOrNullObject();
          ur.load(['rowCount', 'columnCount', 'isNullObject']);
          return ur;
        });

        await context.sync();

        const sheets = worksheets.items.map((sheet: any, i: number) => {
          const ur = usedRanges[i];
          return {
            name: sheet.name,
            rows: ur.isNullObject ? 0 : ur.rowCount,
            columns: ur.isNullObject ? 0 : ur.columnCount,
          };
        });

        return JSON.stringify(
          {
            activeSheet: activeSheet.name,
            selectedRange: selectedRange.address,
            totalSheets: worksheets.items.length,
            sheets,
          },
          null,
          2,
        );
      }),
    );
  } catch {
    return '';
  }
}

/**
 * PowerPoint — presentation metadata: total slides, slide number + first text line per slide.
 */
export async function getPowerPointDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() => {
      const PPT = PowerPoint;
      if (typeof PPT?.run !== 'function') return Promise.resolve('');

      return PPT.run(async (context: any) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        // Batch-load shapes for all slides
        for (const slide of slides.items) {
          slide.shapes.load('items');
        }
        await context.sync();

        // Batch-load textFrame.textRange.text for all shapes
        for (const slide of slides.items) {
          for (const shape of slide.shapes.items) {
            try {
              shape.textFrame.textRange.load('text');
            } catch {
              // Non-text shape — skip
            }
          }
        }
        await context.sync();

        const slideInfo = slides.items.map((slide: any, i: number) => {
          let title = '';
          for (const shape of slide.shapes.items) {
            try {
              const text = (shape.textFrame?.textRange?.text || '').trim();
              if (text) {
                title = text.substring(0, 80);
                break;
              }
            } catch {
              // skip
            }
          }
          return { slideNumber: i + 1, title: title || '<No text>' };
        });

        let activeSlideNumber = 1;
        try {
          if (typeof context.presentation.getSelectedSlides === 'function') {
            const selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load('items/id');
            await context.sync();
            if (selectedSlides.items.length > 0) {
              const selectedId = selectedSlides.items[0].id;
              const idx = slides.items.findIndex((s: any) => s.id === selectedId);
              if (idx !== -1) {
                activeSlideNumber = idx + 1;
              }
            }
          }
        } catch (e) {
          // ignore error if getSelectedSlides fails
        }

        return JSON.stringify(
          {
            activeSlideNumber,
            totalSlides: slides.items.length,
            slides: slideInfo,
          },
          null,
          2,
        );
      });
    });
  } catch {
    return '';
  }
}

/**
 * Outlook — email metadata: subject, sender, recipients, body snippet (first 300 chars).
 */
export function getOutlookDocumentContext(): Promise<string> {
  return new Promise(resolve => {
    try {
      const Office = (window as any).Office;
      const mailbox = Office?.context?.mailbox;
      if (!mailbox?.item) {
        resolve('');
        return;
      }

      const item = mailbox.item;
      const subject = item.subject || '';
      const sender = item.sender
        ? `${item.sender.displayName || ''} <${item.sender.emailAddress || ''}>`.trim()
        : item.from
          ? `${item.from.displayName || ''} <${item.from.emailAddress || ''}>`.trim()
          : '';

      // itemType: 'message' | 'appointment' — compose vs read determined by presence of setAsync
      const itemType: string = item.itemType || (item.body?.setAsync ? 'compose' : 'read');

      const contextObj: Record<string, unknown> = { itemType, subject, sender };

      // Try to read recipients (compose mode only)
      if (item.to?.getAsync) {
        item.to.getAsync(
          (toResult: {
            status: string;
            value?: { emailAddress?: string; displayName?: string }[];
          }) => {
            if (
              toResult.status === Office?.AsyncResultStatus?.Succeeded &&
              Array.isArray(toResult.value)
            ) {
              contextObj.recipients = toResult.value
                .map(r => r.emailAddress || r.displayName || '')
                .slice(0, 5);
            }
            readBody();
          },
        );
      } else {
        readBody();
      }

      function readBody() {
        if (!item.body?.getAsync) {
          resolve(JSON.stringify(contextObj, null, 2));
          return;
        }
        item.body.getAsync(
          Office?.CoercionType?.Text,
          (result: { status: string; value?: string }) => {
            if (result.status === Office?.AsyncResultStatus?.Succeeded) {
              const bodyText = String(result.value || '');
              contextObj.bodySnippet =
                bodyText.substring(0, 300) + (bodyText.length > 300 ? '...' : '');
            }
            resolve(JSON.stringify(contextObj, null, 2));
          },
        );
      }
    } catch {
      resolve('');
    }
  });
}

/**
 * Word — enriched document metadata.
 * Ported and extended from Office Agents adapter.tsx getDocumentMetadata().
 * Includes: page/word count, paragraph/table counts, heading outline,
 * content controls, run-level formatting detection, and style info.
 */
export async function getWordDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() => {
      const Word = (window as any).Word;
      if (typeof Word?.run !== 'function') return Promise.resolve('');

      return Word.run(async (context: any) => {
        const props = context.document.properties;
        props.load(['pageCount', 'wordCount']);

        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load('items');

        const tables = body.tables;
        tables.load('items');

        const inlinePictures = body.inlinePictures;
        inlinePictures.load('items');

        await context.sync();

        // Load paragraph details for heading outline and formatting sample
        // Sample first 30 paragraphs for run-level override detection
        const sampleSize = Math.min(paragraphs.items.length, 30);
        for (let i = 0; i < sampleSize; i++) {
          const para = paragraphs.items[i];
          para.load('text,style,alignment');
          try {
            para.font.load('name,size,color,bold,italic');
          } catch {
            // font load may fail in some contexts
          }
        }

        // Load content controls if available
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let contentControls: any = null;
        try {
          contentControls = body.contentControls;
          contentControls.load('items/title,items/tag,items/type');
        } catch {
          // contentControls may not be available
        }

        await context.sync();

        // Build heading outline — ported from Office Agents get_document_structure
        const headings: { level: number; text: string; paragraphIndex: number }[] = [];
        let hasRunLevelOverrides = false;
        const formatSample: { index: number; style: string; font?: string; size?: number }[] = [];

        for (let i = 0; i < sampleSize; i++) {
          const para = paragraphs.items[i];
          const styleName = para.style || '';
          const text = (para.text || '').trim();

          // Detect headings
          const headingMatch = styleName.match(/[Hh]eading\s*(\d)/);
          if (headingMatch) {
            headings.push({
              level: parseInt(headingMatch[1], 10),
              text: text.substring(0, 80),
              paragraphIndex: i,
            });
          }

          // Detect run-level formatting overrides
          // If a paragraph has direct font formatting different from its style,
          // it means OOXML-based editing is needed to preserve formatting
          try {
            const fontName = para.font?.name;
            const fontSize = para.font?.size;
            if (fontName || fontSize) {
              formatSample.push({
                index: i,
                style: styleName,
                font: fontName,
                size: fontSize,
              });
              // If we see direct formatting on non-heading paragraphs, flag it
              if (!styleName.match(/[Hh]eading/) && (fontName || fontSize)) {
                hasRunLevelOverrides = true;
              }
            }
          } catch {
            // font properties may not be loaded
          }
        }

        // Content control info
        let contentControlInfo: { title: string; tag: string; type: string }[] = [];
        try {
          if (contentControls?.items?.length > 0) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            contentControlInfo = contentControls.items.map((cc: any) => ({
              title: cc.title || '',
              tag: cc.tag || '',
              type: cc.type || '',
            }));
          }
        } catch {
          // ignore
        }

        const result: Record<string, unknown> = {
          pageCount: props.pageCount,
          wordCount: props.wordCount,
          paragraphCount: paragraphs.items.length,
          tableCount: tables.items.length,
          hasImages: inlinePictures.items.length > 0,
        };

        // Only include enriched metadata if there's meaningful content
        if (headings.length > 0) {
          result.headingOutline = headings;
        }
        if (contentControlInfo.length > 0) {
          result.contentControls = contentControlInfo;
        }
        if (hasRunLevelOverrides) {
          result.hasRunLevelOverrides = true;
          result.formattingHint =
            'Document has direct run-level formatting. Use getDocumentOoxml + editDocumentXml for edits that must preserve formatting.';
        }
        if (formatSample.length > 0) {
          result.formatSample = formatSample.slice(0, 10); // First 10 to keep context small
        }

        return JSON.stringify(result, null, 2);
      });
    });
  } catch {
    return '';
  }
}
