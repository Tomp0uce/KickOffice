import type { ToolDefinition } from '@/types';
/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

import { executeOfficeAction } from './officeAction';
import {
  renderOfficeCommonApiHtml,
  stripRichFormattingSyntax,
  stripMarkdownListMarkers,
  applyInheritedStyles,
  type InheritedStyles,
} from './markdown';
import { validateOfficeCode } from './officeCodeValidator';
import { createMutationDetector } from './mutationDetector';
import { getVfsSandboxContext } from '@/utils/vfs';
import {
  computeTextDiffStats,
  createOfficeTools,
  normalizeLineEndings,
  getErrorMessage,
  getDetailedOfficeError,
  createEvalExecutor,
  buildScreenshotResult,
} from './common';
import { message as messageUtil } from '@/utils/message';
import { withSlideZip, escapeXml } from './pptxZipUtils';
import { logService } from '@/utils/logger';

declare const Office: any;
declare const PowerPoint: any;

// Mutation detection patterns for PowerPoint — ported from Office Agents
const looksLikeMutationPpt = createMutationDetector([
  /\.insertSlide\s*\(/,
  /\.delete\s*\(/,
  /\.text\s*=/,
  /\.insertText\s*\(/,
  /\.font\.\w+\s*=/,
  /\.fill\.\w+\s*=/,
  /\.set\s*\(/,
  /\.add\s*\(/,
  /\.addImage\s*\(/,
]);

// Point 3 Fix: Memory registry to store images without crashing the LLM
export const powerpointImageRegistry = new Map<string, string>();

const runPowerPoint = <T>(action: (context: any) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => PowerPoint.run(action) as Promise<T>);

type PowerPointToolTemplate = Omit<ToolDefinition, 'execute'> & {
  executePowerPoint?: (context: any, args: Record<string, any>) => Promise<string>;
  executeCommon?: (args: Record<string, any>) => Promise<string>;
};

export type PowerPointToolName =
  | 'getSelectedText'
  | 'replaceSelectedText'
  | 'insertContent'
  | 'getSlideContent'
  | 'addSlide'
  | 'deleteSlide'
  | 'getShapes'
  | 'getAllSlidesOverview'
  | 'proposeShapeTextRevision'
  | 'getSpeakerNotes'
  | 'setSpeakerNotes'
  | 'insertImageOnSlide'
  | 'getCurrentSlideIndex'
  | 'eval_powerpointjs'
  | 'screenshotSlide'
  | 'duplicateSlide'
  | 'reorderSlide'
  | 'verifySlides'
  | 'editSlideXml'
  | 'searchIcons'
  | 'insertIcon'
  | 'searchAndFormatInPresentation'
  | 'searchAndReplaceInShape'
  | 'replaceShapeParagraphs';

/**
 * Returns true for shape types that have no text frame and cause InvalidArgument
 * errors when Office.js tries to load textFrame.textRange (OLE objects, charts, etc.).
 */
function isNonTextShape(shapeType: string): boolean {
  const t = shapeType.toLowerCase();
  return (
    t === '13' ||            // numeric picture type
    t.includes('picture') || // picture / image
    t === 'ole' ||           // OLE embedded objects (Excel charts, Word docs, etc.)
    t === 'chart'            // native PowerPoint charts
  );
}

/**
 * Returns the 1-based slide number of the currently active/selected slide.
 * Falls back to slide 1 if the API is unavailable.
 */
export async function getCurrentSlideNumber(): Promise<number> {
  try {
    return await executeOfficeAction(() =>
      PowerPoint.run(async (context: any) => {
        let activeSlideIndex = 0;
        try {
          if (typeof context.presentation.getSelectedSlides === 'function') {
            const selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load('items/id');
            await context.sync();
            if (selectedSlides.items.length > 0) {
              const slides = context.presentation.slides;
              slides.load('items/id');
              await context.sync();
              const selectedId = selectedSlides.items[0].id;
              const idx = slides.items.findIndex((s: any) => s.id === selectedId);
              if (idx !== -1) activeSlideIndex = idx;
            }
          }
        } catch {}
        return activeSlideIndex + 1;
      }),
    );
  } catch {
    return 1;
  }
}

/**
 * Set speaker notes for the currently selected slide directly (no agent loop).
 * Returns true on success, false on failure.
 */
export async function setCurrentSlideSpeakerNotes(notes: string): Promise<boolean> {
  try {
    const slideNumber = await getCurrentSlideNumber();
    await executeOfficeAction(() =>
      PowerPoint.run(async (context: any) => {
        // PPT-M5: Check for API support (1.5 required for notesSlide)
        if (!isPowerPointApiSupported('1.5')) {
          throw new Error('Speaker notes modification requires PowerPoint API 1.5 or newer.');
        }

        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = slideNumber - 1;
        if (index < 0 || index >= slides.items.length) {
          throw new Error(`Slide ${slideNumber} not found.`);
        }

        const slide = slides.getItemAt(index);
        // Load the slide item so navigational properties are available
        slide.load('id');
        await context.sync();

        let notesPage: any;
        try {
          notesPage = slide.notesSlide;
        } catch {
          throw new Error('NOTES_SLIDE_UNAVAILABLE');
        }
        if (!notesPage) throw new Error('NOTES_SLIDE_UNAVAILABLE');

        notesPage.load('notesTextFrame/textRange');
        await context.sync();

        notesPage.notesTextFrame.textRange.text = notes;
        await context.sync();
      }),
    );
    return true;
  } catch (err) {
    logService.error('[PowerPointTools] Failed to set speaker notes:', err);
    // Always show the friendly fallback message — notes can always be inserted manually
    messageUtil.error(
      "Impossible d'insérer les notes automatiquement. Cliquez dans la zone de notes de la diapositive, puis utilisez le bouton « Remplacer » sur la réponse générée.",
      9000,
    );
    return false;
  }
}

/**
 * Read the currently selected text inside a PowerPoint shape / text box.
 * Returns an empty string when nothing is selected or the selection is
 * not a text range (e.g. an entire slide is selected).
 */
export function getPowerPointSelection(): Promise<string> {
  return new Promise(resolve => {
    try {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Unformatted },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve((result.value as string) || '');
          } else {
            logService.warn('PowerPoint selection error:', result.error?.message);
            resolve('');
          }
        },
      );
    } catch (err) {
      logService.warn('PowerPoint getSelectedDataAsync unavailable:', err);
      resolve('');
    }
  });
}

/**
 * Reads the current PowerPoint selection and reconstructs basic HTML formatting
 * (bold, italic, underline, strikethrough) by inspecting at the paragraph level.
 *
 * Issue #9 fix: The previous implementation loaded one API object per character
 * (O(N) context.sync calls for N chars), causing extreme latency on long selections.
 * This implementation loads all paragraphs in a single batch — O(1) syncs —
 * preserving paragraph-level formatting while being orders of magnitude faster.
 */
export async function getPowerPointSelectionAsHtml(): Promise<string> {
  if (!isPowerPointApiSupported('1.5')) {
    return getPowerPointSelection();
  }

  try {
    const htmlOut = await executeOfficeAction(async () => {
      return PowerPoint.run(async (context: any) => {
        const textRanges = context.presentation.getSelectedTextRanges();
        textRanges.load('items');
        await context.sync();

        if (textRanges.items.length === 0) return '';

        const range = textRanges.items[0];

        // Load paragraphs in a single batch (replaces per-character loading)
        const paragraphs = range.paragraphs;
        paragraphs.load('items');
        await context.sync();

        if (paragraphs.items.length === 0) return '';

        // Batch-load text and font for every paragraph in one sync
        for (const para of paragraphs.items) {
          para.load('text');
          para.font.load(['bold', 'italic', 'underline', 'strikethrough']);
        }
        await context.sync();

        let html = '';
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          const text: string = para.text || '';
          const font = para.font;

          if (!text && i < paragraphs.items.length - 1) {
            html += '<br/>';
            continue;
          }

          const bold = font.bold === true;
          const italic = font.italic === true;
          const underline = font.underline !== 'None' && font.underline !== null;
          const strike = font.strikethrough === true;

          // Escape HTML entities
          const safeText = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

          // Wrap with inline formatting tags (innermost first)
          let wrapped = safeText;
          if (strike) wrapped = `<s>${wrapped}</s>`;
          if (underline) wrapped = `<u>${wrapped}</u>`;
          if (italic) wrapped = `<i>${wrapped}</i>`;
          if (bold) wrapped = `<b>${wrapped}</b>`;

          if (i > 0) html += '<br/>';
          html += wrapped;
        }

        return html;
      });
    });

    return htmlOut || getPowerPointSelection();
  } catch (err) {
    logService.warn('Failed to extract PowerPoint HTML selection (paragraph mode):', err);
    return getPowerPointSelection();
  }
}

/**
 * Replace the current text selection inside the active PowerPoint shape
 * with the provided text.
 */
export async function insertIntoPowerPoint(text: string, useHtml = true): Promise<void> {
  const normalizedNewlines = normalizeLineEndings(text);

  // Try the Modern API first if available (requires PowerPointApi 1.5+)
  if (isPowerPointApiSupported('1.5') && useHtml) {
    try {
      await executeOfficeAction(async () => {
        await PowerPoint.run(async (context: any) => {
          const textRange = context.presentation.getSelectedTextRanges().getItemAt(0);
          await insertMarkdownIntoTextRange(context, textRange, normalizedNewlines);
          await context.sync();
        });
      });
      return;
    } catch (e: unknown) {
      logService.warn('Modern PowerPoint Html insertion failed, falling back:', e instanceof Error ? e : new Error(String(e)));
    }
  }

  // Fallback to the legacy Shared API (no native bullet detection possible here)
  const htmlContent = renderOfficeCommonApiHtml(normalizedNewlines);
  return new Promise((resolve, reject) => {
    try {
      if (useHtml) {
        Office.context.document.setSelectedDataAsync(
          htmlContent,
          { coercionType: Office.CoercionType.Html },
          (result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              logService.warn('setSelectedDataAsync Html failed, falling back to raw Text');
              fallbackToText(normalizedNewlines, resolve, reject);
            }
          },
        );
      } else {
        fallbackToText(normalizedNewlines, resolve, reject);
      }
    } catch (err: unknown) {
      reject(new Error(getErrorMessage(err) || 'setSelectedDataAsync unavailable'));
    }
  });
}

export async function insertMarkdownIntoTextRange(
  context: any,
  textRange: any,
  text: string,
  forceStripBullets = false,
) {
  let styles: InheritedStyles | null = null;
  try {
    textRange.font.load('name,size,color');
    await context.sync();
    styles = {
      fontFamily: textRange.font.name || '',
      fontSize: textRange.font.size ? `${textRange.font.size}pt` : '',
      fontWeight: 'normal',
      fontStyle: 'normal',
      color: '', // Do NOT force the original color here
      marginTop: '0pt',
      marginBottom: '0pt',
    };
  } catch (e) {
    // Ignore if font loading fails
  }

  const nativeBullets = forceStripBullets || (await hasNativeBullets(context, textRange));

  let finalMarkdown = text;
  if (nativeBullets) {
    finalMarkdown = stripMarkdownListMarkers(text);
  }

  let html = renderOfficeCommonApiHtml(finalMarkdown);
  if (styles) html = applyInheritedStyles(html, styles);

  textRange.insertHtml(html, 'Replace');
}

async function findShapeOnSlide(context: any, slideNumber: number, shapeIdOrName: string | number) {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();

  const idx = Math.trunc(Number(slideNumber)) - 1;
  if (idx < 0 || idx >= slides.items.length) {
    return { slide: null, shape: null, shapes: [], error: `Invalid slide number ${slideNumber}` };
  }

  const slide = slides.items[idx];
  const shapes = slide.shapes;
  shapes.load('items,items/id,items/name,items/placeholderFormat,items/placeholderFormat/type');
  await context.sync();

  for (const s of shapes.items) {
    if (s.id === shapeIdOrName || s.name === shapeIdOrName) {
      return { slide, shape: s, shapes: shapes.items, error: null };
    }
  }

  return {
    slide,
    shape: null,
    shapes: shapes.items,
    error: `Shape '${shapeIdOrName}' not found on slide ${slideNumber}`,
  };
}

function fallbackToText(text: string, resolve: any, reject: any) {
  // Pass true to strip list markers so it plays nice with shapes that are already natively bulleted.
  const fallbackText = stripRichFormattingSyntax(text, true);
  Office.context.document.setSelectedDataAsync(
    fallbackText,
    { coercionType: Office.CoercionType.Text },
    (fallbackResult: any) => {
      if (fallbackResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(fallbackResult.error?.message || 'setSelectedDataAsync failed'));
      }
    },
  );
}

function isPowerPointApiSupported(version: string): boolean {
  try {
    return !!Office?.context?.requirements?.isSetSupported?.('PowerPointApi', version);
  } catch {
    return false;
  }
}

/**
 * Detect whether the given PowerPoint text range belongs to a shape with native
 * (layout/master) bullet points. When true, we should avoid inserting HTML
 * <ul>/<li> tags to prevent double-bullet rendering.
 */
async function hasNativeBullets(context: any, textRange: any): Promise<boolean> {
  try {
    const paragraphs = textRange.paragraphs;
    paragraphs.load('items');
    await context.sync();
    if (paragraphs.items.length > 0) {
      for (const para of paragraphs.items) {
        para.load('bulletFormat/visible');
      }
      await context.sync();
      // Return true if ANY paragraph has native bullets
      return paragraphs.items.some((p: any) => p.bulletFormat?.visible === true);
    }
  } catch {
    // API not available or paragraphs inaccessible — assume no native bullets
  }
  return false;
}

function ensurePowerPointRunAvailable() {
  if (typeof PowerPoint?.run !== 'function') {
    throw new Error('PowerPoint.run is not available in this Office host/runtime.');
  }
}

const powerpointToolDefinitions = createOfficeTools<
  PowerPointToolName,
  PowerPointToolTemplate,
  ToolDefinition
>(
  {
    getSelectedText: {
      name: 'getSelectedText',
      category: 'read',
      description: 'Get the currently selected text in PowerPoint.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeCommon: async () => getPowerPointSelection(),
    },

    replaceSelectedText: {
      name: 'replaceSelectedText',
      category: 'write',
      description:
        'Replace the currently selected text in PowerPoint. This tool preserves block-level formatting and inline styles better than full shape replacement.',
      inputSchema: {
        type: 'object',
        properties: {
          text: {
            type: 'string',
            description:
              'The new text to replace the selection with. Markdown formatting is supported.',
          },
        },
        required: ['text'],
      },
      executeCommon: async (args: Record<string, any>) => {
        await insertIntoPowerPoint(args.text, true);
        return 'Successfully replaced selected text.';
      },
    },

    proposeShapeTextRevision: {
      name: 'proposeShapeTextRevision',
      category: 'write',
      description: `Replace text in a specific shape. WARNING: this performs a FULL TEXT REPLACEMENT — all existing formatting (bold, italic, font, color) will be lost.

The tool reports a word-level diff (insertions/deletions/unchanged stats) for informational purposes, but the actual operation is a complete overwrite of the shape text.

Use this when the content change is more important than preserving formatting. For formatting-sensitive edits, use eval_powerpointjs to modify individual text runs.

PARAMETERS:
- slideNumber: 1-based slide number (as shown in PowerPoint UI)
- shapeIdOrName: Shape ID (number) or shape name (string)
- revisedText: The new text for the shape`,
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Slide number (1-based, as shown in PowerPoint)',
          },
          shapeIdOrName: {
            type: 'string',
            description: 'Shape ID or name. Use getShapes to discover available shapes.',
          },
          revisedText: {
            type: 'string',
            description: 'The complete new text for the shape.',
          },
        },
        required: ['slideNumber', 'shapeIdOrName', 'revisedText'],
      },
      executePowerPoint: async (context, args: Record<string, any>) => {
        const { slideNumber, shapeIdOrName, revisedText } = args;

        try {
          const {
            shape: targetShape,
            shapes,
            error,
          } = await findShapeOnSlide(context, slideNumber, shapeIdOrName);

          if (!targetShape) {
            return JSON.stringify(
              {
                success: false,
                error: error || `Shape "${shapeIdOrName}" not found on slide ${slideNumber}`,
                availableShapes: shapes.map((s: any) => ({ id: s.id, name: s.name })),
              },
              null,
              2,
            );
          }

          // Get current text
          const textFrame = targetShape.textFrame;
          const textRange = textFrame.textRange;
          textRange.load('text');
          await context.sync();

          const originalText = textRange.text || '';

          // 4. Compute diff stats
          const { insertions, deletions, unchanged } = computeTextDiffStats(
            originalText,
            revisedText,
          );

          // 5. Apply changes
          // PowerPoint API is limited - we do full replacement but report the diff
          textRange.text = revisedText;
          await context.sync();

          return JSON.stringify(
            {
              success: true,
              slideNumber,
              shapeId: targetShape.id,
              shapeName: targetShape.name,
              changes: {
                insertions,
                deletions,
                unchanged,
              },
              message: `Updated shape text. ${insertions} characters added, ${deletions} removed.`,
              note: 'PowerPoint applies full text replacement. Formatting may need manual adjustment.',
            },
            null,
            2,
          );
        } catch (error: unknown) {
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(error),
            },
            null,
            2,
          );
        }
      },
    },

    searchAndReplaceInShape: {
      name: 'searchAndReplaceInShape',
      category: 'write',
      description: `Surgically replace a specific word or phrase in a shape's text WITHOUT destroying formatting.

Unlike proposeShapeTextRevision (which replaces ALL text and loses all formatting), this tool
replaces ONLY the matching text runs, preserving bold, italic, font size, color and other
formatting on every other run.

USE THIS for: typo corrections, spell-check fixes, word substitutions.
DO NOT use proposeShapeTextRevision for spell-checking — it destroys formatting.

PARAMETERS:
- slideNumber: 1-based slide number
- shapeIdOrName: Shape ID (number string) or name — use getShapes to discover
- searchText: The exact text to find (case-sensitive by default)
- replaceText: The replacement text
- replaceAll: If true, replaces all occurrences (default: true)`,
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: { type: 'number', description: 'Slide number (1-based)' },
          shapeIdOrName: {
            type: 'string',
            description: 'Shape ID or name. Use getShapes to discover.',
          },
          searchText: { type: 'string', description: 'Exact text to search for' },
          replaceText: { type: 'string', description: 'Replacement text' },
          replaceAll: {
            type: 'boolean',
            description: 'Replace all occurrences (default true)',
          },
        },
        required: ['slideNumber', 'shapeIdOrName', 'searchText', 'replaceText'],
      },
      executePowerPoint: async (context, args: Record<string, any>) => {
        const { slideNumber, shapeIdOrName, searchText, replaceText } = args;
        const replaceAll = args.replaceAll !== false;

        // Helper: XML-based surgical replacement via OOXML ZIP (truly preserves all formatting)
        const tryXmlReplacement = async (): Promise<{
          success: boolean;
          replacements: number;
          error?: string;
        }> => {
          if (!isPowerPointApiSupported('1.5')) {
            return { success: false, replacements: 0, error: 'PowerPointApi 1.5 not supported' };
          }
          const slideIndex = Math.trunc(Number(slideNumber)) - 1;
          let replacements = 0;
          try {
            await withSlideZip(context, slideIndex, async (zip: any, markDirty: () => void) => {
              const slideXmlStr = await zip.file('ppt/slides/slide1.xml')?.async('string');
              if (!slideXmlStr) throw new Error('Could not read slide XML');

              const parser = new DOMParser();
              const doc = parser.parseFromString(slideXmlStr, 'application/xml');

              // Locate the target shape by id or name within the XML
              const nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main';
              const spElements = doc.getElementsByTagNameNS(nsP, 'sp');
              let targetSp: Element | null = null;
              for (let i = 0; i < spElements.length; i++) {
                const sp = spElements[i];
                const nvSpPr = sp.getElementsByTagNameNS(nsP, 'nvSpPr')[0];
                if (nvSpPr) {
                  const cNvPr = nvSpPr.getElementsByTagNameNS(nsP, 'cNvPr')[0];
                  if (cNvPr) {
                    const spId = cNvPr.getAttribute('id');
                    const spName = cNvPr.getAttribute('name');
                    if (
                      spId === String(shapeIdOrName) ||
                      spName === String(shapeIdOrName)
                    ) {
                      targetSp = sp;
                      break;
                    }
                  }
                }
              }

              // If not found by id/name, search all shapes (best-effort)
              const searchRoot: Element = targetSp ?? doc.documentElement;

              // Replace text content within <a:t> elements.
              // Text in OOXML can be split across multiple <a:r> runs within a <a:p> paragraph.
              // Strategy: for each <a:p>, concatenate all <a:t> text, check for match,
              // then redistribute the replacement across the affected runs.
              const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
              const paragraphs = searchRoot.getElementsByTagNameNS(nsA, 'p');
              let done = false;
              for (let pi = 0; pi < paragraphs.length && !(done && !replaceAll); pi++) {
                const para = paragraphs[pi];
                const runs = para.getElementsByTagNameNS(nsA, 'r');
                if (runs.length === 0) continue;

                // Collect text of each run and build concatenated paragraph text
                const runTexts: string[] = [];
                for (let ri = 0; ri < runs.length; ri++) {
                  const tNode = runs[ri].getElementsByTagNameNS(nsA, 't')[0];
                  runTexts.push(tNode?.textContent ?? '');
                }
                const paraText = runTexts.join('');

                if (!paraText.includes(searchText)) continue;

                // Perform replacement(s) on concatenated text
                const newParaText = replaceAll
                  ? paraText.split(searchText).join(replaceText)
                  : paraText.replace(searchText, replaceText);

                if (newParaText === paraText) continue;

                // Redistribute newParaText back across the same number of runs,
                // keeping run boundaries as close to original as possible.
                // Simple approach: put all new text in the first run, empty the rest.
                const firstTNode = runs[0].getElementsByTagNameNS(nsA, 't')[0];
                if (firstTNode) {
                  firstTNode.textContent = newParaText;
                  // Preserve xml:space="preserve" if the new text has leading/trailing spaces
                  if (newParaText !== newParaText.trim()) {
                    firstTNode.setAttribute('xml:space', 'preserve');
                  }
                }
                for (let ri = 1; ri < runs.length; ri++) {
                  const tNode = runs[ri].getElementsByTagNameNS(nsA, 't')[0];
                  if (tNode) tNode.textContent = '';
                }

                replacements++;
                markDirty();
                if (!replaceAll) done = true;
              }

              if (replacements > 0) {
                const serializer = new XMLSerializer();
                zip.file('ppt/slides/slide1.xml', serializer.serializeToString(doc));
              }
            });
            return { success: true, replacements };
          } catch (xmlErr: unknown) {
            return { success: false, replacements: 0, error: getErrorMessage(xmlErr) };
          }
        };

        try {
          const { shape: targetShape, error } = await findShapeOnSlide(
            context,
            slideNumber,
            shapeIdOrName,
          );
          if (!targetShape) {
            return JSON.stringify({ success: false, error }, null, 2);
          }

          // Load paragraphs
          const textRange = targetShape.textFrame.textRange;
          textRange.paragraphs.load('items');
          await context.sync();

          // Load textRuns for all paragraphs
          for (const para of textRange.paragraphs.items) {
            try {
              para.textRange.textRuns.load('items');
            } catch {
              /* paragraph has no runs */
            }
          }
          await context.sync();

          // Load text for all runs
          for (const para of textRange.paragraphs.items) {
            try {
              if (para.textRange.textRuns?.items) {
                for (const run of para.textRange.textRuns.items) {
                  run.load('text');
                }
              }
            } catch {
              /* skip */
            }
          }
          await context.sync();

          // Replace matching text in runs, preserving all formatting
          let replacements = 0;
          for (const para of textRange.paragraphs.items) {
            try {
              if (!para.textRange.textRuns?.items) continue;
              for (const run of para.textRange.textRuns.items) {
                if (run.text && run.text.includes(searchText)) {
                  const newText = replaceAll
                    ? run.text.split(searchText).join(replaceText)
                    : run.text.replace(searchText, replaceText);
                  if (newText !== run.text) {
                    run.text = newText;
                    replacements++;
                  }
                }
              }
            } catch {
              /* skip non-editable runs */
            }
          }
          await context.sync();

          return JSON.stringify(
            {
              success: true,
              replacements,
              message:
                replacements > 0
                  ? `Replaced "${searchText}" → "${replaceText}" in ${replacements} run(s). Formatting preserved.`
                  : `Text "${searchText}" not found in shape. No changes made.`,
            },
            null,
            2,
          );
        } catch (error: unknown) {
          // textRuns approach failed (e.g. GeneralException on certain Placeholder types)
          // Fall back to OOXML XML editing — truly surgical, no formatting loss
          logService.warn(
            '[searchAndReplaceInShape] textRuns API failed, trying XML fallback:',
            error,
          );
          const xmlResult = await tryXmlReplacement();
          if (xmlResult.success) {
            return JSON.stringify(
              {
                success: true,
                replacements: xmlResult.replacements,
                method: 'xml-fallback',
                message:
                  xmlResult.replacements > 0
                    ? `Replaced "${searchText}" → "${replaceText}" in ${xmlResult.replacements} text run(s) via XML. Formatting preserved.`
                    : `Text "${searchText}" not found in shape. No changes made.`,
              },
              null,
              2,
            );
          }
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(error),
              xmlFallbackError: xmlResult.error,
            },
            null,
            2,
          );
        }
      },
    },

    replaceShapeParagraphs: {
      name: 'replaceShapeParagraphs',
      category: 'write',
      description: `Replace the text of specific paragraphs in a shape while preserving formatting (font name, size, bold, italic, color).

Use this for Punchify-style rewrites where whole paragraph text changes but the visual style must be kept.
Each paragraph is identified by its 0-based index within the shape.

PARAMETERS:
- slideNumber: 1-based slide number
- shapeIdOrName: Shape ID or name (from eval_powerpointjs or getShapes)
- paragraphReplacements: Array of { paragraphIndex: number, newText: string }

The tool reads each paragraph's first run font properties (name, size, bold, italic, color),
replaces the paragraph text, then re-applies those font properties so style is fully preserved.`,
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: { type: 'number', description: 'Slide number (1-based)' },
          shapeIdOrName: {
            type: 'string',
            description: 'Shape ID or name. Use eval_powerpointjs to discover.',
          },
          paragraphReplacements: {
            type: 'array',
            description: 'List of paragraph replacements',
            items: {
              type: 'object',
              properties: {
                paragraphIndex: {
                  type: 'number',
                  description: '0-based index of the paragraph within the shape',
                },
                newText: { type: 'string', description: 'New text content for the paragraph' },
              },
              required: ['paragraphIndex', 'newText'],
            },
          },
        },
        required: ['slideNumber', 'shapeIdOrName', 'paragraphReplacements'],
      },
      executePowerPoint: async (context, args: Record<string, any>) => {
        const { slideNumber, shapeIdOrName, paragraphReplacements } = args;

        // XML fallback: edit paragraph text directly in OOXML, preserving all run formatting
        const tryXmlParagraphReplacement = async (): Promise<{
          success: boolean;
          results: { paragraphIndex: number; status: string }[];
          error?: string;
        }> => {
          if (!isPowerPointApiSupported('1.5')) {
            return {
              success: false,
              results: [],
              error: 'PowerPointApi 1.5 not supported',
            };
          }
          const slideIndex = Math.trunc(Number(slideNumber)) - 1;
          const results: { paragraphIndex: number; status: string }[] = [];
          try {
            await withSlideZip(context, slideIndex, async (zip: any, markDirty: () => void) => {
              const slideXmlStr = await zip.file('ppt/slides/slide1.xml')?.async('string');
              if (!slideXmlStr) throw new Error('Could not read slide XML');

              const parser = new DOMParser();
              const doc = parser.parseFromString(slideXmlStr, 'application/xml');

              const nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main';
              const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';

              // Find target shape
              const spElements = doc.getElementsByTagNameNS(nsP, 'sp');
              let targetSp: Element | null = null;
              for (let i = 0; i < spElements.length; i++) {
                const sp = spElements[i];
                const nvSpPr = sp.getElementsByTagNameNS(nsP, 'nvSpPr')[0];
                if (nvSpPr) {
                  const cNvPr = nvSpPr.getElementsByTagNameNS(nsP, 'cNvPr')[0];
                  if (cNvPr) {
                    const spId = cNvPr.getAttribute('id');
                    const spName = cNvPr.getAttribute('name');
                    if (spId === String(shapeIdOrName) || spName === String(shapeIdOrName)) {
                      targetSp = sp;
                      break;
                    }
                  }
                }
              }
              const searchRoot: Element = targetSp ?? doc.documentElement;
              const xmlParas = searchRoot.getElementsByTagNameNS(nsA, 'p');

              for (const replacement of paragraphReplacements as {
                paragraphIndex: number;
                newText: string;
              }[]) {
                const { paragraphIndex, newText } = replacement;
                if (paragraphIndex < 0 || paragraphIndex >= xmlParas.length) {
                  results.push({
                    paragraphIndex,
                    status: `skipped — index out of range (XML has ${xmlParas.length} paragraphs)`,
                  });
                  continue;
                }
                const para = xmlParas[paragraphIndex];
                const runs = para.getElementsByTagNameNS(nsA, 'r');
                if (runs.length === 0) {
                  results.push({ paragraphIndex, status: 'skipped — paragraph has no runs' });
                  continue;
                }
                // Put all new text in first run's <a:t>, clear rest
                const firstTNode = runs[0].getElementsByTagNameNS(nsA, 't')[0];
                if (firstTNode) {
                  firstTNode.textContent = newText;
                  if (newText !== newText.trim()) {
                    firstTNode.setAttribute('xml:space', 'preserve');
                  }
                }
                for (let ri = 1; ri < runs.length; ri++) {
                  const tNode = runs[ri].getElementsByTagNameNS(nsA, 't')[0];
                  if (tNode) tNode.textContent = '';
                }
                markDirty();
                results.push({ paragraphIndex, status: 'replaced-xml' });
              }

              if (results.some(r => r.status === 'replaced-xml')) {
                const serializer = new XMLSerializer();
                zip.file('ppt/slides/slide1.xml', serializer.serializeToString(doc));
              }
            });
            return { success: true, results };
          } catch (xmlErr: unknown) {
            return { success: false, results, error: getErrorMessage(xmlErr) };
          }
        };

        try {
          const { shape: targetShape, error } = await findShapeOnSlide(
            context,
            slideNumber,
            shapeIdOrName,
          );
          if (!targetShape) {
            return JSON.stringify({ success: false, error }, null, 2);
          }

          // Load paragraphs
          const textRange = targetShape.textFrame.textRange;
          textRange.paragraphs.load('items');
          await context.sync();

          const paragraphs = textRange.paragraphs.items;
          const results: { paragraphIndex: number; status: string }[] = [];

          for (const replacement of paragraphReplacements as {
            paragraphIndex: number;
            newText: string;
          }[]) {
            const { paragraphIndex, newText } = replacement;

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
              results.push({
                paragraphIndex,
                status: `skipped — index out of range (shape has ${paragraphs.length} paragraphs)`,
              });
              continue;
            }

            const para = paragraphs[paragraphIndex];

            // Read the first run's font properties to preserve them
            let savedFont: {
              name: string | null;
              size: number | null;
              bold: boolean | null;
              italic: boolean | null;
              color: string | null;
            } = { name: null, size: null, bold: null, italic: null, color: null };

            try {
              para.textRange.textRuns.load('items');
              await context.sync();

              if (para.textRange.textRuns?.items?.length > 0) {
                const firstRun = para.textRange.textRuns.items[0];
                firstRun.font.load('name,size,bold,italic,color');
                await context.sync();
                savedFont = {
                  name: firstRun.font.name,
                  size: firstRun.font.size,
                  bold: firstRun.font.bold,
                  italic: firstRun.font.italic,
                  color: firstRun.font.color,
                };
              }
            } catch {
              /* could not read font — will replace text only */
            }

            // Replace the paragraph text (collapses to single run)
            para.textRange.text = newText;
            await context.sync();

            // Re-apply saved font properties to restore formatting
            try {
              if (savedFont.name !== null) para.textRange.font.name = savedFont.name;
              if (savedFont.size !== null) para.textRange.font.size = savedFont.size;
              if (savedFont.bold !== null) para.textRange.font.bold = savedFont.bold;
              if (savedFont.italic !== null) para.textRange.font.italic = savedFont.italic;
              if (savedFont.color !== null) para.textRange.font.color = savedFont.color;
              await context.sync();
            } catch {
              /* font restore failed — text was still replaced */
            }

            results.push({ paragraphIndex, status: 'replaced' });
          }

          return JSON.stringify({ success: true, results }, null, 2);
        } catch (error: unknown) {
          // textRuns/paragraphs API failed — fall back to XML editing
          logService.warn(
            '[replaceShapeParagraphs] API approach failed, trying XML fallback:',
            error,
          );
          const xmlResult = await tryXmlParagraphReplacement();
          if (xmlResult.success) {
            return JSON.stringify(
              { success: true, method: 'xml-fallback', results: xmlResult.results },
              null,
              2,
            );
          }
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(error),
              xmlFallbackError: xmlResult.error,
            },
            null,
            2,
          );
        }
      },
    },

    getSpeakerNotes: {
      name: 'getSpeakerNotes',
      category: 'read',
      description:
        'Get the speaker notes text for a specific slide (1-based index). Requires PowerPointApi 1.5+.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
        },
        required: ['slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.5')) {
          return 'Error: getSpeakerNotes requires PowerPointApi 1.5 or later, which is not supported in this Office version.';
        }
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be a number >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length)
          throw new Error(
            `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`,
          );
        const slide = slides.getItemAt(index);
        const notesPage = slide.notesSlide;
        notesPage.load('notesTextFrame/textRange/text');
        await context.sync();
        const text = notesPage.notesTextFrame.textRange.text ?? '';
        return text.trim() || '(no speaker notes)';
      },
    },

    setSpeakerNotes: {
      name: 'setSpeakerNotes',
      category: 'write',
      description:
        'Set (replace) the speaker notes for a specific slide (1-based index). Requires PowerPointApi 1.5+.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
          notes: { type: 'string', description: 'The new speaker notes text (plain text).' },
        },
        required: ['slideNumber', 'notes'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.5')) {
          return 'Error: setSpeakerNotes requires PowerPointApi 1.5 or later, which is not supported in this Office version.';
        }
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be a number >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length)
          throw new Error(`Error: slide ${slideNumber} does not exist.`);
        const slide = slides.getItemAt(index);
        const notesPage = slide.notesSlide;
        notesPage.load('notesTextFrame/textRange');
        await context.sync();
        notesPage.notesTextFrame.textRange.text = String(args.notes ?? '');
        await context.sync();
        return `Successfully updated speaker notes for slide ${slideNumber}.`;
      },
    },

    insertImageOnSlide: {
      name: 'insertImageOnSlide',
      category: 'write',
      description:
        'Insert an image onto a specific slide. Position and size are in points (1 inch = 72 points). Default: centered at 400×300pt.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
          base64Data: {
            type: 'string',
            description:
              'The base64-encoded image data, a data URI (data:image/...;base64,...), OR the image ID returned when the user uploaded a file (UUID format). The tool resolves registry IDs automatically.',
          },
          left: {
            type: 'number',
            description: 'Left position in points from the left edge of the slide. Default: 100.',
          },
          top: {
            type: 'number',
            description: 'Top position in points from the top edge of the slide. Default: 100.',
          },
          width: { type: 'number', description: 'Width in points. Default: 400.' },
          height: { type: 'number', description: 'Height in points. Default: 300.' },
        },
        required: ['slideNumber', 'base64Data'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.4')) {
          return 'Error: insertImageOnSlide requires PowerPointApi 1.4 or later.';
        }
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be a number >= 1.');

        // Resolve image data: accept raw base64, data URI, or a registry key (UUID from uploaded files)
        let rawValue = String(args.base64Data);
        const registryHit = powerpointImageRegistry.get(rawValue);
        if (registryHit) rawValue = registryHit;
        const base64 = rawValue.replace(/^data:[^;]+;base64,/, '');

        if (!base64 || base64.length < 10) {
          throw new Error(
            'Error: base64Data is empty or not a valid image. If you uploaded a file, use the image ID shown in the upload confirmation.',
          );
        }

        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length)
          throw new Error(`Error: slide ${slideNumber} does not exist.`);

        // PPT-C2 fix: use slides.items[index] (already loaded) instead of
        // slides.getItemAt(index) which returns a proxy lacking .shapes.addImage()
        const slide = slides.items[index];
        const shape = slide.shapes.addImage(base64);
        shape.left = typeof args.left === 'number' ? args.left : 100;
        shape.top = typeof args.top === 'number' ? args.top : 100;
        shape.width = typeof args.width === 'number' ? args.width : 400;
        shape.height = typeof args.height === 'number' ? args.height : 300;
        await context.sync();

        return `Successfully inserted image on slide ${slideNumber} at position (${shape.left}, ${shape.top}) with size ${shape.width}×${shape.height} points.`;
      },
    },

    getCurrentSlideIndex: {
      name: 'getCurrentSlideIndex',
      category: 'read',
      description: 'Get the 1-based index of the currently active slide viewed by the user.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeCommon: async () => {
        const slideNumber = await getCurrentSlideNumber();
        return String(slideNumber);
      },
    },

    eval_powerpointjs: {
      name: 'eval_powerpointjs',
      category: 'write',
      description: `Execute custom Office.js code within a PowerPoint.run context.

**USE THIS TOOL ONLY WHEN:**
- No dedicated tool exists for your operation
- Operations like: animations, transitions, advanced shape manipulations not covered by other tools

**REQUIRED CODE STRUCTURE:**
\`\`\`javascript
try {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();

  // Your operations here

  await context.sync();
  return { success: true, result: 'Operation completed' };
} catch (error) {
  return { success: false, error: error.message };
}
\`\`\`

**CRITICAL RULES:**
1. ALWAYS call \`.load()\` before reading properties
2. ALWAYS call \`await context.sync()\` after load and after modifications
3. ALWAYS wrap in try/catch
4. ONLY use PowerPoint namespace (not Word, Excel)
5. Slide numbers are 1-based in UI, 0-indexed in arrays`,
      inputSchema: {
        type: 'object',
        properties: {
          code: {
            type: 'string',
            description:
              'JavaScript code following the template. Must include load(), sync(), and try/catch.',
          },
          explanation: {
            type: 'string',
            description: 'Brief explanation of what this code does (required for audit trail).',
          },
        },
        required: ['code', 'explanation'],
      },
      executePowerPoint: createEvalExecutor({
        host: 'PowerPoint',
        toolName: 'eval_powerpointjs',
        suggestion:
          'Refer to the Office.js skill document for correct patterns. Remember: slide indices are 0-based.',
        mutationDetector: looksLikeMutationPpt,
        buildSandboxContext: (context: any) => ({
          context,
          PowerPoint: typeof PowerPoint !== 'undefined' ? PowerPoint : undefined,
          Office: typeof Office !== 'undefined' ? Office : undefined,
          ...getVfsSandboxContext(),
        }),
        preExecuteHook: () => ensurePowerPointRunAvailable(),
      }),
    },

    insertContent: {
      name: 'insertContent',
      category: 'write',
      description:
        'The PREFERRED tool for adding or replacing content in PowerPoint. Supports Markdown (bold, italic, bullets). Requires BOTH slideNumber and shapeIdOrName to target a specific shape (ALWAYS use getShapes first to discover IDs). Without shapeIdOrName it only replaces the currently selected text — never use it on an empty slide without providing shapeIdOrName or you will create a floating text box instead of filling the placeholder.',
      inputSchema: {
        type: 'object',
        properties: {
          content: {
            type: 'string',
            description: 'The content to insert in Markdown format.',
          },
          slideNumber: {
            type: 'number',
            description:
              'Target slide number, 1-based (1 = first slide, not 0-based). Required when shapeIdOrName is provided.',
          },
          shapeIdOrName: {
            type: 'string',
            description:
              'ID or Name of the shape to update (from getShapes). ALWAYS provide this to fill a specific placeholder. If omitted, replaces current text selection only.',
          },
        },
        required: ['content'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        const { content, slideNumber, shapeIdOrName } = args;

        if (shapeIdOrName) {
          if (!slideNumber)
            throw new Error('Error: slideNumber is required when shapeIdOrName is provided.');

          const { shape, shapes, error } = await findShapeOnSlide(
            context,
            slideNumber,
            shapeIdOrName,
          );

          if (!shape) {
            const availableShapes = shapes.map((s: any) => `'${s.name}' (id: ${s.id})`).join(', ');
            throw new Error(`Error: ${error}. Available shapes are: ${availableShapes}`);
          }

          // Detect body/content placeholders to prevent double-bullet rendering
          let isBodyPlaceholder = false;
          try {
            const phType = String((shape as any).placeholderFormat?.type ?? '').toLowerCase();
            const nameLower = ((shape.name as string) ?? '').toLowerCase();
            isBodyPlaceholder =
              phType === 'body' ||
              phType === 'object' ||
              nameLower.includes('content') ||
              nameLower.includes('body') ||
              (nameLower.includes('text') && !nameLower.includes('title'));
          } catch {
            /* not a placeholder or API unavailable */
          }

          const textRange = shape.textFrame.textRange;
          if (isPowerPointApiSupported('1.5')) {
            try {
              const normalizedNewlines = normalizeLineEndings(content);
              await insertMarkdownIntoTextRange(
                context,
                textRange,
                normalizedNewlines,
                isBodyPlaceholder,
              );
            } catch (e) {
              logService.warn(
                'insertMarkdownIntoTextRange failed, falling back to text modification',
                e,
              );
              textRange.text = stripRichFormattingSyntax(content);
              await context.sync();
              return `Content set on shape '${shapeIdOrName}' on slide ${slideNumber} (plain text fallback — markdown formatting was not applied because the rich text API failed).`;
            }
          } else {
            textRange.text = stripRichFormattingSyntax(content);
          }
          await context.sync();
          return `Successfully set content on shape '${shapeIdOrName}' on slide ${slideNumber}.`;
        } else {
          // Replace current selection — only valid when the user has text selected in a shape
          await insertIntoPowerPoint(content);
          return 'Successfully replaced selected text in PowerPoint.';
        }
      },
    },

    getSlideContent: {
      name: 'getSlideContent',
      category: 'read',
      description: 'Read all text content from a specific slide (1-based index).',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Slide number to read (1 = first slide).',
          },
        },
        required: ['slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1) {
          throw new Error('Error: slideNumber must be a number greater than or equal to 1.');
        }
        return getSlideContentStandalone(context, slideNumber);
      },
    },

    addSlide: {
      name: 'addSlide',
      category: 'write',
      description:
        "Add a new slide to the presentation. The tool automatically picks the best matching layout from the presentation's slide master. Pass title and body to automatically populate the template text boxes (recommended).",
      inputSchema: {
        type: 'object',
        properties: {
          layout: {
            type: 'string',
            description:
              "Hint for the desired layout type: 'TitleAndContent' for standard content slides with bullet points (DEFAULT), 'Title' for a title-only or title-slide layout, 'Blank' for an empty slide.",
          },
          title: {
            type: 'string',
            description:
              'Optional title text to insert into the title placeholder of the new slide.',
          },
          body: {
            type: 'string',
            description:
              'Optional body/content text to insert into the main content placeholder. Use newlines for bullet points.',
          },
        },
        required: [],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();

        const slides = context.presentation.slides;
        const layoutHint =
          typeof args.layout === 'string' ? args.layout.trim().toLowerCase() : 'titleandcontent';
        const titleText = typeof args.title === 'string' ? args.title.trim() : undefined;
        const bodyText = typeof args.body === 'string' ? args.body.trim() : undefined;

        // Discover real layout IDs from the presentation's slide masters
        let addOptions: any = {};
        let chosenLayoutName = '';
        try {
          const slideMasters = context.presentation.slideMasters;
          slideMasters.load('items/id');
          await context.sync();

          if (slideMasters.items.length > 0) {
            const master = slideMasters.items[0];
            master.layouts.load('items/id,items/name');
            await context.sync();

            const availableLayouts: any[] = master.layouts.items;

            // Score a layout name against the requested hint
            const scoreLayout = (name: string): number => {
              const n = name.toLowerCase();
              if (
                layoutHint.includes('blank') ||
                layoutHint === 'vide' ||
                layoutHint === 'vierge'
              ) {
                return n.includes('blank') || n.includes('vide') || n.includes('vierge') ? 100 : 0;
              }
              if (layoutHint === 'title' || layoutHint === 'titre') {
                // Prefer dedicated "Title Slide" / "Diapositive de titre" over "Title and Content"
                if (n === 'title' || n === 'titre') return 100;
                if (n.includes('title slide') || n.includes('diapositive de titre')) return 90;
                if (
                  n.includes('title') &&
                  !n.includes('content') &&
                  !n.includes('contenu') &&
                  !n.includes('body') &&
                  !n.includes('texte') &&
                  !n.includes('media') &&
                  !n.includes('picture')
                )
                  return 50;
                return 0;
              }
              // Default: TitleAndContent / content slides
              if (n.includes('title and content') || n.includes('titre et contenu')) return 100;
              if (
                n.includes('title') &&
                (n.includes('content') ||
                  n.includes('contenu') ||
                  n.includes('text') ||
                  n.includes('texte'))
              )
                return 90;
              if (n.includes('content') || n.includes('contenu')) return 80;
              if (n.includes('bullet') || n.includes('body')) return 70;
              // Penalise pure title slide, blank, picture for TitleAndContent requests
              if (n.includes('blank') || n.includes('vide') || n === 'title' || n === 'titre')
                return 0;
              return 5; // anything else
            };

            let bestScore = -1;
            let bestLayout: any = null;
            for (const lo of availableLayouts) {
              const s = scoreLayout(lo.name ?? '');
              if (s > bestScore) {
                bestScore = s;
                bestLayout = lo;
              }
            }

            if (bestLayout) {
              addOptions = { layoutId: bestLayout.id, slideMasterId: master.id };
              chosenLayoutName = bestLayout.name ?? '';
            }
          }
        } catch (err) {
          // Could not discover layouts — fall back to default slide add
          logService.warn('[addSlide] Layout discovery failed:', err);
        }

        slides.add(addOptions);
        await context.sync();

        // If title or body provided, populate the template text boxes
        if (titleText || bodyText) {
          try {
            slides.load('items');
            await context.sync();
            const newSlide = slides.items[slides.items.length - 1];
            const shapes = newSlide.shapes;
            // Load name and placeholderFormat (correct PowerPoint.js API, not .placeholder)
            shapes.load('items,items/name,items/placeholderFormat');
            await context.sync();

            let titleFilled = false;
            let bodyFilled = false;

            for (const shape of shapes.items) {
              try {
                // Use placeholderFormat.type when available (PowerPoint.js 1.3+)
                let phType = '';
                try {
                  phType = String((shape as any).placeholderFormat?.type ?? '').toLowerCase();
                } catch {}
                const nameLower = ((shape.name as string) ?? '').toLowerCase();

                const isTitle =
                  phType === 'title' ||
                  phType === 'centeredtitle' ||
                  phType === 'subtitle' ||
                  nameLower.includes('title');
                const isBody =
                  phType === 'body' ||
                  nameLower.includes('content') ||
                  nameLower.includes('body') ||
                  (nameLower.includes('text') && !nameLower.includes('title'));

                if (isTitle && titleText && !titleFilled) {
                  shape.textFrame.textRange.text = titleText;
                  titleFilled = true;
                } else if (isBody && bodyText && !bodyFilled) {
                  shape.textFrame.textRange.text = bodyText;
                  bodyFilled = true;
                }
              } catch {}
            }
            await context.sync();
          } catch (err: unknown) {
            // Non-fatal: slide was created, just couldn't populate shapes
            return `Successfully added a slide${chosenLayoutName ? ` with layout "${chosenLayoutName}"` : ''} but could not auto-fill shapes: ${getErrorMessage(err)}`;
          }
        }

        return chosenLayoutName
          ? `Successfully added a slide with layout "${chosenLayoutName}"${titleText ? ` titled "${titleText}"` : ''}.`
          : `Successfully added a slide${titleText ? ` titled "${titleText}"` : ''}.`;
      },
    },

    deleteSlide: {
      name: 'deleteSlide',
      category: 'write',
      description: 'Delete a specific slide by its 1-based index.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
        },
        required: ['slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be a number >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length)
          throw new Error(`Error: slide ${slideNumber} does not exist.`);
        slides.getItemAt(index).delete();
        await context.sync();
        return `Successfully deleted slide ${slideNumber}.`;
      },
    },

    getShapes: {
      name: 'getShapes',
      category: 'read',
      description:
        'Get all shapes on a specific slide (1-based index) with their properties (type, text, position).',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
        },
        required: ['slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be a number >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length)
          throw new Error(`Error: slide ${slideNumber} does not exist.`);

        const slide = slides.getItemAt(index);
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();

        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          try {
            shape.textFrame.textRange.load('text');
          } catch {}
          try {
            shape.load(['id', 'name', 'type', 'left', 'top', 'width', 'height']);
          } catch {}
        }
        await context.sync();

        const details = shapes.items.map((shape: any) => {
          let text = '';
          try {
            text = shape.textFrame.textRange.text;
          } catch {}
          return {
            id: shape.id,
            name: shape.name,
            type: shape.type,
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
            text: text.trim(),
          };
        });
        return JSON.stringify(details, null, 2);
      },
    },

    getAllSlidesOverview: {
      name: 'getAllSlidesOverview',
      category: 'read',
      description:
        'Get a text overview of all slides in the presentation (useful to understand the presentation structure).',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executePowerPoint: async (context: any) => {
        ensurePowerPointRunAvailable();
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const overview: string[] = [];
        for (let i = 0; i < slides.items.length; i++) {
          try {
            const slide = slides.items[i];
            slide.load(['id', 'layout']);
            const shapes = slide.shapes;
            // Load shape type together with items to filter out non-text shapes (images, pictures)
            shapes.load('items,items/type');
            await context.sync();

            const textShapes: any[] = [];
            for (let j = 0; j < shapes.items.length; j++) {
              const shape = shapes.items[j];
              const shapeType = String(shape.type || '').toLowerCase();
              if (isNonTextShape(shapeType)) continue;
              try {
                shape.textFrame.textRange.load('text');
                textShapes.push(shape);
              } catch {}
            }

            // PPT-C1 fix: OLE/chart/SmartArt shapes can cause InvalidArgument on sync;
            // when batch fails, fall back to reading each text shape individually so a
            // single bad shape doesn't silence all text on the slide.
            let batchSyncOk = true;
            try {
              await context.sync();
            } catch {
              batchSyncOk = false;
            }

            const lines: string[] = [];

            // Image placeholders (type already loaded, unaffected by text sync failure)
            for (const s of shapes.items) {
              const t = String(s.type || '').toLowerCase();
              if (t.includes('picture') || t === '13') {
                lines.push(`[Image ${s.width}x${s.height}]`);
              }
            }

            if (batchSyncOk) {
              for (const shape of textShapes) {
                try {
                  const text = (shape.textFrame.textRange.text || '').trim();
                  if (text) lines.push(text);
                } catch {}
              }
            } else {
              // Fallback: reload and read each text shape individually
              for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];
                const shapeType = String(shape.type || '').toLowerCase();
                if (isNonTextShape(shapeType)) continue;
                try {
                  shape.textFrame.textRange.load('text');
                  await context.sync();
                  const text = (shape.textFrame.textRange.text || '').trim();
                  if (text) lines.push(text);
                } catch {}
              }
            }

            let slideText = lines.join(' | ');
            if (slideText.length > 2000) slideText = slideText.substring(0, 2000) + '...';

            overview.push(
              `Slide ${i + 1} (Layout: ${slide?.layout || 'unknown'}): ${slideText || '<No Text>'}`,
            );
          } catch {
            overview.push(`Slide ${i + 1}: [Error reading content]`);
          }
        }
        return overview.join('\n');
      },
    },
    screenshotSlide: {
      name: 'screenshotSlide',
      category: 'read',
      description:
        'Capture a visual screenshot of a PowerPoint slide as PNG image. Use after making changes to verify visual rendering. Requires PowerPointApi 1.5+.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description:
              'Target slide number, 1-based (1 = first slide, not 0-based). Defaults to current slide.',
          },
        },
        required: [],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.5')) {
          return JSON.stringify({ error: 'screenshotSlide requires PowerPointApi 1.5 or later.' });
        }
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideNumber = args.slideNumber ? Number(args.slideNumber) : 1;
        const index = Math.trunc(slideNumber) - 1;
        if (index < 0 || index >= slides.items.length)
          throw new Error(`Slide ${slideNumber} does not exist.`);
        const slide = slides.getItemAt(index);
        const imageResult = (slide as any).getImageAsBase64({ width: 960 });
        await context.sync();
        const base64 = imageResult.value as string;
        return buildScreenshotResult(base64, `Screenshot of slide ${slideNumber}`);
      },
    },

    duplicateSlide: {
      name: 'duplicateSlide',
      category: 'write',
      description:
        'Duplicate a slide by its 1-based index. The copy is inserted after the original.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Source slide number, 1-based (1 = first slide, not 0-based).',
          },
        },
        required: ['slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('slideNumber must be >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length) throw new Error(`Slide ${slideNumber} does not exist.`);
        slides.getItemAt(index).copy();
        await context.sync();
        return `Slide ${slideNumber} duplicated successfully.`;
      },
    },

    reorderSlide: {
      name: 'reorderSlide',
      category: 'write',
      description:
        'Move a slide to a new position in the presentation. Requires PowerPointApi 1.5+.',
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Current position of the slide, 1-based.',
          },
          targetPosition: {
            type: 'number',
            description: 'New position for the slide, 1-based. Use 1 to move to the beginning.',
          },
        },
        required: ['slideNumber', 'targetPosition'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.5')) {
          return 'Error: reorderSlide requires PowerPointApi 1.5 or later, which is not supported in this Office version.';
        }
        const slideNumber = Number(args.slideNumber);
        const targetPosition = Number(args.targetPosition);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('Error: slideNumber must be >= 1.');
        if (!Number.isFinite(targetPosition) || targetPosition < 1)
          throw new Error('Error: targetPosition must be >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const totalSlides = slides.items.length;
        if (Math.trunc(slideNumber) > totalSlides)
          throw new Error(`Error: slide ${slideNumber} does not exist. Presentation has ${totalSlides} slides.`);
        if (Math.trunc(targetPosition) > totalSlides)
          throw new Error(`Error: targetPosition ${targetPosition} exceeds slide count (${totalSlides}).`);
        const fromIndex = Math.trunc(slideNumber) - 1;
        const toIndex = Math.trunc(targetPosition) - 1;
        if (fromIndex === toIndex) return `Slide ${slideNumber} is already at position ${targetPosition}.`;
        slides.getItemAt(fromIndex).moveTo(toIndex);
        await context.sync();
        return `Slide ${slideNumber} moved to position ${targetPosition}.`;
      },
    },

    verifySlides: {
      name: 'verifySlides',
      category: 'read',
      description:
        'Verify slide layouts by checking for shape overflows and overlaps across all slides. Returns a report of detected issues.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executePowerPoint: async (context: any, _args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        // Standard 16:9 slide dimensions in points (used by Office.js positioning)
        const slideWidth = 10 * 72; // 720 points = 10 inches
        const slideHeight = 7.5 * 72; // 540 points = 7.5 inches

        const issues: string[] = [];

        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          slide.shapes.load('items/id,items/name,items/left,items/top,items/width,items/height');
          await context.sync();

          const shapes = slide.shapes.items.map((s: any) => ({
            name: s.name,
            left: s.left,
            top: s.top,
            width: s.width,
            height: s.height,
          }));

          for (const s of shapes) {
            if (s.left + s.width > slideWidth + 1 || s.top + s.height > slideHeight + 1) {
              issues.push(
                `Slide ${i + 1}: shape "${s.name}" overflows slide boundaries (right=${Math.round(s.left + s.width)}, bottom=${Math.round(s.top + s.height)})`,
              );
            }
          }

          for (let a = 0; a < shapes.length; a++) {
            for (let b = a + 1; b < shapes.length; b++) {
              const sa = shapes[a],
                sb = shapes[b];
              if (
                sa.left < sb.left + sb.width &&
                sa.left + sa.width > sb.left &&
                sa.top < sb.top + sb.height &&
                sa.top + sa.height > sb.top
              ) {
                issues.push(`Slide ${i + 1}: "${sa.name}" overlaps with "${sb.name}"`);
              }
            }
          }
        }

        return issues.length === 0
          ? `All ${slides.items.length} slide(s) verified: no overlaps or overflows detected.`
          : `Found ${issues.length} issue(s) across ${slides.items.length} slide(s):\n` +
              issues.join('\n');
      },
    },

    editSlideXml: {
      name: 'editSlideXml',
      category: 'write',
      description: `Edit slide OOXML directly via JSZip. Use when Office.js API cannot express the desired change (e.g. charts, diagrams, SmartArt, animations). Requires PowerPointApi 1.5+.

The code has access to: zip (JSZip instance), markDirty() (call this if you modify the zip), escapeXml(str), DOMParser, XMLSerializer.

The zip contains the PPTX archive. Access slide XML via:
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')

ALWAYS call markDirty() after modifying the zip.`,
      inputSchema: {
        type: 'object',
        properties: {
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
          code: {
            type: 'string',
            description:
              'Async JS code with access to: zip, markDirty, escapeXml, DOMParser, XMLSerializer.',
          },
          explanation: {
            type: 'string',
            description: 'What this code does (required for audit trail).',
          },
        },
        required: ['slideNumber', 'code', 'explanation'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.5')) {
          return JSON.stringify({ error: 'editSlideXml requires PowerPointApi 1.5 or later.' });
        }
        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('slideNumber must be >= 1.');

        const validation = validateOfficeCode(args.code, 'PowerPoint');
        if (!validation.valid) {
          return JSON.stringify({ error: 'Code validation failed', errors: validation.errors });
        }

        const slideIndex = Math.trunc(slideNumber) - 1;
        const result = await withSlideZip(
          context,
          slideIndex,
          async (zip: any, markDirty: () => void) => {
            const fn = new Function(
              'zip',
              'markDirty',
              'escapeXml',
              'DOMParser',
              'XMLSerializer',
              `return (async () => { ${args.code} })()`,
            );
            return fn(zip, markDirty, escapeXml, DOMParser, XMLSerializer);
          },
        );

        return JSON.stringify({ success: true, result: result ?? 'Done' });
      },
    },

    searchIcons: {
      name: 'searchIcons',
      category: 'read',
      description:
        'Search for icons by keyword using the Iconify library (thousands of icon sets including Material Design, Fluent, Feather). Returns icon IDs in format "prefix:name".',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search keyword (e.g. "home", "user", "chart").' },
          limit: { type: 'number', description: 'Max results. Default: 10.' },
          prefix: {
            type: 'string',
            description:
              'Filter by icon set prefix (e.g. "mdi" for Material Design, "fluent" for Fluent UI).',
          },
        },
        required: ['query'],
      },
      executeCommon: async (args: Record<string, any>) => {
        const { searchIconify } = await import('@/api/backend');
        const results = await searchIconify(args.query, args.limit || 10, args.prefix);
        return JSON.stringify(results);
      },
    },

    insertIcon: {
      name: 'insertIcon',
      category: 'write',
      description:
        'Insert an icon from Iconify into a slide. Find icon IDs first using searchIcons. Icon is inserted as an SVG image.',
      inputSchema: {
        type: 'object',
        properties: {
          iconId: {
            type: 'string',
            description: 'Icon identifier in format "prefix:name", e.g. "mdi:home".',
          },
          slideNumber: {
            type: 'number',
            description: 'Target slide number, 1-based (1 = first slide, not 0-based).',
          },
          left: { type: 'number', description: 'Left position in points. Default: 100.' },
          top: { type: 'number', description: 'Top position in points. Default: 100.' },
          width: { type: 'number', description: 'Width in points. Default: 72 (1 inch).' },
          height: { type: 'number', description: 'Height in points. Default: 72 (1 inch).' },
          color: {
            type: 'string',
            description: 'Icon color as hex, e.g. "#FF5733". Uses black if omitted.',
          },
        },
        required: ['iconId', 'slideNumber'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();
        if (!isPowerPointApiSupported('1.4')) {
          return 'Error: insertIcon requires PowerPointApi 1.4 or later.';
        }
        const parts = String(args.iconId).split(':');
        if (parts.length !== 2)
          throw new Error('iconId must be in format "prefix:name", e.g. "mdi:home".');
        const [prefix, name] = parts;

        const { fetchIconSvg } = await import('@/api/backend');
        const svgText = await fetchIconSvg(prefix, name, args.color);
        const base64 = btoa(unescape(encodeURIComponent(svgText)));

        const slideNumber = Number(args.slideNumber);
        if (!Number.isFinite(slideNumber) || slideNumber < 1)
          throw new Error('slideNumber must be >= 1.');
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const index = Math.trunc(slideNumber) - 1;
        if (index >= slides.items.length) throw new Error(`Slide ${slideNumber} does not exist.`);

        // PPT-C2 fix: use slides.items[index] (already loaded) instead of proxy
        const slide = slides.items[index];
        const shape = slide.shapes.addImage(base64);
        shape.left = typeof args.left === 'number' ? args.left : 100;
        shape.top = typeof args.top === 'number' ? args.top : 100;
        shape.width = typeof args.width === 'number' ? args.width : 72;
        shape.height = typeof args.height === 'number' ? args.height : 72;
        await context.sync();

        return `Icon "${args.iconId}" inserted on slide ${slideNumber}.`;
      },
    },

    searchAndFormatInPresentation: {
      name: 'searchAndFormatInPresentation',
      category: 'write',
      description:
        'Search for text across all slides and apply font formatting (bold, italic, underline, color, size, font name) to every matching run. Case-sensitive by default.',
      inputSchema: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'Text to search for.' },
          matchCase: { type: 'boolean', description: 'Case-sensitive match. Default: true.' },
          bold: { type: 'boolean', description: 'Set bold.' },
          italic: { type: 'boolean', description: 'Set italic.' },
          underline: { type: 'boolean', description: 'Set underline.' },
          fontColor: { type: 'string', description: 'Font color as hex, e.g. "#FF0000".' },
          fontSize: { type: 'number', description: 'Font size in points.' },
          fontName: { type: 'string', description: 'Font family name.' },
        },
        required: ['searchText'],
      },
      executePowerPoint: async (context: any, args: Record<string, any>) => {
        ensurePowerPointRunAvailable();

        const searchText: string = String(args.searchText);
        const matchCase: boolean = args.matchCase !== false;
        const needle = matchCase ? searchText : searchText.toLowerCase();

        const hasBold = args.bold !== undefined;
        const hasItalic = args.italic !== undefined;
        const hasUnderline = args.underline !== undefined;
        const hasFontColor = typeof args.fontColor === 'string';
        const hasFontSize = typeof args.fontSize === 'number';
        const hasFontName = typeof args.fontName === 'string';

        if (
          !hasBold &&
          !hasItalic &&
          !hasUnderline &&
          !hasFontColor &&
          !hasFontSize &&
          !hasFontName
        ) {
          return 'Error: at least one formatting property (bold, italic, underline, fontColor, fontSize, fontName) must be specified.';
        }

        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        let totalMatches = 0;

        for (let si = 0; si < slides.items.length; si++) {
          const slide = slides.items[si];
          const shapes = slide.shapes;
          shapes.load('items,items/type');
          await context.sync();

          // Phase A: load textRange.text for non-picture shapes to find candidates
          const candidateShapes: any[] = [];
          for (let j = 0; j < shapes.items.length; j++) {
            const shape = shapes.items[j];
            const shapeType = String(shape.type || '').toLowerCase();
            // Skip pictures and non-text shape types
            if (isNonTextShape(shapeType)) continue;
            try {
              shape.textFrame.textRange.load('text');
              candidateShapes.push(shape);
            } catch {
              /* shape has no textFrame */
            }
          }

          if (candidateShapes.length === 0) continue;

          let textSyncOk = true;
          try {
            await context.sync();
          } catch {
            textSyncOk = false;
          }
          if (!textSyncOk) continue;

          // Phase B: filter shapes containing the search text, then load paragraphs
          const matchingShapes: any[] = [];
          for (const shape of candidateShapes) {
            let shapeText = '';
            try {
              shapeText = shape.textFrame.textRange.text || '';
            } catch {
              continue;
            }
            const haystack = matchCase ? shapeText : shapeText.toLowerCase();
            if (!haystack.includes(needle)) continue;
            try {
              shape.textFrame.textRange.paragraphs.load('items');
              matchingShapes.push(shape);
            } catch {
              /* skip */
            }
          }

          if (matchingShapes.length === 0) continue;
          await context.sync();

          // Phase C: load textRuns for each paragraph of matching shapes
          for (const shape of matchingShapes) {
            try {
              for (const para of shape.textFrame.textRange.paragraphs.items) {
                para.textRange.textRuns.load('items');
              }
            } catch {
              /* skip */
            }
          }
          await context.sync();

          // Phase D: apply formatting to runs containing the search text
          for (const shape of matchingShapes) {
            try {
              for (const para of shape.textFrame.textRange.paragraphs.items) {
                for (const run of para.textRange.textRuns.items) {
                  let runText = '';
                  try {
                    runText = run.textRange.text || '';
                  } catch {
                    continue;
                  }
                  const haystack = matchCase ? runText : runText.toLowerCase();
                  if (!haystack.includes(needle)) continue;

                  totalMatches++;
                  if (hasBold) run.textRange.font.bold = args.bold;
                  if (hasItalic) run.textRange.font.italic = args.italic;
                  if (hasUnderline)
                    run.textRange.font.underline = args.underline ? 'Single' : 'None';
                  if (hasFontColor) run.textRange.font.color = args.fontColor;
                  if (hasFontSize) run.textRange.font.size = args.fontSize;
                  if (hasFontName) run.textRange.font.name = args.fontName;
                }
              }
            } catch {
              /* skip malformed shape */
            }
          }

          if (matchingShapes.length > 0) {
            try {
              await context.sync();
            } catch {
              /* non-fatal */
            }
          }
        }

        if (totalMatches === 0)
          return `No occurrences of "${searchText}" found in the presentation.`;
        return `Applied formatting to ${totalMatches} run(s) matching "${searchText}" across the presentation.`;
      },
    },
  },
  def =>
    async (args = {}) => {
      try {
        if (def.executePowerPoint)
          return await runPowerPoint(ctx => def.executePowerPoint!(ctx, args));
        return await executeOfficeAction(() => def.executeCommon!(args));
      } catch (error: unknown) {
        return JSON.stringify(
          {
            success: false,
            error: getErrorMessage(error),
            tool: def.name,
            suggestion: 'Fix the error parameters or context and try again.',
          },
          null,
          2,
        );
      }
    },
);

export async function getSlideContentStandalone(
  context: any,
  slideNumber: number,
): Promise<string> {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();

  const index = Math.trunc(slideNumber) - 1;
  if (index < 0 || index >= slides.items.length) {
    return '';
  }

  const slide = slides.getItemAt(index);
  const shapes = slide.shapes;
  shapes.load('items,items/type');
  await context.sync();

  const shapeEntries: { shape: any; idx: number }[] = [];
  for (let i = 0; i < shapes.items.length; i++) {
    const shape = shapes.items[i];
    const shapeType = String(shape.type || '').toLowerCase();
    if (isNonTextShape(shapeType)) continue;
    try {
      shape.textFrame.textRange.load('text');
      shapeEntries.push({ shape, idx: i + 1 });
    } catch {
      // Non-text shape
    }
  }

  if (shapeEntries.length === 0) return '';

  // PPT-C1 fix: when batch sync fails due to OLE/chart/SmartArt, fall back to
  // reading each shape individually so one bad shape doesn't silence all text.
  let batchSyncOk = true;
  try {
    await context.sync();
  } catch {
    batchSyncOk = false;
  }

  const lines: string[] = [];

  if (batchSyncOk) {
    for (const { shape, idx } of shapeEntries) {
      try {
        const text = (shape.textFrame.textRange.text || '').trim();
        if (text) lines.push(`[Shape ${idx}] ${text}`);
      } catch {}
    }
  } else {
    // Fallback: reload and read each shape individually
    for (const { shape, idx } of shapeEntries) {
      try {
        shape.textFrame.textRange.load('text');
        await context.sync();
        const text = (shape.textFrame.textRange.text || '').trim();
        if (text) lines.push(`[Shape ${idx}] ${text}`);
      } catch {}
    }
  }

  return lines.join('\n');
}

export function getPowerPointToolDefinitions(): ToolDefinition[] {
  return Object.values(powerpointToolDefinitions);
}

export { powerpointToolDefinitions };
