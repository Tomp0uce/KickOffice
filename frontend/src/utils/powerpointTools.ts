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
import { sandboxedEval } from './sandbox';
import { validateOfficeCode } from './officeCodeValidator';
import {
  computeTextDiffStats,
  createOfficeTools,
  normalizeLineEndings,
  getErrorMessage,
} from './common';
import { message as messageUtil } from '@/utils/message';
import { withSlideZip, escapeXml } from './pptxZipUtils';
import { logService } from '@/utils/logger';
import { searchIconify, fetchIconSvg } from '@/api/backend';

declare const Office: any;
declare const PowerPoint: any;

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
  | 'verifySlides'
  | 'editSlideXml'
  | 'searchIcons'
  | 'insertIcon'
  | 'searchAndFormatInPresentation';

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
      executePowerPoint: async (context: any, args: any) => {
        ensurePowerPointRunAvailable();
        const { code, explanation } = args;

        // Validate code BEFORE execution
        const validation = validateOfficeCode(code, 'PowerPoint');

        if (!validation.valid) {
          return JSON.stringify(
            {
              success: false,
              error: 'Code validation failed. Fix the errors below and try again.',
              validationErrors: validation.errors,
              validationWarnings: validation.warnings,
              suggestion:
                'Refer to the Office.js skill document for correct patterns. Remember: slide indices are 0-based.',
              codeReceived: code.slice(0, 300) + (code.length > 300 ? '...' : ''),
            },
            null,
            2,
          );
        }

        // Log warnings but proceed
        if (validation.warnings.length > 0) {
          logService.warn('[eval_powerpointjs] Validation warnings:', validation.warnings);
        }

        try {
          // Execute in sandbox with host restriction
          const result = await sandboxedEval(
            code,
            {
              context,
              PowerPoint: typeof PowerPoint !== 'undefined' ? PowerPoint : undefined,
              Office: typeof Office !== 'undefined' ? Office : undefined,
            },
            'PowerPoint', // Restrict to PowerPoint namespace only
          );

          return JSON.stringify(
            {
              success: true,
              result: result ?? null,
              explanation,
              warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
            },
            null,
            2,
          );
        } catch (err: unknown) {
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(err),
              explanation,
              codeExecuted: code.slice(0, 200) + '...',
              hint: 'Check that all properties are loaded before access, and context.sync() is called.',
            },
            null,
            2,
          );
        }
      },
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

            for (let j = 0; j < shapes.items.length; j++) {
              const shape = shapes.items[j];
              const shapeType = String(shape.type || '').toLowerCase();
              if (shapeType.includes('picture') || shapeType === '13') continue; // keep skipping text load for images
              try {
                shape.textFrame.textRange.load('text');
              } catch {}
            }

            // PPT-C1 fix: OLE/chart/SmartArt shapes can cause InvalidArgument on sync;
            // catch the error so a single bad shape doesn't crash the entire overview.
            let textSyncOk = true;
            try {
              await context.sync();
            } catch {
              textSyncOk = false;
            }

            const lines = shapes.items
              .map((s: any) => {
                const t = String(s.type || '').toLowerCase();
                if (t.includes('picture') || t === '13') {
                  return `[Image ${s.width}x${s.height}]`;
                }
                if (!textSyncOk) return '';
                let text = '';
                try {
                  text = s.textFrame.textRange.text;
                } catch {}
                return text.trim();
              })
              .filter(Boolean);

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
        return JSON.stringify({
          __screenshot__: true,
          base64,
          mimeType: 'image/png',
          description: `Screenshot of slide ${slideNumber}`,
        });
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
            if (shapeType.includes('picture') || shapeType === '13') continue;
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
  shapes.load('items');
  await context.sync();

  const shapeEntries: { shape: any; idx: number }[] = [];
  for (let i = 0; i < shapes.items.length; i++) {
    const shape = shapes.items[i];
    try {
      shape.textFrame.textRange.load('text');
      shapeEntries.push({ shape, idx: i + 1 });
    } catch {
      // Non-text shape
    }
  }

  if (shapeEntries.length === 0) return '';

  await context.sync();

  const lines = shapeEntries
    .map(({ shape, idx }) => {
      const text = (shape.textFrame.textRange.text || '').trim();
      return text ? `[Shape ${idx}] ${text}` : '';
    })
    .filter(Boolean);

  return lines.join('\n');
}

export function getPowerPointToolDefinitions(): ToolDefinition[] {
  return Object.values(powerpointToolDefinitions);
}

export { powerpointToolDefinitions };
