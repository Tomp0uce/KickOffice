import type { ToolDefinition } from '@/types'
/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

import { executeOfficeAction } from './officeAction'
import { renderOfficeCommonApiHtml, stripRichFormattingSyntax, stripMarkdownListMarkers, applyInheritedStyles, type InheritedStyles } from './markdown'
import { sandboxedEval } from './sandbox'
import { validateOfficeCode } from './officeCodeValidator'
import { computeTextDiffStats, createOfficeTools, normalizeLineEndings } from './common'
import { message as messageUtil } from '@/utils/message'

declare const Office: any
declare const PowerPoint: any

// Point 3 Fix: Memory registry to store images without crashing the LLM
export const powerpointImageRegistry = new Map<string, string>()

const runPowerPoint = <T>(action: (context: any) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => PowerPoint.run(action) as Promise<T>)


type PowerPointToolTemplate = Omit<ToolDefinition, 'execute'> & {
  executePowerPoint?: (context: any, args: Record<string, any>) => Promise<string>
  executeCommon?: (args: Record<string, any>) => Promise<string>
}

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
  | 'eval_powerpointjs'

/**
 * Returns the 1-based slide number of the currently active/selected slide.
 * Falls back to slide 1 if the API is unavailable.
 */
export async function getCurrentSlideNumber(): Promise<number> {
  try {
    return await executeOfficeAction(() =>
      PowerPoint.run(async (context: any) => {
        let activeSlideIndex = 0
        try {
          if (typeof context.presentation.getSelectedSlides === 'function') {
            const selectedSlides = context.presentation.getSelectedSlides()
            selectedSlides.load('items/id')
            await context.sync()
            if (selectedSlides.items.length > 0) {
              const slides = context.presentation.slides
              slides.load('items/id')
              await context.sync()
              const selectedId = selectedSlides.items[0].id
              const idx = slides.items.findIndex((s: any) => s.id === selectedId)
              if (idx !== -1) activeSlideIndex = idx
            }
          }
        } catch {}
        return activeSlideIndex + 1
      })
    )
  } catch {
    return 1
  }
}

/**
 * Set speaker notes for the currently selected slide directly (no agent loop).
 * Returns true on success, false on failure.
 */
export async function setCurrentSlideSpeakerNotes(notes: string): Promise<boolean> {
  try {
    const slideNumber = await getCurrentSlideNumber()
    await executeOfficeAction(() =>
      PowerPoint.run(async (context: any) => {
        // PPT-M5: Check for API support (1.5 required for notesSlide)
        if (!isPowerPointApiSupported('1.5')) {
          throw new Error('Speaker notes modification requires PowerPoint API 1.5 or newer.')
        }

        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()
        const index = slideNumber - 1
        if (index < 0 || index >= slides.items.length) {
          throw new Error(`Slide ${slideNumber} not found.`)
        }
        
        const slide = slides.getItemAt(index)
        const notesPage = slide.notesSlide
        // Attempt to load. If it fails, the notes slide might not be initialized.
        notesPage.load('notesTextFrame/textRange')
        await context.sync()
        
        notesPage.notesTextFrame.textRange.text = notes
        await context.sync()
      })
    )
    return true
  } catch (err) {
    console.error('[PowerPointTools] Failed to set speaker notes:', err)
    messageUtil.error(`Impossible d'insérer les notes: ${err instanceof Error ? err.message : 'Erreur API'}`)
    return false
  }
}

/**
 * Read the currently selected text inside a PowerPoint shape / text box.
 * Returns an empty string when nothing is selected or the selection is
 * not a text range (e.g. an entire slide is selected).
 */
export function getPowerPointSelection(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Unformatted },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve((result.value as string) || '')
          } else {
            console.warn('PowerPoint selection error:', result.error?.message)
            resolve('')
          }
        },
      )
    } catch (err) {
      console.warn('PowerPoint getSelectedDataAsync unavailable:', err)
      resolve('')
    }
  })
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
    return getPowerPointSelection()
  }

  try {
    const htmlOut = await executeOfficeAction(async () => {
      return PowerPoint.run(async (context: any) => {
        const textRanges = context.presentation.getSelectedTextRanges()
        textRanges.load('items')
        await context.sync()

        if (textRanges.items.length === 0) return ''

        const range = textRanges.items[0]

        // Load paragraphs in a single batch (replaces per-character loading)
        const paragraphs = range.paragraphs
        paragraphs.load('items')
        await context.sync()

        if (paragraphs.items.length === 0) return ''

        // Batch-load text and font for every paragraph in one sync
        for (const para of paragraphs.items) {
          para.load('text')
          para.font.load(['bold', 'italic', 'underline', 'strikethrough'])
        }
        await context.sync()

        let html = ''
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i]
          const text: string = para.text || ''
          const font = para.font

          if (!text && i < paragraphs.items.length - 1) {
            html += '<br/>'
            continue
          }

          const bold = font.bold === true
          const italic = font.italic === true
          const underline = font.underline !== 'None' && font.underline !== null
          const strike = font.strikethrough === true

          // Escape HTML entities
          const safeText = text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')

          // Wrap with inline formatting tags (innermost first)
          let wrapped = safeText
          if (strike) wrapped = `<s>${wrapped}</s>`
          if (underline) wrapped = `<u>${wrapped}</u>`
          if (italic) wrapped = `<i>${wrapped}</i>`
          if (bold) wrapped = `<b>${wrapped}</b>`

          if (i > 0) html += '<br/>'
          html += wrapped
        }

        return html
      })
    })

    return htmlOut || getPowerPointSelection()
  } catch (err) {
    console.warn('Failed to extract PowerPoint HTML selection (paragraph mode):', err)
    return getPowerPointSelection()
  }
}


/**
 * Replace the current text selection inside the active PowerPoint shape
 * with the provided text.
 */
export async function insertIntoPowerPoint(text: string, useHtml = true): Promise<void> {
  const normalizedNewlines = normalizeLineEndings(text)

  // Try the Modern API first if available (requires PowerPointApi 1.5+)
  if (isPowerPointApiSupported('1.5') && useHtml) {
    try {
      await executeOfficeAction(async () => {
        await PowerPoint.run(async (context: any) => {
          const textRange = context.presentation.getSelectedTextRanges().getItemAt(0)
          await insertMarkdownIntoTextRange(context, textRange, normalizedNewlines)
          await context.sync()
        })
      })
      return
    } catch (e: any) {
      console.warn('Modern PowerPoint Html insertion failed, falling back:', e)
    }
  }

  // Fallback to the legacy Shared API (no native bullet detection possible here)
  const htmlContent = renderOfficeCommonApiHtml(normalizedNewlines)
  return new Promise((resolve, reject) => {
    try {
      if (useHtml) {
        Office.context.document.setSelectedDataAsync(
          htmlContent,
          { coercionType: Office.CoercionType.Html },
          (result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve()
            } else {
              console.warn('setSelectedDataAsync Html failed, falling back to raw Text')
              fallbackToText(normalizedNewlines, resolve, reject)
            }
          },
        )
      } else {
        fallbackToText(normalizedNewlines, resolve, reject)
      }
    } catch (err: any) {
      reject(new Error(err?.message || 'setSelectedDataAsync unavailable'))
    }
  })
}

export async function insertMarkdownIntoTextRange(context: any, textRange: any, text: string) {
  let styles: InheritedStyles | null = null
  try {
    textRange.font.load('name,size,color')
    await context.sync()
    styles = {
      fontFamily: textRange.font.name || '',
      fontSize: textRange.font.size ? `${textRange.font.size}pt` : '',
      fontWeight: 'normal',
      fontStyle: 'normal',
      color: '', // Do NOT force the original color here
      marginTop: '0pt',
      marginBottom: '0pt',
    }
  } catch (e) {
    // Ignore if font loading fails
  }

  const nativeBullets = await hasNativeBullets(context, textRange)

  let finalMarkdown = text
  if (nativeBullets) {
    finalMarkdown = stripMarkdownListMarkers(text)
  }

  let html = renderOfficeCommonApiHtml(finalMarkdown)
  if (styles) html = applyInheritedStyles(html, styles)
  
  textRange.insertHtml(html, 'Replace')
}

async function findShapeOnSlide(context: any, slideNumber: number, shapeIdOrName: string | number) {
  const slides = context.presentation.slides
  slides.load('items')
  await context.sync()

  const idx = Math.trunc(Number(slideNumber)) - 1
  if (idx < 0 || idx >= slides.items.length) {
    return { slide: null, shape: null, shapes: [], error: `Invalid slide number ${slideNumber}` }
  }

  const slide = slides.items[idx]
  const shapes = slide.shapes
  shapes.load('items,items/id,items/name')
  await context.sync()

  for (const s of shapes.items) {
    if (s.id === shapeIdOrName || s.name === shapeIdOrName) {
      return { slide, shape: s, shapes: shapes.items, error: null }
    }
  }

  return { slide, shape: null, shapes: shapes.items, error: `Shape '${shapeIdOrName}' not found on slide ${slideNumber}` }
}

function fallbackToText(text: string, resolve: any, reject: any) {
  // Pass true to strip list markers so it plays nice with shapes that are already natively bulleted.
  const fallbackText = stripRichFormattingSyntax(text, true)
  Office.context.document.setSelectedDataAsync(
    fallbackText,
    { coercionType: Office.CoercionType.Text },
    (fallbackResult: any) => {
      if (fallbackResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve()
      } else {
        reject(new Error(fallbackResult.error?.message || 'setSelectedDataAsync failed'))
      }
    },
  )
}



function isPowerPointApiSupported(version: string): boolean {
  try {
    return !!Office?.context?.requirements?.isSetSupported?.('PowerPointApi', version)
  } catch {
    return false
  }
}

/**
 * Detect whether the given PowerPoint text range belongs to a shape with native
 * (layout/master) bullet points. When true, we should avoid inserting HTML
 * <ul>/<li> tags to prevent double-bullet rendering.
 */
async function hasNativeBullets(context: any, textRange: any): Promise<boolean> {
  try {
    const paragraphs = textRange.paragraphs
    paragraphs.load('items')
    await context.sync()
    if (paragraphs.items.length > 0) {
      for (const para of paragraphs.items) {
        para.load('bulletFormat/visible')
      }
      await context.sync()
      // Return true if ANY paragraph has native bullets
      return paragraphs.items.some((p: any) => p.bulletFormat?.visible === true)
    }
  } catch {
    // API not available or paragraphs inaccessible — assume no native bullets
  }
  return false
}

function ensurePowerPointRunAvailable() {
  if (typeof PowerPoint?.run !== 'function') {
    throw new Error('PowerPoint.run is not available in this Office host/runtime.')
  }
}

const powerpointToolDefinitions = createOfficeTools<PowerPointToolName, PowerPointToolTemplate, ToolDefinition>({
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
    description: 'Replace the currently selected text in PowerPoint. This tool preserves block-level formatting and inline styles better than full shape replacement.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'The new text to replace the selection with. Markdown formatting is supported.' },
      },
      required: ['text'],
    },
    executeCommon: async (args: Record<string, any>) => {
      await insertIntoPowerPoint(args.text, true)
      return 'Successfully replaced selected text.'
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
      const { slideNumber, shapeIdOrName, revisedText } = args

      try {
        const { shape: targetShape, shapes, error } = await findShapeOnSlide(context, slideNumber, shapeIdOrName)

        if (!targetShape) {
          return JSON.stringify({
            success: false,
            error: error || `Shape "${shapeIdOrName}" not found on slide ${slideNumber}`,
            availableShapes: shapes.map((s: any) => ({ id: s.id, name: s.name })),
          }, null, 2)
        }

        // Get current text
        const textFrame = targetShape.textFrame
        const textRange = textFrame.textRange
        textRange.load('text')
        await context.sync()

        const originalText = textRange.text || ''

        // 4. Compute diff stats
        const { insertions, deletions, unchanged } = computeTextDiffStats(originalText, revisedText)

        // 5. Apply changes
        // PowerPoint API is limited - we do full replacement but report the diff
        textRange.text = revisedText
        await context.sync()

        return JSON.stringify({
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
        }, null, 2)
      } catch (error: any) {
        return JSON.stringify({
          success: false,
          error: error.message || String(error),
        }, null, 2)
      }
    },
  },

  getSpeakerNotes: {
    name: 'getSpeakerNotes',
    category: 'read',
    description: 'Get the speaker notes text for a specific slide (1-based index). Requires PowerPointApi 1.5+.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
      },
      required: ['slideNumber'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()
      if (!isPowerPointApiSupported('1.5')) {
        return 'Error: getSpeakerNotes requires PowerPointApi 1.5 or later, which is not supported in this Office version.'
      }
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) throw new Error('Error: slideNumber must be a number >= 1.')
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) throw new Error(`Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`)
      const slide = slides.getItemAt(index)
      const notesPage = slide.notesSlide
      notesPage.load('notesTextFrame/textRange/text')
      await context.sync()
      const text = notesPage.notesTextFrame.textRange.text ?? ''
      return text.trim() || '(no speaker notes)'
    },
  },

  setSpeakerNotes: {
    name: 'setSpeakerNotes',
    category: 'write',
    description: 'Set (replace) the speaker notes for a specific slide (1-based index). Requires PowerPointApi 1.5+.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
        notes: { type: 'string', description: 'The new speaker notes text (plain text).' },
      },
      required: ['slideNumber', 'notes'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()
      if (!isPowerPointApiSupported('1.5')) {
        return 'Error: setSpeakerNotes requires PowerPointApi 1.5 or later, which is not supported in this Office version.'
      }
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) throw new Error('Error: slideNumber must be a number >= 1.')
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) throw new Error(`Error: slide ${slideNumber} does not exist.`)
      const slide = slides.getItemAt(index)
      const notesPage = slide.notesSlide
      notesPage.load('notesTextFrame/textRange')
      await context.sync()
      notesPage.notesTextFrame.textRange.text = String(args.notes ?? '')
      await context.sync()
      return `Successfully updated speaker notes for slide ${slideNumber}.`
    },
  },

  insertImageOnSlide: {
    name: 'insertImageOnSlide',
    category: 'write',
    description: 'Insert an image onto a specific slide. Position and size are in points (1 inch = 72 points). Default: centered at 400×300pt.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
        filename: { type: 'string', description: 'The exact filename of the uploaded image to insert.' },
        left: { type: 'number', description: 'Left position in points from the left edge of the slide. Default: 100.' },
        top: { type: 'number', description: 'Top position in points from the top edge of the slide. Default: 100.' },
        width: { type: 'number', description: 'Width in points. Default: 400.' },
        height: { type: 'number', description: 'Height in points. Default: 300.' },
      },
      required: ['slideNumber', 'filename'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()
      if (!isPowerPointApiSupported('1.4')) {
        return 'Error: insertImageOnSlide requires PowerPointApi 1.4 or later.'
      }
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) throw new Error('Error: slideNumber must be a number >= 1.')
      
      const filename = String(args.filename)
      const base64 = powerpointImageRegistry.get(filename)
      
      if (!base64) {
        throw new Error(`Error: Could not find image data for filename "${filename}". Make sure the name matches exactly. Available in session: ${Array.from(powerpointImageRegistry.keys()).join(', ') || 'none'}`)
      }

      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) throw new Error(`Error: slide ${slideNumber} does not exist.`)

      const slide = slides.getItemAt(index)
      const shape = slide.shapes.addImage(base64)
      shape.left = typeof args.left === 'number' ? args.left : 100
      shape.top = typeof args.top === 'number' ? args.top : 100
      shape.width = typeof args.width === 'number' ? args.width : 400
      shape.height = typeof args.height === 'number' ? args.height : 300
      await context.sync()
      
      return `Successfully inserted image "${filename}" on slide ${slideNumber} at position (${shape.left}, ${shape.top}) with size ${shape.width}×${shape.height} points.`
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
          description: 'JavaScript code following the template. Must include load(), sync(), and try/catch.',
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does (required for audit trail).',
        },
      },
      required: ['code', 'explanation'],
    },
    executePowerPoint: async (context: any, args: any) => {
      ensurePowerPointRunAvailable()
      const { code, explanation } = args

      // Validate code BEFORE execution
      const validation = validateOfficeCode(code, 'PowerPoint')

      if (!validation.valid) {
        return JSON.stringify({
          success: false,
          error: 'Code validation failed. Fix the errors below and try again.',
          validationErrors: validation.errors,
          validationWarnings: validation.warnings,
          suggestion: 'Refer to the Office.js skill document for correct patterns. Remember: slide indices are 0-based.',
          codeReceived: code.slice(0, 300) + (code.length > 300 ? '...' : ''),
        }, null, 2)
      }

      // Log warnings but proceed
      if (validation.warnings.length > 0) {
        console.warn('[eval_powerpointjs] Validation warnings:', validation.warnings)
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
          'PowerPoint'  // Restrict to PowerPoint namespace only
        )

        return JSON.stringify({
          success: true,
          result: result ?? null,
          explanation,
          warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
        }, null, 2)
      } catch (err: any) {
        return JSON.stringify({
          success: false,
          error: err.message || String(err),
          explanation,
          codeExecuted: code.slice(0, 200) + '...',
          hint: 'Check that all properties are loaded before access, and context.sync() is called.',
        }, null, 2)
      }
    },
  },

  insertContent: {
    name: 'insertContent',
    category: 'write',
    description: 'The PREFERRED tool for adding or replacing content in PowerPoint. Supports Markdown (bold, italic, bullets). Can target a specific shape by ID/Name or the current selection. Handles style inheritance automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'The content to insert in Markdown format.',
        },
        slideNumber: {
          type: 'number',
          description: 'Optional: Target slide number (1 = first slide). Required if shapeIdOrName is provided.',
        },
        shapeIdOrName: {
          type: 'string',
          description: 'Optional: ID or Name of the shape to update (from getShapes). If omitted, replaces current selection.',
        },
      },
      required: ['content'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      const { content, slideNumber, shapeIdOrName } = args
      
      if (shapeIdOrName) {
        if (!slideNumber) throw new Error('Error: slideNumber is required when shapeIdOrName is provided.')

        const { shape, shapes, error } = await findShapeOnSlide(context, slideNumber, shapeIdOrName)

        if (!shape) {
          const availableShapes = shapes.map((s: any) => `'${s.name}' (id: ${s.id})`).join(', ')
          throw new Error(`Error: ${error}. Available shapes are: ${availableShapes}`)
        }

        const textRange = shape.textFrame.textRange
        if (isPowerPointApiSupported('1.5')) {
          try {
            const normalizedNewlines = normalizeLineEndings(content)
            await insertMarkdownIntoTextRange(context, textRange, normalizedNewlines)
          } catch (e) {
            console.warn('insertMarkdownIntoTextRange failed, falling back to text modification', e)
            textRange.text = stripRichFormattingSyntax(content)
            await context.sync()
            return `Content set on shape '${shapeIdOrName}' on slide ${slideNumber} (plain text fallback — markdown formatting was not applied because the rich text API failed).`
          }
        } else {
          textRange.text = stripRichFormattingSyntax(content)
        }
        await context.sync()
        return `Successfully set content on shape '${shapeIdOrName}' on slide ${slideNumber}.`
      } else {
        // Replace current selection
        await insertIntoPowerPoint(content)
        return 'Successfully replaced selected text in PowerPoint.'
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
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        throw new Error('Error: slideNumber must be a number greater than or equal to 1.')
      }
      return getSlideContentStandalone(context, slideNumber)
    },
  },

  addSlide: {
    name: 'addSlide',
    category: 'write',
    description: 'Add a new slide to the presentation. Pass title and body to automatically populate the template text boxes (recommended). If title/body are provided, the tool discovers the layout shapes and fills them in one step.',
    inputSchema: {
      type: 'object',
      properties: {
        layout: {
          type: 'string',
          description: 'Optional slide layout name supported by PowerPointApi (e.g., Blank, Title, TitleAndContent).',
        },
        title: {
          type: 'string',
          description: 'Optional title text to insert into the title placeholder of the new slide.',
        },
        body: {
          type: 'string',
          description: 'Optional body/content text to insert into the main content placeholder. Use newlines for bullet points.',
        },
      },
      required: [],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()

      const slides = context.presentation.slides
      const layout = typeof args.layout === 'string' && args.layout.trim().length > 0
        ? args.layout.trim()
        : undefined
      const titleText = typeof args.title === 'string' ? args.title.trim() : undefined
      const bodyText = typeof args.body === 'string' ? args.body.trim() : undefined

      if (layout) {
        ;(slides as any).add({ layout })
      } else {
        slides.add()
      }
      await context.sync()

      // If title or body provided, populate the template text boxes
      if (titleText || bodyText) {
        try {
          slides.load('items')
          await context.sync()
          const newSlide = slides.items[slides.items.length - 1]
          const shapes = newSlide.shapes
          // Load name and placeholderFormat (correct PowerPoint.js API, not .placeholder)
          shapes.load('items,items/name,items/placeholderFormat')
          await context.sync()

          let titleFilled = false
          let bodyFilled = false

          for (const shape of shapes.items) {
            try {
              // Use placeholderFormat.type when available (PowerPoint.js 1.3+)
              let phType = ''
              try { phType = String((shape as any).placeholderFormat?.type ?? '').toLowerCase() } catch {}
              const nameLower = (shape.name as string ?? '').toLowerCase()

              const isTitle = phType === 'title' || phType === 'centeredtitle' || phType === 'subtitle'
                || nameLower.includes('title')
              const isBody = phType === 'body'
                || nameLower.includes('content') || nameLower.includes('body')
                || (nameLower.includes('text') && !nameLower.includes('title'))

              if (isTitle && titleText && !titleFilled) {
                shape.textFrame.textRange.text = titleText
                titleFilled = true
              } else if (isBody && bodyText && !bodyFilled) {
                shape.textFrame.textRange.text = bodyText
                bodyFilled = true
              }
            } catch {}
          }
          await context.sync()
        } catch (err: any) {
          // Non-fatal: slide was created, just couldn't populate shapes
          return `Successfully added a slide${layout ? ` with layout "${layout}"` : ''} but could not auto-fill shapes: ${err?.message ?? 'unknown error'}`
        }
      }

      return layout
        ? `Successfully added a slide with layout "${layout}"${titleText ? ` titled "${titleText}"` : ''}.`
        : `Successfully added a slide${titleText ? ` titled "${titleText}"` : ''}.`
    },
  },



  deleteSlide: {
    name: 'deleteSlide',
    category: 'write',
    description: 'Delete a specific slide by its 1-based index.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
      },
      required: ['slideNumber'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) throw new Error('Error: slideNumber must be a number >= 1.')
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) throw new Error(`Error: slide ${slideNumber} does not exist.`)
      slides.getItemAt(index).delete()
      await context.sync()
      return `Successfully deleted slide ${slideNumber}.`
    },
  },

  getShapes: {
    name: 'getShapes',
    category: 'read',
    description: 'Get all shapes on a specific slide (1-based index) with their properties (type, text, position).',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
      },
      required: ['slideNumber'],
    },
    executePowerPoint: async (context: any, args: Record<string, any>) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) throw new Error('Error: slideNumber must be a number >= 1.')
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) throw new Error(`Error: slide ${slideNumber} does not exist.`)

      const slide = slides.getItemAt(index)
      const shapes = slide.shapes
      shapes.load('items')
      await context.sync()

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i]
        try { shape.textFrame.textRange.load('text') } catch {}
        try { shape.load(['id', 'name', 'type', 'left', 'top', 'width', 'height']) } catch {}
      }
      await context.sync()

      const details = shapes.items.map((shape: any) => {
        let text = ''
        try { text = shape.textFrame.textRange.text } catch {}
        return {
          id: shape.id,
          name: shape.name,
          type: shape.type,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
          text: text.trim()
        }
      })
      return JSON.stringify(details, null, 2)
    },
  },

  getAllSlidesOverview: {
    name: 'getAllSlidesOverview',
    category: 'read',
    description: 'Get a text overview of all slides in the presentation (useful to understand the presentation structure).',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executePowerPoint: async (context: any) => {
      ensurePowerPointRunAvailable()
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()

      const overview: string[] = []
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i]
        slide.load(['id', 'layout'])
        const shapes = slide.shapes
        // Load shape type together with items to filter out non-text shapes (images, pictures)
        shapes.load('items,items/type')
        await context.sync()

        for (let j = 0; j < shapes.items.length; j++) {
          const shape = shapes.items[j]
          const shapeType = String(shape.type || '').toLowerCase()
          if (shapeType.includes('picture') || shapeType === '13') continue // keep skipping text load for images
          try { shape.textFrame.textRange.load('text') } catch {}
        }
        await context.sync()

        const lines = shapes.items
          .map((s: any) => {
            const t = String(s.type || '').toLowerCase()
            if (t.includes('picture') || t === '13') {
               return `[Image ${s.width}x${s.height}]`
            }
            let text = ''
            try { text = s.textFrame.textRange.text } catch {}
            return text.trim()
          })
          .filter(Boolean)

        let slideText = lines.join(' | ')
        if (slideText.length > 2000) slideText = slideText.substring(0, 2000) + '...'

        overview.push(`Slide ${i + 1} (Layout: ${slide?.layout || 'unknown'}): ${slideText || '<No Text>'}`)
      }
      return overview.join('\n')
    },
  },
}, (def) => async (args = {}) => {
  try {
    if (def.executePowerPoint) return await runPowerPoint(ctx => def.executePowerPoint!(ctx, args))
    return await executeOfficeAction(() => def.executeCommon!(args))
  } catch (error: any) {
    return JSON.stringify({
      error: true,
      message: error.message || String(error),
      tool: def.name,
      suggestion: 'Fix the error parameters or context and try again.'
    }, null, 2)
  }
})

export async function getSlideContentStandalone(context: any, slideNumber: number): Promise<string> {
  const slides = context.presentation.slides
  slides.load('items')
  await context.sync()

  const index = Math.trunc(slideNumber) - 1
  if (index < 0 || index >= slides.items.length) {
    return ''
  }

  const slide = slides.getItemAt(index)
  const shapes = slide.shapes
  shapes.load('items')
  await context.sync()

  const shapeEntries: { shape: any; idx: number }[] = []
  for (let i = 0; i < shapes.items.length; i++) {
    const shape = shapes.items[i]
    try {
      shape.textFrame.textRange.load('text')
      shapeEntries.push({ shape, idx: i + 1 })
    } catch {
      // Non-text shape
    }
  }

  if (shapeEntries.length === 0) return ''

  await context.sync()

  const lines = shapeEntries
    .map(({ shape, idx }) => {
      const text = (shape.textFrame.textRange.text || '').trim()
      return text ? `[Shape ${idx}] ${text}` : ''
    })
    .filter(Boolean)

  return lines.join('\n')
}

export function getToolDefinitions(): ToolDefinition[] {
  return Object.values(powerpointToolDefinitions)
}

export const getPowerPointToolDefinitions = getToolDefinitions

export { powerpointToolDefinitions }
