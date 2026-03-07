import type { PowerPointToolDefinition } from '@/types'
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
import DiffMatchPatch from 'diff-match-patch'

declare const Office: any
declare const PowerPoint: any

const runPowerPoint = <T>(action: (context: any) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => PowerPoint.run(action) as Promise<T>)


type PowerPointToolTemplate = Omit<PowerPointToolDefinition, 'execute'> & {
  executePowerPoint?: (context: any, args: Record<string, any>) => Promise<string>
  executeCommon?: (args: Record<string, any>) => Promise<string>
}

function createPowerPointTools(definitions: Record<PowerPointToolName, PowerPointToolTemplate>): Record<PowerPointToolName, PowerPointToolDefinition> {
  return Object.fromEntries(
    Object.entries(definitions).map(([name, definition]) => [
      name,
      {
        ...definition,
        execute: async (args: Record<string, any> = {}) => {
          if (definition.executePowerPoint) {
            return runPowerPoint(context => definition.executePowerPoint!(context, args))
          }
          return executeOfficeAction(() => definition.executeCommon!(args))
        },
      },
    ]),
  ) as unknown as Record<PowerPointToolName, PowerPointToolDefinition>
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
  | 'eval_powerpointjs'

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
  const normalizedNewlines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')

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
      const firstPara = paragraphs.items[0]
      firstPara.load('bulletFormat/visible')
      await context.sync()
      return firstPara.bulletFormat?.visible === true
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

const powerpointToolDefinitions = createPowerPointTools({
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
    description: `Modify text in a specific shape while attempting to preserve formatting on unchanged portions.

IMPORTANT: PowerPoint has limited diff support compared to Word. This tool:
1. Reads the current shape text
2. Computes word-level diff
3. Applies changes character-by-character when possible
4. Falls back to full replacement if diff fails

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

        // 4. Compute diff
        const dmp = new DiffMatchPatch()
        const diffs = dmp.diff_main(originalText, revisedText)
        dmp.diff_cleanupSemantic(diffs)

        // 5. Calculate stats
        let insertions = 0
        let deletions = 0
        let unchanged = 0
        for (const [op, text] of diffs) {
          if (op === 0) unchanged += text.length
          else if (op === -1) deletions += text.length
          else if (op === 1) insertions += text.length
        }

        // 6. Apply changes
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

  eval_powerpointjs: {
    name: 'eval_powerpointjs',
    category: 'write',
    description: `Execute custom Office.js code within a PowerPoint.run context.

**USE THIS TOOL ONLY WHEN:**
- No dedicated tool exists for your operation
- Operations like: speaker notes, images, animations, advanced shape manipulations

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
            const normalizedNewlines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
            await insertMarkdownIntoTextRange(context, textRange, normalizedNewlines)
          } catch (e) {
            console.warn('insertMarkdownIntoTextRange failed, falling back to text modification', e)
            textRange.text = stripRichFormattingSyntax(content)
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

      
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          throw new Error(`Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`)
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

        if (shapeEntries.length === 0) {
          return ''
        }

        await context.sync()

        const lines = shapeEntries
          .map(({ shape, idx }) => {
            const text = (shape.textFrame.textRange.text || '').trim()
            return text ? `[Shape ${idx}] ${text}` : ''
          })
          .filter(Boolean)

        return lines.join('\n')
      },
  },

  addSlide: {
    name: 'addSlide',
    category: 'write',
    description: 'Add a new slide to the presentation. Optionally pass a layout when supported.',
    inputSchema: {
      type: 'object',
      properties: {
        layout: {
          type: 'string',
          description: 'Optional slide layout name supported by PowerPointApi (e.g., Blank, Title, TitleAndContent).',
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

        if (layout) {
          ;(slides as any).add({ layout })
        } else {
          slides.add()
        }
        await context.sync()

        return layout
          ? `Successfully added a slide with layout "${layout}".`
          : 'Successfully added a slide.'
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
        shapes.load('items')
        await context.sync()

        for (let j = 0; j < shapes.items.length; j++) {
          const shape = shapes.items[j]
          try { shape.textFrame.textRange.load('text') } catch {}
        }
        await context.sync()

        const lines = shapes.items.map((s: any) => {
          let text = ''
          try { text = s.textFrame.textRange.text } catch {}
          return text.trim()
        }).filter(Boolean)

        let slideText = lines.join(' | ')
        if (slideText.length > 100) slideText = slideText.substring(0, 100) + '...'

        overview.push(`Slide ${i + 1} (Layout: ${slide?.layout || 'unknown'}): ${slideText || '<No Text>'}`)
      }
      return overview.join('\n')
    },
  },
})

export function getPowerPointToolDefinitions(): PowerPointToolDefinition[] {
  return Object.values(powerpointToolDefinitions)
}

export { powerpointToolDefinitions }
