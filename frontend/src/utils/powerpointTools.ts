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
  | 'insertContent'
  | 'getSlideContent'
  | 'addSlide'
  | 'deleteSlide'
  | 'getShapes'
  | 'getAllSlidesOverview'
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

          let styles: InheritedStyles | null = null
          try {
            textRange.font.load('name,size,color')
            await context.sync()
            styles = {
              fontFamily: textRange.font.name || '',
              fontSize: textRange.font.size ? `${textRange.font.size}pt` : '',
              fontWeight: 'normal',
              fontStyle: 'normal',
              color: '', // Do NOT force the original color here, let PowerPoint default to it or let inline HTML override it
              marginTop: '0pt',
              marginBottom: '0pt',
            }
          } catch (e) {
            // Ignore if font loading fails
          }

          // Check for native bullets to prevent double-bullet rendering
          const nativeBullets = await hasNativeBullets(context, textRange)

          let finalMarkdown = normalizedNewlines
          if (nativeBullets) {
            // Shape has native bullets: strip markdown list markers (-, *, 1.) from the LLM response
            // so they don't get rendered as HTML <ul>/<li> tags, which would cause double-bullets.
            // But we STILL use insertHtml to preserve inline formatting like **bold** and *italic*.
            finalMarkdown = stripMarkdownListMarkers(normalizedNewlines)
          }

          let html = renderOfficeCommonApiHtml(finalMarkdown)
          if (styles) html = applyInheritedStyles(html, styles)
          
          textRange.insertHtml(html, 'Replace')

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

  eval_powerpointjs: {
    name: 'eval_powerpointjs',
    category: 'write',
    description: "Execute arbitrary Office.js code within a PowerPoint.run context (requires PowerPointApi 1.4+). Use this as an escape hatch when existing tools don't cover your use case. The code runs inside `PowerPoint.run(async (context) => { ... })` with `context` available. Return a value to get it back as the result. Always call `await context.sync()` before returning.",
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: "JavaScript code to execute. Has access to `context`. Must be valid async code. Return a value to get it as result.",
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does',
        },
      },
      required: ['code'],
    },
    executePowerPoint: async (context: any, args: any) => {
      ensurePowerPointRunAvailable()
      const { code } = args as Record<string, any>
      try {
        const result = await sandboxedEval(code, { context, PowerPoint: typeof PowerPoint !== 'undefined' ? PowerPoint : undefined })
        return JSON.stringify({ success: true, result: result ?? null }, null, 2)
      } catch (err: any) {
        return JSON.stringify({ success: false, error: err.message || String(err) }, null, 2)
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
        
        const slides = context.presentation.slides
        const idx = Math.trunc(Number(slideNumber)) - 1
        const slide = slides.getItemAt(idx)
        const shape = slide.shapes.getItemOrNullObject(shapeIdOrName)
        shape.load('isNullObject')
        await context.sync()

        if (shape.isNullObject) throw new Error(`Error: Shape '${shapeIdOrName}' not found on slide ${slideNumber}.`)
        
        const textRange = shape.textFrame.textRange
        if (isPowerPointApiSupported('1.5')) {
          textRange.insertHtml(renderOfficeCommonApiHtml(content), 'Replace')
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
