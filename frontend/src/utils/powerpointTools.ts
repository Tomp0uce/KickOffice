/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

import { executeOfficeAction } from './officeAction'
import { renderOfficeCommonApiHtml, stripRichFormattingSyntax, stripMarkdownListMarkers, applyInheritedStyles, type InheritedStyles } from './officeRichText'

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
  | 'getSlideCount'
  | 'getSlideContent'
  | 'addSlide'
  | 'setSlideNotes'
  | 'insertTextBox'
  | 'insertImage'
  | 'deleteSlide'
  | 'getShapes'
  | 'deleteShape'
  | 'setShapeFill'
  | 'moveResizeShape'
  | 'getAllSlidesOverview'

/**
 * Keep list markers in plain text to preserve bullets/numbered lists when
 * the target shape is not already configured as a native bullet paragraph.
 */
export function normalizePowerPointListText(text: string): string {
  const normalizedNewlines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  return stripMarkdownListMarkers(normalizedNewlines)
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
 * Reads the current PowerPoint selection and manually reconstructs basic HTML formatting
 * (bold, italic, underline, strikethrough) by inspecting the TextRange recursively.
 * Used to preserve formatting when sending text to the LLM, as PowerPoint lacks a native getHtml API.
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
        range.load(['text', 'length'])
        await context.sync()
        
        const len = range.length
        if (len === 0) return ''
        
        // Cap length to prevent huge latency on massive selections
        const maxLen = Math.min(len, 3000) 
        
        const charRanges = []
        for (let i = 0; i < maxLen; i++) {
          const charRange = range.getSubstring(i, 1)
          charRange.load('text')
          charRange.font.load(['bold', 'italic', 'underline', 'strikethrough'])
          charRanges.push(charRange)
        }
        
        await context.sync()
        
        let html = ''
        let isBold = false
        let isItalic = false
        let isUnderline = false
        let isStrike = false
        
        for (let i = 0; i < maxLen; i++) {
          const charRange = charRanges[i]
          const text = charRange.text
          const font = charRange.font
          
          const isLineBreak = text === '\r' || text === '\n' || text === '\v'
          
          const bold = font.bold === true
          const italic = font.italic === true
          const underline = font.underline !== 'None' && font.underline !== null
          const strike = font.strikethrough === true
          
          let safeText = text
          if (safeText === '<') safeText = '&lt;'
          else if (safeText === '>') safeText = '&gt;'
          else if (safeText === '&') safeText = '&amp;'
          
          // Re-evaluate styles. If any formatting changes (or line break), close all and reopen to maintain perfectly valid HTML tree hierarchy.
          if (isStrike !== strike || isUnderline !== underline || isItalic !== italic || isBold !== bold || isLineBreak) {
             if (isStrike) html += '</s>'
             if (isUnderline) html += '</u>'
             if (isItalic) html += '</i>'
             if (isBold) html += '</b>'
             isStrike = isUnderline = isItalic = isBold = false
          }
          
          if (!isLineBreak) {
            if (bold && !isBold) { html += '<b>'; isBold = true }
            if (italic && !isItalic) { html += '<i>'; isItalic = true }
            if (underline && !isUnderline) { html += '<u>'; isUnderline = true }
            if (strike && !isStrike) { html += '<s>'; isStrike = true }
            html += safeText
          } else {
             html += '<br/>'
          }
        }
        
        // Close dangling tags
        if (isStrike) html += '</s>'
        if (isUnderline) html += '</u>'
        if (isItalic) html += '</i>'
        if (isBold) html += '</b>'
        
        return html
      })
    })
    
    return htmlOut || getPowerPointSelection()
  } catch (err) {
    console.warn('Failed to extract PowerPoint HTML selection manually:', err)
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
              color: textRange.font.color || '',
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

  replaceSelectedText: {
    name: 'replaceSelectedText',
    category: 'write',
    description: 'Replace the currently selected PowerPoint text with new content. The text will be rendered from Markdown to HTML before insertion. You can use Markdown formatting: **bold**, *italic*, bullet lists (- item), numbered lists (1. item), and headings (## Heading). Indented sub-items are supported for nested lists.',
    inputSchema: {
      type: 'object',
      properties: {
        newText: {
          type: 'string',
          description: 'The replacement text to insert in place of the current selection.',
        },
      },
      required: ['newText'],
    },
    executeCommon: async args => {
      const { newText } = args
      if (!newText || typeof newText !== 'string') {
        return 'Error: newText is required and must be a string.'
      }
      await insertIntoPowerPoint(newText)
      return 'Successfully replaced selected text in PowerPoint.'
    },
  },

  getSlideCount: {
    name: 'getSlideCount',
    category: 'read',
    description: 'Get the total number of slides in the active PowerPoint presentation.',
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
        return String(slides.items.length)
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
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          return `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`
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
    executePowerPoint: async (context: any, args) => {
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

  setSlideNotes: {
    name: 'setSlideNotes',
    category: 'write',
    description: 'Set speaker notes for a given slide (requires PowerPointApi 1.4+).',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: {
          type: 'number',
          description: 'Slide number to update (1 = first slide).',
        },
        notesText: {
          type: 'string',
          description: 'Speaker notes content to place in the notes area.',
        },
      },
      required: ['slideNumber', 'notesText'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()

      if (!isPowerPointApiSupported('1.4')) {
        return 'Error: setSlideNotes requires PowerPointApi 1.4 or newer.'
      }

      const slideNumber = Number(args.slideNumber)
      const notesText = String(args.notesText ?? '')
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          return `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`
        }

        const slide = slides.getItemAt(index)
        const notesSlide = (slide as any).notesSlide
        if (!notesSlide?.shapes?.addTextBox) {
          return 'Error: notesSlide is not available in this PowerPoint runtime.'
        }

        const notesBox = notesSlide.shapes.addTextBox(notesText)
        notesBox.left = 20
        notesBox.top = 20
        notesBox.width = 680
        notesBox.height = 300
        await context.sync()

        return `Successfully updated notes for slide ${slideNumber}.`
      },
  },

  insertTextBox: {
    name: 'insertTextBox',
    category: 'write',
    description: 'Insert a text box into a specific slide with optional position and size. Content supports Markdown formatting for rich text rendering.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: {
          type: 'number',
          description: 'Target slide number (1 = first slide).',
        },
        text: {
          type: 'string',
          description: 'Text to insert in the text box.',
        },
        left: { type: 'number', description: 'Left position in points.' },
        top: { type: 'number', description: 'Top position in points.' },
        width: { type: 'number', description: 'Text box width in points.' },
        height: { type: 'number', description: 'Text box height in points.' },
      },
      required: ['slideNumber', 'text'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      const text = String(args.text ?? '')
      if (!text) {
        return 'Error: text is required.'
      }

      
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          return `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`
        }

        const slide = slides.getItemAt(index)
        // addTextBox requires plain text — create with stripped version initially
        const shape = slide.shapes.addTextBox(stripRichFormattingSyntax(text))
        shape.left = Number.isFinite(args.left) ? Number(args.left) : 50
        shape.top = Number.isFinite(args.top) ? Number(args.top) : 50
        shape.width = Number.isFinite(args.width) ? Number(args.width) : 500
        shape.height = Number.isFinite(args.height) ? Number(args.height) : 120
        await context.sync()

        // Try to upgrade to rich HTML formatting (requires PowerPointApi 1.5+)
        if (isPowerPointApiSupported('1.5')) {
          try {
            shape.textFrame.textRange.insertHtml(renderOfficeCommonApiHtml(text), 'Replace')
            await context.sync()
          } catch {
            // insertHtml not available in this context — plain text already set
          }
        }

        return `Successfully inserted a text box on slide ${slideNumber}.`
      },
  },

  insertImage: {
    name: 'insertImage',
    category: 'write',
    description: 'Insert a base64 image into a specific slide with optional position and size.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: {
          type: 'number',
          description: 'Target slide number (1 = first slide).',
        },
        base64Image: {
          type: 'string',
          description: 'Image payload as raw base64 or data URL.',
        },
        left: { type: 'number', description: 'Left position in points.' },
        top: { type: 'number', description: 'Top position in points.' },
        width: { type: 'number', description: 'Image width in points.' },
        height: { type: 'number', description: 'Image height in points.' },
      },
      required: ['slideNumber', 'base64Image'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      const base64ImageRaw = String(args.base64Image ?? '').trim()
      if (!base64ImageRaw) {
        return 'Error: base64Image is required.'
      }
      const base64Image = base64ImageRaw.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '')

      
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          return `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`
        }

        const slide = slides.getItemAt(index)
        const shape = slide.shapes.addImage(base64Image)
        shape.left = Number.isFinite(args.left) ? Number(args.left) : 50
        shape.top = Number.isFinite(args.top) ? Number(args.top) : 50
        shape.width = Number.isFinite(args.width) ? Number(args.width) : 320
        shape.height = Number.isFinite(args.height) ? Number(args.height) : 180
        await context.sync()

        return `Successfully inserted an image on slide ${slideNumber}.`
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
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) return 'Error: slideNumber must be a number >= 1.'
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) return `Error: slide ${slideNumber} does not exist.`
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
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) return 'Error: slideNumber must be a number >= 1.'
      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) return `Error: slide ${slideNumber} does not exist.`

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

  deleteShape: {
    name: 'deleteShape',
    category: 'write',
    description: 'Delete a shape from a slide by its ID or Name.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
        shapeIdOrName: { type: 'string', description: 'ID or Name of the shape to delete.' },
      },
      required: ['slideNumber', 'shapeIdOrName'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      const shapeIdOrName = String(args.shapeIdOrName ?? '')
      if (!Number.isFinite(slideNumber) || slideNumber < 1) return 'Error: slideNumber must be a number >= 1.'
      if (!shapeIdOrName) return 'Error: shapeIdOrName is required.'

      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) return `Error: slide ${slideNumber} does not exist.`

      const slide = slides.getItemAt(index)
      const shape = slide.shapes.getItemOrNullObject(shapeIdOrName)
      shape.load('isNullObject')
      await context.sync()

      if (shape.isNullObject) return `Error: Shape '${shapeIdOrName}' not found on slide ${slideNumber}.`

      shape.delete()
      await context.sync()
      return `Successfully deleted shape '${shapeIdOrName}' from slide ${slideNumber}.`
    },
  },

  setShapeFill: {
    name: 'setShapeFill',
    category: 'write',
    description: 'Set the fill color of a shape on a slide.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
        shapeIdOrName: { type: 'string', description: 'ID or Name of the shape.' },
        color: { type: 'string', description: 'Hex color code (e.g., "#FF0000").' },
      },
      required: ['slideNumber', 'shapeIdOrName', 'color'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      const shapeIdOrName = String(args.shapeIdOrName ?? '')
      const color = String(args.color ?? '')
      if (!Number.isFinite(slideNumber) || slideNumber < 1) return 'Error: slideNumber must be a number >= 1.'
      if (!shapeIdOrName || !color) return 'Error: shapeIdOrName and color are required.'

      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) return `Error: slide ${slideNumber} does not exist.`

      const slide = slides.getItemAt(index)
      const shape = slide.shapes.getItemOrNullObject(shapeIdOrName)
      shape.load('isNullObject')
      await context.sync()

      if (shape.isNullObject) return `Error: Shape '${shapeIdOrName}' not found.`

      shape.fill.setSolidColor(color)
      await context.sync()
      return `Successfully set fill color of shape '${shapeIdOrName}' to ${color}.`
    },
  },

  moveResizeShape: {
    name: 'moveResizeShape',
    category: 'write',
    description: 'Move and/or resize a shape on a slide. Missing measurements will be left unchanged.',
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: { type: 'number', description: 'Target slide number (1 = first slide).' },
        shapeIdOrName: { type: 'string', description: 'ID or Name of the shape.' },
        left: { type: 'number', description: 'New left position in points.' },
        top: { type: 'number', description: 'New top position in points.' },
        width: { type: 'number', description: 'New width in points.' },
        height: { type: 'number', description: 'New height in points.' },
      },
      required: ['slideNumber', 'shapeIdOrName'],
    },
    executePowerPoint: async (context: any, args) => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      const shapeIdOrName = String(args.shapeIdOrName ?? '')
      if (!Number.isFinite(slideNumber) || slideNumber < 1) return 'Error: slideNumber must be a number >= 1.'
      if (!shapeIdOrName) return 'Error: shapeIdOrName is required.'

      const slides = context.presentation.slides
      slides.load('items')
      await context.sync()
      const index = Math.trunc(slideNumber) - 1
      if (index >= slides.items.length) return `Error: slide ${slideNumber} does not exist.`

      const slide = slides.getItemAt(index)
      const shape = slide.shapes.getItemOrNullObject(shapeIdOrName)
      shape.load('isNullObject')
      await context.sync()

      if (shape.isNullObject) return `Error: Shape '${shapeIdOrName}' not found.`

      if (Number.isFinite(args.left)) shape.left = Number(args.left)
      if (Number.isFinite(args.top)) shape.top = Number(args.top)
      if (Number.isFinite(args.width)) shape.width = Number(args.width)
      if (Number.isFinite(args.height)) shape.height = Number(args.height)

      await context.sync()
      return `Successfully moved/resized shape '${shapeIdOrName}'.`
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

export function getPowerPointTool(name: PowerPointToolName): PowerPointToolDefinition | undefined {
  return powerpointToolDefinitions[name]
}

export { powerpointToolDefinitions }
