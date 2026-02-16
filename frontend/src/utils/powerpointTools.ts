/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

declare const Office: any
declare const PowerPoint: any

export type PowerPointToolName =
  | 'getSelectedText'
  | 'replaceSelectedText'
  | 'getSlideCount'
  | 'getSlideContent'
  | 'addSlide'
  | 'setSlideNotes'
  | 'insertTextBox'
  | 'insertImage'

type ParsedListLine = {
  indentLevel: number
  text: string
}

function parseMarkdownListLine(line: string): ParsedListLine | null {
  const bulletMatch = line.match(/^(\s*)[-*+]\s+(.+)$/)
  if (bulletMatch) {
    const [, indent, itemText] = bulletMatch
    return {
      indentLevel: Math.floor(indent.replace(/\t/g, '  ').length / 2),
      text: itemText.trim(),
    }
  }

  const numberedMatch = line.match(/^(\s*)\d+[.)]\s+(.+)$/)
  if (numberedMatch) {
    const [, indent, itemText] = numberedMatch
    return {
      indentLevel: Math.floor(indent.replace(/\t/g, '  ').length / 2),
      text: itemText.trim(),
    }
  }

  return null
}

/**
 * PowerPoint keeps existing paragraph bullet formatting when replacing text
 * inside a bulleted shape. If we insert markdown markers (-, *, 1.) directly,
 * users can end up with duplicated bullets (native bullet + literal marker).
 *
 * This converter removes list markers while preserving hierarchy via tabs.
 */
export function normalizePowerPointListText(text: string): string {
  const lines = text.split(/\r?\n/)
  let hasListSyntax = false

  const normalizedLines = lines.map((line) => {
    const parsedLine = parseMarkdownListLine(line)
    if (!parsedLine) {
      return line
    }

    hasListSyntax = true
    return `${'\t'.repeat(parsedLine.indentLevel)}${parsedLine.text}`
  })

  return hasListSyntax ? normalizedLines.join('\n') : text
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
 * Replace the current text selection inside the active PowerPoint shape
 * with the provided text.
 */
export function insertIntoPowerPoint(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const normalizedText = normalizePowerPointListText(text)

    try {
      Office.context.document.setSelectedDataAsync(
        normalizedText,
        { coercionType: Office.CoercionType.Text },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve()
          } else {
            reject(new Error(result.error?.message || 'setSelectedDataAsync failed'))
          }
        },
      )
    } catch (err: any) {
      reject(new Error(err?.message || 'setSelectedDataAsync unavailable'))
    }
  })
}

function isPowerPointApiSupported(version: string): boolean {
  try {
    return !!Office?.context?.requirements?.isSetSupported?.('PowerPointApi', version)
  } catch {
    return false
  }
}

function ensurePowerPointRunAvailable() {
  if (typeof PowerPoint?.run !== 'function') {
    throw new Error('PowerPoint.run is not available in this Office host/runtime.')
  }
}

const powerpointToolDefinitions: Record<PowerPointToolName, PowerPointToolDefinition> = {
  getSelectedText: {
    name: 'getSelectedText',
    description: 'Get the currently selected text in PowerPoint.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return getPowerPointSelection()
    },
  },

  replaceSelectedText: {
    name: 'replaceSelectedText',
    description: 'Replace the currently selected PowerPoint text with new text.',
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
    execute: async args => {
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
    description: 'Get the total number of slides in the active PowerPoint presentation.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      ensurePowerPointRunAvailable()
      return PowerPoint.run(async (context: any) => {
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()
        return String(slides.items.length)
      })
    },
  },

  getSlideContent: {
    name: 'getSlideContent',
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
    execute: async args => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      return PowerPoint.run(async (context: any) => {
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
      })
    },
  },

  addSlide: {
    name: 'addSlide',
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
    execute: async args => {
      ensurePowerPointRunAvailable()
      return PowerPoint.run(async (context: any) => {
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
      })
    },
  },

  setSlideNotes: {
    name: 'setSlideNotes',
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
    execute: async args => {
      ensurePowerPointRunAvailable()

      if (!isPowerPointApiSupported('1.4')) {
        return 'Error: setSlideNotes requires PowerPointApi 1.4 or newer.'
      }

      const slideNumber = Number(args.slideNumber)
      const notesText = String(args.notesText ?? '')
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      return PowerPoint.run(async (context: any) => {
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
      })
    },
  },

  insertTextBox: {
    name: 'insertTextBox',
    description: 'Insert a text box into a specific slide with optional position and size.',
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
    execute: async args => {
      ensurePowerPointRunAvailable()
      const slideNumber = Number(args.slideNumber)
      if (!Number.isFinite(slideNumber) || slideNumber < 1) {
        return 'Error: slideNumber must be a number greater than or equal to 1.'
      }

      const text = String(args.text ?? '')
      if (!text) {
        return 'Error: text is required.'
      }

      return PowerPoint.run(async (context: any) => {
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        const index = Math.trunc(slideNumber) - 1
        if (index >= slides.items.length) {
          return `Error: slide ${slideNumber} does not exist. Presentation has ${slides.items.length} slides.`
        }

        const slide = slides.getItemAt(index)
        const shape = slide.shapes.addTextBox(text)
        shape.left = Number.isFinite(args.left) ? Number(args.left) : 50
        shape.top = Number.isFinite(args.top) ? Number(args.top) : 50
        shape.width = Number.isFinite(args.width) ? Number(args.width) : 500
        shape.height = Number.isFinite(args.height) ? Number(args.height) : 120
        await context.sync()

        return `Successfully inserted a text box on slide ${slideNumber}.`
      })
    },
  },

  insertImage: {
    name: 'insertImage',
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
    execute: async args => {
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

      return PowerPoint.run(async (context: any) => {
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
      })
    },
  },
}

export function getPowerPointToolDefinitions(): PowerPointToolDefinition[] {
  return Object.values(powerpointToolDefinitions)
}

export function getPowerPointTool(name: PowerPointToolName): PowerPointToolDefinition | undefined {
  return powerpointToolDefinitions[name]
}

export { powerpointToolDefinitions }
