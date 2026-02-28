import { executeOfficeAction } from './officeAction'
import DiffMatchPatch from 'diff-match-patch'
import { sandboxedEval } from './sandbox'

import { applyInheritedStyles, type InheritedStyles, renderOfficeRichHtml, stripRichFormattingSyntax, htmlToMarkdown } from './officeRichText'

// R17 — Generate a visual diff HTML string (insertions in blue/underline, deletions in red/strikethrough)
function generateVisualDiff(originalText: string, newText: string): string {
  const dmp = new DiffMatchPatch()
  const diffs = dmp.diff_main(originalText, newText)
  dmp.diff_cleanupSemantic(diffs)

  return diffs
    .map(([op, text]) => {
      const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')
      if (op === 1) return `<span style="color:blue;text-decoration:underline;">${escaped}</span>`
      if (op === -1) return `<span style="color:red;text-decoration:line-through;">${escaped}</span>`
      return escaped
    })
    .join('')
}

export type WordToolName =
  | 'getSelectedText'
  | 'getDocumentContent'
  | 'insertText'
  | 'replaceSelectedText'
  | 'appendText'
  | 'insertParagraph'
  | 'formatText'
  | 'searchAndReplace'
  | 'getDocumentProperties'
  | 'insertTable'
  | 'insertList'
  | 'deleteText'
  | 'clearFormatting'
  | 'setFontName'
  | 'insertPageBreak'
  | 'getRangeInfo'
  | 'selectText'
  | 'insertImage'
  | 'getTableInfo'
  | 'insertBookmark'
  | 'goToBookmark'
  | 'insertContentControl'
  | 'findText'
  | 'applyTaggedFormatting'
  | 'setParagraphFormat'
  | 'insertHyperlink'
  | 'getDocumentHtml'
  | 'modifyTableCell'
  | 'addTableRow'
  | 'addTableColumn'
  | 'deleteTableRowColumn'
  | 'formatTableCell'
  | 'insertHeaderFooter'
  | 'insertFootnote'
  | 'addComment'
  | 'getComments'
  | 'setPageSetup'
  | 'getSpecificParagraph'
  | 'insertSectionBreak'
  | 'applyStyle'
  | 'getSelectedTextWithFormatting'
  | 'eval_wordjs'

const runWord = <T>(action: (context: Word.RequestContext) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => Word.run(action))

/**
 * Detect whether the current selection is inside a Word native list
 * (bulleted or numbered paragraph). Used to prevent double-bullet issues
 * when inserting HTML list content into an already-bulleted context.
 */
async function isInsertionInList(context: Word.RequestContext): Promise<boolean> {
  try {
    const selection = context.document.getSelection()
    const para = selection.paragraphs.getFirst()
    para.load('style')
    await context.sync()
    // listItem throws if the paragraph is not part of a list
    para.listItem.load('level')
    await context.sync()
    return true
  } catch {
    return false
  }
}

/**
 * Strip outer <ul>/<ol> wrapper tags from HTML while keeping <li> elements.
 * Prevents double-bullets when inserting list HTML into an existing list context —
 * Word's native list style handles the bullet display for each <li>.
 */
function stripOuterListTags(html: string): string {
  return html
    .replace(/<ul[^>]*>/gi, '')
    .replace(/<\/ul>/gi, '')
    .replace(/<ol[^>]*>/gi, '')
    .replace(/<\/ol>/gi, '')
}

interface InsertionContext {
  inList: boolean
  styles: InheritedStyles
}

/**
 * R1 — Read insertion-point font/spacing styles and list context in 2 syncs.
 * Combines isInsertionInList + style capture to minimise round-trips.
 * Note: paragraph-level spacing (spaceBefore/spaceAfter) is on Paragraph, not Range.
 */
async function readInsertionContext(context: Word.RequestContext): Promise<InsertionContext> {
  // Sync 1: load font from range + spaceBefore/spaceAfter + style from first paragraph
  const selection = context.document.getSelection()
  selection.load('font/name,font/size,font/bold,font/italic,font/color')
  const para = selection.paragraphs.getFirst()
  para.load('style,spaceBefore,spaceAfter')
  await context.sync()

  const styles: InheritedStyles = {
    fontFamily: selection.font.name || '',
    fontSize: selection.font.size ? `${selection.font.size}pt` : '',
    fontWeight: selection.font.bold ? 'bold' : 'normal',
    fontStyle: selection.font.italic ? 'italic' : 'normal',
    color: selection.font.color || '',
    marginTop: `${para.spaceBefore ?? 0}pt`,
    marginBottom: `${para.spaceAfter ?? 0}pt`,
  }

  // Sync 2: detect list context (throws if paragraph is not in a list)
  let inList = false
  try {
    para.listItem.load('level')
    await context.sync()
    inList = true
  } catch {
    inList = false
  }

  return { inList, styles }
}

/**
 * R14 — After inserting HTML, detect paragraphs whose font size matches heading
 * thresholds and apply the corresponding Word builtin heading style.
 * This ensures headings become proper Word styles (h1→Heading 1, etc.) so that
 * features like Table of Contents work correctly.
 * Best-effort: silently skips if the API call fails.
 */
async function applyHeadingBuiltinStyles(
  context: Word.RequestContext,
  insertedRange: Word.Range,
  html: string,
): Promise<void> {
  if (!/\<h[1-6]/i.test(html)) return

  try {
    const paras = insertedRange.paragraphs
    paras.load('items')
    await context.sync()

    for (const p of paras.items) {
      p.load('font/size')
    }
    await context.sync()

    for (const p of paras.items) {
      const size = p.font.size
      if (!size) continue
      if (size >= 20) p.styleBuiltIn = Word.BuiltInStyleName.heading1
      else if (size >= 15) p.styleBuiltIn = Word.BuiltInStyleName.heading2
      else if (size >= 12.5) p.styleBuiltIn = Word.BuiltInStyleName.heading3
    }
    await context.sync()
  } catch {
    // Best-effort only — heading size detection may not be available in all runtimes
  }
}

type WordToolTemplate = Omit<WordToolDefinition, 'execute'> & {
  executeWord: (context: Word.RequestContext, args: Record<string, any>) => Promise<string>
}

function createWordTools(definitions: Record<WordToolName, WordToolTemplate>): Record<WordToolName, WordToolDefinition> {
  return Object.fromEntries(
    Object.entries(definitions).map(([name, definition]) => [
      name,
      {
        ...definition,
        execute: async (args: Record<string, any> = {}) => runWord(context => definition.executeWord(context, args)),
      },
    ]),
  ) as unknown as Record<WordToolName, WordToolDefinition>
}

const wordToolDefinitions = createWordTools({
  getSelectedText: {
    name: 'getSelectedText',
    category: 'read',
    description:
      'Get the currently selected text in the Word document. Returns the selected text or empty string if nothing is selected.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async context => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      return range.text || ''
    },
  },

  getDocumentContent: {
    name: 'getDocumentContent',
    category: 'read',
    description: 'Get the full content of the Word document body as plain text.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async context => {
      const body = context.document.body
      body.load('text')
      await context.sync()
      return body.text || ''
    },
  },

  insertText: {
    name: 'insertText',
    category: 'write',
    description: 'Insert text at the current cursor position in the Word document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to insert',
        },
        location: {
          type: 'string',
          description: 'Where to insert: "Start", "End", "Before", "After", or "Replace"',
          enum: ['Start', 'End', 'Before', 'After', 'Replace'],
        },
      },
      required: ['text'],
    },
    executeWord: async (context, args) => {
      const { text, location = 'End' } = args
      // R1: read insertion context (list detection + document styles) in 2 syncs
      const { inList, styles } = await readInsertionContext(context)
      const html = renderOfficeRichHtml(text)
      const adjustedHtml = inList ? stripOuterListTags(html) : html
      // R1: inject document-level font/spacing into <p> and <li> elements
      const styledHtml = applyInheritedStyles(adjustedHtml, styles)
      const range = context.document.getSelection()
      const insertedRange = range.insertHtml(styledHtml, location as any)
      await context.sync()  // execute the insertHtml
      // R14: post-process headings to apply Word builtin heading styles
      await applyHeadingBuiltinStyles(context, insertedRange, html)
      return `Successfully inserted text at ${location}`
    },
  },

  replaceSelectedText: {
    name: 'replaceSelectedText',
    category: 'write',
    description: 'Replace the currently selected text with new text. CRITICAL: When correcting typos or making small modifications within a large block of text, NEVER use replaceSelectedText to replace the entire text block, as this destroys complex document layouts. ALWAYS use searchAndReplace to make surgical, targeted changes.',
    inputSchema: {
      type: 'object',
      properties: {
        newText: {
          type: 'string',
          description: 'The new text to replace the selection with',
        },
        preserveFormatting: {
          type: 'boolean',
          description: 'Keep the original font styling (name, size, color, bold, italic, underline, highlight) from the selected text. Default: true.',
        },
        diffTracking: {
          type: 'boolean',
          description: 'R17: When true, instead of replacing the selection, show a visual diff: insertions in blue/underline, deletions in red/strikethrough. Useful for proofreading and corrections. Default: false.',
        },
      },
      required: ['newText'],
    },
    executeWord: async (context, args) => {
      const { newText, preserveFormatting = true, diffTracking = false } = args
      const range = context.document.getSelection()
      range.load('text,styleBuiltIn,font/name,font/size,font/bold,font/italic,font/underline,font/color,font/highlightColor')
      // R15: load paragraph-level spacing/alignment from first paragraph (not available on Range)
      const firstPara = range.paragraphs.getFirst()
      firstPara.load('alignment,spaceBefore,spaceAfter')
      await context.sync()

      if (!range.text || range.text.length === 0) {
        return 'Error: No text selected. Select text in the document, then try again.'
      }

      // R17: diffTracking mode — show visual diff instead of replacing directly
      if (diffTracking) {
        const diffHtml = generateVisualDiff(range.text, newText)
        range.insertHtml(diffHtml, 'Replace')
        await context.sync()
        return 'Inserted text with visual diff: insertions shown in blue/underline, deletions in red/strikethrough.'
      }

      // Capture all formatting before additional syncs invalidate cached values
      const savedFontName = range.font.name
      const savedFontSize = range.font.size
      const savedFontColor = range.font.color
      const savedHighlightColor = range.font.highlightColor
      const savedStyleBuiltIn = range.styleBuiltIn
      // R15: paragraph-level properties (from first paragraph)
      const savedAlignment = firstPara.alignment
      const savedSpaceBefore = firstPara.spaceBefore
      const savedSpaceAfter = firstPara.spaceAfter

      // Detect list context to avoid double-bullets
      const inList = await isInsertionInList(context)
      const html = renderOfficeRichHtml(newText)
      
      // We pass the saved font name and size to applyInheritedStyles so the new text
      // matches the selection context natively by default, but inline styles in the HTML
      // (like colors returned by the LLM) will override it locally.
      const inheritedStyles: InheritedStyles = {
        fontFamily: savedFontName || '',
        fontSize: savedFontSize ? `${savedFontSize}pt` : '',
        fontWeight: 'normal',
        fontStyle: 'normal',
        color: '', // Do NOT force the original color here, let Word default to it or let inline HTML override it
        marginTop: savedSpaceBefore ? `${savedSpaceBefore}pt` : '0pt',
        marginBottom: savedSpaceAfter ? `${savedSpaceAfter}pt` : '0pt',
      }
      
      const adjustedHtml = inList ? stripOuterListTags(html) : html
      const styledHtml = applyInheritedStyles(adjustedHtml, inheritedStyles)
      const insertedRange = range.insertHtml(styledHtml, 'Replace')

      if (preserveFormatting) {
        insertedRange.font.name = savedFontName
        insertedRange.font.size = savedFontSize
        // CRITICAL FIX: We do NOT restore font.color or font.highlightColor onto the entire range natively.
        // Doing so forces Word to override any internal <span> colors generated by the LLM. 
      }

      await context.sync()  // execute insertHtml + font restoration

      // R15: restore paragraph spacing/alignment on each inserted paragraph
      if (preserveFormatting) {
        try {
          const insertedParas = insertedRange.paragraphs
          insertedParas.load('items')
          await context.sync()
          for (const p of insertedParas.items) {
            if (savedStyleBuiltIn && savedStyleBuiltIn !== 'Normal') p.styleBuiltIn = savedStyleBuiltIn as any
            if (savedAlignment) p.alignment = savedAlignment
            if (savedSpaceBefore != null) p.spaceBefore = savedSpaceBefore
            if (savedSpaceAfter != null) p.spaceAfter = savedSpaceAfter
          }
          await context.sync()
        } catch {
          // Best-effort: paragraph spacing restoration may fail in some runtimes
        }
      }

      // R14: post-process headings to apply Word builtin heading styles
      await applyHeadingBuiltinStyles(context, insertedRange, html)
      return preserveFormatting
        ? 'Successfully replaced selected text while preserving layout formatting'
        : 'Successfully replaced selected text'
    },
  },

  appendText: {
    name: 'appendText',
    category: 'write',
    description: 'Append text to the end of the document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to append',
        },
      },
      required: ['text'],
    },
    executeWord: async (context, args) => {
      const { text } = args
      const body = context.document.body
      body.insertHtml(renderOfficeRichHtml(text), 'End')
      await context.sync()
      return 'Successfully appended text to document'
    },
  },

  insertParagraph: {
    name: 'insertParagraph',
    category: 'format',
    description: 'Insert a new paragraph at the specified location.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The paragraph text',
        },
        location: {
          type: 'string',
          description:
            'Where to insert: "After" (after cursor/selection), "Before" (before cursor), "Start" (start of doc), or "End" (end of doc). Default is "After".',
          enum: ['After', 'Before', 'Start', 'End'],
        },
        style: {
          type: 'string',
          description: 'Optional Word built-in style: Normal, Heading1, Heading2, Heading3, Quote, etc.',
          enum: [
            'Normal',
            'Heading1',
            'Heading2',
            'Heading3',
            'Heading4',
            'Quote',
            'IntenseQuote',
            'Title',
            'Subtitle',
          ],
        },
      },
      required: ['text'],
    },
    executeWord: async (context, args) => {
      const { text, location = 'After', style } = args
      let range
      const htmlText = renderOfficeRichHtml(text)

      if (location === 'Start' || location === 'End') {
        const body = context.document.body
        range = body.insertHtml(htmlText, location)
      } else {
        const selectionRange = context.document.getSelection()
        range = selectionRange.insertHtml(htmlText, location as 'After' | 'Before')
      }
      if (style) {
        range.styleBuiltIn = style as any
      }
      await context.sync()
      return `Successfully inserted paragraph at ${location}`
    },
  },

  formatText: {
    name: 'formatText',
    category: 'format',
    description: 'Apply formatting to the currently selected text. At least one text character must be selected.',
    inputSchema: {
      type: 'object',
      properties: {
        bold: {
          type: 'boolean',
          description: 'Make text bold',
        },
        italic: {
          type: 'boolean',
          description: 'Make text italic',
        },
        underline: {
          type: 'boolean',
          description: 'Underline text',
        },
        fontSize: {
          type: 'number',
          description: 'Font size in points',
        },
        fontColor: {
          type: 'string',
          description: 'Font color as hex (e.g., "#FF0000" for red)',
        },
        highlightColor: {
          type: 'string',
          description:
            'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { bold, italic, underline, fontSize, fontColor, highlightColor } = args
      
        const range = context.document.getSelection()
        range.load('text')
        await context.sync()

        if (!range.text || range.text.length === 0) {
          return 'Error: No text selected. Select text in the document, then try again.'
        }

        if (bold !== undefined) range.font.bold = bold
        if (italic !== undefined) range.font.italic = italic
        if (underline !== undefined) range.font.underline = underline ? 'Single' : 'None'
        if (fontSize !== undefined) range.font.size = fontSize
        if (fontColor !== undefined) range.font.color = fontColor
        if (highlightColor !== undefined) range.font.highlightColor = highlightColor

        await context.sync()
        return 'Successfully applied formatting'
      },
  },

  searchAndReplace: {
    name: 'searchAndReplace',
    category: 'read',
    description: 'Search for text in the document and replace it with new text. PREFERRED METHOD for correcting typos, modifying phrasing, or making targeted changes within paragraphs without destroying the surrounding document layout. Be surgical.',
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'The text to search for',
        },
        replaceText: {
          type: 'string',
          description: 'The text to replace with',
        },
        matchCase: {
          type: 'boolean',
          description: 'Whether to match case (default: false)',
        },
        matchWholeWord: {
          type: 'boolean',
          description: 'Whether to match whole word only (default: false)',
        },
      },
      required: ['searchText', 'replaceText'],
    },
    executeWord: async (context, args) => {
      const { searchText, replaceText, matchCase = false, matchWholeWord = false } = args
      
        const body = context.document.body
        const searchResults = body.search(searchText, {
          matchCase,
          matchWholeWord,
        })
        searchResults.load('items')
        await context.sync()

        const count = searchResults.items.length
        for (const item of searchResults.items) {
          item.insertText(replaceText, 'Replace')
        }
        await context.sync()
        return `Replaced ${count} occurrence(s) of "${searchText}" with "${replaceText}"`
      },
  },

  getDocumentProperties: {
    name: 'getDocumentProperties',
    category: 'read',
    description: 'Get document properties including paragraph count, word count, and character count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const body = context.document.body
        body.load(['text'])

        const paragraphs = body.paragraphs
        paragraphs.load('items')

        await context.sync()

        const text = body.text || ''
        const wordCount = text.split(/\s+/).filter(word => word.length > 0).length
        const charCount = text.length
        const paragraphCount = paragraphs.items.length

        return JSON.stringify(
          {
            paragraphCount,
            wordCount,
            characterCount: charCount,
          },
          null,
          2,
        )
      },
  },

  insertTable: {
    name: 'insertTable',
    category: 'write',
    description: 'Insert a table at the current cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        rows: {
          type: 'number',
          description: 'Number of rows',
        },
        columns: {
          type: 'number',
          description: 'Number of columns',
        },
        data: {
          type: 'array',
          description: 'Optional 2D array of cell values',
          items: {
            type: 'array',
            items: { type: 'string' },
          },
        },
      },
      required: ['rows', 'columns'],
    },
    executeWord: async (context, args) => {
      const { rows, columns, data } = args
      
        const range = context.document.getSelection()

        // Create table data with markdown stripped for raw cells
        const tableData: string[][] = (data || Array(rows).fill(null).map(() => Array(columns).fill('')))
          .map((row: string[]) => row.map((cell: string) => stripRichFormattingSyntax(cell || '')))

        const table = range.insertTable(rows, columns, 'After', tableData)
        table.styleBuiltIn = 'GridTable1Light'

        await context.sync()
        return `Successfully inserted ${rows}x${columns} table`
      },
  },

  insertList: {
    name: 'insertList',
    category: 'write',
    description: 'Insert a bulleted or numbered list at the current position. Each item can be a plain string or an object {text, children} for nested sub-items.',
    inputSchema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          description:
            'Array of list items. Each entry is either a plain string or an object {text: string, children?: string[]} for nested sub-items.',
          items: { type: 'object' },
        },
        listType: {
          type: 'string',
          description: 'Type of list: "bullet" or "number"',
          enum: ['bullet', 'number'],
        },
      },
      required: ['items', 'listType'],
    },
    executeWord: async (context, args) => {
      const { items, listType } = args
      const range = context.document.getSelection()

      let topLevelIndex = 0
      const markdownLines: string[] = []

      for (const item of items) {
        const text = typeof item === 'string' ? item : (item.text ?? '')
        const marker = listType === 'bullet' ? '-' : `${++topLevelIndex}.`
        markdownLines.push(`${marker} ${text}`)

        const children: string[] = typeof item === 'object' && Array.isArray(item.children) ? item.children : []
        for (const child of children) {
          markdownLines.push(listType === 'bullet' ? `  - ${child}` : `  1. ${child}`)
        }
      }

      range.insertHtml(renderOfficeRichHtml(markdownLines.join('\n')), 'After')
      await context.sync()
      return `Successfully inserted ${listType} list with ${items.length} items`
    },
  },

  deleteText: {
    name: 'deleteText',
    category: 'write',
    description:
      'Delete the currently selected text or a specific range. If no text is selected, this will delete at the cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        direction: {
          type: 'string',
          description: 'Direction to delete if nothing selected: "Before" (backspace) or "After" (delete key)',
          enum: ['Before', 'After'],
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { direction = 'After' } = args
      
        const range = context.document.getSelection()
        range.load('text')
        await context.sync()

        if (range.text && range.text.length > 0) {
          range.delete()
        } else {
          if (direction === 'After') {
            range.insertText('', 'After')
          } else {
            range.insertText('', 'Before')
          }
        }
        await context.sync()
        return 'Successfully deleted text'
      },
  },

  clearFormatting: {
    name: 'clearFormatting',
    category: 'format',
    description: 'Clear all formatting from the selected text, returning it to default style.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const range = context.document.getSelection()
        range.font.bold = false
        range.font.italic = false
        range.font.underline = 'None'
        range.styleBuiltIn = 'Normal'
        await context.sync()
        return 'Successfully cleared formatting'
      },
  },

  setFontName: {
    name: 'setFontName',
    category: 'format',
    description: 'Set the font name/family for the selected text (e.g., Arial, Times New Roman, Calibri).',
    inputSchema: {
      type: 'object',
      properties: {
        fontName: {
          type: 'string',
          description: 'The font name to apply (e.g., "Arial", "Times New Roman", "Calibri", "Consolas")',
        },
      },
      required: ['fontName'],
    },
    executeWord: async (context, args) => {
      const { fontName } = args
      
        const range = context.document.getSelection()
        range.font.name = fontName
        await context.sync()
        return `Successfully set font to ${fontName}`
      },
  },

  insertPageBreak: {
    name: 'insertPageBreak',
    category: 'write',
    description: 'Insert a page break at the current cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        location: {
          type: 'string',
          description: 'Where to insert: "Before", "After", "Start", or "End"',
          enum: ['Before', 'After', 'Start', 'End'],
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { location = 'After' } = args
      
        const range = context.document.getSelection()
        // insertBreak only supports Before and After for page breaks
        const insertLoc = location === 'Start' || location === 'Before' ? 'Before' : 'After'
        range.insertBreak('Page', insertLoc)
        await context.sync()
        return `Successfully inserted page break ${location.toLowerCase()}`
      },
  },

  getRangeInfo: {
    name: 'getRangeInfo',
    category: 'read',
    description: 'Get detailed information about the current selection including text, formatting, and position.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const range = context.document.getSelection()
        range.load([
          'text',
          'style',
          'font/name',
          'font/size',
          'font/bold',
          'font/italic',
          'font/underline',
          'font/color',
        ])
        await context.sync()

        return JSON.stringify(
          {
            text: range.text || '',
            style: range.style,
            font: {
              name: range.font.name,
              size: range.font.size,
              bold: range.font.bold,
              italic: range.font.italic,
              underline: range.font.underline,
              color: range.font.color,
            },
          },
          null,
          2,
        )
      },
  },

  selectText: {
    name: 'selectText',
    category: 'write',
    description: 'Select all text in the document or specific location.',
    inputSchema: {
      type: 'object',
      properties: {
        scope: {
          type: 'string',
          description: 'What to select: "All" for entire document',
          enum: ['All'],
        },
      },
      required: ['scope'],
    },
    executeWord: async (context, args) => {
      const { scope } = args
      
        if (scope === 'All') {
          const body = context.document.body
          body.select()
          await context.sync()
          return 'Successfully selected all text'
        }
        return 'Invalid scope'
      },
  },

  insertImage: {
    name: 'insertImage',
    category: 'write',
    description: 'Insert an image from a URL at the current cursor position. The image URL must be accessible.',
    inputSchema: {
      type: 'object',
      properties: {
        imageUrl: {
          type: 'string',
          description: 'The URL of the image to insert',
        },
        width: {
          type: 'number',
          description: 'Optional width in points',
        },
        height: {
          type: 'number',
          description: 'Optional height in points',
        },
        location: {
          type: 'string',
          description: 'Where to insert: "Before", "After", "Start", "End", or "Replace"',
          enum: ['Before', 'After', 'Start', 'End', 'Replace'],
        },
      },
      required: ['imageUrl'],
    },
    executeWord: async (context, args) => {
      const { imageUrl, width, height, location = 'After' } = args
      
        const range = context.document.getSelection()
        const image = range.insertInlinePictureFromBase64(imageUrl, location as any)

        if (width) image.width = width
        if (height) image.height = height

        await context.sync()
        return `Successfully inserted image at ${location}`
      },
  },

  getTableInfo: {
    name: 'getTableInfo',
    category: 'read',
    description: 'Get information about tables in the document, including row and column counts.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const tables = context.document.body.tables
        tables.load(['items'])
        await context.sync()

        const tableInfos = []
        for (let i = 0; i < tables.items.length; i++) {
          const table = tables.items[i]
          table.load(['rowCount', 'values'])
          await context.sync()

          const columnCount = table.values && table.values[0] ? table.values[0].length : 0

          tableInfos.push({
            index: i,
            rowCount: table.rowCount,
            columnCount,
          })
        }

        return JSON.stringify(
          {
            tableCount: tables.items.length,
            tables: tableInfos,
          },
          null,
          2,
        )
      },
  },

  insertBookmark: {
    name: 'insertBookmark',
    category: 'write',
    description: 'Insert a bookmark at the current selection to mark a location in the document.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'The name of the bookmark (must be unique, no spaces allowed)',
        },
      },
      required: ['name'],
    },
    executeWord: async (context, args) => {
      const { name } = args
      
        const range = context.document.getSelection()

        const bookmarkName = name.replace(/\s+/g, '_')

        const contentControl = range.insertContentControl()
        contentControl.tag = `bookmark_${bookmarkName}`
        contentControl.title = bookmarkName
        contentControl.appearance = 'Tags'

        await context.sync()
        return `Successfully inserted bookmark: ${bookmarkName}`
      },
  },

  goToBookmark: {
    name: 'goToBookmark',
    category: 'write',
    description: 'Navigate to a previously created bookmark in the document.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'The name of the bookmark to navigate to',
        },
      },
      required: ['name'],
    },
    executeWord: async (context, args) => {
      const { name } = args
      
        const bookmarkName = name.replace(/\s+/g, '_')
        const contentControls = context.document.contentControls
        contentControls.load(['items'])
        await context.sync()

        for (const cc of contentControls.items) {
          cc.load(['tag', 'title'])
          await context.sync()

          if (cc.tag === `bookmark_${bookmarkName}` || cc.title === bookmarkName) {
            cc.select()
            await context.sync()
            return `Successfully navigated to bookmark: ${bookmarkName}`
          }
        }

        return `Bookmark not found: ${bookmarkName}`
      },
  },

  insertContentControl: {
    name: 'insertContentControl',
    category: 'write',
    description:
      'Insert a content control (a container for content) at the current selection. Useful for creating structured documents.',
    inputSchema: {
      type: 'object',
      properties: {
        title: {
          type: 'string',
          description: 'The title of the content control',
        },
        tag: {
          type: 'string',
          description: 'Optional tag for programmatic identification',
        },
        appearance: {
          type: 'string',
          description: 'Visual appearance of the control',
          enum: ['BoundingBox', 'Tags', 'Hidden'],
        },
      },
      required: ['title'],
    },
    executeWord: async (context, args) => {
      const { title, tag, appearance = 'BoundingBox' } = args
      
        const range = context.document.getSelection()
        const contentControl = range.insertContentControl()
        contentControl.title = title
        if (tag) contentControl.tag = tag
        contentControl.appearance = appearance as Word.ContentControlAppearance

        await context.sync()
        return `Successfully inserted content control: ${title}`
      },
  },

  findText: {
    name: 'findText',
    category: 'read',
    description: 'Find text in the document and return information about matches. Does not modify the document.',
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'The text to search for',
        },
        matchCase: {
          type: 'boolean',
          description: 'Whether to match case (default: false)',
        },
        matchWholeWord: {
          type: 'boolean',
          description: 'Whether to match whole word only (default: false)',
        },
      },
      required: ['searchText'],
    },
    executeWord: async (context, args) => {
      const { searchText, matchCase = false, matchWholeWord = false } = args
      
        const body = context.document.body
        const searchResults = body.search(searchText, {
          matchCase,
          matchWholeWord,
        })
        searchResults.load(['items'])
        await context.sync()

        const count = searchResults.items.length
        return JSON.stringify(
          {
            searchText,
            matchCount: count,
            found: count > 0,
          },
          null,
          2,
        )
      },
  },

  applyTaggedFormatting: {
    name: 'applyTaggedFormatting',
    category: 'format',
    description:
      'Convert inline formatting tags in the document into real Word formatting (e.g., <format>text</format> can apply size, font, italic, bold, underline, strike, highlight, color, and other font settings).',
    inputSchema: {
      type: 'object',
      properties: {
        tagName: {
          type: 'string',
          description: 'Tag name to process (default: "format")',
        },
        fontName: {
          type: 'string',
          description: 'Font family name (e.g., "Calibri", "Arial")',
        },
        fontSize: {
          type: 'number',
          description: 'Font size in points',
        },
        color: {
          type: 'string',
          description: 'Font color to apply as hex or named color',
        },
        highlightColor: {
          type: 'string',
          description:
            'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
        },
        bold: {
          type: 'boolean',
          description: 'Whether to apply bold formatting',
        },
        italic: {
          type: 'boolean',
          description: 'Whether to apply italic formatting',
        },
        underline: {
          type: 'boolean',
          description: 'Whether to apply underline formatting',
        },
        strikethrough: {
          type: 'boolean',
          description: 'Whether to apply strikethrough formatting',
        },
        allCaps: {
          type: 'boolean',
          description: 'Whether to format text in all caps',
        },
        subscript: {
          type: 'boolean',
          description: 'Whether to apply subscript formatting',
        },
        superscript: {
          type: 'boolean',
          description: 'Whether to apply superscript formatting',
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const tagName = typeof args.tagName === 'string' && args.tagName.trim() ? args.tagName.trim() : 'format'
      const fontName = typeof args.fontName === 'string' && args.fontName.trim() ? args.fontName.trim() : undefined
      const fontSize = typeof args.fontSize === 'number' ? args.fontSize : undefined
      const color = typeof args.color === 'string' && args.color.trim() ? args.color.trim() : undefined
      const highlightColor =
        typeof args.highlightColor === 'string' && args.highlightColor.trim() ? args.highlightColor.trim() : undefined
      const bold = args.bold !== undefined ? Boolean(args.bold) : undefined
      const italic = args.italic !== undefined ? Boolean(args.italic) : undefined
      const underline = args.underline !== undefined ? Boolean(args.underline) : undefined
      const strikethrough = args.strikethrough !== undefined ? Boolean(args.strikethrough) : undefined
      const allCaps = args.allCaps !== undefined ? Boolean(args.allCaps) : undefined
      const subscript = args.subscript !== undefined ? Boolean(args.subscript) : undefined
      const superscript = args.superscript !== undefined ? Boolean(args.superscript) : undefined

      
        const body = context.document.body
        body.load('text')
        await context.sync()

        const openingTag = `<${tagName}>`
        const closingTag = `</${tagName}>`
        const openingTagMatches = body.search(openingTag, { matchCase: true, matchWholeWord: false })
        openingTagMatches.load('items')
        await context.sync()

        if (openingTagMatches.items.length === 0) {
          return `No <${tagName}>...</${tagName}> tags found.`
        }

        let replacedCount = 0
        for (let i = openingTagMatches.items.length - 1; i >= 0; i--) {
          const openingRange = openingTagMatches.items[i]
          const afterOpeningRange = openingRange.getRange('After')
          const closingTagMatches = afterOpeningRange.search(closingTag, {
            matchCase: true,
            matchWholeWord: false,
          })
          closingTagMatches.load('items')
          await context.sync()

          if (closingTagMatches.items.length === 0) {
            continue
          }

          const taggedRange = openingRange.expandTo(closingTagMatches.items[0])
          taggedRange.load('text')
          await context.sync()

          const taggedText = taggedRange.text || ''
          if (!taggedText.endsWith(closingTag)) {
            continue
          }

          const innerText = taggedText.slice(openingTag.length, taggedText.length - closingTag.length)
          const formattedRange = taggedRange.insertText(innerText, 'Replace')

          if (fontName !== undefined) formattedRange.font.name = fontName
          if (fontSize !== undefined) formattedRange.font.size = fontSize
          if (color !== undefined) formattedRange.font.color = color
          if (highlightColor !== undefined) formattedRange.font.highlightColor = highlightColor
          if (bold !== undefined) formattedRange.font.bold = bold
          if (italic !== undefined) formattedRange.font.italic = italic
          if (underline !== undefined) formattedRange.font.underline = underline ? 'Single' : 'None'
          if (strikethrough !== undefined) formattedRange.font.strikeThrough = strikethrough
          if (allCaps !== undefined) formattedRange.font.allCaps = allCaps
          if (subscript !== undefined) formattedRange.font.subscript = subscript
          if (superscript !== undefined) formattedRange.font.superscript = superscript

          replacedCount++
          await context.sync()
        }

        return `Converted ${replacedCount} tagged occurrence(s) using <${tagName}>...</${tagName}>.`
      },
  },

  setParagraphFormat: {
    name: 'setParagraphFormat',
    category: 'format',
    description:
      'Set paragraph formatting on the current selection (alignment, spacing before/after, line spacing, and indentation).',
    inputSchema: {
      type: 'object',
      properties: {
        alignment: {
          type: 'string',
          description: 'Paragraph alignment: Left, Centered, Right, or Justified.',
          enum: ['Left', 'Centered', 'Right', 'Justified'],
        },
        lineSpacing: {
          type: 'number',
          description: 'Line spacing value in points.',
        },
        spaceBefore: {
          type: 'number',
          description: 'Paragraph spacing before in points.',
        },
        spaceAfter: {
          type: 'number',
          description: 'Paragraph spacing after in points.',
        },
        leftIndent: {
          type: 'number',
          description: 'Left indent in points.',
        },
        firstLineIndent: {
          type: 'number',
          description: 'First line indent in points.',
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { alignment, lineSpacing, spaceBefore, spaceAfter, leftIndent, firstLineIndent } = args
      
        const selection = context.document.getSelection()
        const paragraphs = selection.paragraphs
        paragraphs.load('items')
        await context.sync()

        if (paragraphs.items.length === 0) {
          return 'Error: No paragraph found in the current selection.'
        }

        for (const paragraph of paragraphs.items) {
          if (alignment !== undefined) paragraph.alignment = alignment as Word.Alignment
          if (lineSpacing !== undefined) paragraph.lineSpacing = lineSpacing
          if (spaceBefore !== undefined) paragraph.spaceBefore = spaceBefore
          if (spaceAfter !== undefined) paragraph.spaceAfter = spaceAfter
          if (leftIndent !== undefined) paragraph.leftIndent = leftIndent
          if (firstLineIndent !== undefined) paragraph.firstLineIndent = firstLineIndent
        }

        await context.sync()
        return `Successfully formatted ${paragraphs.items.length} paragraph(s)`
      },
  },

  insertHyperlink: {
    name: 'insertHyperlink',
    category: 'write',
    description: 'Insert a clickable hyperlink at the current selection.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'URL address of the hyperlink.',
        },
        textToDisplay: {
          type: 'string',
          description: 'Optional display text for the hyperlink. If omitted, uses the current selection.',
        },
      },
      required: ['address'],
    },
    executeWord: async (context, args) => {
      const { address, textToDisplay } = args
      
        const range = context.document.getSelection()
        const linkRange = typeof textToDisplay === 'string' ? range.insertText(textToDisplay, 'Replace') : range
        linkRange.hyperlink = address
        await context.sync()
        return 'Successfully inserted hyperlink'
      },
  },

  getDocumentHtml: {
    name: 'getDocumentHtml',
    category: 'read',
    description: 'Get the full document body as HTML to preserve structure and rich formatting context.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const htmlResult = context.document.body.getHtml()
        await context.sync()
        return htmlResult.value || ''
      },
  },

  modifyTableCell: {
    name: 'modifyTableCell',
    category: 'write',
    description: 'Replace content of a specific table cell in an existing table.',
    inputSchema: {
      type: 'object',
      properties: {
        row: { type: 'number', description: 'Zero-based row index.' },
        column: { type: 'number', description: 'Zero-based column index.' },
        text: { type: 'string', description: 'Text to set in the table cell.' },
        tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
      },
      required: ['row', 'column', 'text'],
    },
    executeWord: async (context, args) => {
      const { row, column, text, tableIndex = 0 } = args
      
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          return 'Error: Table not found at the requested index.'
        }

        const targetTable = tables.items[tableIndex]
        const cell = targetTable.getCell(row, column)
        cell.body.insertText(text, 'Replace')
        await context.sync()
        return `Successfully updated table ${tableIndex}, cell (${row}, ${column})`
      },
  },

  addTableRow: {
    name: 'addTableRow',
    category: 'write',
    description: 'Add row(s) to an existing table.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
        location: {
          type: 'string',
          description: 'Insert location relative to the current selection in the table.',
          enum: ['Before', 'After'],
        },
        count: { type: 'number', description: 'Number of rows to add.' },
        values: {
          type: 'array',
          description: 'Optional row values as a 2D array.',
          items: { type: 'array', items: { type: 'string' } },
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { tableIndex = 0, location = 'After', count = 1, values } = args
      
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          return 'Error: Table not found at the requested index.'
        }

        const targetTable = tables.items[tableIndex] as any
        targetTable.addRows(location, count, values)
        await context.sync()
        return `Successfully added ${count} row(s) to table ${tableIndex}`
      },
  },

  addTableColumn: {
    name: 'addTableColumn',
    category: 'write',
    description: 'Add column(s) to an existing table.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
        location: {
          type: 'string',
          description: 'Insert location relative to the current selection in the table.',
          enum: ['Before', 'After'],
        },
        count: { type: 'number', description: 'Number of columns to add.' },
        values: {
          type: 'array',
          description: 'Optional column values as a 2D array.',
          items: { type: 'array', items: { type: 'string' } },
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { tableIndex = 0, location = 'After', count = 1, values } = args
      
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          return 'Error: Table not found at the requested index.'
        }

        const targetTable = tables.items[tableIndex] as any
        targetTable.addColumns(location, count, values)
        await context.sync()
        return `Successfully added ${count} column(s) to table ${tableIndex}`
      },
  },

  deleteTableRowColumn: {
    name: 'deleteTableRowColumn',
    category: 'write',
    description: 'Delete row(s) or column(s) from an existing table.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
        target: {
          type: 'string',
          description: 'Whether to delete rows or columns.',
          enum: ['row', 'column'],
        },
        index: { type: 'number', description: 'Zero-based row or column index.' },
        count: { type: 'number', description: 'Number of rows or columns to delete.' },
      },
      required: ['target', 'index'],
    },
    executeWord: async (context, args) => {
      const { tableIndex = 0, target, index, count = 1 } = args
      
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          return 'Error: Table not found at the requested index.'
        }

        const targetTable = tables.items[tableIndex] as any
        if (target === 'row') {
          targetTable.deleteRows(index, count)
          await context.sync()
          return `Successfully deleted ${count} row(s) starting at index ${index}`
        }

        targetTable.deleteColumns(index, count)
        await context.sync()
        return `Successfully deleted ${count} column(s) starting at index ${index}`
      },
  },

  formatTableCell: {
    name: 'formatTableCell',
    category: 'format',
    description: 'Apply formatting to a specific table cell (background and font style).',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
        row: { type: 'number', description: 'Zero-based row index.' },
        column: { type: 'number', description: 'Zero-based column index.' },
        shadingColor: { type: 'string', description: 'Cell background color (hex or color name).' },
        fontName: { type: 'string', description: 'Font family to apply.' },
        fontSize: { type: 'number', description: 'Font size in points.' },
        fontColor: { type: 'string', description: 'Font color (hex or color name).' },
        bold: { type: 'boolean', description: 'Whether font is bold.' },
        italic: { type: 'boolean', description: 'Whether font is italic.' },
      },
      required: ['row', 'column'],
    },
    executeWord: async (context, args) => {
      const { tableIndex = 0, row, column, shadingColor, fontName, fontSize, fontColor, bold, italic } = args
      
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          return 'Error: Table not found at the requested index.'
        }

        const cell = tables.items[tableIndex].getCell(row, column) as any
        if (shadingColor !== undefined) cell.shadingColor = shadingColor

        const cellBody = cell.body
        if (fontName !== undefined) cellBody.font.name = fontName
        if (fontSize !== undefined) cellBody.font.size = fontSize
        if (fontColor !== undefined) cellBody.font.color = fontColor
        if (bold !== undefined) cellBody.font.bold = bold
        if (italic !== undefined) cellBody.font.italic = italic

        await context.sync()
        return `Successfully formatted table ${tableIndex}, cell (${row}, ${column})`
      },
  },

  insertHeaderFooter: {
    name: 'insertHeaderFooter',
    category: 'write',
    description: 'Insert text into the document header or footer of the first section.',
    inputSchema: {
      type: 'object',
      properties: {
        target: {
          type: 'string',
          description: 'Where to insert text.',
          enum: ['header', 'footer'],
        },
        type: {
          type: 'string',
          description: 'Header/footer type: Primary, FirstPage, or EvenPages.',
          enum: ['Primary', 'FirstPage', 'EvenPages'],
        },
        text: {
          type: 'string',
          description: 'Text to insert into header/footer.',
        },
      },
      required: ['target', 'text'],
    },
    executeWord: async (context, args) => {
      const { target, type = 'Primary', text } = args
      
        const section = context.document.sections.getFirst() as any
        const container = target === 'header' ? section.getHeader(type) : section.getFooter(type)
        container.insertText(text, 'Replace')
        await context.sync()
        return `Successfully inserted text into ${target} (${type})`
      },
  },

  insertFootnote: {
    name: 'insertFootnote',
    category: 'write',
    description: 'Insert a footnote at the current selection.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Footnote content.' },
      },
      required: ['text'],
    },
    executeWord: async (context, args) => {
      const { text } = args
      
        const range = context.document.getSelection() as any
        range.insertFootnote(text)
        await context.sync()
        return 'Successfully inserted footnote'
      },
  },

  addComment: {
    name: 'addComment',
    category: 'write',
    description: 'Add a review comment bubble to a specific segment within the current selection. Use this to suggest changes, provide feedback, or alert the user during proofreading without modifying the original text directly.',
    inputSchema: {
      type: 'object',
      properties: {
        textSegment: { type: 'string', description: 'The specific word or phrase in the selection that has an error.' },
        comment: { type: 'string', description: 'The feedback or correction for the text segment.' },
      },
      required: ['textSegment', 'comment'],
    },
    executeWord: async (context, args) => {
      const { textSegment, comment } = args
      
      const range = context.document.getSelection() as any
      // Perform a search within the selected range
      const searchResults = range.search(textSegment, { matchCase: false })
      searchResults.load('items')
      
      await context.sync()
      
      if (searchResults.items && searchResults.items.length > 0) {
        // Apply the comment exclusively to the first match within the selection
        searchResults.items[0].insertComment(comment)
      } else {
        // Fallback: apply to entire selection if the specific segment string is somehow not found
        range.insertComment(`[On: "${textSegment}"]\n${comment}`)
      }
      
      await context.sync()
      return `Successfully added comment to segment: "${textSegment}"`
    },
  },

  getComments: {
    name: 'getComments',
    category: 'read',
    description: 'Get comments from the document body with author and content.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      
        const body = context.document.body as any
        const comments = body.getComments()
        comments.load('items/content,items/authorName')
        await context.sync()

        return JSON.stringify(
          {
            count: comments.items.length,
            comments: comments.items.map((comment: any) => ({
              authorName: comment.authorName || '',
              content: comment.content || '',
            })),
          },
          null,
          2,
        )
      },
  },

  setPageSetup: {
    name: 'setPageSetup',
    category: 'format',
    description: 'Configure page setup on the first section (margins, orientation, paper size).',
    inputSchema: {
      type: 'object',
      properties: {
        topMargin: { type: 'number', description: 'Top margin in points.' },
        bottomMargin: { type: 'number', description: 'Bottom margin in points.' },
        leftMargin: { type: 'number', description: 'Left margin in points.' },
        rightMargin: { type: 'number', description: 'Right margin in points.' },
        orientation: {
          type: 'string',
          description: 'Page orientation: Portrait or Landscape.',
          enum: ['Portrait', 'Landscape'],
        },
        paperSize: {
          type: 'string',
          description: 'Paper size (for example Letter, A4, Legal).',
          enum: ['Letter', 'A4', 'Legal'],
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { topMargin, bottomMargin, leftMargin, rightMargin, orientation, paperSize } = args
      
        const section = context.document.sections.getFirst() as any
        const pageSetup = section.pageSetup

        if (topMargin !== undefined) pageSetup.topMargin = topMargin
        if (bottomMargin !== undefined) pageSetup.bottomMargin = bottomMargin
        if (leftMargin !== undefined) pageSetup.leftMargin = leftMargin
        if (rightMargin !== undefined) pageSetup.rightMargin = rightMargin
        if (orientation !== undefined) pageSetup.orientation = orientation
        if (paperSize !== undefined) pageSetup.paperSize = paperSize

        await context.sync()
        return 'Successfully updated page setup'
      },
  },

  getSpecificParagraph: {
    name: 'getSpecificParagraph',
    category: 'format',
    description: 'Get a paragraph by index without reading all document content.',
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'number',
          description: 'Zero-based paragraph index.',
        },
      },
      required: ['index'],
    },
    executeWord: async (context, args) => {
      const { index } = args
      
        const paragraphs = context.document.body.paragraphs
        paragraphs.load('items')
        await context.sync()

        if (index < 0 || index >= paragraphs.items.length) {
          return `Error: Paragraph index out of bounds. Range is 0 to ${Math.max(paragraphs.items.length - 1, 0)}.`
        }

        const paragraph = paragraphs.items[index]
        paragraph.load('text,style,font/name,font/size,font/bold,font/italic,font/underline,font/color')
        await context.sync()

        return JSON.stringify(
          {
            index,
            text: paragraph.text || '',
            style: paragraph.style,
            font: {
              name: paragraph.font.name,
              size: paragraph.font.size,
              bold: paragraph.font.bold,
              italic: paragraph.font.italic,
              underline: paragraph.font.underline,
              color: paragraph.font.color,
            },
          },
          null,
          2,
        )
      },
  },

  insertSectionBreak: {
    name: 'insertSectionBreak',
    category: 'write',
    description: 'Insert a section break at the current selection.',
    inputSchema: {
      type: 'object',
      properties: {
        location: {
          type: 'string',
          description: 'Where to insert section break: Before or After current selection.',
          enum: ['Before', 'After'],
        },
      },
      required: [],
    },
    executeWord: async (context, args) => {
      const { location = 'After' } = args

        const range = context.document.getSelection() as any
        range.insertBreak('SectionNext', location)
        await context.sync()
        return `Successfully inserted section break ${location.toLowerCase()} selection`
      },
  },

  applyStyle: {
    name: 'applyStyle',
    category: 'format',
    description: 'Apply a Word builtin paragraph style to the current selection or paragraph. '
      + 'Use this to set headings (Heading1-9), body text (Normal), special styles (Title, Subtitle, Quote, etc.). '
      + 'Builtin styles work across all languages and appear in the Word Style Gallery. '
      + 'Target "current-paragraph" to style only the paragraph containing the cursor.',
    inputSchema: {
      type: 'object',
      properties: {
        styleBuiltIn: {
          type: 'string',
          description: 'Word builtin style name.',
          enum: [
            'Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5',
            'Heading6', 'Heading7', 'Heading8', 'Heading9',
            'Title', 'Subtitle', 'Strong', 'Emphasis',
            'ListBullet', 'ListBullet2', 'ListBullet3',
            'ListNumber', 'ListNumber2', 'ListNumber3',
            'Quote', 'IntenseQuote', 'SubtleEmphasis', 'IntenseEmphasis',
            'SubtleReference', 'IntenseReference', 'BookTitle',
            'NoSpacing', 'ListParagraph',
          ],
        },
        target: {
          type: 'string',
          description: 'Where to apply: "selection" (whole selection) or "current-paragraph" (paragraph at cursor). Default: "selection".',
          enum: ['selection', 'current-paragraph'],
        },
      },
      required: ['styleBuiltIn'],
    },
    executeWord: async (context, args) => {
      const { styleBuiltIn, target = 'selection' } = args
      const selection = context.document.getSelection()

      if (target === 'current-paragraph') {
        const para = selection.paragraphs.getFirst()
        para.styleBuiltIn = styleBuiltIn as Word.BuiltInStyleName
      } else {
        selection.styleBuiltIn = styleBuiltIn as Word.BuiltInStyleName
      }

      await context.sync()
      return `Applied style "${styleBuiltIn}" to ${target}.`
    },
  },

  getSelectedTextWithFormatting: {
    name: 'getSelectedTextWithFormatting',
    category: 'read',
    description: 'R16: Get the currently selected text as Markdown, preserving all formatting '
      + '(bold, italic, underline, headings, lists, hyperlinks). '
      + 'Use this instead of getSelectedText when you need to modify text while preserving its rich formatting. '
      + 'The LLM can edit the Markdown and pass the result back to replaceSelectedText.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeWord: async (context) => {
      const selection = context.document.getSelection()
      const htmlResult = selection.getHtml()
      await context.sync()

      if (!htmlResult.value || htmlResult.value.trim() === '') {
        return 'No text selected or selection is empty.'
      }

      const markdown = htmlToMarkdown(htmlResult.value)
      return markdown || 'Selection contains no convertible content.'
    },
  },

  eval_wordjs: {
    name: 'eval_wordjs',
    category: 'write',
    description: "Execute arbitrary Office.js code within a Word.run context. Use this as an escape hatch when existing tools don't cover your use case. The code runs inside `Word.run(async (context) => { ... })` with `context` available. Return a value to get it back as the result. Always call `await context.sync()` before returning.",
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: "JavaScript code to execute. Has access to `context` (Word.RequestContext). Must be valid async code. Return a value to get it as result. Example: `const doc = context.document; const paras = doc.paragraphs; paras.load('text'); await context.sync(); return paras.items[0].text;`",
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does',
        },
      },
      required: ['code'],
    },
    executeWord: async (context, args) => {
      const { code } = args
      try {
        const result = await sandboxedEval(code, { context, Word: typeof Word !== 'undefined' ? Word : undefined })
        return JSON.stringify({ success: true, result: result ?? null }, null, 2)
      } catch (err: any) {
        return JSON.stringify({ success: false, error: err.message || String(err) }, null, 2)
      }
    },
  },
})

export function getWordToolDefinitions(): WordToolDefinition[] {
  return Object.values(wordToolDefinitions)
}

export function getWordTool(name: WordToolName): WordToolDefinition | undefined {
  return wordToolDefinitions[name]
}

export { wordToolDefinitions }
