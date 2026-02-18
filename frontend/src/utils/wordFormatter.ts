import { Ref } from 'vue'

import { detectOfficeHost } from './hostDetection'
import { renderOfficeRichHtml, stripRichFormattingSyntax } from './officeRichText'

type FormatPartStyle =
  | 'paragraph'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'heading5'
  | 'heading6'
  | 'quote'
  | 'codeBlock'

export interface FormatPart {
  text: string
  style?: FormatPartStyle | 'bold' | 'italic' | 'code'
  hyperlink?: string
  listType?: 'bullet' | 'number'
  listLevel: number
  paragraphIndex: number
  bold?: boolean
  italic?: boolean
  code?: boolean
}

function normalizeTextInput(value: string): string {
  return value.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim()
}

function normalizeListIndentation(value: string): number {
  const normalizedIndent = value.replace(/\t/g, '  ')
  return Math.max(0, Math.floor(normalizedIndent.length / 2))
}

function parseInlineMarkdown(text: string, base: Omit<FormatPart, 'text'>): FormatPart[] {
  const inlinePattern = /(\[[^\]]+\]\([^\)\s]+\)|\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_|`[^`]+`)/g
  const parts: FormatPart[] = []
  let cursor = 0
  let match = inlinePattern.exec(text)

  while (match) {
    if (match.index > cursor) {
      parts.push({ ...base, text: text.slice(cursor, match.index), style: base.code ? 'code' : base.style })
    }

    const token = match[0]
    if (token.startsWith('[')) {
      const linkMatch = token.match(/^\[([^\]]+)\]\(([^\)\s]+)\)$/)
      if (linkMatch) {
        parts.push({ ...base, text: linkMatch[1], hyperlink: linkMatch[2], style: base.code ? 'code' : base.style })
      } else {
        parts.push({ ...base, text: token, style: base.code ? 'code' : base.style })
      }
    } else if (token.startsWith('**') || token.startsWith('__')) {
      parts.push({ ...base, text: token.slice(2, -2), style: 'bold', bold: true })
    } else if (token.startsWith('*') || token.startsWith('_')) {
      parts.push({ ...base, text: token.slice(1, -1), style: 'italic', italic: true })
    } else if (token.startsWith('`')) {
      parts.push({ ...base, text: token.slice(1, -1), style: 'code', code: true })
    }

    cursor = match.index + token.length
    match = inlinePattern.exec(text)
  }

  if (cursor < text.length) {
    parts.push({ ...base, text: text.slice(cursor), style: base.code ? 'code' : base.style })
  }

  return parts.length > 0 ? parts : [{ text, ...base }]
}

export function parseMarkdown(content: string): FormatPart[] {
  const lines = normalizeTextInput(content).split('\n')
  const parts: FormatPart[] = []
  let inCodeBlock = false
  let paragraphIndex = 0

  for (const rawLine of lines) {
    const line = rawLine ?? ''

    if (/^```/.test(line.trim())) {
      inCodeBlock = !inCodeBlock
      continue
    }

    if (inCodeBlock) {
      parts.push({
        text: line,
        style: 'codeBlock',
        code: true,
        listLevel: 0,
        paragraphIndex,
      })
      paragraphIndex += 1
      continue
    }

    if (!line.trim()) {
      parts.push({ text: '', style: 'paragraph', listLevel: 0, paragraphIndex })
      paragraphIndex += 1
      continue
    }

    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/)
    if (headingMatch) {
      const level = headingMatch[1].length as 1 | 2 | 3 | 4 | 5 | 6
      const style = `heading${level}` as FormatPartStyle
      parts.push(...parseInlineMarkdown(headingMatch[2], { listLevel: 0, paragraphIndex, style }))
      paragraphIndex += 1
      continue
    }

    const quoteMatch = line.match(/^>\s?(.*)$/)
    if (quoteMatch) {
      parts.push(...parseInlineMarkdown(quoteMatch[1], { listLevel: 0, paragraphIndex, style: 'quote' }))
      paragraphIndex += 1
      continue
    }

    const bulletMatch = line.match(/^(\s*)[-*+]\s+(.+)$/)
    if (bulletMatch) {
      parts.push(...parseInlineMarkdown(bulletMatch[2], {
        listType: 'bullet',
        listLevel: normalizeListIndentation(bulletMatch[1]),
        paragraphIndex,
        style: 'paragraph',
      }))
      paragraphIndex += 1
      continue
    }

    const numberMatch = line.match(/^(\s*)\d+[.)]\s+(.+)$/)
    if (numberMatch) {
      parts.push(...parseInlineMarkdown(numberMatch[2], {
        listType: 'number',
        listLevel: normalizeListIndentation(numberMatch[1]),
        paragraphIndex,
        style: 'paragraph',
      }))
      paragraphIndex += 1
      continue
    }

    parts.push(...parseInlineMarkdown(line, { listLevel: 0, paragraphIndex, style: 'paragraph' }))
    paragraphIndex += 1
  }

  return parts
}

function toBuiltInStyle(style?: FormatPartStyle): Word.BuiltInStyleName | undefined {
  switch (style) {
    case 'heading1':
      return 'Heading1'
    case 'heading2':
      return 'Heading2'
    case 'heading3':
      return 'Heading3'
    case 'heading4':
      return 'Heading4'
    case 'heading5':
      return 'Heading5'
    case 'heading6':
      return 'Heading6'
    case 'quote':
      return 'Quote'
    default:
      return undefined
  }
}


function insertHtmlWithCommonApi(html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve()
        } else {
          reject(result.error)
        }
      },
    )
  })
}

class WordFormatter {
  static async insertFormattedResult(result: string, insertType: Ref): Promise<void> {
    const normalizedResult = normalizeTextInput(result)
    if (!normalizedResult) return

    const host = detectOfficeHost()
    if (host !== 'Word') {
      const html = renderOfficeRichHtml(normalizedResult)
      const htmlWithLinePrefix = insertType.value === 'newLine' ? `<br/>${html}` : html
      await insertHtmlWithCommonApi(htmlWithLinePrefix)
      return
    }

    const parsed = parseMarkdown(normalizedResult)
    if (parsed.length === 0) return

    await Word.run(async context => {
      const selection = context.document.getSelection()
      let cursorRange = selection
      let firstInsertLocation: Word.InsertLocation = 'Replace'

      if (insertType.value === 'NoAction') return
      if (insertType.value === 'append') firstInsertLocation = 'End'
      if (insertType.value === 'newLine') {
        selection.insertParagraph('', 'After')
        cursorRange = selection.getRange('After')
        firstInsertLocation = 'After'
      }

      const groupedParagraphs = new Map<number, FormatPart[]>()
      parsed.forEach(part => {
        const group = groupedParagraphs.get(part.paragraphIndex) ?? []
        group.push(part)
        groupedParagraphs.set(part.paragraphIndex, group)
      })

      const orderedParagraphIndexes = [...groupedParagraphs.keys()].sort((a, b) => a - b)

      orderedParagraphIndexes.forEach((index, indexInLoop) => {
        const segments = groupedParagraphs.get(index)
        if (!segments || segments.length === 0) return

        const firstSegment = segments[0]
        const paragraph = cursorRange.insertParagraph('', indexInLoop === 0 ? firstInsertLocation : 'After')

        const paragraphStyle = toBuiltInStyle((firstSegment.style as FormatPartStyle | undefined) ?? 'paragraph')
        if (paragraphStyle) paragraph.styleBuiltIn = paragraphStyle

        if (firstSegment.listType) {
          paragraph.listItem.level = Math.max(0, firstSegment.listLevel)
        }

        if (firstSegment.style === 'codeBlock' || firstSegment.code) {
          paragraph.font.name = 'Consolas'
          paragraph.font.color = '#1F2937'
          paragraph.font.highlightColor = '#F3F4F6'
        }

        segments.forEach(segment => {
          const insertedRange = paragraph.insertText(segment.text, 'End')
          if (segment.bold) insertedRange.font.bold = true
          if (segment.italic) insertedRange.font.italic = true
          if (segment.hyperlink) insertedRange.hyperlink = segment.hyperlink
          if (segment.code) {
            insertedRange.font.name = 'Consolas'
            insertedRange.font.color = '#1F2937'
            insertedRange.font.highlightColor = '#F3F4F6'
          }
        })

        cursorRange = paragraph.getRange('End')
      })

      await context.sync()
    })
  }

  static async insertPlainResult(result: string, insertType: Ref): Promise<void> {
    const normalizedResult = stripRichFormattingSyntax(result.replace(/\r\n/g, '\n').replace(/\r/g, '\n'))

    if (!normalizedResult.trim()) return

    switch (insertType.value) {
      case 'replace':
        await Word.run(async context => {
          const range = context.document.getSelection()
          range.insertText(normalizedResult, 'Replace')
          await context.sync()
        })
        break
      case 'append':
        await Word.run(async context => {
          const range = context.document.getSelection()
          range.insertText(normalizedResult, 'End')
          await context.sync()
        })
        break
      case 'newLine':
        await Word.run(async context => {
          const range = context.document.getSelection()
          range.insertParagraph(normalizedResult, 'After')
          await context.sync()
        })
        break
      case 'NoAction':
        break
    }
  }
}

export { WordFormatter }
