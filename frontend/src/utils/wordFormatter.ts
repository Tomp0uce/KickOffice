import { Ref } from 'vue'

import { renderOfficeCommonApiHtml, stripRichFormattingSyntax } from './officeRichText'

type WordRangeFontSnapshot = {
  name?: string
  size?: number
  color?: string
  highlightColor?: string
}

function isValidCssColor(value: unknown): value is string {
  if (typeof value !== 'string') return false
  const normalized = value.trim().toLowerCase()
  return normalized.length > 0 && normalized !== 'none' && normalized !== 'nohighlight'
}

function toInlineCss(font: WordRangeFontSnapshot): string {
  const css: string[] = []
  if (font.name && typeof font.name === 'string') css.push(`font-family: ${font.name.replace(/"/g, '\\"')};`)
  if (typeof font.size === 'number' && Number.isFinite(font.size) && font.size > 0) css.push(`font-size: ${font.size}pt;`)
  if (isValidCssColor(font.color)) css.push(`color: ${font.color};`)
  if (isValidCssColor(font.highlightColor)) css.push(`background-color: ${font.highlightColor};`)
  return css.join(' ')
}

class WordFormatter {
  static async insertFormattedResult(result: string, insertType: Ref): Promise<void> {
    const normalizedResult = result.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim()
    if (!normalizedResult) return

    const html = renderOfficeCommonApiHtml(normalizedResult)

    await Word.run(async context => {
      const range = context.document.getSelection()
      range.load('font/name,font/size,font/color,font/highlightColor')
      await context.sync()

      const baseStyle = toInlineCss({
        name: range.font?.name,
        size: range.font?.size,
        color: range.font?.color,
        highlightColor: range.font?.highlightColor,
      })

      const scopedHtml = baseStyle ? `<div style="${baseStyle}">${html}</div>` : html

      switch (insertType.value) {
        case 'replace':
          range.insertHtml(scopedHtml, 'Replace')
          break
        case 'append':
          range.insertHtml(scopedHtml, 'End')
          break
        case 'newLine':
          range.insertParagraph('', 'After')
          range.getRange('After').insertHtml(scopedHtml, 'After')
          break
        case 'NoAction':
          break
      }

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
