import { Ref } from 'vue'

import { detectOfficeHost } from './hostDetection'
import { renderOfficeRichHtml, stripRichFormattingSyntax } from './officeRichText'

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
    if (!result || !result.trim()) return

    const html = renderOfficeRichHtml(result)
    const host = detectOfficeHost()

    if (host !== 'Word') {
      const htmlWithLinePrefix = insertType.value === 'newLine' ? `<br/>${html}` : html
      await insertHtmlWithCommonApi(htmlWithLinePrefix)
      return
    }

    await Word.run(async context => {
      const selection = context.document.getSelection()
      if (insertType.value === 'NoAction') return

      let location: any = 'Replace'
      let finalHtml = html

      if (insertType.value === 'append') {
        location = 'End'
      } else if (insertType.value === 'newLine') {
        location = 'After'
        finalHtml = `<br/>${html}`
      }

      selection.insertHtml(finalHtml, location)
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
