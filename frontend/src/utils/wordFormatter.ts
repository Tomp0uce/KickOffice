import { Ref } from 'vue'

import { renderOfficeRichHtml } from './officeRichText'

class WordFormatter {
  static async insertFormattedResult(result: string, insertType: Ref): Promise<void> {
    const normalizedResult = result.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim()
    if (!normalizedResult) return

    const html = renderOfficeRichHtml(normalizedResult)

    await Word.run(async context => {
      const range = context.document.getSelection()

      switch (insertType.value) {
        case 'replace':
          range.insertHtml(html, 'Replace')
          break
        case 'append':
          range.insertHtml(html, 'End')
          break
        case 'newLine':
          range.insertParagraph('', 'After')
          range.getRange('After').insertHtml(html, 'After')
          break
        case 'NoAction':
          break
      }

      await context.sync()
    })
  }

  static async insertPlainResult(result: string, insertType: Ref): Promise<void> {
    const normalizedResult = result.replace(/\r\n/g, '\n').replace(/\r/g, '\n')

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
