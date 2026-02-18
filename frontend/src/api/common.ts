import { Ref } from 'vue'

import { stripRichFormattingSyntax } from '@/utils/officeRichText'
import { WordFormatter } from '@/utils/wordFormatter'

export async function insertResult(result: string, insertType: Ref): Promise<void> {
  const normalizedResult = stripRichFormattingSyntax(result
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n'))

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

export async function insertFormattedResult(result: string, insertType: Ref): Promise<void> {
  try {
    await WordFormatter.insertFormattedResult(result, insertType)
  } catch (error) {
    console.warn('Formatted insertion failed, falling back to plain text:', error)
    await WordFormatter.insertPlainResult(result, insertType)
  }
}
