/**
 * officeDocumentContext.ts
 *
 * Fetches lightweight document metadata for each Office host and returns it as a
 * JSON string to be injected as <doc_context> into every agent request.
 *
 * This mirrors Open_Excel's `getWorkbookMetadata()` pattern (Issue #1 of AGENT_MODE_ANALYSIS.md):
 * the model receives the workbook/document structure automatically without needing to call
 * a discovery tool first.
 */

import { executeOfficeAction } from './officeAction'

declare const Excel: any
declare const PowerPoint: any

/**
 * Excel — workbook metadata: all sheet names + usedRange dimensions, active sheet, selected range.
 */
export async function getExcelDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() =>
      Excel.run(async (context: any) => {
        const workbook = context.workbook
        const worksheets = workbook.worksheets

        worksheets.load('items/name')

        const activeSheet = worksheets.getActiveWorksheet()
        activeSheet.load('name')

        const selectedRange = workbook.getSelectedRange()
        selectedRange.load('address')

        await context.sync()

        // Batch-load usedRange for every sheet in a single sync
        const usedRanges = worksheets.items.map((sheet: any) => {
          const ur = sheet.getUsedRangeOrNullObject()
          ur.load(['rowCount', 'columnCount', 'isNullObject'])
          return ur
        })

        await context.sync()

        const sheets = worksheets.items.map((sheet: any, i: number) => {
          const ur = usedRanges[i]
          return {
            name: sheet.name,
            rows: ur.isNullObject ? 0 : ur.rowCount,
            columns: ur.isNullObject ? 0 : ur.columnCount,
          }
        })

        return JSON.stringify(
          {
            activeSheet: activeSheet.name,
            selectedRange: selectedRange.address,
            totalSheets: worksheets.items.length,
            sheets,
          },
          null,
          2,
        )
      }),
    )
  } catch {
    return ''
  }
}

/**
 * PowerPoint — presentation metadata: total slides, slide number + first text line per slide.
 */
export async function getPowerPointDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() => {
      const PPT = PowerPoint
      if (typeof PPT?.run !== 'function') return Promise.resolve('')

      return PPT.run(async (context: any) => {
        const slides = context.presentation.slides
        slides.load('items')
        await context.sync()

        // Batch-load shapes for all slides
        for (const slide of slides.items) {
          slide.shapes.load('items')
        }
        await context.sync()

        // Batch-load textFrame.textRange.text for all shapes
        for (const slide of slides.items) {
          for (const shape of slide.shapes.items) {
            try {
              shape.textFrame.textRange.load('text')
            } catch {
              // Non-text shape — skip
            }
          }
        }
        await context.sync()

        const slideInfo = slides.items.map((slide: any, i: number) => {
          let title = ''
          for (const shape of slide.shapes.items) {
            try {
              const text = (shape.textFrame?.textRange?.text || '').trim()
              if (text) {
                title = text.substring(0, 80)
                break
              }
            } catch {
              // skip
            }
          }
          return { slideNumber: i + 1, title: title || '<No text>' }
        })

        return JSON.stringify(
          {
            totalSlides: slides.items.length,
            slides: slideInfo,
          },
          null,
          2,
        )
      })
    })
  } catch {
    return ''
  }
}

/**
 * Outlook — email metadata: subject, sender, recipients, body snippet (first 300 chars).
 */
export function getOutlookDocumentContext(): Promise<string> {
  return new Promise((resolve) => {
    try {
      const Office = (window as any).Office
      const mailbox = Office?.context?.mailbox
      if (!mailbox?.item) {
        resolve('')
        return
      }

      const item = mailbox.item
      const subject = item.subject || ''
      const sender = item.sender
        ? `${item.sender.displayName || ''} <${item.sender.emailAddress || ''}>`.trim()
        : item.from
          ? `${item.from.displayName || ''} <${item.from.emailAddress || ''}>`.trim()
          : ''

      const contextObj: Record<string, any> = { subject, sender }

      // Try to read recipients (compose mode only)
      if (item.to?.getAsync) {
        item.to.getAsync((toResult: any) => {
          if (toResult.status === Office?.AsyncResultStatus?.Succeeded && Array.isArray(toResult.value)) {
            contextObj.recipients = toResult.value.map((r: any) => r.emailAddress || r.displayName || '').slice(0, 5)
          }
          readBody()
        })
      } else {
        readBody()
      }

      function readBody() {
        if (!item.body?.getAsync) {
          resolve(JSON.stringify(contextObj, null, 2))
          return
        }
        item.body.getAsync(Office?.CoercionType?.Text, (result: any) => {
          if (result.status === Office?.AsyncResultStatus?.Succeeded) {
            const bodyText = String(result.value || '')
            contextObj.bodySnippet =
              bodyText.substring(0, 300) + (bodyText.length > 300 ? '...' : '')
          }
          resolve(JSON.stringify(contextObj, null, 2))
        })
      }
    } catch {
      resolve('')
    }
  })
}

/**
 * Word — lightweight document stats: page count, word count.
 */
export async function getWordDocumentContext(): Promise<string> {
  try {
    return await executeOfficeAction(() => {
      const Word = (window as any).Word
      if (typeof Word?.run !== 'function') return Promise.resolve('')

      return Word.run(async (context: any) => {
        const props = context.document.properties
        props.load(['pageCount', 'wordCount'])
        await context.sync()

        return JSON.stringify(
          {
            pageCount: props.pageCount,
            wordCount: props.wordCount,
          },
          null,
          2,
        )
      })
    })
  } catch {
    return ''
  }
}
