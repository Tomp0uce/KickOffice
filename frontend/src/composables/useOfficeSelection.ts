import { getPowerPointSelection } from '@/utils/powerpointTools'
import { getOfficeTextCoercionType, getOutlookMailbox, isOfficeAsyncSucceeded, type OfficeAsyncResult } from '@/utils/officeOutlook'

export interface UseOfficeSelectionOptions {
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
}

export function useOfficeSelection(options: UseOfficeSelectionOptions) {
  const { hostIsOutlook, hostIsPowerPoint, hostIsExcel } = options

  const getOutlookMailBody = (): Promise<string> => {
    return Promise.race([
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item) return resolve('')
          mailbox.item.body.getAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<string>) => resolve(isOfficeAsyncSucceeded(result.status) ? (result.value || '') : ''))
        } catch { resolve('') }
      }),
      new Promise<string>(resolve => setTimeout(() => resolve(''), 3000))
    ])
  }

  const getOutlookSelectedText = (): Promise<string> => {
    return Promise.race([
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function') return resolve('')
          mailbox.item.getSelectedDataAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<{ data?: string }>) => resolve(isOfficeAsyncSucceeded(result.status) && result.value?.data ? result.value.data : ''))
        } catch { resolve('') }
      }),
      new Promise<string>(resolve => setTimeout(() => resolve(''), 3000))
    ])
  }

  async function getOfficeSelection(selectionOptions?: { includeOutlookSelectedText?: boolean }): Promise<string> {
    if (hostIsOutlook) {
      if (selectionOptions?.includeOutlookSelectedText) {
        const selected = await getOutlookSelectedText()
        if (selected) return selected
      }
      return getOutlookMailBody()
    }
    if (hostIsPowerPoint) return getPowerPointSelection()
    if (hostIsExcel) {
      return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange()
        range.load('values, address')
        await ctx.sync()
        return `[${range.address}]\n${range.values.map((row: any[]) => row.join('\t')).join('\n')}`
      })
    }
    return Word.run(async (ctx) => {
      const range = ctx.document.getSelection()
      range.load('text')
      await ctx.sync()
      return range.text
    })
  }

  return { getOfficeSelection }
}
