import { getPowerPointSelection, getPowerPointSelectionAsHtml } from '@/utils/powerpointTools'
import { getOfficeTextCoercionType, getOfficeHtmlCoercionType, getOutlookMailbox, isOfficeAsyncSucceeded, type OfficeAsyncResult } from '@/utils/officeOutlook'

export interface UseOfficeSelectionOptions {
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
}

export function useOfficeSelection(options: UseOfficeSelectionOptions) {
  const { hostIsOutlook, hostIsPowerPoint, hostIsExcel } = options

  function withTimeout<T>(promise: Promise<T>, ms: number, fallbackValue: T): Promise<T> {
    return new Promise((resolve) => {
      let resolved = false
      const timeoutId = setTimeout(() => {
        if (!resolved) {
          resolved = true
          resolve(fallbackValue)
        }
      }, ms)
      
      promise.then((val) => {
        if (!resolved) {
          resolved = true
          clearTimeout(timeoutId)
          resolve(val)
        }
      }).catch(() => {
        if (!resolved) {
          resolved = true
          clearTimeout(timeoutId)
          resolve(fallbackValue)
        }
      })
    })
  }

  const getOutlookMailBody = (): Promise<string> => {
    return withTimeout(
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item) return resolve('')
          mailbox.item.body.getAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<string>) => resolve(isOfficeAsyncSucceeded(result.status) ? (result.value || '') : ''))
        } catch { resolve('') }
      }),
      3000,
      ''
    )
  }

  const getOutlookMailBodyAsHtml = (): Promise<string> => {
    return withTimeout(
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item) return resolve('')
          const htmlType = getOfficeHtmlCoercionType()
          if (!htmlType) return resolve('')
          mailbox.item.body.getAsync(htmlType, (result: OfficeAsyncResult<string>) => resolve(isOfficeAsyncSucceeded(result.status) ? (result.value || '') : ''))
        } catch { resolve('') }
      }),
      3000,
      ''
    )
  }

  const getOutlookSelectedText = (): Promise<string> => {
    return withTimeout(
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function') return resolve('')
          mailbox.item.getSelectedDataAsync(getOfficeTextCoercionType(), (result: OfficeAsyncResult<{ data?: string }>) => resolve(isOfficeAsyncSucceeded(result.status) && result.value?.data ? result.value.data : ''))
        } catch { resolve('') }
      }),
      3000,
      ''
    )
  }

  const getOutlookSelectedHtml = (): Promise<string> => {
    return withTimeout(
      new Promise<string>((resolve) => {
        try {
          const mailbox = getOutlookMailbox()
          if (!mailbox?.item || typeof mailbox.item.getSelectedDataAsync !== 'function') return resolve('')
          const htmlType = getOfficeHtmlCoercionType()
          if (!htmlType) return resolve('')
          mailbox.item.getSelectedDataAsync(htmlType, (result: OfficeAsyncResult<{ data?: string }>) => resolve(isOfficeAsyncSucceeded(result.status) && result.value?.data ? result.value.data : ''))
        } catch { resolve('') }
      }),
      3000,
      ''
    )
  }

  async function getOfficeSelection(selectionOptions?: { includeOutlookSelectedText?: boolean, actionKey?: string }): Promise<string> {
    if (hostIsOutlook) {
      if (selectionOptions?.includeOutlookSelectedText) {
        const selected = await getOutlookSelectedText()
        if (selected) return selected
      }
      
      // Only the 'reply' quick action is allowed to fall back to reading the entire thread.
      // All other actions (proofread, formal, polite, etc.) require an active text selection.
      if (selectionOptions?.actionKey === 'reply') {
        return getOutlookMailBody()
      }
      
      return ''
    }
    if (hostIsPowerPoint) return getPowerPointSelection()
    if (hostIsExcel) {
      return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange()
        range.load('values, address')
        await ctx.sync()
        
        // Escape tabs and newlines in cell values to prevent ambiguous output
        const escapeCell = (val: any) => {
          if (val === null || val === undefined) return ''
          const str = String(val)
          if (str.includes('\t') || str.includes('\n') || str.includes('\r')) {
            return `"${str.replace(/"/g, '""')}"`
          }
          return str
        }
        
        return `[${range.address}]\n${range.values.map((row: any[]) => row.map(escapeCell).join('\t')).join('\n')}`
      })
    }
    return Word.run(async (ctx) => {
      const range = ctx.document.getSelection()
      range.load('text')
      await ctx.sync()
      return range.text
    })
  }

  /**
   * Get the selection as HTML to preserve rich content (images, formatting, etc.).
   * Falls back to plain text if HTML is not available.
   * Used by quick actions to preserve non-text elements during LLM processing.
   */
  async function getOfficeSelectionAsHtml(selectionOptions?: { includeOutlookSelectedText?: boolean, actionKey?: string }): Promise<string> {
    if (hostIsOutlook) {
      if (selectionOptions?.includeOutlookSelectedText) {
        const selectedHtml = await getOutlookSelectedHtml()
        if (selectedHtml) return selectedHtml
      }
      
      if (selectionOptions?.actionKey === 'reply') {
        const html = await getOutlookMailBodyAsHtml()
        return html || getOutlookMailBody()
      }
      
      return ''
    }
    if (hostIsExcel) {
      // Excel cells don't have meaningful HTML content
      return ''
    }
    if (hostIsPowerPoint) {
      return getPowerPointSelectionAsHtml()
    }
    // Word: get selection as HTML
    try {
      return await Word.run(async (ctx) => {
        const range = ctx.document.getSelection()
        const htmlResult = range.getHtml()
        await ctx.sync()
        return htmlResult.value || ''
      })
    } catch (err) {
      console.warn('[useOfficeSelection] Word getHtml failed', err)
      return ''
    }
  }

  return { getOfficeSelection, getOfficeSelectionAsHtml }
}
