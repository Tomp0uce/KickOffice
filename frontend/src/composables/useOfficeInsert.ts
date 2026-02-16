import type { Ref } from 'vue'

import { insertFormattedResult, insertResult } from '@/api/common'
import { message as messageUtil } from '@/utils/message'
import { getOfficeTextCoercionType, getOutlookMailbox, isOfficeAsyncSucceeded, type OfficeAsyncResult } from '@/utils/officeOutlook'
import { insertIntoPowerPoint } from '@/utils/powerpointTools'

interface UseOfficeInsertOptions {
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
  hostIsWord: boolean
  useWordFormatting: Ref<boolean>
  insertType: Ref<insertTypes>
  t: (key: string) => string
  shouldTreatMessageAsImage: (message: any) => boolean
  getMessageActionPayload: (message: any) => string
  copyImageToClipboard: (imageSrc: string, fallback?: boolean) => Promise<void>
  insertImageToWord: (imageSrc: string, type: insertTypes) => Promise<void>
  insertImageToPowerPoint: (imageSrc: string, type: insertTypes) => Promise<void>
}

export function useOfficeInsert(options: UseOfficeInsertOptions) {
  const {
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
    hostIsWord,
    useWordFormatting,
    insertType,
    t,
    shouldTreatMessageAsImage,
    getMessageActionPayload,
    copyImageToClipboard,
    insertImageToWord,
    insertImageToPowerPoint,
  } = options

  async function copyToClipboard(text: string, fallback = false) {
    if (!text.trim()) return
    const notifySuccess = () => messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))
    try {
      await navigator.clipboard.writeText(text)
      notifySuccess()
      return
    } catch {}
    try {
      const textarea = document.createElement('textarea')
      textarea.value = text
      textarea.setAttribute('readonly', '')
      textarea.style.position = 'fixed'
      textarea.style.opacity = '0'
      document.body.appendChild(textarea)
      textarea.select()
      const copied = document.execCommand('copy')
      document.body.removeChild(textarea)
      if (copied) notifySuccess()
      else messageUtil.error(t('failedToInsert'))
    } catch {
      messageUtil.error(t('failedToInsert'))
    }
  }

  async function insertToDocument(content: string, type: insertTypes) {
    if (!content.trim()) return
    if (hostIsOutlook) {
      try {
        const mailbox = getOutlookMailbox()
        const item = mailbox?.item
        if (item?.body?.setAsync) {
          await new Promise<void>((resolve, reject) => {
            item.body.setAsync(content, { coercionType: getOfficeTextCoercionType() }, (result: OfficeAsyncResult) => {
              if (isOfficeAsyncSucceeded(result.status)) resolve()
              else reject(new Error(result.error?.message || 'setAsync failed'))
            })
          })
          messageUtil.success(t('insertedToEmail'))
        } else {
          await copyToClipboard(content, true)
        }
      } catch {
        await copyToClipboard(content, true)
      }
      return
    }

    if (hostIsPowerPoint) {
      try {
        await insertIntoPowerPoint(content)
        messageUtil.success(t('insertedToSlide'))
      } catch {
        await copyToClipboard(content, true)
      }
      return
    }

    if (hostIsExcel) {
      try {
        await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange()
          range.values = [[content]]
          await ctx.sync()
        })
        messageUtil.success(t('insertedToCell'))
      } catch {
        await copyToClipboard(content, true)
      }
      return
    }

    try {
      insertType.value = type
      if (useWordFormatting.value) await insertFormattedResult(content, insertType)
      else await insertResult(content, insertType)
      messageUtil.success(t('inserted'))
    } catch {
      await copyToClipboard(content, true)
    }
  }

  async function copyMessageToClipboard(message: any, fallback = false) {
    if (shouldTreatMessageAsImage(message) && message.imageSrc) {
      await copyImageToClipboard(message.imageSrc, fallback)
      return
    }
    await copyToClipboard(getMessageActionPayload(message), fallback)
  }

  async function insertMessageToDocument(message: any, type: insertTypes) {
    if (shouldTreatMessageAsImage(message) && message.imageSrc) {
      if (hostIsWord) {
        try {
          await insertImageToWord(message.imageSrc, type)
          messageUtil.success(t('inserted'))
        } catch {
          await copyImageToClipboard(message.imageSrc, true)
        }
        return
      }
      if (hostIsPowerPoint) {
        try {
          await insertImageToPowerPoint(message.imageSrc, type)
          messageUtil.success(t('insertedToSlide'))
        } catch {
          await copyImageToClipboard(message.imageSrc, true)
        }
        return
      }
      if (hostIsExcel) {
        messageUtil.info(t('imageInsertExcelNotSupported'))
        return
      }
      await copyImageToClipboard(message.imageSrc, true)
      messageUtil.info(t('imageInsertWordOnly'))
      return
    }
    await insertToDocument(getMessageActionPayload(message), type)
  }

  return { copyToClipboard, copyMessageToClipboard, insertToDocument, insertMessageToDocument }
}
