import { type Ref, ref } from 'vue'

import type { DisplayMessage } from '@/types/chat'
import { insertFormattedResult, insertResult } from '@/api/common'
import { message as messageUtil } from '@/utils/message'
import { getOfficeHtmlCoercionType, getOutlookMailbox, isOfficeAsyncSucceeded, type OfficeAsyncResult } from '@/utils/officeOutlook'
import { insertIntoPowerPoint } from '@/utils/powerpointTools'
import { renderOfficeCommonApiHtml } from '@/utils/officeRichText'
import DOMPurify from 'dompurify'

const VERBOSE_LOGGING_ENABLED = import.meta.env.VITE_VERBOSE_LOGGING === 'true'
const verboseLog = VERBOSE_LOGGING_ENABLED ? console.info.bind(console, '[KO-INSERT]') : () => {}

interface UseOfficeInsertOptions {
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
  hostIsWord: boolean
  useWordFormatting: Ref<boolean>
  insertType: Ref<insertTypes>
  t: (key: string) => string
  shouldTreatMessageAsImage: (message: DisplayMessage) => boolean
  getMessageActionPayload: (message: DisplayMessage) => string
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

  function normalizeInsertionContent(rawContent: string): string {
    return rawContent
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .trim()
  }

  async function copyToClipboard(text: string, fallback = false) {
    if (!text.trim()) return
    const notifySuccess = () => messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))
    try {
      await navigator.clipboard.writeText(text)
      notifySuccess()
      return
    } catch (err) {
      console.warn('Clipboard API writeText failed, trying fallback:', err)
    }
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

  async function insertToDocument(content: string, type: insertTypes, richHtml?: string) {
    const normalizedContent = normalizeInsertionContent(content)
    if (!normalizedContent) return

    verboseLog('insertToDocument', {
      host: hostIsOutlook ? 'outlook' : hostIsPowerPoint ? 'powerpoint' : hostIsExcel ? 'excel' : 'word',
      type,
      contentLength: normalizedContent.length,
      lineCount: normalizedContent.split('\n').length,
      hasRichHtml: !!richHtml,
    })

    if (hostIsOutlook) {
      try {
        const mailbox = getOutlookMailbox()
        const item = mailbox?.item
        if (item?.body?.setSelectedDataAsync) {
          // F1: Use rich HTML if available (preserves images/formatting), otherwise render from markdown
          const rawHtmlBody = richHtml || renderOfficeCommonApiHtml(normalizedContent)
          const htmlBody = DOMPurify.sanitize(rawHtmlBody, { USE_PROFILES: { html: true } })
          await new Promise<void>((resolve, reject) => {
            item.body.setSelectedDataAsync!(htmlBody, { coercionType: getOfficeHtmlCoercionType() }, (result: OfficeAsyncResult) => {
              if (isOfficeAsyncSucceeded(result.status)) resolve()
              else reject(new Error(result.error?.message || 'setSelectedDataAsync failed'))
            })
          })
          messageUtil.success(t('insertedToEmail'))
        } else {
          await copyToClipboard(content, true)
        }
      } catch (err) {
        console.warn('[useOfficeInsert] Outlook error/fallback to clipboard', err)
        await copyToClipboard(content, true)
      }
      return
    }

    if (hostIsPowerPoint) {
      try {
        await insertIntoPowerPoint(normalizedContent)
        messageUtil.success(t('insertedToSlide'))
      } catch (err) {
        console.warn('[useOfficeInsert] PowerPoint error/fallback to clipboard', err)
        await copyToClipboard(normalizedContent, true)
      }
      return
    }

    if (hostIsExcel) {
      try {
        await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange()
          range.values = [[normalizedContent]]
          range.format.wrapText = true
          await ctx.sync()
        })
        messageUtil.success(t('insertedToCell'))
      } catch (err) {
        console.warn('[useOfficeInsert] Excel error/fallback to clipboard', err)
        await copyToClipboard(normalizedContent, true)
      }
      return
    }

    // Word insertion
    try {
      // F1: Use rich HTML if available (preserves images/formatting from original selection)
      if (richHtml) {
        const sanitizedHtml = DOMPurify.sanitize(richHtml, { USE_PROFILES: { html: true } })
        await Word.run(async (context) => {
          const range = context.document.getSelection()
          range.insertHtml(sanitizedHtml, type === 'newLine' ? 'After' : 'Replace')
          await context.sync()
        })
        messageUtil.success(t('inserted'))
      } else if (useWordFormatting.value) {
        await insertFormattedResult(normalizedContent, ref(type))
        messageUtil.success(t('inserted'))
      } else {
        await insertResult(normalizedContent, ref(type))
        messageUtil.success(t('inserted'))
      }
    } catch (err) {
      console.warn('[useOfficeInsert] Word error/fallback to clipboard', err)
      await copyToClipboard(normalizedContent, true)
    }
  }

  async function copyMessageToClipboard(message: DisplayMessage, fallback = false) {
    if (shouldTreatMessageAsImage(message) && message.imageSrc) {
      await copyImageToClipboard(message.imageSrc, fallback)
      return
    }
    await copyToClipboard(getMessageActionPayload(message), fallback)
  }

  async function insertMessageToDocument(message: DisplayMessage, type: insertTypes) {
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
    await insertToDocument(getMessageActionPayload(message), type, message.richHtml)
  }

  return { copyToClipboard, copyMessageToClipboard, insertToDocument, insertMessageToDocument }
}
