import type { Ref } from 'vue'

import { message as messageUtil } from '@/utils/message'

import type { DisplayMessage, RenderSegment } from '@/types/chat'

const THINK_TAG = '<think>'
const THINK_TAG_END = '</think>'
// Cached regex for think tag removal (compiled once at module load)
const THINK_TAG_REGEX = new RegExp(`${THINK_TAG}[\\s\\S]*?${THINK_TAG_END}`, 'g')

export function useImageActions(t: (key: string) => string) {
  function splitThinkSegments(text: string): RenderSegment[] {
    if (!text) return []
    const segments: RenderSegment[] = []
    let cursor = 0
    while (cursor < text.length) {
      const start = text.indexOf(THINK_TAG, cursor)
      if (start === -1) {
        segments.push({ type: 'text', text: text.slice(cursor) })
        break
      }
      if (start > cursor) segments.push({ type: 'text', text: text.slice(cursor, start) })
      const end = text.indexOf(THINK_TAG_END, start + THINK_TAG.length)
      if (end === -1) {
        segments.push({ type: 'think', text: text.slice(start + THINK_TAG.length) })
        break
      }
      segments.push({ type: 'think', text: text.slice(start + THINK_TAG.length, end) })
      cursor = end + THINK_TAG_END.length
    }
    return segments.filter(s => s.text)
  }

  function createDisplayMessage(role: DisplayMessage['role'], content: string, imageSrc?: string): DisplayMessage {
    const id = globalThis.crypto?.randomUUID?.() || `message-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
    return { id, role, content, imageSrc }
  }

  function cleanContent(content: string): string {
    return content.replace(THINK_TAG_REGEX, '').trim()
  }

  function getMessageActionPayload(message: DisplayMessage): string {
    const cleanedText = cleanContent(message.content)
    return cleanedText || message.imageSrc || ''
  }

  function shouldTreatMessageAsImage(message: DisplayMessage): boolean {
    return !cleanContent(message.content) && !!message.imageSrc
  }

  async function copyImageToClipboard(imageSrc: string, fallback = false) {
    const notifySuccess = () => messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))
    try {
      const response = await fetch(imageSrc)
      const blob = await response.blob()
      if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
        await navigator.clipboard.write([new ClipboardItem({ [blob.type || 'image/png']: blob })])
        notifySuccess()
        return
      }
    } catch (err) {
      console.warn('Image clipboard write failed:', err)
    }

    // Fallback using execCommand for environments lacking ClipboardItem support (e.g. Office Webview)
    try {
      const img = document.createElement('img')
      img.src = imageSrc
      const div = document.createElement('div')
      div.contentEditable = 'true'
      div.appendChild(img)
      div.style.position = 'fixed'
      div.style.opacity = '0'
      document.body.appendChild(div)

      const range = document.createRange()
      range.selectNode(img)
      const selection = window.getSelection()
      selection?.removeAllRanges()
      selection?.addRange(range)

      const successful = document.execCommand('copy')

      selection?.removeAllRanges()
      document.body.removeChild(div)

      if (successful) {
        notifySuccess()
        return
      }
    } catch (err) {
      console.warn('Image clipboard execCommand fallback failed:', err)
    }

    messageUtil.error(t('imageClipboardNotSupported'))
  }

  async function insertImageToWord(imageSrc: string, type: insertTypes) {
    // Use regex to safely extract base64 payload from data URL
    const base64Payload = imageSrc.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '').trim()
    if (!base64Payload) throw new Error('Image base64 payload is empty')
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection()
      range.insertInlinePictureFromBase64(base64Payload, type === 'replace' ? 'Replace' : 'After')
      await ctx.sync()
    })
  }

  async function insertImageToPowerPoint(imageSrc: string, type: insertTypes) {
    const base64Payload = imageSrc.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '').trim()
    if (!base64Payload) throw new Error('Image base64 payload is empty')

    let insertionSuccess = false

    if (type === 'replace') {
      try {
        await new Promise<void>((resolve) => {
          Office.context.document.setSelectedDataAsync('', { coercionType: Office.CoercionType.Text }, () => resolve())
        })
      } catch (err) {
        console.warn('Failed to clear selection for replace', err)
      }
    }

    try {
      if (typeof PowerPoint !== 'undefined' && PowerPoint.run) {
        await PowerPoint.run(async (context: any) => {
          const slides = context.presentation.getSelectedSlides()
          slides.load('items')
          await context.sync()
          if (!slides.items.length) {
            throw new Error('No PowerPoint slide selected')
          }
          const targetSlide = type === 'append' ? slides.items[slides.items.length - 1] : slides.items[0]
          targetSlide.shapes.addImage(base64Payload)
          await context.sync()
        })
        insertionSuccess = true
      }
    } catch (err) {
      console.warn('PowerPoint.run addImage failed, falling back to Common API', err)
    }

    if (!insertionSuccess) {
      await new Promise<void>((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(base64Payload, { coercionType: Office.CoercionType.Image }, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message))
          } else {
            resolve()
          }
        })
      })
    }
  }

  function historyWithSegments(history: Ref<DisplayMessage[]>) {
    return history.value.map(message => ({ message, key: message.id, segments: splitThinkSegments(message.content) }))
  }

  return {
    createDisplayMessage,
    splitThinkSegments,
    cleanContent,
    getMessageActionPayload,
    shouldTreatMessageAsImage,
    copyImageToClipboard,
    insertImageToWord,
    insertImageToPowerPoint,
    historyWithSegments,
  }
}
