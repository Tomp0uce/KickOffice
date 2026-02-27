/**
 * Rich Content Preserver
 *
 * Utility for preserving non-text elements (images, tables, SVGs, embedded objects)
 * during LLM text processing. Extracts text from HTML while replacing non-text
 * elements with unique placeholders, then reassembles the HTML after LLM processing.
 *
 * Used by quick actions (translate, polish, academic, etc.) across all Office hosts
 * to prevent data loss when processing rich content.
 */

import type { InheritedStyles } from './officeRichText'

const PLACEHOLDER_PREFIX = '{{PRESERVE_'
const PLACEHOLDER_SUFFIX = '}}'
const PLACEHOLDER_REGEX = /\{\{PRESERVE_(\d+)\}\}/g

/**
 * Tags considered as non-text elements that must be preserved as-is.
 * These are replaced with placeholders before sending text to the LLM.
 */
const PRESERVE_TAG_NAMES = new Set([
  'img', 'svg', 'table', 'video', 'audio', 'canvas',
  'object', 'embed', 'iframe', 'picture', 'figure',
])

/**
 * Check if an HTML element should be preserved (not sent to LLM).
 */
function shouldPreserveElement(el: Element): boolean {
  const tagName = el.tagName.toLowerCase()

  // Preserve known non-text tags
  if (PRESERVE_TAG_NAMES.has(tagName)) return true

  // Preserve images encoded as background-image in style
  const style = el.getAttribute('style') || ''
  if (style.includes('background-image') && style.includes('url(')) return true

  // Preserve elements with data-uri src (base64 images)
  const src = el.getAttribute('src') || ''
  if (src.startsWith('data:image/')) return true

  return false
}

export interface RichContentContext {
  /** The clean text with placeholders replacing non-text elements */
  cleanText: string
  /** Map of placeholder ID to original HTML fragment */
  fragments: Map<number, string>
  /** The original full HTML before extraction */
  originalHtml: string
  /** Whether any non-text elements were found */
  hasRichContent: boolean
  /** Extracted inline styles if present in the HTML (useful for Outlook) */
  extractedStyles?: InheritedStyles
}

/**
 * Extract text from HTML, replacing non-text elements with numbered placeholders.
 *
 * @param html - Raw HTML content (from range.getHtml() or getAsync(CoercionType.Html))
 * @returns RichContentContext with clean text and preserved fragments
 */
export function extractTextFromHtml(html: string): RichContentContext {
  const fragments = new Map<number, string>()

  if (!html || !html.includes('<')) {
    return { cleanText: html || '', fragments, originalHtml: html || '', hasRichContent: false }
  }

  try {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')

    let extractedStyles: InheritedStyles | undefined
    const firstStyledElement = doc.querySelector('[style*="font-family"], [style*="font-size"]')
    if (firstStyledElement) {
      const styleAttr = firstStyledElement.getAttribute('style') || ''
      const matchFamily = styleAttr.match(/font-family:\s*([^;]+)/i)
      const matchSize = styleAttr.match(/font-size:\s*([^;]+)/i)
      const matchColor = styleAttr.match(/(?:^|;)\s*color:\s*([^;]+)/i)
      extractedStyles = {
        fontFamily: matchFamily ? matchFamily[1].trim().replace(/['"]/g, '') : '',
        fontSize: matchSize ? matchSize[1].trim() : '',
        fontWeight: 'normal',
        fontStyle: 'normal',
        color: matchColor ? matchColor[1].trim() : '',
        marginTop: '0pt',
        marginBottom: '0pt',
      }
    }

    let counter = 0

    // Walk the DOM and replace non-text elements with placeholders
    const processNode = (node: Node): void => {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as Element
        if (shouldPreserveElement(el)) {
          const id = counter++
          const placeholder = `${PLACEHOLDER_PREFIX}${id}${PLACEHOLDER_SUFFIX}`
          fragments.set(id, el.outerHTML)
          // Replace the element with a text placeholder
          const textNode = doc.createTextNode(placeholder)
          el.parentNode?.replaceChild(textNode, el)
          return // Don't process children of preserved elements
        }
      }

      // Process children (iterate on a copy since we may modify the DOM)
      const children = Array.from(node.childNodes)
      for (const child of children) {
        processNode(child)
      }
    }

    processNode(doc.body)

    const cleanText = doc.body.textContent || ''

    return {
      cleanText,
      fragments,
      originalHtml: html,
      hasRichContent: fragments.size > 0,
      extractedStyles,
    }
  } catch (err) {
    console.warn('[RichContentPreserver] Failed to parse HTML, falling back to plain text', err)
    return { cleanText: html.replace(/<[^>]+>/g, ''), fragments, originalHtml: html, hasRichContent: false }
  }
}

/**
 * Reassemble processed text back into HTML with preserved non-text elements.
 *
 * Takes the LLM-processed text (which should contain {{PRESERVE_N}} placeholders)
 * and replaces them with the original HTML fragments.
 *
 * @param processedText - Text returned by the LLM (may contain placeholders)
 * @param context - The RichContentContext from extractTextFromHtml
 * @returns HTML string with non-text elements restored
 */
export function reassembleWithFragments(processedText: string, context: RichContentContext): string {
  if (!context.hasRichContent) return processedText

  let result = processedText

  // Replace all placeholders with their original HTML fragments
  result = result.replace(PLACEHOLDER_REGEX, (_match, idStr) => {
    const id = parseInt(idStr, 10)
    return context.fragments.get(id) || ''
  })

  return result
}

/**
 * Build the placeholder preservation instruction to append to LLM prompts.
 * Only returns the instruction if the content actually has rich elements.
 */
export function getPreservationInstruction(context: RichContentContext): string {
  if (!context.hasRichContent) return ''
  return `\nCRITICAL: The text contains preservation placeholders ({{PRESERVE_0}}, {{PRESERVE_1}}, etc.) representing images and other non-text elements. You MUST keep these placeholders EXACTLY as-is in their original positions in your output. Do NOT translate, remove, or modify them.`
}
