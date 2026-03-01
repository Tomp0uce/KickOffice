import DOMPurify from 'dompurify'
import MarkdownIt from 'markdown-it'
import markdownItDeflist from 'markdown-it-deflist'
import markdownItFootnote from 'markdown-it-footnote'
import markdownItTaskLists from 'markdown-it-task-lists'
import TurndownService from 'turndown'

// R16 — Reusable TurndownService instance configured for Office markdown output
const turndownService = new TurndownService({ headingStyle: 'atx', bulletListMarker: '-', codeBlockStyle: 'fenced' })
turndownService.addRule('underline', {
  filter: ['u', 'ins'],
  replacement: (content) => `__${content}__`,
})
turndownService.addRule('strikethrough', {
  filter: ['del', 's', 'strike'],
  replacement: (content) => `~~${content}~~`,
})

// Turndown rule to convert color styles into custom markup [color:#HEX]text[/color]
turndownService.addRule('color', {
  filter: (node: HTMLElement) => {
    return (
      node.nodeType === 1 &&
      (node.getAttribute('style')?.includes('color:') || node.hasAttribute('color'))
    )
  },
  replacement: (content: string, node: any) => {
    let color = '';
    const style = (node as HTMLElement).getAttribute('style') || '';
    const match = style.match(/color:\s*([^;"]+)/i); // Added " to avoid matching until end of style
    if (match) {
      color = match[1].trim();
    } else {
      color = (node as HTMLElement).getAttribute('color') || '';
    }
    // Only wrap if it's an actual hex or named color and content is not empty
    if (color && content.trim()) {
      return `[color:${color}]${content}[/color]`;
    }
    return content;
  }
})

export function htmlToMarkdown(html: string): string {
  if (!html) return ''
  
  try {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')
    
    // Pre-process inline styles to semantic tags for better Markdown conversion
    const styledElements = doc.querySelectorAll('[style]')
    styledElements.forEach((node) => {
      const style = node.getAttribute('style') || ''
      let wrapStart = ''
      let wrapEnd = ''
      
      if (/font-weight:\s*(bold|[7-9]\d{2})/i.test(style)) { wrapStart += '<b>'; wrapEnd = '</b>' + wrapEnd }
      if (/font-style:\s*italic/i.test(style)) { wrapStart += '<i>'; wrapEnd = '</i>' + wrapEnd }
      if (/text-decoration(?:-line)?:\s*underline/i.test(style)) { wrapStart += '<u>'; wrapEnd = '</u>' + wrapEnd }
      if (/text-decoration(?:-line)?:\s*line-through/i.test(style)) { wrapStart += '<s>'; wrapEnd = '</s>' + wrapEnd }
      
      if (wrapStart && node.innerHTML) {
        node.innerHTML = wrapStart + node.innerHTML + wrapEnd
      }
    })
    
    return turndownService.turndown(doc.body.innerHTML)
  } catch (err) {
    console.warn('[htmlToMarkdown] DOM parsing failed, falling back to raw html', err)
    return turndownService.turndown(html)
  }
}

const officeMarkdownParser = new MarkdownIt({
  breaks: true,
  html: true,
  linkify: true,
  typographer: true,
})
  .use(markdownItTaskLists, { enabled: true })
  .use(markdownItDeflist)
  .use(markdownItFootnote)

/**
 * R1 — Font and spacing properties captured from the insertion point in Word.
 * Used to inject document-level styles into generated HTML so inserted content
 * visually matches the surrounding text.
 */
export interface InheritedStyles {
  fontFamily: string
  fontSize: string
  fontWeight: string
  fontStyle: string
  color: string
  marginTop: string
  marginBottom: string
}

/**
 * R1 — Inject document-level CSS into bare <p> and <li> elements so that
 * inserted content inherits the font and spacing of the insertion context.
 * Headings and already-styled elements are left untouched.
 */
export function applyInheritedStyles(html: string, styles: InheritedStyles): string {
  const cssRules: string[] = []
  if (styles.fontFamily) cssRules.push(`font-family:'${styles.fontFamily}'`)
  if (styles.fontSize && styles.fontSize !== '0pt') cssRules.push(`font-size:${styles.fontSize}`)
  if (styles.color) cssRules.push(`color:${styles.color}`)

  if (cssRules.length === 0) return html

  const fontCss = cssRules.join('; ')
  const marginCss = `margin:${styles.marginTop || '0pt'} 0 ${styles.marginBottom || '0pt'} 0`

  return html
    .replace(/<p>/gi, `<p style="${fontCss}; ${marginCss}">`)
    .replace(/<li>/gi, `<li style="${fontCss}">`)
}

type StyleDefinition = {
  fontSize?: string
  fontWeight?: string
  fontStyle?: string
  marginBottom?: string
}

const STYLE_ALIASES: Record<string, StyleDefinition> = {
  // Word default-ish families (EN + FR variants)
  title: { fontSize: '2em', fontWeight: '700', marginBottom: '10px' },
  titre: { fontSize: '2em', fontWeight: '700', marginBottom: '10px' },
  subtitle: { fontSize: '1.3em', fontStyle: 'italic', marginBottom: '8px' },
  sous_titre: { fontSize: '1.3em', fontStyle: 'italic', marginBottom: '8px' },
  heading1: { fontSize: '2em', fontWeight: '700', marginBottom: '8px' },
  heading2: { fontSize: '1.5em', fontWeight: '700', marginBottom: '6px' },
  heading3: { fontSize: '1.17em', fontWeight: '700', marginBottom: '6px' },
  heading4: { fontSize: '1.05em', fontWeight: '700', marginBottom: '4px' },
  heading5: { fontSize: '0.95em', fontWeight: '700', marginBottom: '4px' },
  heading6: { fontSize: '0.9em', fontWeight: '700', marginBottom: '4px' },
  heading7: { fontSize: '0.88em', fontWeight: '700', marginBottom: '4px' },
  heading8: { fontSize: '0.86em', fontWeight: '700', marginBottom: '4px' },
  heading9: { fontSize: '0.84em', fontWeight: '700', marginBottom: '4px' },
  titre1: { fontSize: '2em', fontWeight: '700', marginBottom: '8px' },
  titre2: { fontSize: '1.5em', fontWeight: '700', marginBottom: '6px' },
  titre3: { fontSize: '1.17em', fontWeight: '700', marginBottom: '6px' },
  titre4: { fontSize: '1.05em', fontWeight: '700', marginBottom: '4px' },
  titre5: { fontSize: '0.95em', fontWeight: '700', marginBottom: '4px' },
  titre6: { fontSize: '0.9em', fontWeight: '700', marginBottom: '4px' },
  titre7: { fontSize: '0.88em', fontWeight: '700', marginBottom: '4px' },
  titre8: { fontSize: '0.86em', fontWeight: '700', marginBottom: '4px' },
  titre9: { fontSize: '0.84em', fontWeight: '700', marginBottom: '4px' },
  normal: { marginBottom: '6px' },
  paragraph: { marginBottom: '6px' },
  quote: { fontStyle: 'italic', marginBottom: '6px' },
  citation: { fontStyle: 'italic', marginBottom: '6px' },
  intense_quote: { fontStyle: 'italic', fontWeight: '700', marginBottom: '6px' },
  list_paragraph: { marginBottom: '2px' },
  no_spacing: { marginBottom: '0' },
}

function normalizeStyleKey(styleName: string): string {
  return styleName
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_')
}

function styleDefinitionToInlineCss(definition: StyleDefinition): string {
  const rules: string[] = []
  if (definition.fontSize) rules.push(`font-size:${definition.fontSize}`)
  if (definition.fontWeight) rules.push(`font-weight:${definition.fontWeight}`)
  if (definition.fontStyle) rules.push(`font-style:${definition.fontStyle}`)
  if (definition.marginBottom !== undefined) rules.push(`margin:0 0 ${definition.marginBottom} 0`)
  return rules.join(';')
}

/**
 * R5 — Insert a non-breaking space on otherwise blank lines so the Markdown
 * parser emits an explicit empty paragraph instead of silently collapsing
 * sequences of 3+ newlines into a single paragraph break.
 */
function preserveMultipleLineBreaks(content: string): string {
  return content.replace(/\n[ \t]*\n[ \t]*\n/g, '\n\n&nbsp;\n\n')
}

/**
 * R4 — Split <br> tags that appear inside <li> or <p> elements into
 * separate child paragraphs. Prevents Word and PowerPoint from rendering a
 * line-break-inside-a-list-item as a plain newline that visually breaks
 * bullet alignment.
 *
 * Works in all Office add-in runtimes (browser-based DOMParser).
 */
function splitBrInListItems(html: string): string {
  try {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')

    // Split <br> inside <li> into sibling <p> elements
    doc.querySelectorAll('li').forEach((li) => {
      if (!/<br\s*\/?>/i.test(li.innerHTML)) return
      const segments = li.innerHTML.split(/<br\s*\/?>/i).map(s => s.trim()).filter(Boolean)
      if (segments.length <= 1) return
      li.innerHTML = '<p>' + segments.join('</p><p>') + '</p>'
    })

    // Split <br> inside <p> into sibling <p> elements
    // Iterate over a snapshot to avoid live-collection issues
    const paragraphs = Array.from(doc.querySelectorAll('p'))
    for (const p of paragraphs) {
      if (!/<br\s*\/?>/i.test(p.innerHTML)) continue
      const segments = p.innerHTML.split(/<br\s*\/?>/i).map(s => s.trim()).filter(Boolean)
      if (segments.length <= 1) continue
      const parent = p.parentNode
      if (!parent) continue
      const style = p.getAttribute('style') ?? ''
      for (const seg of segments) {
        const newP = document.createElement('p')
        newP.innerHTML = seg
        if (style) newP.setAttribute('style', style)
        parent.insertBefore(newP, p)
      }
      parent.removeChild(p)
    }

    return doc.body.innerHTML
  } catch {
    return html
  }
}

/**
 * R19 — Post-process markdown-it-footnote HTML output for cleaner Office rendering.
 *
 * The plugin generates a <section class="footnotes"> block with back-reference links.
 * Office hosts don't render the `<section>` element natively, so we convert it to:
 * - A horizontal-rule separator paragraph
 * - A numbered list of footnote texts (back-reference links stripped)
 *
 * Inline footnote references (<sup class="footnote-ref">) are kept as superscript
 * numbers so they remain visible in the body text.
 */
function processFootnotes(html: string): string {
  if (!html.includes('class="footnotes"')) return html

  try {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')
    const section = doc.querySelector('section.footnotes')
    if (!section) return html

    // Build replacement: separator + numbered footnote list
    const items = Array.from(section.querySelectorAll('li.footnote-item'))
    if (items.length === 0) {
      section.remove()
      return doc.body.innerHTML
    }

    const separator = doc.createElement('p')
    separator.setAttribute('style', 'border-bottom:1px solid #999; margin:12px 0;')
    separator.innerHTML = '&nbsp;'

    const ol = doc.createElement('ol')
    ol.setAttribute('style', 'font-size:0.85em; color:#555; margin:4px 0; padding-left:1.5em;')

    for (const item of items) {
      // Strip back-reference links (<a class="footnote-backref">)
      item.querySelectorAll('a.footnote-backref').forEach(a => a.remove())
      const li = doc.createElement('li')
      li.innerHTML = item.innerHTML.trim()
      ol.appendChild(li)
    }

    section.replaceWith(separator, ol)
    return doc.body.innerHTML
  } catch {
    return html
  }
}

function normalizeUnderlineMarkdown(rawContent: string): string {
  // Many model prompts use __text__ to ask for underline. Convert it to <u>...</u>
  // before markdown parsing so Office hosts render the expected style.
  return rawContent.replace(/(^|[^*])__(.+?)__(?!\*)/g, '$1<u>$2</u>')
}

function normalizeSuperAndSubScript(rawContent: string): string {
  return rawContent
    // ^^texte^^ or ^texte^ => superscript
    .replace(/\^\^(.+?)\^\^/g, '<sup>$1</sup>')
    .replace(/\^([^\^\n]+?)\^/g, '<sup>$1</sup>')
    // ~texte~ => subscript (while preserving markdown strikethrough ~~texte~~)
    .replace(/(^|[^~])~([^~\n]+?)~(?=[^~]|$)/g, '$1<sub>$2</sub>')
}

function normalizeNamedStyles(rawContent: string): string {
  const byTag = rawContent
    .replace(/<\s*(title|titre)\s*>([\s\S]*?)<\s*\/\s*(title|titre)\s*>/gi, '<h1>$2</h1>')
    .replace(/<\s*(subtitle|sous[-\s]?titre)\s*>([\s\S]*?)<\s*\/\s*(subtitle|sous[-\s]?titre)\s*>/gi, '<h2><em>$2</em></h2>')
    .replace(/<\s*(normal|paragraph|style-normal)\s*>([\s\S]*?)<\s*\/\s*(normal|paragraph|style-normal)\s*>/gi, '<p>$2</p>')
    .replace(/<\s*(quote|citation)\s*>([\s\S]*?)<\s*\/\s*(quote|citation)\s*>/gi, '<blockquote>$2</blockquote>')
    .replace(/<\s*(intense[-\s]?quote)\s*>([\s\S]*?)<\s*\/\s*(intense[-\s]?quote)\s*>/gi, '<blockquote><strong><em>$2</em></strong></blockquote>')
    .replace(/<\s*(heading|titre)\s*([1-9])\s*>([\s\S]*?)<\s*\/\s*(heading|titre)\s*\2\s*>/gi, (_, _prefix, level, text) => {
      const htmlLevel = Math.min(6, Number(level))
      return `<h${htmlLevel}>${text}</h${htmlLevel}>`
    })

  const byPrefix = byTag
    .replace(/^\s*(title|titre)\s*:\s+(.+)$/gim, '# $2')
    .replace(/^\s*(subtitle|sous[-\s]?titre)\s*:\s+(.+)$/gim, '## *$2*')
    .replace(/^\s*(heading|titre)\s*([1-9])\s*:\s+(.+)$/gim, (_m, _h, level, text) => `${'#'.repeat(Math.min(6, Number(level)))} ${text}`)
    .replace(/^\s*(normal|paragraph|style\s*normal)\s*:\s+(.+)$/gim, '$2')

  // Explicit wrapper form: [style:Heading 3]text[/style]
  return byPrefix.replace(/\[style\s*:\s*([^\]]+)\]([\s\S]*?)\[\/style\]/gi, (_m, styleName, text) => {
    const key = normalizeStyleKey(styleName)
    const style = STYLE_ALIASES[key]
    if (!style) return text
    const css = styleDefinitionToInlineCss(style)
    return `<p style="${css}">${text}</p>`
  })
}

function applyOfficeBlockStyles(html: string): string {
  // Two-pass code styling: protect <code> inside <pre> from inline-code rules
  // by temporarily marking it, then style standalone <code> separately.
  const withPreCode = html
    .replace(/<pre>/gi, '<pre style="font-family:Consolas,\'Courier New\',monospace; font-size:10pt; background:#f4f4f4; padding:8px; margin:6px 0; border-left:3px solid #ccc;">')
    .replace(/(<pre[^>]*>)(<code>)/gi, '$1<code data-pre="1">')

  const withCode = withPreCode
    .replace(/<code(?! data-pre)/gi, '<code style="font-family:Consolas,\'Courier New\',monospace; font-size:0.9em; background:#f0f0f0; padding:1px 4px;">')
    .replace(/ data-pre="1"/gi, '')
    
  // Transform [color:HEX]text[/color] syntax into <span style="color:HEX">text</span>
  const withColor = withCode.replace(/\[color:\s*([^\]]+)\]([\s\S]*?)\[\/color\]/gi, '<span style="color:$1">$2</span>')

  return withColor
    .replace(/<hr\s*\/?>/gi, '<p style="border-bottom:1px solid #999; margin:8px 0;">&nbsp;</p>')
    .replace(/<h1>/gi, '<h1 style="margin:0 0 8px 0; font-size:2em; font-weight:700;">')
    .replace(/<h2>/gi, '<h2 style="margin:0 0 6px 0; font-size:1.5em; font-weight:700;">')
    .replace(/<h3>/gi, '<h3 style="margin:0 0 6px 0; font-size:1.17em; font-weight:700;">')
    .replace(/<h4>/gi, '<h4 style="margin:0 0 4px 0; font-size:1.05em; font-weight:700;">')
    .replace(/<h5>/gi, '<h5 style="margin:0 0 4px 0; font-size:0.95em; font-weight:700;">')
    .replace(/<h6>/gi, '<h6 style="margin:0 0 4px 0; font-size:0.9em; font-weight:700;">')
    .replace(/<p>/gi, '<p style="margin:0 0 6px 0;">')
    .replace(/<ul>/gi, '<ul style="margin:0 0 6px 0; padding-left:1.25em;">')
    .replace(/<ol>/gi, '<ol style="margin:0 0 6px 0; padding-left:1.25em;">')
    .replace(/<li>/gi, '<li style="margin:0 0 2px 0;">')
}

function normalizeNumberedListItem(marker: string): string {
  return marker.endsWith(')') ? `${marker.slice(0, -1)}.` : marker
}

function normalizeListIndentationForPlainText(content: string): string {
  return content
    .split(/\r?\n/)
    .map((line) => {
      const match = line.match(/^(\s*)([-*+]|\d+[.)])\s+(.+)$/)
      if (!match) return line

      const [, leading, marker, itemText] = match
      const indentLevel = Math.floor(leading.replace(/\t/g, '  ').length / 2)
      const prefix = /^\d+[.)]$/.test(marker) ? normalizeNumberedListItem(marker) : '•'
      return `${'\t'.repeat(indentLevel)}${prefix} ${itemText}`
    })
    .join('\n')
}

/**
 * Strips markdown list markers (-, *, +, 1.) from the beginning of lines
 * while preserving the hierarchical indentation.
 * Useful for PowerPoint where shapes often natively apply their own bullet points.
 */
export function stripMarkdownListMarkers(content: string): string {
  return content
    .split(/\r?\n/)
    .map((line) => {
      const match = line.match(/^(\s*)(?:[-*+]|\d+[.)])\s+(.+)$/)
      if (!match) return line

      const [, leading, itemText] = match
      const indentLevel = Math.floor(leading.replace(/\t/g, '  ').length / 2)
      return `${'\t'.repeat(indentLevel)}${itemText}`
    })
    .join('\n')
}

export function renderOfficeRichHtml(content: string): string {
  const withStyleAliases = normalizeNamedStyles(content?.trim() ?? '')
  const withPreservedBreaks = preserveMultipleLineBreaks(withStyleAliases)  // R5
  const withUnderline = normalizeUnderlineMarkdown(withPreservedBreaks)
  const normalizedContent = normalizeSuperAndSubScript(withUnderline)
  const unsafeHtml = officeMarkdownParser.render(normalizedContent)

  const sanitized = DOMPurify.sanitize(unsafeHtml, {
    ALLOWED_TAGS: [
      'a', 'b', 'blockquote', 'br', 'code', 'dd', 'del', 'div', 'dl', 'dt', 'em',
      'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr', 'i',
      'input', 'li', 'ol', 'p', 'pre', 'section', 'span', 'strong', 'sub', 'sup', 'u', 'ul',
    ],
    ALLOWED_ATTR: ['checked', 'class', 'disabled', 'href', 'id', 'rel', 'target', 'type', 'style'],
  })

  const withFootnotes = processFootnotes(sanitized)  // R19
  return splitBrInListItems(withFootnotes)  // R4
}

/**
 * Keep one coherent markdown->HTML pipeline across Word/Outlook/PowerPoint.
 * Common API hosts can be stricter, so we keep semantic tags but normalize block styles.
 */
export function renderOfficeCommonApiHtml(content: string): string {
  const richHtml = renderOfficeRichHtml(content)
  const styledHtml = applyOfficeBlockStyles(richHtml)

  return styledHtml.trim() || content
}

export function stripRichFormattingSyntax(content: string, stripListMarkers = false): string {
  const withLineBreaks = content
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/\s*(p|div|li|h1|h2|h3|h4|h5|h6|blockquote|pre|ul|ol)\s*>/gi, '\n')

  const stripped = withLineBreaks
    .replace(/<[^>]+>/g, '')
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/__(.+?)__/g, '$1')
    .replace(/\*(.+?)\*/g, '$1')
    .replace(/_(.+?)_/g, '$1')
    .replace(/~~(.+?)~~/g, '$1')
    .replace(/\^\^(.+?)\^\^/g, '$1')
    .replace(/\^([^\^\n]+?)\^/g, '$1')
    .replace(/(^|[^~])~([^~\n]+?)~(?=[^~]|$)/g, '$1$2')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

  return stripListMarkers ? stripMarkdownListMarkers(stripped) : normalizeListIndentationForPlainText(stripped)
}
