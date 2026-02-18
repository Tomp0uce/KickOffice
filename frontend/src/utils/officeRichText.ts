import DOMPurify from 'dompurify'
import MarkdownIt from 'markdown-it'

const officeMarkdownParser = new MarkdownIt({
  breaks: true,
  html: true,
  linkify: true,
  typographer: true,
})

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
  return html
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

export function renderOfficeRichHtml(content: string): string {
  const withStyleAliases = normalizeNamedStyles(content?.trim() ?? '')
  const withUnderline = normalizeUnderlineMarkdown(withStyleAliases)
  const normalizedContent = normalizeSuperAndSubScript(withUnderline)
  const unsafeHtml = officeMarkdownParser.render(normalizedContent)

  return DOMPurify.sanitize(unsafeHtml, {
    ALLOWED_TAGS: [
      'a', 'b', 'blockquote', 'br', 'code', 'del', 'div', 'em', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr', 'i', 'li',
      'ol', 'p', 'pre', 'span', 'strong', 'sub', 'sup', 'u', 'ul',
    ],
    ALLOWED_ATTR: ['href', 'rel', 'style', 'target'],
  })
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

export function stripRichFormattingSyntax(content: string): string {
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

  return normalizeListIndentationForPlainText(stripped)
}
