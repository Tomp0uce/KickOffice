import DOMPurify from 'dompurify'
import MarkdownIt from 'markdown-it'

const officeMarkdownParser = new MarkdownIt({
  breaks: true,
  html: true,
  linkify: true,
  typographer: true,
})

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

export function renderOfficeRichHtml(content: string): string {
  const withUnderline = normalizeUnderlineMarkdown(content?.trim() ?? '')
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
