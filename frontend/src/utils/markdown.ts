import DOMPurify from 'dompurify'
import MarkdownIt from 'markdown-it'

const markdownParser = new MarkdownIt({
  breaks: true,
  html: false,
  linkify: true,
  typographer: true,
})

const defaultLinkRender = markdownParser.renderer.rules.link_open
  ?? ((tokens: any[], idx: number, options: any, _env: any, self: any) => self.renderToken(tokens, idx, options))

markdownParser.renderer.rules.link_open = (tokens: any[], idx: number, options: any, env: any, self: any) => {
  const token = tokens[idx]

  token.attrSet('target', '_blank')
  token.attrSet('rel', 'noopener noreferrer')

  return defaultLinkRender(tokens, idx, options, env, self)
}

// Strict allowlist of safe HTML tags for markdown rendering
const ALLOWED_TAGS = [
  'p', 'br', 'hr',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'ul', 'ol', 'li',
  'blockquote', 'pre', 'code',
  'a', 'strong', 'em', 'b', 'i', 's', 'del', 'ins', 'mark',
  'table', 'thead', 'tbody', 'tr', 'th', 'td',
  'span', 'div',
]

// Strict allowlist of safe HTML attributes
const ALLOWED_ATTR = [
  'href', 'target', 'rel', 'title',
  'class', 'id',
  'colspan', 'rowspan',
]

export function renderSanitizedMarkdown(content: string): string {
  const rawMarkdown = content?.trim() ?? ''
  const unsafeHtml = markdownParser.render(rawMarkdown)

  return DOMPurify.sanitize(unsafeHtml, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOW_DATA_ATTR: false,
    ALLOW_ARIA_ATTR: false,
  })
}
