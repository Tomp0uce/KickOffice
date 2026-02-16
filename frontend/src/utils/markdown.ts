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

export function renderSanitizedMarkdown(content: string): string {
  const rawMarkdown = content?.trim() ?? ''
  const unsafeHtml = markdownParser.render(rawMarkdown)

  return DOMPurify.sanitize(unsafeHtml, {
    ADD_ATTR: ['target', 'rel'],
  })
}
