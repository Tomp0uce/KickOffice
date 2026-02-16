import DOMPurify from 'dompurify'
import MarkdownIt from 'markdown-it'

type MarkdownItRendererRule = NonNullable<MarkdownIt.Renderer['rules']['link_open']>

const markdownParser = new MarkdownIt({
  breaks: true,
  html: false,
  linkify: true,
  typographer: true,
})

const defaultLinkRender: MarkdownItRendererRule = markdownParser.renderer.rules.link_open
  ?? ((tokens, idx, options, _env, self) => self.renderToken(tokens, idx, options))

markdownParser.renderer.rules.link_open = (tokens, idx, options, env, self) => {
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
