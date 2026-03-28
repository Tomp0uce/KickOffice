import { sanitizeHtml, createBaseMarkdownParser } from './markdown';

const markdownParser = createBaseMarkdownParser(false);

// MarkdownIt renderer rules use `any` types per the library's own type definitions (env: any).
// The types below match markdown-it's RenderRule signature exactly.
/* eslint-disable @typescript-eslint/no-explicit-any */
const defaultLinkRender =
  markdownParser.renderer.rules.link_open ??
  ((tokens: any[], idx: number, options: any, _env: any, self: any) =>
    self.renderToken(tokens, idx, options));

markdownParser.renderer.rules.link_open = (
  tokens: any[],
  idx: number,
  options: any,
  env: any,
  self: any,
) => {
  const token = tokens[idx];

  token.attrSet('target', '_blank');
  token.attrSet('rel', 'noopener noreferrer');

  return defaultLinkRender(tokens, idx, options, env, self);
};
/* eslint-enable @typescript-eslint/no-explicit-any */

// Strict allowlist of safe HTML tags for markdown rendering
const ALLOWED_TAGS = [
  'p',
  'br',
  'hr',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'ul',
  'ol',
  'li',
  'blockquote',
  'pre',
  'code',
  'a',
  'strong',
  'em',
  'b',
  'i',
  's',
  'del',
  'ins',
  'mark',
  'table',
  'thead',
  'tbody',
  'tr',
  'th',
  'td',
  'span',
  'div',
];

// Strict allowlist of safe HTML attributes
const ALLOWED_ATTR = ['href', 'target', 'rel', 'title', 'class', 'id', 'colspan', 'rowspan'];

export function renderSanitizedMarkdown(content: string): string {
  const rawMarkdown = content?.trim() ?? '';
  const unsafeHtml = markdownParser.render(rawMarkdown);

  return sanitizeHtml(unsafeHtml, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOW_DATA_ATTR: false,
    ALLOW_ARIA_ATTR: false,
  });
}
