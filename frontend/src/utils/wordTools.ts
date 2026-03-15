import type { ToolDefinition } from '@/types';
import { logService } from '@/utils/logger';
import { executeOfficeAction } from './officeAction';
import { sandboxedEval } from './sandbox';
import { validateOfficeCode } from './officeCodeValidator';
import { applyRevisionToSelection, applyRevisionToDocument } from './wordDiffUtils';

import {
  applyInheritedStyles,
  type InheritedStyles,
  renderOfficeRichHtml,
  htmlToMarkdown,
} from './markdown';

import {
  createOfficeTools,
  truncateString,
  buildExecuteWrapper,
  type OfficeToolTemplate,
  getErrorMessage,
} from './common';
import { escapeXml } from './pptxZipUtils';
import {
  WORD_SEARCH_TEXT_MAX_LENGTH,
  WORD_HEADING_1_FONT_SIZE,
  WORD_HEADING_2_FONT_SIZE,
  WORD_HEADING_3_FONT_SIZE,
  WORD_CODE_TRUNCATE_SHORT,
  WORD_CODE_TRUNCATE_LONG,
} from '@/constants/limits';

export type WordToolName =
  | 'getSelectedText'
  | 'getDocumentContent'
  | 'getDocumentHtml'
  | 'getDocumentProperties'
  | 'insertContent'
  | 'formatText'
  | 'searchAndReplace'
  | 'searchAndFormat'
  | 'addComment'
  | 'getComments'
  | 'getSpecificParagraph'
  | 'applyStyle'
  | 'getSelectedTextWithFormatting'
  | 'findText'
  | 'applyTaggedFormatting'
  | 'setParagraphFormat'
  | 'insertHyperlink'
  | 'modifyTableCell'
  | 'addTableRow'
  | 'addTableColumn'
  | 'deleteTableRowColumn'
  | 'formatTableCell'
  | 'insertHeaderFooter'
  | 'insertFootnote'
  | 'setPageSetup'
  | 'insertSectionBreak'
  | 'proposeRevision'
  | 'proposeDocumentRevision'
  | 'editDocumentXml'
  | 'eval_wordjs';

const runWord = <T>(action: (context: Word.RequestContext) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => Word.run(action));

/**
 * Strip outer <ul>/<ol> wrapper tags from HTML while keeping <li> elements.
 * Prevents double-bullets when inserting list HTML into an existing list context —
 * Word's native list style handles the bullet display for each <li>.
 */
function stripOuterListTags(html: string): string {
  return html
    .replace(/<ul[^>]*>/gi, '')
    .replace(/<\/ul>/gi, '')
    .replace(/<ol[^>]*>/gi, '')
    .replace(/<\/ol>/gi, '');
}

interface InsertionContext {
  inList: boolean;
  styles: InheritedStyles;
}

/**
 * R1 — Read insertion-point font/spacing styles and list context in 2 syncs.
 * Combines isInsertionInList + style capture to minimise round-trips.
 * Note: paragraph-level spacing (spaceBefore/spaceAfter) is on Paragraph, not Range.
 */
async function readInsertionContext(context: Word.RequestContext): Promise<InsertionContext> {
  // Sync 1: load font from range + spaceBefore/spaceAfter + style from first paragraph
  const selection = context.document.getSelection();
  selection.load('font/name,font/size,font/bold,font/italic,font/color');
  const para = selection.paragraphs.getFirst();
  para.load('style,spaceBefore,spaceAfter');
  await context.sync();

  const styles: InheritedStyles = {
    fontFamily: selection.font.name || '',
    fontSize: selection.font.size ? `${selection.font.size}pt` : '',
    fontWeight: selection.font.bold ? 'bold' : 'normal',
    fontStyle: selection.font.italic ? 'italic' : 'normal',
    color: selection.font.color || '',
    marginTop: `${para.spaceBefore ?? 0}pt`,
    marginBottom: `${para.spaceAfter ?? 0}pt`,
  };

  // Sync 2: detect list context (throws if paragraph is not in a list)
  let inList = false;
  try {
    para.listItem.load('level');
    await context.sync();
    inList = true;
  } catch {
    inList = false;
  }

  return { inList, styles };
}

/**
 * R14 — After inserting HTML, detect paragraphs whose font size matches heading
 * thresholds and apply the corresponding Word builtin heading style.
 * This ensures headings become proper Word styles (h1→Heading 1, etc.) so that
 * features like Table of Contents work correctly.
 * Best-effort: silently skips if the API call fails.
 */
async function applyHeadingBuiltinStyles(
  context: Word.RequestContext,
  insertedRange: Word.Range,
  html: string,
): Promise<void> {
  if (!/\<h[1-6]/i.test(html)) return;

  try {
    const paras = insertedRange.paragraphs;
    paras.load('items');
    await context.sync();

    for (const p of paras.items) {
      p.load('font/size');
    }
    await context.sync();

    for (const p of paras.items) {
      const size = p.font.size;
      if (!size) continue;
      if (size >= WORD_HEADING_1_FONT_SIZE) p.styleBuiltIn = Word.BuiltInStyleName.heading1;
      else if (size >= WORD_HEADING_2_FONT_SIZE) p.styleBuiltIn = Word.BuiltInStyleName.heading2;
      else if (size >= WORD_HEADING_3_FONT_SIZE) p.styleBuiltIn = Word.BuiltInStyleName.heading3;
    }
    await context.sync();
  } catch {
    // Best-effort only — heading size detection may not be available in all runtimes
  }
}

type WordToolTemplate = OfficeToolTemplate<Word.RequestContext> & {
  executeWord: (context: Word.RequestContext, args: Record<string, any>) => Promise<string>;
};

const wordToolDefinitions = createOfficeTools<WordToolName, WordToolTemplate, ToolDefinition>(
  {
    getSelectedText: {
      name: 'getSelectedText',
      category: 'read',
      description:
        'Get the currently selected text in the Word document. Returns the selected text or empty string if nothing is selected.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        const range = context.document.getSelection();
        range.load('text');
        await context.sync();
        return range.text || '';
      },
    },

    getDocumentContent: {
      name: 'getDocumentContent',
      category: 'read',
      description: 'Get the full content of the Word document body as plain text.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        return body.text || '';
      },
    },

    insertContent: {
      name: 'insertContent',
      category: 'write',
      description:
        'The PREFERRED tool for adding any content to Word. Supports text, tables, and lists via Markdown. Can insert at cursor, start/end of document, or replace selection. Handles styling and list structure automatically.',
      inputSchema: {
        type: 'object',
        properties: {
          content: {
            type: 'string',
            description:
              'The content to insert in Markdown format. Tables and lists are supported.',
          },
          location: {
            type: 'string',
            description:
              'Where to insert relative to the selection: "Start", "End", "Before", "After", or "Replace" (default). Use "End" to append to the document.',
            enum: ['Start', 'End', 'Before', 'After', 'Replace'],
          },
          target: {
            type: 'string',
            description:
              'Target for insertion: "Selection" (default) or "Body" (entire document). Use "Body" with location "End" to append to document.',
            enum: ['Selection', 'Body'],
          },
          preserveFormatting: {
            type: 'boolean',
            description: 'When replacing selection, keep the original font styles. Default: true.',
          },
        },
        required: ['content'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const {
          content,
          location = 'Replace',
          target = 'Selection',
          preserveFormatting = true,
        } = args;

        const range =
          target === 'Body' ? context.document.body.getRange() : context.document.getSelection();

        // Load context styles for inheritance
        const { inList, styles } = await readInsertionContext(context);

        // If replacing and preserving formatting, we save the font props
        let savedFont: any = null;
        if (location === 'Replace' && preserveFormatting && target === 'Selection') {
          range.load('font/name,font/size,styleBuiltIn');
          await context.sync();
          savedFont = {
            name: range.font.name,
            size: range.font.size,
            style: range.styleBuiltIn,
          };
        }

        const html = renderOfficeRichHtml(content);
        const adjustedHtml = inList ? stripOuterListTags(html) : html;
        const styledHtml = applyInheritedStyles(adjustedHtml, styles);

        const insertedRange = range.insertHtml(styledHtml, location as any);

        if (savedFont && preserveFormatting) {
          insertedRange.font.name = savedFont.name;
          insertedRange.font.size = savedFont.size;
          // We don't force color here to allow markdown bold/italic/color to work
        }

        await context.sync();

        // Post-process headings
        await applyHeadingBuiltinStyles(context, insertedRange, html);

        // Auto-select the inserted range so formatText can be called immediately after
        insertedRange.select();
        await context.sync();

        return `Successfully inserted content at ${location} of ${target}`;
      },
    },

    formatText: {
      name: 'formatText',
      category: 'format',
      description:
        'Apply formatting to the currently selected text. WARNING: NEVER use this tool unless the user explicitly asks to format their CURRENT SELECTION. If you need to format specific text that is not selected (e.g., text extracted from a PDF or identified in the document), you MUST use "searchAndFormat" instead. At least one text character must be manually selected by the user for this to work.',
      inputSchema: {
        type: 'object',
        properties: {
          bold: {
            type: 'boolean',
            description: 'Make text bold',
          },
          italic: {
            type: 'boolean',
            description: 'Make text italic',
          },
          underline: {
            type: 'boolean',
            description: 'Underline text',
          },
          fontSize: {
            type: 'number',
            description: 'Font size in points',
          },
          fontColor: {
            type: 'string',
            description: 'Font color as hex (e.g., "#FF0000" for red)',
          },
          highlightColor: {
            type: 'string',
            description:
              'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { bold, italic, underline, fontSize, fontColor, highlightColor } = args as Record<
          string,
          any
        >;

        const range = context.document.getSelection();
        range.load('text');
        await context.sync();

        if (!range.text || range.text.length === 0) {
          throw new Error('Error: No text selected. Select text in the document, then try again.');
        }

        if (bold !== undefined) range.font.bold = bold;
        if (italic !== undefined) range.font.italic = italic;
        if (underline !== undefined) range.font.underline = underline ? 'Single' : 'None';
        if (fontSize !== undefined) range.font.size = fontSize;
        if (fontColor !== undefined) range.font.color = fontColor;
        if (highlightColor !== undefined) range.font.highlightColor = highlightColor;

        await context.sync();
        return 'Successfully applied formatting';
      },
    },

    searchAndReplace: {
      name: 'searchAndReplace',
      category: 'write',
      description:
        'Search for text in the document and replace it with new text. PREFERRED METHOD for correcting typos, modifying phrasing, or making targeted changes within paragraphs without destroying the surrounding document layout. Be surgical.',
      inputSchema: {
        type: 'object',
        properties: {
          searchText: {
            type: 'string',
            description: 'The text to search for',
          },
          replaceText: {
            type: 'string',
            description: 'The text to replace with',
          },
          matchCase: {
            type: 'boolean',
            description: 'Whether to match case (default: false)',
          },
          matchWholeWord: {
            type: 'boolean',
            description: 'Whether to match whole word only (default: false)',
          },
        },
        required: ['searchText', 'replaceText'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const {
          searchText,
          replaceText,
          matchCase = false,
          matchWholeWord = false,
        } = args as Record<string, any>;
        if (typeof searchText === 'string' && searchText.length > WORD_SEARCH_TEXT_MAX_LENGTH) {
          throw new Error(
            'Error: searchText cannot exceed 255 characters in Word. Please search for a smaller distinctive phrase (e.g., 5-10 words) instead of selecting entire paragraphs.',
          );
        }

        const body = context.document.body;
        const searchResults = body.search(searchText, {
          matchCase,
          matchWholeWord,
        });
        searchResults.load('items');
        await context.sync();

        const count = searchResults.items.length;
        for (const item of searchResults.items) {
          item.insertText(replaceText, 'Replace');
        }
        await context.sync();
        return `Replaced ${count} occurrence(s) of "${searchText}" with "${replaceText}"`;
      },
    },

    searchAndFormat: {
      name: 'searchAndFormat',
      category: 'format',
      description:
        'DEFAULT tool for applying formatting to existing text in the document. Finds text by content and applies formatting WITHOUT changing the text. Use this whenever you need to format text that is NOT the user\'s current selection — including text extracted from a PDF, text identified in the document, or any specific words/phrases. PREFERRED for requests like "color verbs in green", "highlight errors", "bold all names". Multiple calls expected — one per word/phrase to format.',
      inputSchema: {
        type: 'object',
        properties: {
          searchText: {
            type: 'string',
            description: 'The text to search for',
          },
          matchCase: {
            type: 'boolean',
            description: 'Whether to match case (default: false)',
          },
          matchWholeWord: {
            type: 'boolean',
            description: 'Whether to match whole word only (default: false)',
          },
          bold: {
            type: 'boolean',
            description: 'Apply bold formatting',
          },
          italic: {
            type: 'boolean',
            description: 'Apply italic formatting',
          },
          underline: {
            type: 'boolean',
            description: 'Apply underline formatting',
          },
          strikethrough: {
            type: 'boolean',
            description: 'Apply strikethrough formatting',
          },
          fontColor: {
            type: 'string',
            description: 'Font color as hex (e.g., "#228B22" for green, "#CC0000" for red)',
          },
          highlightColor: {
            type: 'string',
            description:
              'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
          },
          fontSize: {
            type: 'number',
            description: 'Font size in points',
          },
          fontName: {
            type: 'string',
            description: 'Font family name (e.g., "Calibri", "Arial")',
          },
        },
        required: ['searchText'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const {
          searchText,
          matchCase = false,
          matchWholeWord = false,
          bold,
          italic,
          underline,
          strikethrough,
          fontColor,
          highlightColor,
          fontSize,
          fontName,
        } = args;

        if (typeof searchText === 'string' && searchText.length > WORD_SEARCH_TEXT_MAX_LENGTH) {
          throw new Error('Error: searchText cannot exceed 255 characters.');
        }

        const body = context.document.body;
        const searchResults = body.search(searchText, { matchCase, matchWholeWord });
        searchResults.load('items');
        await context.sync();

        const count = searchResults.items.length;
        if (count === 0) {
          return `No occurrences of "${searchText}" found in the document.`;
        }

        for (const item of searchResults.items) {
          if (bold !== undefined) item.font.bold = bold;
          if (italic !== undefined) item.font.italic = italic;
          if (underline !== undefined) item.font.underline = underline ? 'Single' : 'None';
          if (strikethrough !== undefined) item.font.strikeThrough = strikethrough;
          if (fontColor !== undefined) item.font.color = fontColor;
          if (highlightColor !== undefined) item.font.highlightColor = highlightColor;
          if (fontSize !== undefined) item.font.size = fontSize;
          if (fontName !== undefined) item.font.name = fontName;
        }
        await context.sync();

        const formats: string[] = [];
        if (fontColor !== undefined) formats.push(`color: ${fontColor}`);
        if (highlightColor !== undefined) formats.push(`highlight: ${highlightColor}`);
        if (bold !== undefined) formats.push(bold ? 'bold' : 'not bold');
        if (italic !== undefined) formats.push(italic ? 'italic' : 'not italic');
        if (underline !== undefined) formats.push(underline ? 'underlined' : 'not underlined');
        if (fontSize !== undefined) formats.push(`size: ${fontSize}pt`);
        if (fontName !== undefined) formats.push(`font: ${fontName}`);

        return `Applied formatting (${formats.join(', ')}) to ${count} occurrence(s) of "${searchText}".`;
      },
    },

    getDocumentProperties: {
      name: 'getDocumentProperties',
      category: 'read',
      description:
        'Get document properties including paragraph count, word count, and character count.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        const body = context.document.body;
        body.load(['text']);

        const paragraphs = body.paragraphs;
        paragraphs.load('items');

        await context.sync();

        const text = body.text || '';
        const wordCount = text.split(/\s+/).filter(word => word.length > 0).length;
        const charCount = text.length;
        const paragraphCount = paragraphs.items.length;

        return JSON.stringify(
          {
            paragraphCount,
            wordCount,
            characterCount: charCount,
          },
          null,
          2,
        );
      },
    },

    findText: {
      name: 'findText',
      category: 'read',
      description:
        'Find text in the document and return information about matches. Does not modify the document.',
      inputSchema: {
        type: 'object',
        properties: {
          searchText: {
            type: 'string',
            description: 'The text to search for',
          },
          matchCase: {
            type: 'boolean',
            description: 'Whether to match case (default: false)',
          },
          matchWholeWord: {
            type: 'boolean',
            description: 'Whether to match whole word only (default: false)',
          },
        },
        required: ['searchText'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const {
          searchText,
          matchCase = false,
          matchWholeWord = false,
        } = args as Record<string, any>;
        if (typeof searchText === 'string' && searchText.length > WORD_SEARCH_TEXT_MAX_LENGTH) {
          throw new Error(
            'Error: searchText cannot exceed 255 characters in Word. Please search for a smaller distinctive phrase (e.g., 5-10 words) instead of selecting entire paragraphs.',
          );
        }

        const body = context.document.body;
        const searchResults = body.search(searchText, {
          matchCase,
          matchWholeWord,
        });
        searchResults.load(['items']);
        await context.sync();

        const count = searchResults.items.length;
        return JSON.stringify(
          {
            searchText,
            matchCount: count,
            found: count > 0,
          },
          null,
          2,
        );
      },
    },

    applyTaggedFormatting: {
      name: 'applyTaggedFormatting',
      category: 'format',
      description:
        'Convert inline formatting tags in the document into real Word formatting (e.g., <format>text</format> can apply size, font, italic, bold, underline, strike, highlight, color, and other font settings).',
      inputSchema: {
        type: 'object',
        properties: {
          tagName: {
            type: 'string',
            description: 'Tag name to process (default: "format")',
          },
          fontName: {
            type: 'string',
            description: 'Font family name (e.g., "Calibri", "Arial")',
          },
          fontSize: {
            type: 'number',
            description: 'Font size in points',
          },
          color: {
            type: 'string',
            description: 'Font color to apply as hex or named color',
          },
          highlightColor: {
            type: 'string',
            description:
              'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
          },
          bold: {
            type: 'boolean',
            description: 'Whether to apply bold formatting',
          },
          italic: {
            type: 'boolean',
            description: 'Whether to apply italic formatting',
          },
          underline: {
            type: 'boolean',
            description: 'Whether to apply underline formatting',
          },
          strikethrough: {
            type: 'boolean',
            description: 'Whether to apply strikethrough formatting',
          },
          allCaps: {
            type: 'boolean',
            description: 'Whether to format text in all caps',
          },
          subscript: {
            type: 'boolean',
            description: 'Whether to apply subscript formatting',
          },
          superscript: {
            type: 'boolean',
            description: 'Whether to apply superscript formatting',
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const tagName =
          typeof args.tagName === 'string' && args.tagName.trim() ? args.tagName.trim() : 'format';
        const fontName =
          typeof args.fontName === 'string' && args.fontName.trim()
            ? args.fontName.trim()
            : undefined;
        const fontSize = typeof args.fontSize === 'number' ? args.fontSize : undefined;
        const color =
          typeof args.color === 'string' && args.color.trim() ? args.color.trim() : undefined;
        const highlightColor =
          typeof args.highlightColor === 'string' && args.highlightColor.trim()
            ? args.highlightColor.trim()
            : undefined;
        const bold = args.bold !== undefined ? Boolean(args.bold) : undefined;
        const italic = args.italic !== undefined ? Boolean(args.italic) : undefined;
        const underline = args.underline !== undefined ? Boolean(args.underline) : undefined;
        const strikethrough =
          args.strikethrough !== undefined ? Boolean(args.strikethrough) : undefined;
        const allCaps = args.allCaps !== undefined ? Boolean(args.allCaps) : undefined;
        const subscript = args.subscript !== undefined ? Boolean(args.subscript) : undefined;
        const superscript = args.superscript !== undefined ? Boolean(args.superscript) : undefined;

        const body = context.document.body;
        body.load('text');
        await context.sync();

        const openingTag = `<${tagName}>`;
        const closingTag = `</${tagName}>`;
        const openingTagMatches = body.search(openingTag, {
          matchCase: true,
          matchWholeWord: false,
        });
        openingTagMatches.load('items');
        await context.sync();

        if (openingTagMatches.items.length === 0) {
          return `No <${tagName}>...</${tagName}> tags found.`;
        }

        let replacedCount = 0;
        for (let i = openingTagMatches.items.length - 1; i >= 0; i--) {
          const openingRange = openingTagMatches.items[i];
          const afterOpeningRange = openingRange.getRange('After');
          const closingTagMatches = afterOpeningRange.search(closingTag, {
            matchCase: true,
            matchWholeWord: false,
          });
          closingTagMatches.load('items');
          await context.sync();

          if (closingTagMatches.items.length === 0) {
            continue;
          }

          const taggedRange = openingRange.expandTo(closingTagMatches.items[0]);
          taggedRange.load('text');
          await context.sync();

          const taggedText = taggedRange.text || '';
          if (!taggedText.endsWith(closingTag)) {
            continue;
          }

          const innerText = taggedText.slice(
            openingTag.length,
            taggedText.length - closingTag.length,
          );
          const formattedRange = taggedRange.insertText(innerText, 'Replace');

          if (fontName !== undefined) formattedRange.font.name = fontName;
          if (fontSize !== undefined) formattedRange.font.size = fontSize;
          if (color !== undefined) formattedRange.font.color = color;
          if (highlightColor !== undefined) formattedRange.font.highlightColor = highlightColor;
          if (bold !== undefined) formattedRange.font.bold = bold;
          if (italic !== undefined) formattedRange.font.italic = italic;
          if (underline !== undefined)
            formattedRange.font.underline = underline ? 'Single' : 'None';
          if (strikethrough !== undefined) formattedRange.font.strikeThrough = strikethrough;
          if (allCaps !== undefined) formattedRange.font.allCaps = allCaps;
          if (subscript !== undefined) formattedRange.font.subscript = subscript;
          if (superscript !== undefined) formattedRange.font.superscript = superscript;

          replacedCount++;
          await context.sync();
        }

        return `Converted ${replacedCount} tagged occurrence(s) using <${tagName}>...</${tagName}>.`;
      },
    },

    setParagraphFormat: {
      name: 'setParagraphFormat',
      category: 'format',
      description:
        'Set paragraph formatting on the current selection (alignment, spacing before/after, line spacing, and indentation).',
      inputSchema: {
        type: 'object',
        properties: {
          alignment: {
            type: 'string',
            description: 'Paragraph alignment: Left, Centered, Right, or Justified.',
            enum: ['Left', 'Centered', 'Right', 'Justified'],
          },
          lineSpacing: {
            type: 'number',
            description: 'Line spacing value in points.',
          },
          spaceBefore: {
            type: 'number',
            description: 'Paragraph spacing before in points.',
          },
          spaceAfter: {
            type: 'number',
            description: 'Paragraph spacing after in points.',
          },
          leftIndent: {
            type: 'number',
            description: 'Left indent in points.',
          },
          firstLineIndent: {
            type: 'number',
            description: 'First line indent in points.',
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { alignment, lineSpacing, spaceBefore, spaceAfter, leftIndent, firstLineIndent } =
          args as Record<string, any>;

        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load('items');
        await context.sync();

        if (paragraphs.items.length === 0) {
          throw new Error('Error: No paragraph found in the current selection.');
        }

        for (const paragraph of paragraphs.items) {
          if (alignment !== undefined) paragraph.alignment = alignment as Word.Alignment;
          if (lineSpacing !== undefined) paragraph.lineSpacing = lineSpacing;
          if (spaceBefore !== undefined) paragraph.spaceBefore = spaceBefore;
          if (spaceAfter !== undefined) paragraph.spaceAfter = spaceAfter;
          if (leftIndent !== undefined) paragraph.leftIndent = leftIndent;
          if (firstLineIndent !== undefined) paragraph.firstLineIndent = firstLineIndent;
        }

        await context.sync();
        return `Successfully formatted ${paragraphs.items.length} paragraph(s)`;
      },
    },

    insertHyperlink: {
      name: 'insertHyperlink',
      category: 'write',
      description: 'Insert a clickable hyperlink at the current selection.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'URL address of the hyperlink.',
          },
          textToDisplay: {
            type: 'string',
            description:
              'Optional display text for the hyperlink. If omitted, uses the current selection.',
          },
        },
        required: ['address'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { address, textToDisplay } = args as Record<string, any>;

        const range = context.document.getSelection();
        const linkRange =
          typeof textToDisplay === 'string' ? range.insertText(textToDisplay, 'Replace') : range;
        linkRange.hyperlink = address;
        await context.sync();
        return 'Successfully inserted hyperlink';
      },
    },

    getDocumentHtml: {
      name: 'getDocumentHtml',
      category: 'read',
      description:
        'Get the full document body as HTML to preserve structure and rich formatting context.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        const htmlResult = context.document.body.getHtml();
        await context.sync();
        return htmlResult.value || '';
      },
    },

    modifyTableCell: {
      name: 'modifyTableCell',
      category: 'write',
      description: 'Replace content of a specific table cell in an existing table.',
      inputSchema: {
        type: 'object',
        properties: {
          row: { type: 'number', description: 'Zero-based row index.' },
          column: { type: 'number', description: 'Zero-based column index.' },
          text: { type: 'string', description: 'Text to set in the table cell.' },
          tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
        },
        required: ['row', 'column', 'text'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { row, column, text, tableIndex = 0 } = args as Record<string, any>;

        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error('Error: Table not found at the requested index.');
        }

        const targetTable = tables.items[tableIndex];
        const cell = targetTable.getCell(row, column);
        cell.body.insertText(text, 'Replace');
        await context.sync();
        return `Successfully updated table ${tableIndex}, cell (${row}, ${column})`;
      },
    },

    addTableRow: {
      name: 'addTableRow',
      category: 'write',
      description: 'Add row(s) to an existing table.',
      inputSchema: {
        type: 'object',
        properties: {
          tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
          location: {
            type: 'string',
            description: 'Insert location relative to the current selection in the table.',
            enum: ['Before', 'After'],
          },
          count: { type: 'number', description: 'Number of rows to add.' },
          values: {
            type: 'array',
            description: 'Optional row values as a 2D array.',
            items: { type: 'array', items: { type: 'string' } },
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { tableIndex = 0, location = 'After', count = 1, values } = args;

        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error('Error: Table not found at the requested index.');
        }

        const targetTable = tables.items[tableIndex] as any;
        targetTable.addRows(location, count, values);
        await context.sync();
        return `Successfully added ${count} row(s) to table ${tableIndex}`;
      },
    },

    addTableColumn: {
      name: 'addTableColumn',
      category: 'write',
      description: 'Add column(s) to an existing table.',
      inputSchema: {
        type: 'object',
        properties: {
          tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
          location: {
            type: 'string',
            description: 'Insert location relative to the current selection in the table.',
            enum: ['Before', 'After'],
          },
          count: { type: 'number', description: 'Number of columns to add.' },
          values: {
            type: 'array',
            description: 'Optional column values as a 2D array.',
            items: { type: 'array', items: { type: 'string' } },
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { tableIndex = 0, location = 'After', count = 1, values } = args;

        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error('Error: Table not found at the requested index.');
        }

        const targetTable = tables.items[tableIndex] as any;
        targetTable.addColumns(location, count, values);
        await context.sync();
        return `Successfully added ${count} column(s) to table ${tableIndex}`;
      },
    },

    deleteTableRowColumn: {
      name: 'deleteTableRowColumn',
      category: 'write',
      description: 'Delete row(s) or column(s) from an existing table.',
      inputSchema: {
        type: 'object',
        properties: {
          tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
          target: {
            type: 'string',
            description: 'Whether to delete rows or columns.',
            enum: ['row', 'column'],
          },
          index: { type: 'number', description: 'Zero-based row or column index.' },
          count: { type: 'number', description: 'Number of rows or columns to delete.' },
        },
        required: ['target', 'index'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { tableIndex = 0, target, index, count = 1 } = args as Record<string, any>;

        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error('Error: Table not found at the requested index.');
        }

        const targetTable = tables.items[tableIndex] as any;
        if (target === 'row') {
          targetTable.deleteRows(index, count);
          await context.sync();
          return `Successfully deleted ${count} row(s) starting at index ${index}`;
        }

        targetTable.deleteColumns(index, count);
        await context.sync();
        return `Successfully deleted ${count} column(s) starting at index ${index}`;
      },
    },

    formatTableCell: {
      name: 'formatTableCell',
      category: 'format',
      description: 'Apply formatting to a specific table cell (background and font style).',
      inputSchema: {
        type: 'object',
        properties: {
          tableIndex: { type: 'number', description: 'Zero-based table index (default: 0).' },
          row: { type: 'number', description: 'Zero-based row index.' },
          column: { type: 'number', description: 'Zero-based column index.' },
          shadingColor: {
            type: 'string',
            description: 'Cell background color (hex or color name).',
          },
          fontName: { type: 'string', description: 'Font family to apply.' },
          fontSize: { type: 'number', description: 'Font size in points.' },
          fontColor: { type: 'string', description: 'Font color (hex or color name).' },
          bold: { type: 'boolean', description: 'Whether font is bold.' },
          italic: { type: 'boolean', description: 'Whether font is italic.' },
        },
        required: ['row', 'column'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const {
          tableIndex = 0,
          row,
          column,
          shadingColor,
          fontName,
          fontSize,
          fontColor,
          bold,
          italic,
        } = args as Record<string, any>;

        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        if (tables.items.length === 0 || tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error('Error: Table not found at the requested index.');
        }

        const cell = tables.items[tableIndex].getCell(row, column) as any;
        if (shadingColor !== undefined) cell.shadingColor = shadingColor;

        const cellBody = cell.body;
        if (fontName !== undefined) cellBody.font.name = fontName;
        if (fontSize !== undefined) cellBody.font.size = fontSize;
        if (fontColor !== undefined) cellBody.font.color = fontColor;
        if (bold !== undefined) cellBody.font.bold = bold;
        if (italic !== undefined) cellBody.font.italic = italic;

        await context.sync();
        return `Successfully formatted table ${tableIndex}, cell (${row}, ${column})`;
      },
    },

    insertHeaderFooter: {
      name: 'insertHeaderFooter',
      category: 'write',
      description: 'Insert text into the document header or footer of the first section.',
      inputSchema: {
        type: 'object',
        properties: {
          target: {
            type: 'string',
            description: 'Where to insert text.',
            enum: ['header', 'footer'],
          },
          type: {
            type: 'string',
            description: 'Header/footer type: Primary, FirstPage, or EvenPages.',
            enum: ['Primary', 'FirstPage', 'EvenPages'],
          },
          text: {
            type: 'string',
            description: 'Text to insert into header/footer.',
          },
        },
        required: ['target', 'text'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { target, type = 'Primary', text } = args;

        const section = context.document.sections.getFirst() as any;
        const container = target === 'header' ? section.getHeader(type) : section.getFooter(type);
        container.insertText(text, 'Replace');
        await context.sync();
        return `Successfully inserted text into ${target} (${type})`;
      },
    },

    insertFootnote: {
      name: 'insertFootnote',
      category: 'write',
      description: 'Insert a footnote at the current selection.',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'Footnote content.' },
        },
        required: ['text'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { text } = args as Record<string, any>;

        const range = context.document.getSelection() as any;
        range.insertFootnote(text);
        await context.sync();
        return 'Successfully inserted footnote';
      },
    },

    addComment: {
      name: 'addComment',
      category: 'write',
      description:
        'Add a review comment bubble to a specific segment within the current selection. Use this to suggest changes, provide feedback, or alert the user during proofreading without modifying the original text directly.',
      inputSchema: {
        type: 'object',
        properties: {
          textSegment: {
            type: 'string',
            description: 'The specific word or phrase in the selection that has an error.',
          },
          comment: {
            type: 'string',
            description: 'The feedback or correction for the text segment.',
          },
        },
        required: ['textSegment', 'comment'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { textSegment, comment } = args as Record<string, any>;

        const range = context.document.getSelection() as any;
        // Perform a search within the selected range
        const searchResults = range.search(textSegment, { matchCase: false });
        searchResults.load('items');

        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          // Apply the comment exclusively to the first match within the selection
          searchResults.items[0].insertComment(comment);
        } else {
          // Fallback: apply to entire selection if the specific segment string is somehow not found
          range.insertComment(`[On: "${textSegment}"]\n${comment}`);
        }

        await context.sync();
        return `Successfully added comment to segment: "${textSegment}"`;
      },
    },

    getComments: {
      name: 'getComments',
      category: 'read',
      description: 'Get comments from the document body with author and content.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        const body = context.document.body as any;
        const comments = body.getComments();
        comments.load('items/content,items/authorName');
        await context.sync();

        return JSON.stringify(
          {
            count: comments.items.length,
            comments: comments.items.map((comment: any) => ({
              authorName: comment.authorName || '',
              content: comment.content || '',
            })),
          },
          null,
          2,
        );
      },
    },

    setPageSetup: {
      name: 'setPageSetup',
      category: 'format',
      description: 'Configure page setup on the first section (margins, orientation, paper size).',
      inputSchema: {
        type: 'object',
        properties: {
          topMargin: { type: 'number', description: 'Top margin in points.' },
          bottomMargin: { type: 'number', description: 'Bottom margin in points.' },
          leftMargin: { type: 'number', description: 'Left margin in points.' },
          rightMargin: { type: 'number', description: 'Right margin in points.' },
          orientation: {
            type: 'string',
            description: 'Page orientation: Portrait or Landscape.',
            enum: ['Portrait', 'Landscape'],
          },
          paperSize: {
            type: 'string',
            description: 'Paper size (for example Letter, A4, Legal).',
            enum: ['Letter', 'A4', 'Legal'],
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { topMargin, bottomMargin, leftMargin, rightMargin, orientation, paperSize } =
          args as Record<string, any>;

        const section = context.document.sections.getFirst() as any;
        const pageSetup = section.pageSetup;

        if (topMargin !== undefined) pageSetup.topMargin = topMargin;
        if (bottomMargin !== undefined) pageSetup.bottomMargin = bottomMargin;
        if (leftMargin !== undefined) pageSetup.leftMargin = leftMargin;
        if (rightMargin !== undefined) pageSetup.rightMargin = rightMargin;
        if (orientation !== undefined) pageSetup.orientation = orientation;
        if (paperSize !== undefined) pageSetup.paperSize = paperSize;

        await context.sync();
        return 'Successfully updated page setup';
      },
    },

    getSpecificParagraph: {
      name: 'getSpecificParagraph',
      category: 'format',
      description: 'Get a paragraph by index without reading all document content.',
      inputSchema: {
        type: 'object',
        properties: {
          index: {
            type: 'number',
            description: 'Zero-based paragraph index.',
          },
        },
        required: ['index'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { index } = args as Record<string, any>;

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        if (index < 0 || index >= paragraphs.items.length) {
          throw new Error(
            `Error: Paragraph index out of bounds. Range is 0 to ${Math.max(paragraphs.items.length - 1, 0)}.`,
          );
        }

        const paragraph = paragraphs.items[index];
        paragraph.load(
          'text,style,font/name,font/size,font/bold,font/italic,font/underline,font/color',
        );
        await context.sync();

        return JSON.stringify(
          {
            index,
            text: paragraph.text || '',
            style: paragraph.style,
            font: {
              name: paragraph.font.name,
              size: paragraph.font.size,
              bold: paragraph.font.bold,
              italic: paragraph.font.italic,
              underline: paragraph.font.underline,
              color: paragraph.font.color,
            },
          },
          null,
          2,
        );
      },
    },

    insertSectionBreak: {
      name: 'insertSectionBreak',
      category: 'write',
      description: 'Insert a section break at the current selection.',
      inputSchema: {
        type: 'object',
        properties: {
          location: {
            type: 'string',
            description: 'Where to insert section break: Before or After current selection.',
            enum: ['Before', 'After'],
          },
        },
        required: [],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { location = 'After' } = args;

        const range = context.document.getSelection() as any;
        range.insertBreak('SectionNext', location);
        await context.sync();
        return `Successfully inserted section break ${location.toLowerCase()} selection`;
      },
    },

    applyStyle: {
      name: 'applyStyle',
      category: 'format',
      description:
        'Apply a Word builtin paragraph style to the current selection, a specific paragraph by index, or the paragraph at cursor. ' +
        'Use this to set headings (Heading1-9), body text (Normal), special styles (Title, Subtitle, Quote, etc.). ' +
        'Builtin styles work across all languages and appear in the Word Style Gallery. ' +
        'Use paragraphIndex (from getDocumentContent or getSpecificParagraph) to target a paragraph without requiring user selection.',
      inputSchema: {
        type: 'object',
        properties: {
          styleBuiltIn: {
            type: 'string',
            description: 'Word builtin style name.',
            enum: [
              'Normal',
              'Heading1',
              'Heading2',
              'Heading3',
              'Heading4',
              'Heading5',
              'Heading6',
              'Heading7',
              'Heading8',
              'Heading9',
              'Title',
              'Subtitle',
              'Strong',
              'Emphasis',
              'ListBullet',
              'ListBullet2',
              'ListBullet3',
              'ListNumber',
              'ListNumber2',
              'ListNumber3',
              'Quote',
              'IntenseQuote',
              'SubtleEmphasis',
              'IntenseEmphasis',
              'SubtleReference',
              'IntenseReference',
              'BookTitle',
              'NoSpacing',
              'ListParagraph',
            ],
          },
          paragraphIndex: {
            type: 'number',
            description:
              'Zero-based index of the paragraph to style (from getSpecificParagraph or getDocumentContent). When provided, targets that paragraph directly without requiring a user selection.',
          },
          target: {
            type: 'string',
            description:
              'Where to apply when paragraphIndex is not provided: "selection" (whole selection) or "current-paragraph" (paragraph at cursor). Default: "selection".',
            enum: ['selection', 'current-paragraph'],
          },
        },
        required: ['styleBuiltIn'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { styleBuiltIn, paragraphIndex, target = 'selection' } = args;

        if (paragraphIndex !== undefined && paragraphIndex !== null) {
          const idx = Number(paragraphIndex);
          const paragraphs = context.document.body.paragraphs;
          paragraphs.load('items');
          await context.sync();

          if (idx < 0 || idx >= paragraphs.items.length) {
            return `Error: paragraphIndex ${idx} is out of bounds. Document has ${paragraphs.items.length} paragraphs (0–${paragraphs.items.length - 1}).`;
          }

          paragraphs.items[idx].styleBuiltIn = styleBuiltIn as Word.BuiltInStyleName;
          await context.sync();
          return `Applied style "${styleBuiltIn}" to paragraph at index ${idx}.`;
        }

        const selection = context.document.getSelection();
        if (target === 'current-paragraph') {
          const para = selection.paragraphs.getFirst();
          para.styleBuiltIn = styleBuiltIn as Word.BuiltInStyleName;
        } else {
          selection.styleBuiltIn = styleBuiltIn as Word.BuiltInStyleName;
        }

        await context.sync();
        return `Applied style "${styleBuiltIn}" to ${target}.`;
      },
    },

    getSelectedTextWithFormatting: {
      name: 'getSelectedTextWithFormatting',
      category: 'read',
      description:
        'R16: Get the currently selected text as Markdown, preserving all formatting ' +
        '(bold, italic, underline, headings, lists, hyperlinks). ' +
        'Use this instead of getSelectedText when you need to modify text while preserving its rich formatting. ' +
        'The LLM can edit the Markdown and pass the result back to replaceSelectedText.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeWord: async context => {
        try {
          const selection = context.document.getSelection();
          const htmlResult = selection.getHtml();
          await context.sync();

          if (!htmlResult.value || htmlResult.value.trim() === '') {
            return 'No text selected. Use getDocumentContent to read the full document, then use searchAndFormat to apply formatting to specific words or passages.';
          }

          const markdown = htmlToMarkdown(htmlResult.value);
          return markdown || 'Selection contains no convertible content.';
        } catch {
          return 'No text selected. Use getDocumentContent to read the full document, then use searchAndFormat to apply formatting to specific words or passages.';
        }
      },
    },

    proposeRevision: {
      name: 'proposeRevision',
      category: 'write',
      description: `**PREFERRED TOOL** for modifying existing text.

Generates native Word Track Changes (redlines) using OOXML revision markup.
The user can accept/reject each change individually in Word's Review pane.

Changes are attributed to a configurable author (default: "KickOffice AI")
visible in the Track Changes panel, distinguishable from human edits.

**Input**: The COMPLETE revised version of the selected text.
**Output**: The selection is replaced with tracked insertions/deletions.

**Requirements**: Text must be selected in the document before calling.
**Track Changes**: Enabled by default. Set enableTrackChanges=false for silent replacement.`,

      inputSchema: {
        type: 'object',
        properties: {
          revisedText: {
            type: 'string',
            description:
              'The complete revised version of the selected text. Must contain ALL text, not just changes.',
          },
          enableTrackChanges: {
            type: 'boolean',
            description:
              'Show changes in Word Track Changes panel (default: true). Set false for silent replacement.',
          },
        },
        required: ['revisedText'],
      },

      executeWord: async (context, args: Record<string, any>) => {
        const { revisedText, enableTrackChanges = true } = args;

        const result = await applyRevisionToSelection(context, revisedText, enableTrackChanges);

        return JSON.stringify(
          {
            success: result.success,
            strategy: result.strategy,
            author: result.author,
            message: result.message,
          },
          null,
          2,
        );
      },
    },

    proposeDocumentRevision: {
      name: 'proposeDocumentRevision',
      category: 'write',
      description: `Apply paragraph-level Track Changes revisions across the **entire document**.

Use when you need to revise multiple paragraphs throughout the document with native Word
Track Changes (redlines) — the user can then accept/reject each change individually.

**Workflow**:
1. Call \`getDocumentContent\` to read the full document text.
2. Identify which paragraphs need revision.
3. Call this tool with an array of \`{ originalText, revisedText }\` pairs.

Each entry is matched to the first paragraph whose text equals \`originalText\` (trimmed).
Provide the **complete** revised paragraph text in \`revisedText\`, not just the diff.

**vs proposeRevision**: \`proposeRevision\` requires a selection; this tool operates on
the full document without requiring the user to select anything.

**Track Changes**: Enabled by default. Set \`enableTrackChanges=false\` for silent replacement.`,

      inputSchema: {
        type: 'object',
        properties: {
          revisions: {
            type: 'array',
            description: 'List of paragraph revisions to apply.',
            items: {
              type: 'object',
              properties: {
                originalText: {
                  type: 'string',
                  description:
                    'The exact current text of the paragraph to revise (used to locate it).',
                },
                revisedText: {
                  type: 'string',
                  description:
                    'The complete revised version of the paragraph. Must contain ALL text, not just changes.',
                },
              },
              required: ['originalText', 'revisedText'],
            },
            minItems: 1,
          },
          enableTrackChanges: {
            type: 'boolean',
            description:
              'Show changes in Word Track Changes panel (default: true). Set false for silent replacement.',
          },
        },
        required: ['revisions'],
      },

      executeWord: async (context, args: Record<string, any>) => {
        const { revisions, enableTrackChanges = true } = args;

        const result = await applyRevisionToDocument(context, revisions, enableTrackChanges);

        return JSON.stringify(
          {
            success: result.success,
            applied: result.applied,
            failed: result.failed,
            skipped: result.skipped,
            author: result.author,
            details: result.details,
          },
          null,
          2,
        );
      },
    },

    editDocumentXml: {
      name: 'editDocumentXml',
      category: 'write',
      description: `Edit Word document OOXML directly for precision operations.

**USE WHEN**: You need to modify text while preserving exact formatting
(fonts, colors, sizes, styles) that would be lost with insertText/insertHtml.

**DO NOT USE FOR**: Track Changes (use proposeRevision instead).

The code receives:
- \`ooxml\`: string — the raw Flat OPC XML of the target range
- \`DOMParser\`, \`XMLSerializer\` — XML manipulation
- \`escapeXml(str)\` — safely escape XML special chars
- \`setResult(modifiedXml)\` — call this to write back the modified XML

Your code should:
1. Parse the ooxml with DOMParser
2. Find and modify the desired elements
3. Serialize back and call setResult()`,

      inputSchema: {
        type: 'object',
        properties: {
          target: {
            type: 'string',
            enum: ['selection', 'paragraph'],
            description: 'What to target: current selection or specific paragraph',
          },
          paragraphIndex: {
            type: 'number',
            description: 'If target=paragraph, the 0-based paragraph index',
          },
          code: {
            type: 'string',
            description: 'JavaScript code to manipulate the OOXML',
          },
          explanation: {
            type: 'string',
            description: 'What this code does (required for audit trail)',
          },
        },
        required: ['code', 'explanation'],
      },

      executeWord: async (context, args: Record<string, any>) => {
        const { target = 'selection', paragraphIndex, code, explanation } = args;

        // 1. Get the target range
        let range;
        if (target === 'paragraph' && paragraphIndex !== undefined) {
          const paragraphs = context.document.body.paragraphs;
          paragraphs.load('items');
          await context.sync();
          if (paragraphIndex >= paragraphs.items.length) {
            return JSON.stringify(
              {
                success: false,
                error: `Paragraph ${paragraphIndex} out of bounds`,
              },
              null,
              2,
            );
          }
          range = paragraphs.items[paragraphIndex].getRange();
        } else {
          range = context.document.getSelection();
        }

        // 2. Get OOXML
        const ooxmlResult = range.getOoxml();
        await context.sync();
        const ooxml = ooxmlResult.value;

        // 3. Execute code in sandbox
        let modifiedXml: string | null = null;
        const setResult = (xml: string) => {
          modifiedXml = xml;
        };

        try {
          await sandboxedEval(
            code,
            {
              ooxml,
              DOMParser,
              XMLSerializer,
              escapeXml,
              setResult,
            },
            'Word',
          );
        } catch (err: unknown) {
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(err),
              explanation,
            },
            null,
            2,
          );
        }

        // 4. Write back if modified
        if (modifiedXml) {
          try {
            range.insertOoxml(modifiedXml, 'Replace');
            await context.sync();
            return JSON.stringify(
              {
                success: true,
                explanation,
                action: 'OOXML modified and reinserted',
              },
              null,
              2,
            );
          } catch (insertError: any) {
            return JSON.stringify(
              {
                success: false,
                error: `insertOoxml failed: ${insertError.message || String(insertError)}`,
                explanation,
              },
              null,
              2,
            );
          }
        }

        return JSON.stringify(
          {
            success: true,
            explanation,
            action: 'No modifications applied (setResult not called)',
          },
          null,
          2,
        );
      },
    },

    eval_wordjs: {
      name: 'eval_wordjs',
      category: 'write',
      description: `Execute custom Office.js code within a Word.run context.

**USE THIS TOOL ONLY WHEN:**
- No dedicated tool exists for your operation
- You need to perform a complex multi-step operation
- You're doing something unusual not covered by other tools

**REQUIRED CODE STRUCTURE:**
Your code MUST follow this template:

\`\`\`javascript
try {
  // 1. Get reference to document/range
  const range = context.document.getSelection();

  // 2. Load required properties BEFORE reading them
  range.load('text,font/bold,font/size');
  await context.sync();

  // 3. Check for valid state
  if (!range.text) {
    return { success: false, error: 'No text selected' };
  }

  // 4. Perform your operations
  range.font.bold = true;

  // 5. Commit changes with sync
  await context.sync();

  // 6. Return result
  return { success: true, result: 'Operation completed' };
} catch (error) {
  return { success: false, error: error.message };
}
\`\`\`

**CRITICAL RULES:**
1. ALWAYS call \`.load()\` before reading any property
2. ALWAYS call \`await context.sync()\` after load and after modifications
3. ALWAYS wrap in try/catch
4. ONLY use Word namespace (not Excel, PowerPoint)`,
      inputSchema: {
        type: 'object',
        properties: {
          code: {
            type: 'string',
            description:
              'JavaScript code following the template above. Must include load(), sync(), and try/catch.',
          },
          explanation: {
            type: 'string',
            description: 'Brief explanation of what this code does (required for audit trail).',
          },
        },
        required: ['code', 'explanation'],
      },
      executeWord: async (context, args: Record<string, any>) => {
        const { code, explanation } = args;

        // Validate code BEFORE execution
        const validation = validateOfficeCode(code, 'Word');

        if (!validation.valid) {
          return JSON.stringify(
            {
              success: false,
              error: 'Code validation failed. Fix the errors below and try again.',
              validationErrors: validation.errors,
              validationWarnings: validation.warnings,
              suggestion:
                'Refer to the Office.js skill document for correct patterns. Common issues: missing load() before reading properties, missing context.sync() to commit changes.',
              codeReceived: truncateString(code, WORD_CODE_TRUNCATE_LONG),
            },
            null,
            2,
          );
        }

        // Log warnings but proceed
        if (validation.warnings.length > 0) {
          logService.warn('[eval_wordjs] Validation warnings:', validation.warnings);
        }

        try {
          // Execute in sandbox with host restriction
          const result = await sandboxedEval(
            code,
            {
              context,
              Word: typeof Word !== 'undefined' ? Word : undefined,
              Office: typeof Office !== 'undefined' ? Office : undefined,
            },
            'Word', // Restrict to Word namespace only
          );

          return JSON.stringify(
            {
              success: true,
              result: result ?? null,
              explanation,
              warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
            },
            null,
            2,
          );
        } catch (err: unknown) {
          return JSON.stringify(
            {
              success: false,
              error: getErrorMessage(err),
              explanation,
              codeExecuted: truncateString(code, WORD_CODE_TRUNCATE_SHORT),
              hint: 'Check that all properties are loaded before access, and context.sync() is called.',
            },
            null,
            2,
          );
        }
      },
    },
  },
  buildExecuteWrapper<WordToolTemplate>('executeWord', runWord),
);

export function getWordToolDefinitions(): ToolDefinition[] {
  return Object.values(wordToolDefinitions);
}

export { wordToolDefinitions };
