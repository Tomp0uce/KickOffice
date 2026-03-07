# Word Office.js Skill

## AVAILABLE TOOLS — Use in this priority order

### For READING content:

| Tool                            | When to use                                       |
| ------------------------------- | ------------------------------------------------- |
| `getSelectedText`               | Get plain text of selection                       |
| `getSelectedTextWithFormatting` | **PREFERRED** — Get Markdown with bold/italic/etc |
| `getDocumentContent`            | Read entire document as plain text                |
| `getDocumentHtml`               | Read document as HTML (for complex analysis)      |
| `getSpecificParagraph`          | Read one paragraph by index                       |
| `findText`                      | Search for text occurrences                       |
| `getComments`                   | List all comments in document                     |
| `getDocumentProperties`         | Get word count, paragraph count, etc.             |

### For WRITING/EDITING content:

| Tool                 | When to use                                                              |
| -------------------- | ------------------------------------------------------------------------ |
| `proposeRevision`    | **PREFERRED** for editing existing text. Preserves formatting! Uses diff |
| `searchAndReplace`   | Fix specific words/phrases throughout document                           |
| `insertContent`      | Add NEW content only (Markdown + inline color/style syntax)              |
| `insertHyperlink`    | Add clickable links                                                      |
| `addComment`         | Add review comments                                                      |
| `insertHeaderFooter` | Add headers/footers                                                      |
| `insertFootnote`     | Add footnotes                                                            |

### For FORMATTING:

**CRITICAL RULE**: The `formatText` tool ONLY works when text is already selected by the user. If you just inserted text via `insertContent`, it is NOT selected — you CANNOT color/bold it with `formatText` after the fact.

To apply formatting to specific words or to newly inserted text, use one of these three workflows (in priority order):

---

#### WORKFLOW C — `searchAndFormat` (PREFERRED for formatting specific words)

The simplest way to format specific words. Does NOT modify text content, no Track Changes impact.

Example: "mettre les verbes en vert"

1. Read the document with `getDocumentContent` or `getSelectedTextWithFormatting`
2. Identify the target words (verbs, names, errors, etc.)
3. Call `searchAndFormat` for each word:

```json
{ "searchText": "mange", "fontColor": "#228B22" }
```

```json
{ "searchText": "court", "fontColor": "#228B22" }
```

Result: only those words are colored, nothing else changes.

---

#### WORKFLOW A — Inline syntax in `insertContent` (for full rewrites with formatting)

Use ONLY when writing NEW content. Embed formatting directly into the `content` string:

| Effect        | Syntax                             | Example                                         |
| ------------- | ---------------------------------- | ----------------------------------------------- |
| **Color**     | `[color:#HEX]text[/color]`         | `[color:#228B22]important[/color]` → green text |
| **Bold**      | `**text**`                         | `**critical**`                                  |
| **Italic**    | `*text*`                           | `*note*`                                        |
| **Underline** | `__text__`                         | `__key term__`                                  |
| **Highlight** | Not in markdown — use Workflow B   |                                                 |
| **Combined**  | `[color:#CC0000]**error**[/color]` | Red + bold                                      |

Common hex colors: green `#228B22`, red `#CC0000`, blue `#1F4E79`, orange `#D86000`, purple `#7030A0`

---

#### WORKFLOW B — `applyTaggedFormatting` (advanced 2-step workflow)

Use this when Workflow C is not sufficient (e.g., formatting complex tagged spans with mixed styles).

**Step 1** — Insert content with `<yourTag>` around words to format:

```json
{
  "content": "La <highlight>conquête spatiale</highlight> a souvent été racontée…",
  "location": "Replace",
  "target": "Body"
}
```

**Step 2** — Call `applyTaggedFormatting` to convert the tags to real formatting:

```json
{
  "tagName": "highlight",
  "color": "#228B22",
  "bold": true
}
```

You can pass any combination of: `color`, `bold`, `italic`, `underline`, `strikethrough`, `fontSize`, `fontName`, `highlightColor`, `allCaps`, `superscript`, `subscript`.

---

> ⚠️ **NEVER substitute bold/italic for a requested color.** If the user says "mettre en vert", use `searchAndFormat` with `fontColor`, or `[color:#228B22]` in insertContent. Bold is NOT an acceptable replacement for color.

| Tool                    | When to use                                                          |
| ----------------------- | -------------------------------------------------------------------- |
| `searchAndFormat`       | **PREFERRED** — Apply formatting to specific words without replacing |
| `formatText`            | Apply formatting to the user's current selection only                |
| `applyTaggedFormatting` | Apply formatting to tagged spans (advanced 2-step workflow)          |
| `applyStyle`            | Apply Word named styles (Heading 1, Title, Quote…)                   |
| `setParagraphFormat`    | Set alignment, spacing, indentation                                  |

### For TABLES:

| Tool              | When to use       |
| ----------------- | ----------------- |
| `modifyTableCell` | Edit cell content |
| `addTableRow`     | Insert new row    |
| `addTableColumn`  | Insert new column |
| `formatTableCell` | Style table cells |

### ESCAPE HATCH:

| Tool          | When to use                                          |
| ------------- | ---------------------------------------------------- |
| `eval_wordjs` | **LAST RESORT** — Only when no dedicated tool exists |

## TOOL SELECTION DECISION TREE

```
User wants to apply formatting to specific words (color, bold, highlight...)?
  YES → Use `searchAndFormat` (Workflow C — one call per word/phrase)

User wants to modify existing TEXT content?
  YES: Is it a simple word/phrase replacement?
    YES → Use `searchAndReplace`
    NO (rewriting paragraphs) → Use `proposeRevision`

User wants to add NEW content?
  YES → Use `insertContent` with Markdown syntax (Workflow A for formatting)

Other:
  Formatting on user's active selection only → `formatText`
  Comments → `addComment`
  Tables → table tools
  None of above → `eval_wordjs`
```

## searchAndFormat vs proposeRevision vs insertContent

**Use searchAndFormat when:**

- User wants to apply formatting (color, bold, highlight, etc.) to specific words
- Examples: "mettre les verbes en vert", "surligner les erreurs", "mettre en gras les dates"
- The TEXT content does not change, only the formatting
- Call once per word/phrase to format

**Use proposeRevision when:**

- Editing existing text content (fix, correct, improve, rewrite, edit)
- You want to preserve existing formatting on unchanged portions
- Track Changes will show what was modified

**Use insertContent when:**

- Adding completely new content that doesn't exist yet
- Creating tables or lists from scratch
- User says "add", "insert", "create", "write"
- NEVER use to modify existing text (causes full replacement visible in Track Changes)

## WORD-SPECIFIC API PATTERNS

### Getting selection

```javascript
const range = context.document.getSelection()
range.load('text,font/bold,font/size,font/color,font/name')
await context.sync()
```

### Searching text

```javascript
const results = context.document.body.search('find this', {
  matchCase: false,
  matchWholeWord: true,
})
results.load('items')
await context.sync()

for (const item of results.items) {
  item.insertText('replace with', 'Replace')
}
await context.sync()
```

### Working with paragraphs

```javascript
const paragraphs = context.document.body.paragraphs
paragraphs.load('items,items/text')
await context.sync()

const firstPara = paragraphs.items[0]
```

### Inserting HTML (for complex content)

```javascript
const range = context.document.getSelection()
range.insertHtml('<b>Bold</b> and <i>italic</i>', 'Replace')
await context.sync()
```

## COMMON ERRORS AND FIXES

### Error: "The property 'text' is not available"

**Fix**: Add `range.load('text')` before `context.sync()`

### Error: "Cannot read items of undefined"

**Fix**: Add `.load('items')` to the collection

### Error: "The operation failed because the object doesn't exist"

**Fix**: Re-acquire the range reference after structural changes

### Error: Empty text when selection exists

**Fix**: Check and inform user to select text content
