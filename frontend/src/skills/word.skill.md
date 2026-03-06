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

| Tool                 | When to use                                                    |
| -------------------- | -------------------------------------------------------------- |
| `proposeRevision`    | **PREFERRED** for editing existing text. Preserves formatting! |
| `searchAndReplace`   | Fix specific words/phrases throughout document                 |
| `insertContent`      | Add new content (Markdown + inline color/style syntax)         |
| `insertHyperlink`    | Add clickable links                                            |
| `addComment`         | Add review comments                                            |
| `insertHeaderFooter` | Add headers/footers                                            |
| `insertFootnote`     | Add footnotes                                                  |

### For FORMATTING:

**CRITICAL RULE**: The `formatText` tool ONLY works when text is already selected by the user. If you just inserted text via `insertContent`, it is NOT selected — you CANNOT color/bold it with `formatText` after the fact.

To apply any formatting (color, bold, italic, underline, highlight, font size…) to newly inserted or existing text, use one of these two workflows:

---

#### WORKFLOW A — Inline syntax in `insertContent` (PREFERRED for full rewrites with formatting)

Embed formatting directly into the `content` string:

| Effect        | Syntax                             | Example                                         |
| ------------- | ---------------------------------- | ----------------------------------------------- |
| **Color**     | `[color:#HEX]text[/color]`         | `[color:#228B22]important[/color]` → green text |
| **Bold**      | `**text**`                         | `**critical**`                                  |
| **Italic**    | `*text*`                           | `*note*`                                        |
| **Underline** | `__text__`                         | `__key term__`                                  |
| **Highlight** | Not in markdown — use Workflow B   |                                                 |
| **Combined**  | `[color:#CC0000]**error**[/color]` | Red + bold                                      |

Common hex colors: green `#228B22`, red `#CC0000`, blue `#1F4E79`, orange `#D86000`, purple `#7030A0`

Example:

```json
{
  "content": "La [color:#228B22]conquête spatiale[/color] a souvent été **racontée** comme une aventure.",
  "location": "Replace",
  "target": "Body"
}
```

---

#### WORKFLOW B — `applyTaggedFormatting` (PREFERRED when not rewriting the whole text)

Use this to apply any formatting to specific words **already in the document** without replacing everything.

**Step 1** — Insert the document with `<yourTag>` around words to format:

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

> ⚠️ **NEVER substitute bold/italic for a requested color.** If the user says "mettre en vert", use `[color:#228B22]` or `applyTaggedFormatting` with `color`. Bold is NOT an acceptable replacement for color.

| Tool                    | When to use                                           |
| ----------------------- | ----------------------------------------------------- |
| `formatText`            | Apply formatting to the user's current selection only |
| `applyTaggedFormatting` | Apply formatting to tagged spans across the document  |
| `applyStyle`            | Apply Word named styles (Heading 1, Title, Quote…)    |
| `setParagraphFormat`    | Set alignment, spacing, indentation                   |

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
User wants to modify existing text?
├─ YES: Is it a simple word/phrase replacement?
│   ├─ YES → Use `searchAndReplace`
│   └─ NO (rewriting paragraphs) → Use `proposeRevision`
└─ NO: Adding new content or rewriting WITH formatting?
    ├─ YES, with color/bold/etc → Use `insertContent` with [color:] / **bold** inline syntax
    ├─ YES, apply formatting to existing doc → Use `applyTaggedFormatting` (Workflow B)
    ├─ Formatting on user's active selection only → Use `formatText`
    ├─ Comments → Use `addComment`
    ├─ Tables → Use table tools
    └─ None of above → Use `eval_wordjs`
```

## proposeRevision vs insertContent

**Use proposeRevision when:**

- Editing existing text (fix, correct, improve, rewrite, edit)
- You want to preserve existing formatting on unchanged portions

**Use insertContent when:**

- Adding completely new content
- Creating tables or lists from scratch
- User says "add", "insert", "create", "write"
- User wants a rewrite **with color/formatting** (use inline syntax)

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
