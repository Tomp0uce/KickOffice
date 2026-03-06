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
| `insertContent`      | Add new content (supports Markdown tables, lists)              |
| `insertHyperlink`    | Add clickable links                                            |
| `addComment`         | Add review comments                                            |
| `insertHeaderFooter` | Add headers/footers                                            |
| `insertFootnote`     | Add footnotes                                                  |

### For FORMATTING:

**CRITICAL RULE FOR FORMATTING**: The `formatText` tool ONLY works on the user's active selection. If you just inserted text (e.g., from a generated response or a PDF), it is NOT selected. You CANNOT format it with `formatText`.
To format newly inserted text or apply colors without a selection:

1. Use `applyTaggedFormatting` by inserting text wrapped in `<format>` tags: `<format color='red'>text</format>` or `<format bold='true' size='14'>Heading</format>`.
2. Or use `searchAndReplace` to find the text and apply formatting by passing the `format` object.

| Tool                 | When to use                                       |
| -------------------- | ------------------------------------------------- |
| `formatText`         | Apply bold, italic, color, highlight to selection |
| `applyStyle`         | Apply Word styles (Heading 1, Title, Quote...)    |
| `setParagraphFormat` | Set alignment, spacing, indentation               |

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
│   └─ NO (rewriting sentences/paragraphs) → Use `proposeRevision`
└─ NO: Adding new content?
    ├─ YES → Use `insertContent`
    └─ NO: Something else?
        ├─ Formatting → Use `formatText` or `applyStyle`
        ├─ Comments → Use `addComment`
        ├─ Tables → Use table tools
        └─ None of above → Use `eval_wordjs`
```

## proposeRevision vs insertContent

**Use proposeRevision when:**

- Editing existing text
- User says "fix", "correct", "improve", "rewrite", "edit"
- You want to preserve formatting on unchanged portions
- Track Changes would be helpful for user review

**Use insertContent when:**

- Adding completely new content
- Creating tables, lists from scratch
- Appending to document
- User says "add", "insert", "create", "write"

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

// Access by index
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

**Cause**: Forgot to load() before reading
**Fix**: Add `range.load('text')` before `context.sync()`

### Error: "Cannot read items of undefined"

**Cause**: Trying to access collection without loading
**Fix**: Add `.load('items')` to the collection

### Error: "The operation failed because the object doesn't exist"

**Cause**: Range was deleted or document structure changed
**Fix**: Re-acquire the range reference after structural changes

### Error: Empty text when selection exists

**Cause**: User selected image/table, not text
**Fix**: Check and inform user to select text content
