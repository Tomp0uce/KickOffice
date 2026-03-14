# Office.js Common Rules

These rules apply to ALL Office hosts (Word, Excel, PowerPoint, Outlook).

## THE PROXY PATTERN — UNDERSTAND THIS FIRST

Office.js uses a **proxy pattern**. Objects returned by the API are NOT real objects with data.
They are proxies that queue operations. You MUST follow this pattern:

```javascript
// STEP 1: Get a proxy object
const range = context.document.getSelection(); // This is a PROXY, not real data

// STEP 2: Tell Office.js what properties you need
range.load('text,font/bold'); // Queue a request to load these properties

// STEP 3: Execute the queue and wait
await context.sync(); // NOW the data is fetched from Office

// STEP 4: Now you can read the properties
console.log(range.text); // Works! Data is loaded
console.log(range.font.bold); // Works!
```

## CRITICAL RULE 1: Always load() before reading

**WRONG** — Property is undefined:

```javascript
const range = context.document.getSelection();
console.log(range.text); // UNDEFINED! Not loaded yet
```

**CORRECT** — Load first, sync, then read:

```javascript
const range = context.document.getSelection();
range.load('text');
await context.sync();
console.log(range.text); // Now it works
```

## CRITICAL RULE 2: Always sync() after writing

**WRONG** — Changes don't apply:

```javascript
range.font.bold = true;
return 'Done'; // Font NOT changed! sync() never called
```

**CORRECT** — Sync commits changes:

```javascript
range.font.bold = true;
await context.sync(); // Changes committed to Office
return 'Done';
```

## CRITICAL RULE 3: Use try/catch for ALL Office.js code

**WRONG** — Errors crash silently:

```javascript
const range = context.document.getSelection();
range.load('text');
await context.sync();
// If something fails, no error handling
```

**CORRECT** — Catch and report errors:

```javascript
try {
  const range = context.document.getSelection();
  range.load('text');
  await context.sync();
  return { success: true, text: range.text };
} catch (error) {
  return { success: false, error: error.message };
}
```

## CRITICAL RULE 4: Check for empty selections

Many operations fail on empty selections. Always check:

```javascript
const range = context.document.getSelection();
range.load('text');
await context.sync();

if (!range.text || range.text.trim() === '') {
  return { error: 'No text selected. Please select text first.' };
}
```

## CRITICAL RULE 5: Prefer dedicated tools over eval

Priority order when choosing tools:

1. **Dedicated read tools** (getDocumentContent, getSelectedText, etc.)
2. **Dedicated write tools** (insertContent, searchAndReplace, etc.)
3. **Format tools** (formatText, applyStyle, etc.)
4. **eval\_\* tools** — ONLY when no dedicated tool exists

## NEVER DO THESE THINGS

1. **Never use percentages for font sizes** — Use points (pt)
   - WRONG: `range.font.size = '150%'`
   - CORRECT: `range.font.size = 18`

2. **Never insert Unicode bullets for lists**
   - WRONG: `range.insertText('• Item 1\n• Item 2')`
   - CORRECT: Use HTML lists via insertHtml() or insertContent with Markdown

3. **Never assume collections are loaded**
   - WRONG: `body.paragraphs.items[0]` without load
   - CORRECT: `paragraphs.load('items'); await context.sync();`

4. **Never modify while iterating**
   - WRONG: `for (const item of items) { item.delete() }`
   - CORRECT: Collect items first, then modify in reverse order

## CODE TEMPLATE FOR eval\_\* TOOLS

Always use this structure:

```javascript
try {
  // 1. Get reference
  const target = context.document.getSelection(); // or worksheet, etc.

  // 2. Load required properties
  target.load('text,font/bold,font/size');
  await context.sync();

  // 3. Validate
  if (!target.text) {
    return { success: false, error: 'No selection' };
  }

  // 4. Perform operations
  target.font.bold = true;

  // 5. Commit and return
  await context.sync();
  return { success: true, result: 'Formatting applied' };
} catch (error) {
  return { success: false, error: error.message };
}
```
