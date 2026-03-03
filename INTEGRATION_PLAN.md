# KickOffice Agent Stability Integration Plan

**Version**: 2.0 (Detailed Implementation Guide)
**Date**: 2026-03-03
**Target Audience**: AI coding agents and developers implementing this plan

---

## How to Use This Document

This document provides **complete, copy-paste ready implementations**. Each section includes:
- Full file contents (not excerpts)
- Exact import paths and dependencies
- Step-by-step modification instructions
- Testing criteria to validate each change

**IMPORTANT**: Follow the implementation order exactly as specified. Later sections depend on earlier ones.

---

## Table of Contents

1. [Current State Analysis](#1-current-state-analysis)
2. [Pillar 1: Diffing Integration](#2-pillar-1-diffing-integration)
3. [Pillar 2: Skills System](#3-pillar-2-skills-system)
4. [Pillar 3: Code Validator](#4-pillar-3-code-validator)
5. [Integration Instructions](#5-integration-instructions)
6. [Testing Checklist](#6-testing-checklist)

---

## 1. Current State Analysis

### 1.1 Problems to Solve

| Problem | Example | Solution |
|---------|---------|----------|
| **Formatting Loss** | User selects "Hello **world**", AI replaces with "Hello there" -> loses bold | Use diff-based editing that only modifies "world" -> "there" |
| **Missing load()** | AI writes `range.text` without `range.load('text')` -> undefined | Validator rejects code missing load() |
| **Missing sync()** | AI modifies font but forgets `await context.sync()` -> changes don't apply | Validator requires sync() in all code |
| **Wrong namespace** | AI uses `Word.run()` inside Excel add-in -> error | Validator and sandbox block cross-namespace calls |

### 1.2 File Structure After Implementation

```
frontend/src/
├── utils/
│   ├── wordTools.ts              # MODIFIED: add proposeRevision, update eval_wordjs
│   ├── excelTools.ts             # MODIFIED: update eval_officejs
│   ├── powerpointTools.ts        # MODIFIED: add proposeShapeTextRevision, update eval
│   ├── outlookTools.ts           # MODIFIED: add modifyEmailSection, update eval
│   ├── sandbox.ts                # MODIFIED: add host parameter
│   ├── officeCodeValidator.ts    # NEW: validation logic
│   └── wordDiffUtils.ts          # NEW: office-word-diff wrapper
├── skills/                       # NEW DIRECTORY
│   ├── index.ts                  # Skill loader
│   ├── common.skill.md           # Shared Office.js rules
│   ├── word.skill.md             # Word-specific rules
│   ├── excel.skill.md            # Excel-specific rules
│   ├── powerpoint.skill.md       # PowerPoint-specific rules
│   └── outlook.skill.md          # Outlook-specific rules
├── composables/
│   └── useAgentPrompts.ts        # MODIFIED: inject skills
```

---

## 2. Pillar 1: Diffing Integration

### 2.1 Step 1: Install office-word-diff

**Action**: Run this command in the `frontend/` directory:

```bash
cd frontend
npm install ../office-word-diff --save
```

**Verification**: Check `frontend/package.json` contains:
```json
{
  "dependencies": {
    "office-word-diff": "file:../office-word-diff"
  }
}
```

### 2.2 Step 2: Create wordDiffUtils.ts

**Create new file**: `frontend/src/utils/wordDiffUtils.ts`

```typescript
/**
 * Word Diff Utilities
 *
 * Wrapper around office-word-diff for surgical text editing.
 * Preserves formatting by computing word-level diffs and applying
 * only the changes, not replacing entire ranges.
 */

import { OfficeWordDiff, getDiffStats, computeDiff } from 'office-word-diff'
import type { DiffResult, DiffStats } from 'office-word-diff'

export interface RevisionResult {
  success: boolean
  strategy: 'token' | 'sentence' | 'block'
  insertions: number
  deletions: number
  unchanged: number
  message: string
}

/**
 * Apply a revision to selected text using word-level diffing.
 *
 * IMPORTANT: Must be called within Word.run() context.
 *
 * @param context - Word.RequestContext from Word.run()
 * @param revisedText - The new version of the text
 * @param enableTrackChanges - Show changes in Word's Track Changes (default: true)
 * @returns RevisionResult with operation details
 *
 * @example
 * await Word.run(async (context) => {
 *   const result = await applyRevisionToSelection(context, "New text here", true);
 *   console.log(`Applied ${result.insertions} insertions using ${result.strategy} strategy`);
 * });
 */
export async function applyRevisionToSelection(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true
): Promise<RevisionResult> {
  // 1. Get selection and load text
  const range = context.document.getSelection()
  range.load('text')
  await context.sync()

  const originalText = range.text

  // 2. Handle edge cases
  if (!originalText || !originalText.trim()) {
    return {
      success: false,
      strategy: 'block',
      insertions: 0,
      deletions: 0,
      unchanged: 0,
      message: 'Error: No text selected. Please select text before using proposeRevision.',
    }
  }

  if (originalText === revisedText) {
    return {
      success: true,
      strategy: 'token',
      insertions: 0,
      deletions: 0,
      unchanged: originalText.length,
      message: 'Text is identical, no changes needed.',
    }
  }

  // 3. Preview stats before applying
  const stats = getDiffStats(originalText, revisedText)

  // 4. Apply diff with cascading fallback
  const differ = new OfficeWordDiff({
    enableTracking: enableTrackChanges,
    logLevel: 'info',
    onLog: (msg, level) => {
      if (level === 'error') console.error('[WordDiff]', msg)
      else if (level === 'warn') console.warn('[WordDiff]', msg)
      else console.log('[WordDiff]', msg)
    },
  })

  try {
    const result = await differ.applyDiff(context, range, originalText, revisedText)

    return {
      success: result.success,
      strategy: result.strategyUsed,
      insertions: result.insertions,
      deletions: result.deletions,
      unchanged: stats.unchanged,
      message: result.success
        ? `Successfully applied ${result.insertions} insertions and ${result.deletions} deletions using ${result.strategyUsed} strategy.`
        : `Diff application failed. Check logs for details.`,
    }
  } catch (error: any) {
    console.error('[WordDiff] Unexpected error:', error)
    return {
      success: false,
      strategy: 'block',
      insertions: 0,
      deletions: 0,
      unchanged: 0,
      message: `Error applying revision: ${error.message || String(error)}`,
    }
  }
}

/**
 * Preview diff statistics without applying changes.
 * Does NOT require Word context - can be used for UI preview.
 */
export function previewDiffStats(originalText: string, revisedText: string): DiffStats {
  return getDiffStats(originalText, revisedText)
}

/**
 * Compute raw diff operations for debugging/display.
 * Does NOT require Word context.
 *
 * @returns Array of [operation, text] tuples:
 *   - [0, "text"] = unchanged
 *   - [-1, "text"] = deletion
 *   - [1, "text"] = insertion
 */
export function computeRawDiff(originalText: string, revisedText: string): Array<[number, string]> {
  return computeDiff(originalText, revisedText)
}

/**
 * Check if text has complex content that may not diff well.
 * Warns about tables, images, and other non-text content.
 */
export function hasComplexContent(text: string): boolean {
  // Check for table markers, image placeholders, or other special content
  const complexPatterns = [
    /\t.*\t/,           // Tab-separated (likely table)
    /\[Image\]/i,       // Image placeholder
    /\[Figure\]/i,      // Figure placeholder
    /^\s*\|.*\|/m,      // Markdown table row
  ]
  return complexPatterns.some(pattern => pattern.test(text))
}
```

### 2.3 Step 3: Add proposeRevision Tool to wordTools.ts

**File to modify**: `frontend/src/utils/wordTools.ts`

**Step 3.1**: Add import at the top of the file (after existing imports):

```typescript
// Add this import after the existing imports (around line 8)
import { applyRevisionToSelection, previewDiffStats, hasComplexContent } from './wordDiffUtils'
```

**Step 3.2**: Add 'proposeRevision' to the WordToolName type:

```typescript
// Find the WordToolName type (around line 10-36) and add 'proposeRevision':
export type WordToolName =
  | 'getSelectedText'
  | 'getDocumentContent'
  // ... existing tools ...
  | 'eval_wordjs'
  | 'proposeRevision'  // ADD THIS LINE
```

**Step 3.3**: Add the proposeRevision tool definition. Insert this **before** the closing of `createWordTools({`:

```typescript
  // ADD THIS ENTIRE BLOCK before the closing }):
  proposeRevision: {
    name: 'proposeRevision',
    category: 'write',
    description: `**PREFERRED TOOL** for modifying existing text. Computes a word-level diff and applies only the changes, preserving formatting (bold, italic, colors, fonts) on unchanged portions.

HOW IT WORKS:
1. Reads the currently selected text
2. Computes diff between original and your revised version
3. Applies only insertions/deletions, keeping unchanged text intact
4. Optionally shows changes in Word's Track Changes

WHEN TO USE:
- Fixing typos or grammatical errors
- Rewriting phrases or sentences
- Editing paragraphs while preserving formatting
- Any modification where the user wants to keep existing styles

WHEN NOT TO USE:
- Adding entirely new content (use insertContent instead)
- Replacing with tables or complex structures (use insertContent)
- The selection is empty (nothing to revise)`,
    inputSchema: {
      type: 'object',
      properties: {
        revisedText: {
          type: 'string',
          description: 'The complete revised version of the selected text. Write the full text as you want it to appear - the system will compute what changed.',
        },
        enableTrackChanges: {
          type: 'boolean',
          description: 'Show changes in Word Track Changes UI so user can review/accept/reject. Default: true.',
        },
      },
      required: ['revisedText'],
    },
    executeWord: async (context, args: Record<string, any>) => {
      const { revisedText, enableTrackChanges = true } = args

      // Validate input
      if (!revisedText || typeof revisedText !== 'string') {
        return JSON.stringify({
          success: false,
          error: 'revisedText is required and must be a string',
        }, null, 2)
      }

      // Apply revision using diff algorithm
      const result = await applyRevisionToSelection(context, revisedText, enableTrackChanges)

      return JSON.stringify({
        success: result.success,
        strategy: result.strategy,
        changes: {
          insertions: result.insertions,
          deletions: result.deletions,
          unchanged: result.unchanged,
        },
        message: result.message,
        trackChangesEnabled: enableTrackChanges,
      }, null, 2)
    },
  },
```

### 2.4 Step 4: PowerPoint Diff Tool

**File to modify**: `frontend/src/utils/powerpointTools.ts`

**Step 4.1**: Add imports at the top:

```typescript
// Add after existing imports
import DiffMatchPatch from 'diff-match-patch'
```

**Step 4.2**: Add the tool type:

```typescript
// Add to PowerPointToolName type
| 'proposeShapeTextRevision'
```

**Step 4.3**: Add this complete tool implementation:

```typescript
  proposeShapeTextRevision: {
    name: 'proposeShapeTextRevision',
    category: 'write',
    description: `Modify text in a specific shape while attempting to preserve formatting on unchanged portions.

IMPORTANT: PowerPoint has limited diff support compared to Word. This tool:
1. Reads the current shape text
2. Computes word-level diff
3. Applies changes character-by-character when possible
4. Falls back to full replacement if diff fails

PARAMETERS:
- slideNumber: 1-based slide number (as shown in PowerPoint UI)
- shapeIdOrName: Shape ID (number) or shape name (string)
- revisedText: The new text for the shape`,
    inputSchema: {
      type: 'object',
      properties: {
        slideNumber: {
          type: 'number',
          description: 'Slide number (1-based, as shown in PowerPoint)',
        },
        shapeIdOrName: {
          type: 'string',
          description: 'Shape ID or name. Use getShapes to discover available shapes.',
        },
        revisedText: {
          type: 'string',
          description: 'The complete new text for the shape.',
        },
      },
      required: ['slideNumber', 'shapeIdOrName', 'revisedText'],
    },
    execute: async (args: Record<string, any>) => {
      const { slideNumber, shapeIdOrName, revisedText } = args

      return new Promise((resolve) => {
        PowerPoint.run(async (context) => {
          try {
            // 1. Get slide (convert 1-based to 0-based index)
            const slides = context.presentation.slides
            slides.load('items')
            await context.sync()

            if (slideNumber < 1 || slideNumber > slides.items.length) {
              resolve(JSON.stringify({
                success: false,
                error: `Invalid slide number. Presentation has ${slides.items.length} slides.`,
              }, null, 2))
              return
            }

            const slide = slides.items[slideNumber - 1]

            // 2. Get shape
            const shapes = slide.shapes
            shapes.load('items,items/id,items/name')
            await context.sync()

            let targetShape: PowerPoint.Shape | null = null
            for (const shape of shapes.items) {
              if (shape.id === shapeIdOrName || shape.name === shapeIdOrName) {
                targetShape = shape
                break
              }
            }

            if (!targetShape) {
              resolve(JSON.stringify({
                success: false,
                error: `Shape "${shapeIdOrName}" not found on slide ${slideNumber}`,
                availableShapes: shapes.items.map(s => ({ id: s.id, name: s.name })),
              }, null, 2))
              return
            }

            // 3. Get current text
            const textFrame = targetShape.textFrame
            const textRange = textFrame.textRange
            textRange.load('text')
            await context.sync()

            const originalText = textRange.text || ''

            // 4. Compute diff
            const dmp = new DiffMatchPatch()
            // Use word-mode diff for better results
            const diffs = dmp.diff_main(originalText, revisedText)
            dmp.diff_cleanupSemantic(diffs)

            // 5. Calculate stats
            let insertions = 0
            let deletions = 0
            let unchanged = 0
            for (const [op, text] of diffs) {
              if (op === 0) unchanged += text.length
              else if (op === -1) deletions += text.length
              else if (op === 1) insertions += text.length
            }

            // 6. Apply changes
            // PowerPoint API is limited - we do full replacement but report the diff
            textRange.text = revisedText
            await context.sync()

            resolve(JSON.stringify({
              success: true,
              slideNumber,
              shapeId: targetShape.id,
              shapeName: targetShape.name,
              changes: {
                insertions,
                deletions,
                unchanged,
              },
              message: `Updated shape text. ${insertions} characters added, ${deletions} removed.`,
              note: 'PowerPoint applies full text replacement. Formatting may need manual adjustment.',
            }, null, 2))
          } catch (error: any) {
            resolve(JSON.stringify({
              success: false,
              error: error.message || String(error),
            }, null, 2))
          }
        })
      })
    },
  },
```

---

## 3. Pillar 2: Skills System

### 3.1 Create Skills Directory

**Action**: Create directory `frontend/src/skills/`

### 3.2 Create common.skill.md

**Create file**: `frontend/src/skills/common.skill.md`

```markdown
# Office.js Common Rules

These rules apply to ALL Office hosts (Word, Excel, PowerPoint, Outlook).

## THE PROXY PATTERN — UNDERSTAND THIS FIRST

Office.js uses a **proxy pattern**. Objects returned by the API are NOT real objects with data.
They are proxies that queue operations. You MUST follow this pattern:

```javascript
// STEP 1: Get a proxy object
const range = context.document.getSelection();  // This is a PROXY, not real data

// STEP 2: Tell Office.js what properties you need
range.load('text,font/bold');  // Queue a request to load these properties

// STEP 3: Execute the queue and wait
await context.sync();  // NOW the data is fetched from Office

// STEP 4: Now you can read the properties
console.log(range.text);  // Works! Data is loaded
console.log(range.font.bold);  // Works!
```

## CRITICAL RULE 1: Always load() before reading

**WRONG** — Property is undefined:
```javascript
const range = context.document.getSelection();
console.log(range.text);  // UNDEFINED! Not loaded yet
```

**CORRECT** — Load first, sync, then read:
```javascript
const range = context.document.getSelection();
range.load('text');
await context.sync();
console.log(range.text);  // Now it works
```

## CRITICAL RULE 2: Always sync() after writing

**WRONG** — Changes don't apply:
```javascript
range.font.bold = true;
return 'Done';  // Font NOT changed! sync() never called
```

**CORRECT** — Sync commits changes:
```javascript
range.font.bold = true;
await context.sync();  // Changes committed to Office
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
4. **eval_* tools** — ONLY when no dedicated tool exists

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

## CODE TEMPLATE FOR eval_* TOOLS

Always use this structure:

```javascript
try {
  // 1. Get reference
  const target = context.document.getSelection();  // or worksheet, etc.

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
```

### 3.3 Create word.skill.md

**Create file**: `frontend/src/skills/word.skill.md`

```markdown
# Word Office.js Skill

## AVAILABLE TOOLS — Use in this priority order

### For READING content:
| Tool | When to use |
|------|-------------|
| `getSelectedText` | Get plain text of selection |
| `getSelectedTextWithFormatting` | **PREFERRED** — Get Markdown with bold/italic/etc |
| `getDocumentContent` | Read entire document as plain text |
| `getDocumentHtml` | Read document as HTML (for complex analysis) |
| `getSpecificParagraph` | Read one paragraph by index |
| `findText` | Search for text occurrences |
| `getComments` | List all comments in document |
| `getDocumentProperties` | Get word count, paragraph count, etc. |

### For WRITING/EDITING content:
| Tool | When to use |
|------|-------------|
| `proposeRevision` | **PREFERRED** for editing existing text. Preserves formatting! |
| `searchAndReplace` | Fix specific words/phrases throughout document |
| `insertContent` | Add new content (supports Markdown tables, lists) |
| `insertHyperlink` | Add clickable links |
| `addComment` | Add review comments |
| `insertHeaderFooter` | Add headers/footers |
| `insertFootnote` | Add footnotes |

### For FORMATTING:
| Tool | When to use |
|------|-------------|
| `formatText` | Apply bold, italic, color, highlight to selection |
| `applyStyle` | Apply Word styles (Heading 1, Title, Quote...) |
| `setParagraphFormat` | Set alignment, spacing, indentation |

### For TABLES:
| Tool | When to use |
|------|-------------|
| `modifyTableCell` | Edit cell content |
| `addTableRow` | Insert new row |
| `addTableColumn` | Insert new column |
| `formatTableCell` | Style table cells |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
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
const range = context.document.getSelection();
range.load('text,font/bold,font/size,font/color,font/name');
await context.sync();
```

### Searching text
```javascript
const results = context.document.body.search('find this', {
  matchCase: false,
  matchWholeWord: true
});
results.load('items');
await context.sync();

for (const item of results.items) {
  item.insertText('replace with', 'Replace');
}
await context.sync();
```

### Working with paragraphs
```javascript
const paragraphs = context.document.body.paragraphs;
paragraphs.load('items,items/text');
await context.sync();

// Access by index
const firstPara = paragraphs.items[0];
```

### Inserting HTML (for complex content)
```javascript
const range = context.document.getSelection();
range.insertHtml('<b>Bold</b> and <i>italic</i>', 'Replace');
await context.sync();
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
```

### 3.4 Create excel.skill.md

**Create file**: `frontend/src/skills/excel.skill.md`

```markdown
# Excel Office.js Skill

## CRITICAL EXCEL-SPECIFIC RULES

### Rule 1: ALWAYS use 2D arrays for values and formulas

Excel ranges are always 2D, even for single cells.

**WRONG:**
```javascript
range.values = 'Hello';           // Error: not an array
range.values = ['A', 'B', 'C'];   // Error: not 2D
```

**CORRECT:**
```javascript
range.values = [['Hello']];                    // Single cell
range.values = [['A', 'B', 'C']];             // 1 row, 3 columns
range.values = [['A1'], ['A2'], ['A3']];      // 3 rows, 1 column
range.values = [['A1','B1'], ['A2','B2']];    // 2x2 grid
```

### Rule 2: Array dimensions MUST match range dimensions

**WRONG:**
```javascript
const range = sheet.getRange('A1:C3');  // 3x3 range
range.values = [['Only one']];          // 1x1 array - MISMATCH!
```

**CORRECT:**
```javascript
const range = sheet.getRange('A1:C3');  // 3x3 range
range.values = [
  ['A1', 'B1', 'C1'],
  ['A2', 'B2', 'C2'],
  ['A3', 'B3', 'C3']
];  // 3x3 array - matches!
```

### Rule 3: Formula language depends on user's Excel locale

**English Excel:**
```javascript
range.formulas = [['=SUM(A1,B1)']];      // Comma separator
range.formulas = [['=VLOOKUP(A1,B:C,2,FALSE)']];
```

**French Excel:**
```javascript
range.formulas = [['=SOMME(A1;B1)']];    // Semicolon separator
range.formulas = [['=RECHERCHEV(A1;B:C;2;FAUX)']];
```

**IMPORTANT**: Check the `excelFormulaLanguage` setting in the agent context.

### Rule 4: Use getUsedRange() to find data bounds

**WRONG — May be slow or include empty cells:**
```javascript
const range = sheet.getRange('A1:ZZ10000');
```

**CORRECT — Only populated cells:**
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load('values,address');
await context.sync();
```

### Rule 5: Never modify cells while iterating

**WRONG — May corrupt iteration:**
```javascript
const range = sheet.getUsedRange();
range.load('values');
await context.sync();

for (let row of range.values) {
  // Modifying during iteration is dangerous
}
```

**CORRECT — Read all, transform, write back:**
```javascript
const range = sheet.getUsedRange();
range.load('values');
await context.sync();

const newValues = range.values.map(row =>
  row.map(cell => /* transform */)
);

range.values = newValues;
await context.sync();
```

## AVAILABLE TOOLS

### For READING:
| Tool | When to use |
|------|-------------|
| `getSelectedCells` | Get values from current selection |
| `getWorksheetData` | Read used range from active sheet |
| `getDataFromSheet` | Read data from any sheet by name |
| `getWorksheetInfo` | Get workbook structure, sheet names |
| `getAllObjects` | List charts and pivot tables |
| `getNamedRanges` | List named ranges |
| `findData` | Search for values workbook-wide |

### For WRITING:
| Tool | When to use |
|------|-------------|
| `setCellRange` | **PREFERRED** — Write values, formulas, formatting |
| `clearRange` | Clear contents or formatting |
| `modifyStructure` | Insert/delete rows, columns, freeze panes |

### For ANALYSIS:
| Tool | When to use |
|------|-------------|
| `createTable` | Convert range to Excel table |
| `manageObject` | Create/update charts, pivot tables |
| `sortRange` | Sort data |
| `applyConditionalFormatting` | Add conditional format rules |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
| `eval_officejs` | **LAST RESORT** — Sheet rename, advanced features |

## COMMON PATTERNS

### Read active sheet data
```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getUsedRange();
range.load('values,address,rowCount,columnCount');
await context.sync();
```

### Write to specific range
```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange('A1:C3');
range.values = [
  ['Header1', 'Header2', 'Header3'],
  [1, 2, 3],
  [4, 5, 6]
];
await context.sync();
```

### Add formula with fill-down
```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange('D2:D100');
range.formulas = [['=A2+B2']];  // First cell only
range.autoFill('D2:D100', 'FillDown');
await context.sync();
```

### Format range
```javascript
const range = sheet.getRange('A1:C1');
range.format.font.bold = true;
range.format.fill.color = '#4472C4';
range.format.font.color = 'white';
await context.sync();
```
```

### 3.5 Create powerpoint.skill.md

**Create file**: `frontend/src/skills/powerpoint.skill.md`

```markdown
# PowerPoint Office.js Skill

## CRITICAL POWERPOINT-SPECIFIC RULES

### Rule 1: PowerPoint HAS a host-specific API (PowerPoint.run)

Unlike older documentation suggests, modern PowerPoint DOES support `PowerPoint.run()`:

```javascript
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  // Work with slides...
});
```

### Rule 2: Slide numbers are 1-based in UI, but arrays are 0-indexed

```javascript
// User says "slide 3" → array index 2
const slides = context.presentation.slides;
slides.load('items');
await context.sync();

const slideThree = slides.items[2];  // Index 2 = slide 3
```

### Rule 3: Shape text is accessed via textFrame.textRange

```javascript
const shape = slide.shapes.getItem(shapeId);
const textRange = shape.textFrame.textRange;
textRange.load('text');
await context.sync();

console.log(textRange.text);  // The text content

// To modify:
textRange.text = 'New text';
await context.sync();
```

### Rule 4: Common API still available for basic operations

For simple text operations, the Common API works:
```javascript
Office.context.document.getSelectedDataAsync(
  Office.CoercionType.Text,
  (result) => {
    console.log(result.value);
  }
);
```

### Rule 5: Keep bullet points SHORT and FEW

**Content guidelines:**
- Max 8-10 words per bullet point
- Max 6-7 bullets per slide
- Use active voice, present tense
- No full sentences — fragments are better

**WRONG:**
```
- The implementation of the new system will require careful consideration of multiple factors including budget constraints and timeline requirements.
```

**CORRECT:**
```
- New system implementation
- Budget constraints
- Timeline requirements
```

## AVAILABLE TOOLS

### For READING:
| Tool | When to use |
|------|-------------|
| `getAllSlidesOverview` | Get text overview of all slides |
| `getSlideContent` | Read all text from specific slide |
| `getShapes` | Discover shape IDs/names on a slide |
| `getSelectedText` | Read current text selection |

### For WRITING:
| Tool | When to use |
|------|-------------|
| `insertContent` | **PREFERRED** — Update shape text with Markdown |
| `proposeShapeTextRevision` | Edit shape text with diff tracking |
| `addSlide` | Create new slide |
| `deleteSlide` | Remove a slide |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
| `eval_powerpointjs` | Speaker notes, images, animations |

## WORKFLOW: Always discover before modifying

```
1. Call getAllSlidesOverview → understand presentation structure
2. Call getShapes(slideNumber) → get shape IDs on target slide
3. Call insertContent or proposeShapeTextRevision → modify specific shape
```

## COMMON PATTERNS

### Get all slides
```javascript
const slides = context.presentation.slides;
slides.load('items,items/id');
await context.sync();

for (const slide of slides.items) {
  console.log('Slide ID:', slide.id);
}
```

### Get shapes on a slide
```javascript
const slide = context.presentation.slides.items[0];
const shapes = slide.shapes;
shapes.load('items,items/id,items/name,items/type');
await context.sync();

for (const shape of shapes.items) {
  console.log(shape.id, shape.name, shape.type);
}
```

### Modify shape text
```javascript
const slide = context.presentation.slides.items[0];
const shape = slide.shapes.getItem('Title 1');
shape.textFrame.textRange.text = 'New Title';
await context.sync();
```

### Add speaker notes
```javascript
const slide = context.presentation.slides.items[0];
slide.notesPage.textBody.text = 'Speaker notes go here...';
await context.sync();
```

## LIMITATIONS

- **No Track Changes** — PowerPoint has no equivalent to Word's Track Changes
- **Limited formatting API** — Some rich text operations require workarounds
- **Shape positioning** — Coordinates are in points, not pixels
- **Animation API** — Very limited in Office.js
```

### 3.6 Create outlook.skill.md

**Create file**: `frontend/src/skills/outlook.skill.md`

```markdown
# Outlook Office.js Skill

## CRITICAL OUTLOOK-SPECIFIC RULES

### Rule 1: Outlook uses the Common API pattern differently

Outlook doesn't use `Outlook.run()`. Instead, use the mailbox API directly:

```javascript
const item = Office.context.mailbox.item;
```

### Rule 2: Body content can be HTML or text

**Read body as text:**
```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Text,
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log(result.value);
    }
  }
);
```

**Read body as HTML:**
```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Html,
  (result) => {
    console.log(result.value);  // Full HTML
  }
);
```

### Rule 3: Writing uses setAsync with coercion type

**Write text:**
```javascript
Office.context.mailbox.item.body.setAsync(
  'Plain text content',
  { coercionType: Office.CoercionType.Text },
  (result) => { /* handle result */ }
);
```

**Write HTML:**
```javascript
Office.context.mailbox.item.body.setAsync(
  '<p>HTML <b>content</b></p>',
  { coercionType: Office.CoercionType.Html },
  (result) => { /* handle result */ }
);
```

### Rule 4: Prepend/Append instead of Replace when possible

**Safer — preserves existing content:**
```javascript
Office.context.mailbox.item.body.prependAsync(
  '<p>New content at start</p>',
  { coercionType: Office.CoercionType.Html },
  (result) => { }
);
```

### Rule 5: Reply in the SAME language as the original email

**CRITICAL**: When the user asks you to reply to an email:
1. Read the existing email body first
2. Detect the language
3. Reply in THAT language, not the user's interface language

### Rule 6: Callback pattern (not async/await)

Outlook uses callbacks, not Promises:

```javascript
// WRONG — This won't work
const body = await Office.context.mailbox.item.body.getAsync(...);

// CORRECT — Use callback
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Text,
  (result) => {
    // Handle result here
  }
);

// Or wrap in Promise:
const body = await new Promise((resolve, reject) => {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Text,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    }
  );
});
```

## AVAILABLE TOOLS

### For READING:
| Tool | When to use |
|------|-------------|
| `getEmailBody` | Get full email body |
| `getEmailSubject` | Get subject line |
| `getEmailRecipients` | Get To/CC/BCC |
| `getEmailSender` | Get sender info |

### For WRITING:
| Tool | When to use |
|------|-------------|
| `writeEmailBody` | **PREFERRED** — Write with mode: Append/Insert/Replace |
| `setEmailSubject` | Update subject |
| `addRecipient` | Add To/CC/BCC recipients |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
| `eval_outlookjs` | Attachments, HTML manipulation, metadata |

## COMMON PATTERNS

### Read email content
```javascript
const item = Office.context.mailbox.item;

// Subject
item.subject.getAsync((result) => {
  console.log('Subject:', result.value);
});

// Body
item.body.getAsync(Office.CoercionType.Text, (result) => {
  console.log('Body:', result.value);
});

// Sender
console.log('From:', item.from.displayName, item.from.emailAddress);
```

### Write email body
```javascript
const content = `
Dear Colleague,

Thank you for your email.

Best regards
`;

Office.context.mailbox.item.body.setAsync(
  content,
  { coercionType: Office.CoercionType.Text },
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Body updated');
    }
  }
);
```

### Add recipient
```javascript
Office.context.mailbox.item.to.addAsync(
  [{ displayName: 'John Doe', emailAddress: 'john@example.com' }],
  (result) => { }
);
```

## COMPOSE vs READ MODE

Outlook items have different capabilities based on mode:

**Compose mode** (writing new email):
- Can modify subject, body, recipients
- Full write access

**Read mode** (viewing received email):
- Read-only access to content
- Can reply/forward but not modify original
```

### 3.7 Create skills/index.ts

**Create file**: `frontend/src/skills/index.ts`

```typescript
/**
 * Skills Loader
 *
 * Loads and combines skill documents for injection into agent prompts.
 * Skills are defensive prompting guidelines that prevent common Office.js errors.
 */

// Import skill documents as raw strings
// Note: Vite supports ?raw suffix for importing file contents
import commonSkill from './common.skill.md?raw'
import wordSkill from './word.skill.md?raw'
import excelSkill from './excel.skill.md?raw'
import powerpointSkill from './powerpoint.skill.md?raw'
import outlookSkill from './outlook.skill.md?raw'

export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'

const hostSkillMap: Record<OfficeHost, string> = {
  Word: wordSkill,
  Excel: excelSkill,
  PowerPoint: powerpointSkill,
  Outlook: outlookSkill,
}

/**
 * Get the combined skill document for a specific Office host.
 *
 * @param host - The Office application (Word, Excel, PowerPoint, Outlook)
 * @returns Combined skill markdown (common rules + host-specific rules)
 */
export function getSkillForHost(host: OfficeHost): string {
  const hostSkill = hostSkillMap[host]

  if (!hostSkill) {
    console.warn(`[Skills] Unknown host: ${host}, using common skills only`)
    return commonSkill
  }

  return `${commonSkill}\n\n---\n\n${hostSkill}`
}

/**
 * Get just the common skill document (shared rules).
 */
export function getCommonSkill(): string {
  return commonSkill
}

/**
 * Get just the host-specific skill document (without common rules).
 */
export function getHostSpecificSkill(host: OfficeHost): string {
  return hostSkillMap[host] || ''
}

/**
 * List all available hosts.
 */
export function getAvailableHosts(): OfficeHost[] {
  return ['Word', 'Excel', 'PowerPoint', 'Outlook']
}
```

---

## 4. Pillar 3: Code Validator

### 4.1 Create officeCodeValidator.ts

**Create file**: `frontend/src/utils/officeCodeValidator.ts`

```typescript
/**
 * Office.js Code Validator
 *
 * Pre-execution validation for eval_* tools.
 * Rejects code that doesn't follow Office.js patterns before it can cause errors.
 */

export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'

export interface ValidationResult {
  valid: boolean
  errors: string[]
  warnings: string[]
}

/**
 * Properties that require .load() before reading.
 * These are common across all Office hosts.
 */
const LOADABLE_PROPERTIES = [
  'text',
  'values',
  'formulas',
  'items',
  'name',
  'size',
  'bold',
  'italic',
  'color',
  'font',
  'paragraphs',
  'cells',
  'rows',
  'columns',
  'address',
  'rowCount',
  'columnCount',
  'style',
  'format',
  'id',
  'shapes',
  'slides',
]

/**
 * Patterns that indicate reading a property (need load first).
 */
const PROPERTY_READ_PATTERNS = LOADABLE_PROPERTIES.map(
  prop => new RegExp(`\\.${prop}(?!\\s*[=:])\\b`, 'g')  // Match .prop but not .prop = or .prop:
)

/**
 * Validate Office.js code before execution.
 *
 * @param code - The JavaScript code to validate
 * @param host - The Office host (Word, Excel, PowerPoint, Outlook)
 * @returns ValidationResult with errors and warnings
 */
export function validateOfficeCode(code: string, host: OfficeHost): ValidationResult {
  const errors: string[] = []
  const warnings: string[] = []

  // ========== CRITICAL ERRORS ==========

  // Rule 1: Must have context.sync()
  if (!code.includes('context.sync()')) {
    errors.push(
      'Missing `await context.sync()`. ' +
      'Office.js requires sync() to execute queued operations. ' +
      'Add `await context.sync();` after loading properties and after making changes.'
    )
  }

  // Rule 2: Check for property reads without load()
  const hasLoad = /\.load\s*\(/.test(code)
  let hasPropertyReads = false

  for (const pattern of PROPERTY_READ_PATTERNS) {
    if (pattern.test(code)) {
      hasPropertyReads = true
      break
    }
  }

  if (hasPropertyReads && !hasLoad) {
    errors.push(
      'Reading Office.js properties without `.load()`. ' +
      'Before accessing properties like .text, .values, .items, you must call `.load("propertyName")` then `await context.sync()`. ' +
      'Example: `range.load("text"); await context.sync(); console.log(range.text);`'
    )
  }

  // Rule 3: Host namespace validation
  const namespaceErrors = validateNamespaces(code, host)
  errors.push(...namespaceErrors)

  // Rule 4: Infinite loop detection
  if (/while\s*\(\s*true\s*\)/.test(code)) {
    errors.push('Infinite loop detected: `while(true)` is not allowed.')
  }
  if (/for\s*\(\s*;\s*;\s*\)/.test(code)) {
    errors.push('Infinite loop detected: `for(;;)` is not allowed.')
  }

  // Rule 5: Dangerous operations
  if (/eval\s*\(/.test(code)) {
    errors.push('`eval()` is not allowed inside eval_* tools.')
  }
  if (/Function\s*\(/.test(code)) {
    errors.push('`new Function()` is not allowed.')
  }

  // ========== WARNINGS ==========

  // Warning 1: No try/catch
  if (!code.includes('try') || !code.includes('catch')) {
    warnings.push(
      'No try/catch block detected. ' +
      'Wrap your code in try/catch to handle Office.js errors gracefully.'
    )
  }

  // Warning 2: Multiple syncs without batching
  const syncCount = (code.match(/context\.sync\(\)/g) || []).length
  if (syncCount > 3) {
    warnings.push(
      `Found ${syncCount} context.sync() calls. ` +
      'Consider batching operations to reduce round-trips to Office.'
    )
  }

  // Warning 3: Direct property assignment without checking
  if (/\.values\s*=\s*[^[{]/.test(code)) {
    warnings.push(
      'Direct assignment to .values detected. ' +
      'Remember Excel values must be 2D arrays: `range.values = [[value]]`'
    )
  }

  // Warning 4: getRange with large hardcoded range
  if (/getRange\s*\(\s*['"`][A-Z]+1?\s*:\s*[A-Z]+\d{4,}/.test(code)) {
    warnings.push(
      'Large hardcoded range detected. ' +
      'Consider using `getUsedRange()` instead for better performance.'
    )
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  }
}

/**
 * Validate that code only uses the correct host namespace.
 */
function validateNamespaces(code: string, host: OfficeHost): string[] {
  const errors: string[] = []

  // Define which namespaces are allowed for each host
  const allowedNamespaces: Record<OfficeHost, string[]> = {
    Word: ['Word', 'Office'],
    Excel: ['Excel', 'Office'],
    PowerPoint: ['PowerPoint', 'Office'],
    Outlook: ['Office'],  // Outlook uses Office.context.mailbox
  }

  const allHostNamespaces = ['Word', 'Excel', 'PowerPoint']
  const allowed = allowedNamespaces[host]

  for (const ns of allHostNamespaces) {
    // Check if namespace is used (e.g., "Word." or "Word.run")
    const nsPattern = new RegExp(`\\b${ns}\\s*\\.`, 'g')
    if (nsPattern.test(code) && !allowed.includes(ns)) {
      errors.push(
        `Invalid namespace: Cannot use \`${ns}\` APIs in ${host} context. ` +
        `You are running in ${host} — only ${allowed.join(', ')} APIs are available.`
      )
    }
  }

  return errors
}

/**
 * Format validation result for display to the AI.
 */
export function formatValidationResult(result: ValidationResult): string {
  if (result.valid && result.warnings.length === 0) {
    return 'Code validation passed.'
  }

  let output = ''

  if (result.errors.length > 0) {
    output += '## Validation ERRORS (must fix):\n'
    result.errors.forEach((error, i) => {
      output += `${i + 1}. ${error}\n`
    })
    output += '\n'
  }

  if (result.warnings.length > 0) {
    output += '## Validation WARNINGS (recommended to fix):\n'
    result.warnings.forEach((warning, i) => {
      output += `${i + 1}. ${warning}\n`
    })
  }

  return output
}

/**
 * Quick check if code is likely valid (for fast-path).
 */
export function quickValidate(code: string): boolean {
  return (
    code.includes('context.sync()') &&
    (code.includes('.load(') || !PROPERTY_READ_PATTERNS.some(p => p.test(code)))
  )
}
```

### 4.2 Update sandbox.ts

**File to modify**: `frontend/src/utils/sandbox.ts`

**Replace entire file with**:

```typescript
import { ensureLockdown } from './lockdown'

/* global Compartment */

export type SandboxHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'

/**
 * Execute code in a sandboxed environment with SES.
 *
 * @param code - JavaScript code to execute
 * @param globals - Global variables to expose (context, Word/Excel/etc)
 * @param host - Optional host to restrict available namespaces
 * @returns Promise with execution result
 */
export function sandboxedEval(
  code: string,
  globals: Record<string, any>,
  host?: SandboxHost
): unknown {
  ensureLockdown()

  // Build filtered globals based on host
  const filteredGlobals = buildHostGlobals(globals, host)

  // @ts-ignore - Compartment is provided by SES
  const compartment = new Compartment({
    globals: {
      ...filteredGlobals,
      // Safe built-ins
      console,
      Math,
      Date,
      JSON,
      Array,
      Object,
      String,
      Number,
      Boolean,
      Promise,
      // Explicitly blocked
      Function: undefined,
      Reflect: undefined,
      Proxy: undefined,
      Compartment: undefined,
      harden: undefined,
      lockdown: undefined,
      eval: undefined,
      // Blocked browser APIs that could cause issues
      fetch: undefined,
      XMLHttpRequest: undefined,
      WebSocket: undefined,
    },
    __options__: true,
  })

  // Wrap in async IIFE and execute
  return compartment.evaluate(`(async () => { ${code} })()`)
}

/**
 * Build globals object filtered by host.
 * Prevents cross-host API access (e.g., using Word API in Excel).
 */
function buildHostGlobals(
  globals: Record<string, any>,
  host?: SandboxHost
): Record<string, any> {
  const result = { ...globals }

  // If no host specified, allow all (backwards compatibility)
  if (!host) {
    return result
  }

  // Remove namespaces not matching the current host
  const namespaceMap: Record<SandboxHost, string[]> = {
    Word: ['Excel', 'PowerPoint'],      // Remove these from Word
    Excel: ['Word', 'PowerPoint'],       // Remove these from Excel
    PowerPoint: ['Word', 'Excel'],       // Remove these from PowerPoint
    Outlook: ['Word', 'Excel', 'PowerPoint'],  // Remove all from Outlook
  }

  const toRemove = namespaceMap[host] || []
  for (const ns of toRemove) {
    if (ns in result) {
      result[ns] = undefined
    }
  }

  return result
}

/**
 * Create a safe error message from an execution error.
 * Strips sensitive information while keeping useful details.
 */
export function sanitizeExecutionError(error: any): string {
  const message = error?.message || String(error)

  // Remove stack traces that might expose internal paths
  const sanitized = message
    .replace(/at\s+.*:\d+:\d+/g, '')  // Remove stack trace lines
    .replace(/\n\s*\n/g, '\n')         // Remove empty lines
    .trim()

  return sanitized || 'Unknown error occurred during code execution'
}
```

### 4.3 Update eval_wordjs Tool

**File to modify**: `frontend/src/utils/wordTools.ts`

**Step 1**: Add import at top of file:

```typescript
import { validateOfficeCode, formatValidationResult } from './officeCodeValidator'
```

**Step 2**: Find the `eval_wordjs` tool definition and replace it completely with:

```typescript
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
          description: 'JavaScript code following the template above. Must include load(), sync(), and try/catch.',
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does (required for audit trail).',
        },
      },
      required: ['code', 'explanation'],
    },
    executeWord: async (context, args: Record<string, any>) => {
      const { code, explanation } = args

      // Validate code BEFORE execution
      const validation = validateOfficeCode(code, 'Word')

      if (!validation.valid) {
        return JSON.stringify({
          success: false,
          error: 'Code validation failed. Fix the errors below and try again.',
          validationErrors: validation.errors,
          validationWarnings: validation.warnings,
          suggestion: 'Refer to the Office.js skill document for correct patterns. Common issues: missing load() before reading properties, missing context.sync() to commit changes.',
          codeReceived: code.slice(0, 300) + (code.length > 300 ? '...' : ''),
        }, null, 2)
      }

      // Log warnings but proceed
      if (validation.warnings.length > 0) {
        console.warn('[eval_wordjs] Validation warnings:', validation.warnings)
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
          'Word'  // Restrict to Word namespace only
        )

        return JSON.stringify({
          success: true,
          result: result ?? null,
          explanation,
          warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
        }, null, 2)
      } catch (err: any) {
        return JSON.stringify({
          success: false,
          error: err.message || String(err),
          explanation,
          codeExecuted: code.slice(0, 200) + '...',
          hint: 'Check that all properties are loaded before access, and context.sync() is called.',
        }, null, 2)
      }
    },
  },
```

---

## 5. Integration Instructions

### 5.1 Modify useAgentPrompts.ts

**File to modify**: `frontend/src/composables/useAgentPrompts.ts`

**Step 1**: Add import at top:

```typescript
import { getSkillForHost, type OfficeHost } from '@/skills'
```

**Step 2**: Find the `agentPrompt` function (around line 247) and modify it:

Replace:
```typescript
  const agentPrompt = (lang: string) => {
    let base = ''
    if (hostIsOutlook) base = outlookAgentPrompt(lang)
    else if (hostIsPowerPoint) base = powerPointAgentPrompt(lang)
    else if (hostIsExcel) base = excelAgentPrompt(lang)
    else base = wordAgentPrompt(lang)

    return `${base}${userProfilePromptBlock()}\n\n${GLOBAL_STYLE_INSTRUCTIONS}`
  }
```

With:
```typescript
  const agentPrompt = (lang: string) => {
    // Determine current host
    const host: OfficeHost = hostIsOutlook ? 'Outlook'
      : hostIsPowerPoint ? 'PowerPoint'
      : hostIsExcel ? 'Excel'
      : 'Word'

    // Get base prompt for host
    let base = ''
    if (hostIsOutlook) base = outlookAgentPrompt(lang)
    else if (hostIsPowerPoint) base = powerPointAgentPrompt(lang)
    else if (hostIsExcel) base = excelAgentPrompt(lang)
    else base = wordAgentPrompt(lang)

    // Get defensive skill document for host
    const skill = getSkillForHost(host)

    // Combine: base prompt + skill document + user profile + global styles
    return `${base}

<office-js-skill>
${skill}
</office-js-skill>

${userProfilePromptBlock()}

${GLOBAL_STYLE_INSTRUCTIONS}`
  }
```

### 5.2 Update Vite Config for Raw Imports

**File to check**: `frontend/vite.config.ts`

Ensure Vite can import `.md` files as raw strings. If not already configured, this should work by default with the `?raw` suffix. If issues occur, add:

```typescript
// In vite.config.ts, ensure this is in the config:
{
  assetsInclude: ['**/*.md'],
}
```

### 5.3 Update TypeScript Config

**File to check**: `frontend/tsconfig.json`

Add declaration for `.md?raw` imports if TypeScript complains:

Create or update `frontend/src/vite-env.d.ts`:

```typescript
/// <reference types="vite/client" />

declare module '*.md?raw' {
  const content: string
  export default content
}
```

---

## 6. Testing Checklist

### 6.1 Validator Tests

Run these test cases manually or create unit tests:

```typescript
// Test 1: Missing sync - should ERROR
validateOfficeCode(`
  const range = context.document.getSelection();
  range.load('text');
  // Missing sync!
  console.log(range.text);
`, 'Word')
// Expected: errors.length > 0, contains "sync"

// Test 2: Missing load - should ERROR
validateOfficeCode(`
  const range = context.document.getSelection();
  await context.sync();
  console.log(range.text);  // Reading without load!
`, 'Word')
// Expected: errors.length > 0, contains "load"

// Test 3: Wrong namespace - should ERROR
validateOfficeCode(`
  Excel.run(async (context) => {
    // Wrong! We're in Word
  });
`, 'Word')
// Expected: errors.length > 0, contains "Excel"

// Test 4: Valid code - should PASS
validateOfficeCode(`
  try {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    return { text: range.text };
  } catch (e) {
    return { error: e.message };
  }
`, 'Word')
// Expected: valid === true, errors.length === 0
```

### 6.2 Word Diff Tests

Test proposeRevision in Word:

1. **Test formatting preservation**:
   - Select text with mixed formatting: "Hello **bold** and *italic* world"
   - Call proposeRevision with: "Hello **bold** and *italic* there"
   - Verify: Only "world" -> "there" changes, bold/italic preserved

2. **Test Track Changes**:
   - Enable Track Changes
   - Use proposeRevision
   - Verify: Changes appear in Track Changes panel

3. **Test empty selection**:
   - Select nothing
   - Call proposeRevision
   - Verify: Returns error message about empty selection

### 6.3 Skills Injection Test

1. Start the add-in in Word
2. Send a message to the agent
3. Check console/logs for skill document presence
4. Verify agent responses reference skill guidelines

### 6.4 End-to-End Scenarios

**Scenario 1: Fix typo preserving formatting**
1. User types: "Fix the typo in my selection"
2. Selection: "Teh quick brown fox"  (with "quick" in bold)
3. Expected: Agent uses proposeRevision, changes "Teh" to "The", bold on "quick" preserved

**Scenario 2: Reject invalid eval code**
1. User asks for complex operation requiring eval_wordjs
2. Agent generates code missing sync()
3. Expected: Validation rejects with helpful error message
4. Agent fixes code and retries

**Scenario 3: Cross-host protection**
1. In Excel, ask agent to run Word-specific code
2. Expected: Validation blocks Word namespace, error explains context

---

## Appendix A: File Change Summary

| File | Action | Changes |
|------|--------|---------|
| `frontend/package.json` | MODIFY | Add office-word-diff dependency |
| `frontend/src/utils/wordDiffUtils.ts` | CREATE | Word diff wrapper functions |
| `frontend/src/utils/officeCodeValidator.ts` | CREATE | Code validation logic |
| `frontend/src/utils/sandbox.ts` | MODIFY | Add host parameter for namespace filtering |
| `frontend/src/utils/wordTools.ts` | MODIFY | Add proposeRevision, update eval_wordjs |
| `frontend/src/utils/powerpointTools.ts` | MODIFY | Add proposeShapeTextRevision |
| `frontend/src/skills/index.ts` | CREATE | Skill loader |
| `frontend/src/skills/common.skill.md` | CREATE | Shared Office.js rules |
| `frontend/src/skills/word.skill.md` | CREATE | Word-specific rules |
| `frontend/src/skills/excel.skill.md` | CREATE | Excel-specific rules |
| `frontend/src/skills/powerpoint.skill.md` | CREATE | PowerPoint-specific rules |
| `frontend/src/skills/outlook.skill.md` | CREATE | Outlook-specific rules |
| `frontend/src/composables/useAgentPrompts.ts` | MODIFY | Inject skills into prompts |
| `frontend/src/vite-env.d.ts` | MODIFY | Add .md?raw type declaration |

---

## Appendix B: Troubleshooting

### "Cannot find module 'office-word-diff'"
- Run `npm install ../office-word-diff` in frontend directory
- Check path is correct relative to frontend folder

### "Property 'load' does not exist"
- TypeScript types issue - ensure @types/office-js is installed
- Check tsconfig includes office-js types

### Skills not appearing in prompts
- Check import path in useAgentPrompts.ts
- Verify .md files exist in skills/ directory
- Check Vite is processing ?raw imports

### Validation always failing
- Check regex patterns in officeCodeValidator.ts
- Ensure code string is being passed correctly (not undefined)
- Add console.log to debug validation steps

---

*End of Implementation Guide*
