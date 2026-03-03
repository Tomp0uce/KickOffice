# office-word-diff

A library for applying word-level text diffs to Microsoft Word documents using the Office.js API, preserving formatting and enabling granular track changes.

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)

> **Note:** This library is experimental and has not been thoroughly tested in production environments. Use at your own risk and please report any issues you encounter.
> **Project history**: This library was extracted from a private codebase and open-sourced as a standalone project in Jan 2026.

## Documentation

- [Overview](README.md) - You are here
- [Token Map Strategy Deep Dive](TOKEN-MAP-STRATEGY.md) - How word-level diffing is implemented using the MS JS API.

## Requirements

This library **requires** the Microsoft Office JavaScript API (Office.js) and must run within:

- A Word Add-in (task pane, content add-in, or dialog)
- Office Online (Word on the web)
- Office desktop applications with add-in support

**This is NOT a general-purpose diff library.** It directly manipulates Word documents through the Office.js API.

## Features

- **Word-level granular diffs** — Preserve formatting while applying precise changes
- **Track changes support** — Changes appear in Word's native Track Changes UI
- **Cascading fallback** — Token → Sentence → Block replacement strategies
- **Preview support** — Compute diffs without Office.js context
- **Detailed operation logs** — Full transparency into applied changes

## Installation

### From GitHub

```bash
npm install github:yuch85/office-word-diff
```

Or add to your `package.json`:

```json
{
  "dependencies": {
    "office-word-diff": "github:yuch85/office-word-diff"
  }
}
```

You can also pin to a specific commit or tag:

```bash
# Specific tag
npm install github:yuch85/office-word-diff#v1.0.0

# Specific commit
npm install github:yuch85/office-word-diff#abc1234
```

## Quick Start

```javascript
import { OfficeWordDiff } from 'office-word-diff';

// Inside your Word Add-in
await Word.run(async (context) => {
  const range = context.document.getSelection();
  range.load('text');
  await context.sync();
  
  const originalText = range.text;
  const newText = "Your modified text here";
  
  const differ = new OfficeWordDiff({
    enableTracking: true,  // Show changes in Track Changes
    logLevel: 'info'       // 'silent', 'error', 'warn', 'info', 'debug'
  });
  
  const result = await differ.applyDiff(context, range, originalText, newText);
  
  console.log(`Applied ${result.insertions} insertions and ${result.deletions} deletions`);
  console.log(`Strategy used: ${result.strategyUsed}`);
  console.log(`Completed in ${result.duration}ms`);
});
```

## Strategy Cascade

The library uses a cascading fallback approach:

```
1. Token Map Strategy (word-level precision)
         ↓ (if fails)
2. Sentence Diff Strategy (sentence-level)
         ↓ (if fails)
3. Block Replace Strategy (full replacement)
```

### Token Map Strategy
- Maps individual words/tokens 1:1 to their Word.Range objects
- Preserves character-level formatting
- Best for: Minor edits, word substitutions, small changes

### Sentence Diff Strategy
- Tokenizes by sentence boundaries (`. ` or `.  `)
- More robust for structural changes
- Best for: Paragraph rewrites, moderate restructuring

### Block Replace Strategy
- Deletes entire range, inserts new text
- Final fallback when precision isn't possible
- Best for: Complete rewrites where formatting preservation isn't critical

## API Reference

### OfficeWordDiff Class

#### Constructor

```javascript
const differ = new OfficeWordDiff(options);
```

**Options:**
| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `enableTracking` | `boolean` | `true` | Enable Word track changes |
| `logLevel` | `string` | `'info'` | Log level: `'silent'`, `'error'`, `'warn'`, `'info'`, `'debug'` |
| `onLog` | `function` | `null` | Custom log handler: `(message, level) => void` |

#### applyDiff(context, range, originalText, newText)

Apply diff to a Word range. **Must be called within `Word.run()`.**

**Returns:** `Promise<DiffResult>`

```typescript
interface DiffResult {
  success: boolean;
  strategyUsed: 'token' | 'sentence' | 'block';
  insertions: number;
  deletions: number;
  duration: number;
  logs: Array<{ timestamp: number, level: string, message: string }>;
}
```

#### computeDiff(text1, text2)

Compute a word-level diff without applying to document. **No Office.js context required.**

```javascript
const diffs = differ.computeDiff('Hello world', 'Hello there');
// Returns: [[0, 'Hello '], [-1, 'world'], [1, 'there']]
```

**Returns:** `Array<[number, string]>` where:
- `[0, text]` = unchanged
- `[-1, text]` = deletion
- `[1, text]` = insertion

#### getDiffStats(text1, text2)

Get statistics about the diff. **No Office.js context required.**

```javascript
const stats = differ.getDiffStats('Hello world', 'Hello there');
// { insertions: 1, deletions: 1, unchanged: 1, totalChanges: 2 }
```

#### getLogs() / clearLogs() / setLogLevel(level)

Log management methods.

### Convenience Functions

```javascript
import { applyWordDiff, computeDiff, getDiffStats } from 'office-word-diff';

// Apply diff (shorthand)
const result = await applyWordDiff(context, range, oldText, newText, options);

// Compute diff (no Office.js needed)
const diffs = computeDiff(text1, text2);

// Get stats (no Office.js needed)
const stats = getDiffStats(text1, text2);
```

### Advanced: Direct Strategy Access

```javascript
import { 
  applyTokenMapStrategy, 
  applySentenceDiffStrategy, 
  applyBlockReplaceStrategy 
} from 'office-word-diff';

// Use a specific strategy directly
await applyTokenMapStrategy(context, range, originalText, newText, logCallback);
```

## Use Cases

- **AI text editing** — Apply GPT/Claude suggestions while preserving formatting
- **Collaborative editing** — Sync external edits with track changes
- **Grammar/style checkers** — Apply corrections granularly
- **Template processing** — Update document sections programmatically

## Example: AI Text Editor

```javascript
import { OfficeWordDiff } from 'office-word-diff';

async function applyAISuggestion(aiGeneratedText) {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    
    if (!range.text.trim()) {
      console.error('Please select some text first');
      return;
    }
    
    const differ = new OfficeWordDiff({ 
      enableTracking: true,
      logLevel: 'debug',
      onLog: (msg, level) => console.log(`[${level}] ${msg}`)
    });
    
    // Preview changes first
    const stats = differ.getDiffStats(range.text, aiGeneratedText);
    console.log(`Preview: ${stats.insertions} insertions, ${stats.deletions} deletions`);
    
    // Apply the diff
    const result = await differ.applyDiff(context, range, range.text, aiGeneratedText);
    
    if (result.success) {
      console.log(`Changes applied using ${result.strategyUsed} strategy`);
    } else {
      console.error('Failed to apply changes');
    }
  });
}
```

## How It Works

1. **Diff Computation**: Uses Google's diff-match-patch algorithm extended with word-level tokenization
2. **Token Mapping**: Maps each word to its corresponding `Word.Range` object
3. **Two-Pass Application**: 
   - Pass 1: Identify and queue deletions
   - Pass 2: Identify and queue insertions
4. **Atomic Execution**: Apply all queued operations with track changes enabled
5. **Fallback Cascade**: If mapping fails, try less granular strategies

## Limitations

- Requires Office.js (Word JavaScript API 1.4+)
- Best results with plain text or simple formatting
- Complex formatting (nested tables, SmartArt) may need manual review
- Track changes requires supported Office environment

## Dependencies

- **diff-match-patch** — Google's diff algorithm (bundled with word-mode extension)

## License

This project is licensed under the [Apache License 2.0](LICENSE).

## Contributing

Contributions are welcome! Please ensure any contributions are compatible with the Apache-2.0 license.

## Related

- [Office.js Word API Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
- [diff-match-patch](https://github.com/google/diff-match-patch)
