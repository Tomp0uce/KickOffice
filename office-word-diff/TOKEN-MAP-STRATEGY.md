# Library Overview

This library applies **word-level diffs** to Microsoft Word documents using the **Diff Match Patch (DMP)** algorithm and the **MS Word JavaScript API**. It aims to preserve formatting while tracking granular changes (insertions/deletions) at the word level.

---

## High-Level Flow

1. **Compute diff** using DMP in "word mode" (tokenizes by words instead of characters)
2. **Build a token map** using two mapping strategies (coarse → refined)
3. **Execute changes** in two passes:
   - **Pass 1**: Identify deletions
   - **Pass 2**: Identify insertions
4. **Apply changes atomically** with Track Changes enabled
5. **Fallback** to sentence-level diff if word-level fails

---

## The Two Mapping Strategies

### **1. Coarse Map Strategy**

```javascript
const coarseRanges = range.getTextRanges([" "], false);
```

**Purpose**: Break the document into "coarse" chunks separated by spaces.

**How it works**:
- `getTextRanges([" "], false)` splits the range at every space character
- Returns an array of `Word.Range` objects, each containing text between spaces
- Example: `"Hello world, how are you?"` → `["Hello", "world,", "how", "are", "you?"]`

**Why this step exists**:
- The Word API doesn't provide direct word-level ranges
- This gives approximate word boundaries to work with
- **Limitation**: Punctuation sticks to words (e.g., `"world,"` not `"world"`)

---

### **2. Refined Map Strategy (Batched)**

This is the sophisticated part that fixes the coarse map's limitations.

#### **Step 2a: Tokenize each coarse range**

```javascript
const dmpRegex = /(\w+|[^\w\s]+|\s+)/g;
while ((match = dmpRegex.exec(coarseText)) !== null) {
    const tokenText = match[0];
    // Queue search for this token
}
```

**Purpose**: Break coarse chunks into finer tokens matching DMP's tokenization.

**Regex breakdown**:
- `\w+` → Word characters (e.g., `"Hello"`)
- `[^\w\s]+` → Punctuation (e.g., `","`)
- `\s+` → Whitespace (e.g., `" "`)

**Example refinement**:
- Coarse: `"world,"` 
- Refined: `["world", ","]`

This matches how DMP tokenizes text, ensuring 1:1 alignment.

---

#### **Step 2b: Batch all searches**

```javascript
const searchResults = coarseRange.search(tokenText, { matchCase: true });
searchProxies.push({
    text: tokenText,
    results: searchResults,
    coarseText: coarseText
});
```

**Purpose**: Queue search operations for performance.

**How it works**:
- For each refined token (e.g., `"world"`), search **within its coarse range** only
- Store search proxy objects without executing yet
- **Critical optimization**: Batch all searches, then execute with one `context.sync()`

**Why batching matters**:
- Each `context.sync()` is a network round-trip to Word
- Batching 100 searches → 1 round-trip instead of 100
- Massive performance improvement

---

#### **Step 2c: Execute and build final map**

```javascript
await context.sync(); // Execute all searches at once

for (const proxy of searchProxies) {
    if (proxy.results.items.length > 0) {
        fineTokens.push({
            text: proxy.text,
            range: proxy.results.items[0] // Actual Word.Range object
        });
    }
}
```

**Result**: An ordered array of `{text, range, index}` objects where:
- `text` is the token string (e.g., `"world"`)
- `range` is a live `Word.Range` object pointing to that exact token in the document
- `index` is the position in the array

**This is the "refined token map"** - a precise, ordered list of every token with its actual document location.

---

## Why Two Strategies?

| Strategy | Granularity | Accuracy | Use |
|----------|-------------|----------|-----|
| **Coarse** | Space-delimited chunks | ~70% | Initial subdivision |
| **Refined** | Regex tokens | 99%+ | Final 1:1 mapping |

**The two-step process**:
1. **Coarse** reduces search space (search within chunks, not entire document)
2. **Refined** achieves exact token boundaries matching DMP's diff output

Without the coarse step, searching for common tokens like `"the"` might match hundreds of wrong locations in a large document.

---

## The Two-Pass Execution

### **Pass 1: Identify Deletions**

```javascript
for (const [op, chunk] of diffs) {
    if (op === -1) { // DELETE operation
        const chunkTokens = chunk.match(/(\w+|[^\w\s]+|\s+)/g) || [];
        for (let i = 0; i < count; i++) {
            deleteTargets.push(fineTokens[tokenIndex]);
            tokenIndex++;
        }
    }
}
```

**Purpose**: Walk through diff operations and collect tokens to delete.

**How it works**:
- Iterate through DMP diff output (array of `[operation, text]` tuples)
- `op === -1` means "delete this text"
- Count how many tokens the deleted chunk contains
- Mark that many entries in `fineTokens` for deletion
- **Key insight**: We're building a DELETE QUEUE, not executing yet

---

### **Pass 2: Identify Insertions**

```javascript
const tokensAfterDeletes = fineTokens.filter(t => !deletedIndices.has(t.index));

for (const [op, chunk] of diffs) {
    if (op === 0) { // EQUAL - advance pointer
        lastAnchorRange = token.range;
    } else if (op === 1) { // INSERT
        insertOps.push({
            anchor: lastAnchorRange,
            location: Word.InsertLocation.after,
            text: chunk
        });
    }
}
```

**Purpose**: Determine WHERE to insert new text.

**How it works**:
1. Create a "post-deletion" view of remaining tokens
2. Walk through diffs again, tracking the "last anchor" (last unchanged token)
3. When encountering an INSERT operation, record:
   - **What to insert**: `text: chunk`
   - **Where**: `anchor: lastAnchorRange` + `location: after`
4. Build an INSERT QUEUE

**Why track "last anchor"**:
- Insertions are positioned relative to unchanged tokens
- Example: `"Hello world"` → `"Hello beautiful world"`
  - Last anchor before insert = `"Hello"`
  - Insert `"beautiful "` AFTER `"Hello"`

---

## Atomic Execution

```javascript
// Enable tracking
context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

// Apply deletes (reverse order)
deleteTargets.sort((a, b) => b.index - a.index);
deleteTargets.forEach(token => token.range.delete());

// Apply inserts
insertOps.forEach(op => op.anchor.insertText(op.text, op.location));

// Disable tracking + commit
context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
await context.sync();
```

**Critical details**:

1. **Reverse deletion order**: Delete from END → START to avoid invalidating ranges
2. **Track Changes enabled**: All edits are recorded as Word's tracked changes
3. **Single sync**: All operations batched into one round-trip
4. **"Atomic-ish"**: If `context.sync()` fails, none of the changes persist

---

## Fallback Strategy

```javascript
catch (e) {
    range.insertText(originalText, Word.InsertLocation.replace);
    await context.sync();
    return await applySentenceDiffStrategy(context, range, originalText, newText, log);
}
```

**When it triggers**:
- Token mapping fails (can't find a token in its coarse range)
- Token mismatch during insertion phase

**How it works**:
1. **Reset**: Replace entire range with original text (clean slate)
2. **Fallback**: Use sentence-level diff strategy (coarser granularity, more reliable)

---

## Summary

**Coarse Map**: Space-delimited chunks → reduces search scope  
**Refined Map**: Regex tokens within chunks → exact 1:1 alignment with DMP  
**Two Passes**: Separate planning (identify changes) from execution (apply changes)  
**Batching**: Minimize network round-trips for performance  
**Result**: Granular tracked changes preserving formatting, with robust fallback
