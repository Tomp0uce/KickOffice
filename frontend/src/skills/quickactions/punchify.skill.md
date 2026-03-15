# Punchify Quick Action Skill (PowerPoint Agent)

## Purpose

Make every text shape on the active slide more impactful, concise, and engaging while
**preserving all visual formatting** (font size, font name, bold, italic, color).

---

## Workflow — MUST follow these steps in order

### Step 1 — Identify the active slide

Call `getCurrentSlideIndex` to get the current 1-based slide number.

### Step 2 — Read every text shape on the slide

Use `eval_powerpointjs` to enumerate the shapes and retrieve their text.
This is required to avoid `InvalidArgument` errors on OLE or chart shapes.

Example code to pass to `eval_powerpointjs`:

```javascript
const slide = context.presentation.slides.getItemAt(SLIDE_INDEX); // 0-based
slide.shapes.load('items/id,items/name,items/type');
await context.sync();
const result = [];
for (const shape of slide.shapes.items) {
  const t = (shape.type || '').toString().toLowerCase();
  if (t === '13' || t.includes('picture') || t === 'ole' || t === 'chart') continue;
  try {
    shape.textFrame.textRange.paragraphs.load('items/textRange/text');
    await context.sync();
    const paragraphs = shape.textFrame.textRange.paragraphs.items.map((p, i) => ({
      index: i,
      text: p.textRange.text,
    }));
    result.push({ id: shape.id, name: shape.name, paragraphs });
  } catch (e) {
    // skip shapes with no text frame
  }
}
return JSON.stringify(result);
```

Replace `SLIDE_INDEX` with `slideNumber - 1` (0-based).

### Step 3 — Generate punchified text

For each paragraph in each text shape, produce a punchified version:

- **Conciseness**: Reduce word count by 30–50% when possible
- **Impact**: Active voice, strong verbs, concrete nouns
- **No em-dashes (—) or semicolons (;)** — use commas or split into separate bullets
- **Numbers over words**: "three benefits" → "3 benefits"
- **Language**: MUST match the original language exactly — never translate
- **Already short / punchy text**: Leave unchanged (return identical text)

### Step 4 — Apply changes with `replaceShapeParagraphs`

For **each shape that has at least one changed paragraph**, call `replaceShapeParagraphs` once,
passing only the paragraphs whose text actually changed.

```json
{
  "slideNumber": <1-based slide number>,
  "shapeIdOrName": "<shape ID or name>",
  "paragraphReplacements": [
    { "paragraphIndex": 0, "newText": "punchified text here" },
    { "paragraphIndex": 2, "newText": "another punchified bullet" }
  ]
}
```

`replaceShapeParagraphs` preserves font name, size, bold, italic and color for each paragraph.
**Do NOT call `proposeShapeTextRevision`** — it destroys all formatting.
**Do NOT call `insertContent`** — it inserts new content instead of replacing existing text.

### Step 5 — Report

After all `replaceShapeParagraphs` calls succeed, output a brief summary of what was changed.
Do NOT ask for confirmation before applying — apply directly.

---

## Example transformation

**Before** (paragraph text):
> "We are going to make improvements to the customer experience by implementing a new feedback system"

**After**:
> "Improve customer experience with new feedback system"

**Before**:
> "Notre équipe a travaillé très dur pour essayer d'augmenter les performances commerciales"

**After**:
> "Booster les performances commerciales grâce à l'équipe"
