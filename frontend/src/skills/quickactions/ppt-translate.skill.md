---
name: Traduire la slide
description: "Traduit toutes les shapes textuelles de la slide active entre français et anglais en modifiant directement le contenu via les outils PowerPoint. Préserve la mise en forme."
host: powerpoint
executionMode: agent
icon: Globe
actionKey: ppt-translate
---

# PowerPoint Translate Quick Action Skill

## Purpose

Translate the selected text on the current PowerPoint slide and inject the translation **directly into the slide** — not in the chat — while **preserving all formatting** (bold, colors, font sizes, bullet levels, indentation).

## Target Language

**Detect the language of the selected text** — do NOT rely on the `[UI language]` header for direction.

- If the text is **primarily in French** → translate to **English**
- If the text is **primarily in English** → translate to **French**
- If in another language → translate to **French** (default)

The `[UI language: X]` header tells you which language to use for your chat confirmation message.

## Required Workflow — OOXML-based (preserves all formatting)

### Step 1 — Get current slide index

```json
{ "tool": "getCurrentSlideIndex" }
```

Returns: `{ "slideIndex": N }` (1-based)

### Step 2 — Find the target shape and get its XML

Use `eval_powerpointjs` to enumerate text shapes and get the XML of the matching shape:

```javascript
try {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  const slide = slides.items[slideIndex - 1]; // 0-based
  const shapes = slide.shapes;
  shapes.load('items/id,items/name,items/type');
  await context.sync();

  const result = [];
  for (const sh of shapes.items) {
    let text = '';
    try {
      sh.textFrame.textRange.load('text');
      await context.sync();
      text = sh.textFrame.textRange.text || '';
    } catch { /* non-text shape */ }
    if (text.trim()) result.push({ id: sh.id, name: sh.name, text });
  }
  return { success: true, shapes: result };
} catch (e) { return { success: false, error: e.message }; }
```

Match the shape whose text most closely matches the selected text from `<document_content>`.

### Step 3 — Get the shape XML

Use `eval_powerpointjs` to extract the raw XML of the matched shape:

```javascript
try {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  const slide = slides.items[slideIndex - 1];
  const shapes = slide.shapes;
  shapes.load('items/id');
  await context.sync();

  // Find shape by id
  const shape = shapes.items.find(s => s.id === targetShapeId);
  if (!shape) return { success: false, error: 'Shape not found' };

  shape.load('id');
  await context.sync();

  // Get the shape XML via OOXML
  const range = shape.textFrame.textRange;
  range.load('ooxml');
  await context.sync();
  return { success: true, ooxml: range.ooxml };
} catch (e) { return { success: false, error: e.message }; }
```

### Step 4 — Translate by modifying XML text nodes only

Parse the OOXML and translate **only the text content** inside `<a:t>` tags. Preserve everything else:
- Keep `<a:r>` (text runs) structure intact
- Keep `<a:rPr>` (run properties: bold, italic, color, font size, etc.) untouched
- Keep `<a:pPr>` (paragraph properties: bullet level, indentation) untouched
- Keep `<a:p>` (paragraph) structure intact — same number of paragraphs

**Translation strategy:**
- Collect all `<a:t>` text content grouped by paragraph (each `<a:p>` = one paragraph)
- Translate paragraph by paragraph, preserving the 1:1 paragraph mapping
- Do NOT merge or split paragraphs
- The translated text must have exactly the **same number of paragraphs** as the original

### Step 5 — Build the modified XML

Replace only the text content of each `<a:t>` element in the OOXML with the translated text.

If a paragraph has multiple runs (`<a:r>`) and you translate the full paragraph, distribute the translated text across the runs proportionally, or put it all in the first run of the paragraph and clear the rest (keeping empty `<a:t/>` in other runs).

**Simpler approach**: if each paragraph has only one run, just replace `<a:t>original text</a:t>` with `<a:t>translated text</a:t>`.

### Step 6 — Write back the modified XML

Use `eval_powerpointjs` to set the modified OOXML back to the shape:

```javascript
try {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  const slide = slides.items[slideIndex - 1];
  const shapes = slide.shapes;
  shapes.load('items/id');
  await context.sync();

  const shape = shapes.items.find(s => s.id === targetShapeId);
  if (!shape) return { success: false, error: 'Shape not found' };

  // Set modified OOXML back
  shape.textFrame.textRange.ooxml = modifiedOoxml;
  await context.sync();
  return { success: true };
} catch (e) { return { success: false, error: e.message }; }
```

### Step 7 — Confirm in chat

Briefly confirm in the UI language: e.g. "✅ Translated 3 paragraphs to English — formatting preserved."

Do NOT show the full translated text in the chat — it is already in the slide.

---

## Fallback — if OOXML approach fails

If the OOXML read/write fails (e.g., the property is not available in this version of Office):

1. Use `searchAndReplaceInShape` **one call per paragraph** — never per full multi-paragraph block
2. Preserve bullet structure: if original has 5 bullet points, translation must have 5 bullet points
3. Note in the confirmation that formatting may not be fully preserved

---

## Rules

- **OOXML first** — always try the XML approach to preserve formatting
- **Inject into the slide** — do NOT just show the translation in the chat
- **One paragraph = one XML `<a:p>`** — never merge paragraphs during translation
- **Preserve run properties** — never touch `<a:rPr>`, `<a:pPr>`, `<a:lstStyle>`, `<a:bodyPr>`
- **Skip non-text shapes**: charts, images, OLE objects
- **Translate full matched shape** — not just the selected substring

## Example

**User message:**
```
[UI language: French → Target language: English]

<document_content>
Proposition Kickmaker
Fort de son expérience en intégration d'outils d'IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude.
</document_content>
```

**Workflow:**
1. `getCurrentSlideIndex` → slide 3
2. `eval_powerpointjs` → find shape id=36 with matching text
3. `eval_powerpointjs` → get shape OOXML (contains `<a:t>Proposition Kickmaker</a:t>` etc.)
4. Translate paragraph by paragraph:
   - "Proposition Kickmaker" → "Kickmaker Proposal"
   - "Fort de son expérience..." → "Drawing on its experience integrating cutting-edge AI tools..."
5. Replace `<a:t>` text nodes in XML, keep all `<a:rPr>` and `<a:pPr>` unchanged
6. Write modified OOXML back via `eval_powerpointjs`
7. Confirm: "✅ Translated 2 paragraphs to English — formatting preserved."
