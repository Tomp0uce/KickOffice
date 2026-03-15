# PowerPoint Translate Quick Action Skill

## Purpose

Translate the selected text on the current PowerPoint slide and inject the translation **directly into the slide** — not in the chat.

## Target Language

**Detect the language of the selected text** — do NOT rely on the `[UI language]` header for direction.

- If the text is **primarily in French** → translate to **English**
- If the text is **primarily in English** → translate to **French**
- If in another language → translate to **French** (default)

The `[UI language: X]` header tells you which language to use for your chat confirmation message.

## Required Workflow

### Step 1 — Get current slide

```json
{ "tool": "getCurrentSlideIndex" }
```

Returns: `{ "slideIndex": N }` (1-based)

### Step 2 — Get shapes with their text

Use `eval_powerpointjs` to enumerate text shapes on the current slide:

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

### Step 3 — Find the target shape

The user message contains the selected text inside `<document_content>` tags.
Match it against the shapes found in Step 2:
- Pick the shape whose text most closely matches the selected text.
- If the selection is a substring, pick the shape that contains it.

### Step 4 — Translate

Translate the **entire text of the matched shape** to the target language:
- Preserve paragraph structure (bullet points, numbered lists, line breaks)
- Preserve tone and formality level
- Keep proper nouns, brand names, and technical terms as-is

### Step 5 — Inject into the slide

Use `searchAndReplaceInShape` once per **paragraph** — replace each original paragraph with its translation:

```json
{
  "slideNumber": 3,
  "shapeIdOrName": "42",
  "searchText": "Original paragraph text here",
  "replaceText": "Translated paragraph text here"
}
```

One call per paragraph. **Do NOT try to replace the entire multi-paragraph block in a single call** — split by paragraph for reliable matching.

## Rules

- **Always inject into the slide** — do NOT just show the translation in the chat.
- **One `searchAndReplaceInShape` call per paragraph** — never try to replace multi-paragraph text in a single call.
- **Preserve paragraph structure**: if the original has 5 bullet points, the translation must also have 5 bullet points.
- **Translate the full matched shape**, not just the literal selection substring.
- **Skip non-text shapes**: charts, images, OLE objects.
- **After all replacements**, briefly confirm in the chat what was translated (e.g. "✅ Translated 3 paragraphs to English."). Do NOT show the full translated text in the chat — it is already in the slide.

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
2. `eval_powerpointjs` → shape id=36 contains "Proposition Kickmaker\nFort de son expérience..."
3. Translate:
   - "Proposition Kickmaker" → "Kickmaker Proposal"
   - "Fort de son expérience en intégration..." → "Drawing on its experience integrating..."
4. Call `searchAndReplaceInShape` for each paragraph:
   ```json
   { "slideNumber": 3, "shapeIdOrName": "36", "searchText": "Proposition Kickmaker", "replaceText": "Kickmaker Proposal" }
   ```
   ```json
   { "slideNumber": 3, "shapeIdOrName": "36", "searchText": "Fort de son expérience en intégration d'outils d'IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude.", "replaceText": "Drawing on its experience integrating cutting-edge AI tools in a secure environment, Kickmaker proposes conducting an initial study." }
   ```
5. Confirm: "✅ Translated 2 paragraphs to English."
