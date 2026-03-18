---
name: Corriger la slide
description: "Corrige les fautes d'orthographe et grammaire sur la slide active en modifiant directement chaque shape concernée via les outils PowerPoint. Préserve tout le formatage visuel."
host: powerpoint
executionMode: agent
icon: CheckCircle
actionKey: ppt-proofread
---

# PowerPoint Proofread Quick Action Skill

## Purpose

Find and fix spelling and grammar errors in all text shapes on the current slide — surgically, without destroying any formatting.

## When to Use

- User clicks "Proofread" Quick Action in PowerPoint
- Works on the CURRENT slide (no text selection required)
- Goal: Correct only the specific erroneous words, preserving ALL formatting (bold, italic, font, colors, etc.)

## CRITICAL: Surgical Replacement Only

**NEVER use `proposeShapeTextRevision` for proofread corrections** — it replaces the entire shape text with a single unstyled run, destroying all bold, italic, font sizes, and colors.

**ALWAYS use `searchAndReplaceInShape`** — it finds and replaces only the specific wrong words within text runs, preserving all other formatting.

## Required Workflow

### Step 1 — Get current slide

```json
{ "tool": "getCurrentSlideIndex" }
```
Returns: `{ "slideIndex": 5 }` (1-based)

### Step 2 — Get shapes on the slide

Use `eval_powerpointjs` to read all shapes including their text (needed because `getShapes` may fail on slides with OLE/chart objects):

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

### Step 3 — Identify errors and apply surgical fixes

For **each typo or grammar error found**, call `searchAndReplaceInShape` once per correction:

```json
{
  "slideNumber": 5,
  "shapeIdOrName": "36",
  "searchText": "expérienc",
  "replaceText": "expérience"
}
```

```json
{
  "slideNumber": 5,
  "shapeIdOrName": "36",
  "searchText": "gestion strict",
  "replaceText": "gestion stricte"
}
```

**One call per correction** — do NOT try to fix multiple errors in one call.

### Step 4 — Screenshot to verify (optional)

```json
{ "tool": "screenshotSlide", "slideNumber": 5 }
```

## Rules

- **Only fix real errors**: spelling, grammar, agreement, accents. Do NOT rewrite.
- **Use `searchAndReplaceInShape`** for every correction — not `proposeShapeTextRevision`, not `eval_powerpointjs` with `tr.text = "..."`.
- **Skip non-text shapes**: charts, images, OLE objects — they have no editable text.
- **Preserve everything else**: do not change wording, structure, or formatting.
- **Match language**: French text stays French, English stays English.
- **XML fallback is automatic**: `searchAndReplaceInShape` now automatically falls back to OOXML XML editing when the `textRuns` API fails (GeneralException on Placeholder shapes). This is fully transparent — just call the tool normally. If it returns `"method": "xml-fallback"` in the result, the correction succeeded via XML (formatting still preserved).

## What counts as an error to fix

- Typos: "expérienc" → "expérience", "environneme" → "environnement"
- Agreement: "gestion strict" → "gestion stricte", "code importante" → "code important"
- Spelling: "recieve" → "receive"
- Accents: "recu" → "reçu"

## What NOT to change

- Anglicismes volontaires (Inputs/Outputs if part of the slide's intentional style)
- Sentence structure (unless there's a clear grammar error)
- Abbreviations, technical terms, brand names
