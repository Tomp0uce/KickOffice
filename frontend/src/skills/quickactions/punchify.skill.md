---
name: Punchifier le texte
description: "Retravaille le texte de la slide pour le rendre plus percutant et mémorable. Renforce les verbes, élimine le jargon superflu et maximise l'impact en minimum de mots. Idéal pour les slides de pitch."
host: powerpoint
executionMode: immediate
icon: Zap
actionKey: punchify
---

# Punchify Quick Action Skill (PowerPoint Agent)

## Purpose

Make every text shape on the active slide more impactful, concise, and engaging while
**preserving all visual formatting** (font size, font name, bold, italic, color).

---

## Workflow — MUST follow these steps in order

### Step 1 — Identify the active slide

Call `getCurrentSlideIndex` to get the current 1-based slide number.

### Step 2 — Read the slide content

Use `getSlideContent` to get the text of all shapes on the slide (this is the most reliable method
and works on all shape types including Placeholders):

```json
{ "slideNumber": <1-based slide number> }
```

`getSlideContent` returns shapes as `[Shape N] <id>/<name>\n<text>` blocks. Parse each block to get:
- The shape's ID or name (use the ID, e.g. `36`)
- The text lines (each line = one paragraph)

### Step 3 — Generate punchified text

For each paragraph in each text shape, produce a punchified version:

- **Conciseness**: Reduce word count by 30–50% when possible
- **Impact**: Active voice, strong verbs, concrete nouns
- **No em-dashes (—) or semicolons (;)** — use commas or split into separate bullets
- **Numbers over words**: "three benefits" → "3 benefits"
- **Language**: MUST match the original language exactly — never translate
- **Already short / punchy text**: Leave unchanged (return identical text)

Only produce punchified replacements for paragraphs that actually changed.

### Step 4 — Apply changes with `searchAndReplaceInShape`

For **each paragraph whose text changed**, call `searchAndReplaceInShape` once with:
- `searchText`: the original paragraph text (exact match)
- `replaceText`: the punchified version

```json
{
  "slideNumber": <1-based slide number>,
  "shapeIdOrName": "<shape ID>",
  "searchText": "original paragraph text here",
  "replaceText": "punchified bullet"
}
```

**One call per changed paragraph.** Do NOT batch multiple paragraphs into a single call.

`searchAndReplaceInShape` has an automatic XML fallback for Placeholder shapes (GeneralException)
— just call it normally; formatting is preserved in all cases.

**Do NOT call `replaceShapeParagraphs`** — it fails with GeneralException on Placeholder shapes.
**Do NOT call `proposeShapeTextRevision`** — it destroys all formatting.
**Do NOT call `insertContent`** — it inserts new content instead of replacing existing text.

### Step 5 — Report

After all replacements succeed, output a brief summary of what was changed.
Do NOT ask for confirmation before applying — apply directly.

---

## Example transformation

**`getSlideContent` returns:**
```
[Shape 4] 36/Content Placeholder 35
Besoin client
Stago souhaite explorer le potentiel des outils d'IA générative pour automatiser la génération de documents d'architecture logicielle de leurs instruments médicaux.
Proposition Kickmaker
Fort de son expérience en intégration d'outils IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude.
```

**Punchified replacements for shape 36:**
- "Stago souhaite explorer le potentiel des outils d'IA générative pour automatiser la génération de documents d'architecture logicielle de leurs instruments médicaux." → "Stago veut exploiter l'IA générative pour auto-générer ses documents d'architecture logicielle."
- "Fort de son expérience en intégration d'outils IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude." → "Kickmaker mène une 1re étude : comprendre et documenter des logiciels complexes grâce à l'IA générative."

**Calls:**
```json
{ "slideNumber": 5, "shapeIdOrName": "36", "searchText": "Stago souhaite explorer le potentiel des outils d'IA générative pour automatiser la génération de documents d'architecture logicielle de leurs instruments médicaux.", "replaceText": "Stago veut exploiter l'IA générative pour auto-générer ses documents d'architecture logicielle." }
```
```json
{ "slideNumber": 5, "shapeIdOrName": "36", "searchText": "Fort de son expérience en intégration d'outils IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude.", "replaceText": "Kickmaker mène une 1re étude : comprendre et documenter des logiciels complexes grâce à l'IA générative." }
```

---

## Edge cases

- **Title shapes** (short text): Leave unchanged if already punchy.
- **Footer / date shapes**: Leave unchanged.
- **Single word / number paragraphs**: Leave unchanged.
- **Empty paragraphs**: Skip (do not replace).
