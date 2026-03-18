---
name: Corriger (Word)
description: "Corrige l'orthographe et la grammaire du document Word sélectionné via Track Changes. Chaque correction est proposée avec proposeRevision — acceptable ou rejectable individuellement."
host: word
executionMode: agent
icon: CheckSquare
actionKey: word-proofread
---

# Word Proofread Quick Action Skill

## Purpose

Find and fix spelling, grammar, and punctuation errors in the selected Word text — **chirurgically,
directly in the document**, with full native Track Changes so the user can review every correction.

## Key Principle: Chirurgical Paragraph-Level Corrections

- Only fix paragraphs that contain real errors — leave correct paragraphs untouched
- Each corrected paragraph becomes **one independent Track Change** in the Review pane
- Preserve ALL formatting: bold, italic, font sizes, colors, and structure
- Never rewrite — only correct actual errors

## Required Workflow

### Step 1 — Read the selected text

Call `getSelectedTextWithFormatting` to get the selected text with formatting context:

```json
{ "tool": "getSelectedTextWithFormatting" }
```

### Step 2 — Identify errors per paragraph

Go through each paragraph and identify:
- **Spelling**: typos, missing accents (`expérienc` → `expérience`)
- **Grammar**: agreement errors (`gestion strict` → `gestion stricte`), tense, subject-verb
- **Punctuation**: missing or wrong punctuation, incorrect apostrophes
- **Capitalization**: proper nouns, sentence starts

**Do NOT fix**:
- Intentional stylistic choices
- Brand names, proper nouns, technical terms
- Code snippets or technical syntax
- Deliberately informal language (unless clearly wrong)

### Step 3 — Apply corrections with Track Changes

Call `proposeDocumentRevision` with one entry per **paragraph that contains corrections**.

**CRITICAL — originalText rules:**
- `originalText` MUST be **plain text only** — the exact paragraph text as it appears in the document
- **NEVER include formatting markers** like `[color:#CC0000]`, `**bold**`, `*italic*`, `__underline__` in `originalText`
- Strip ALL `[color:...]` markers from `originalText` — they are display annotations, not real document text
- Only the actual readable words go in `originalText`

**CRITICAL — revisedText rules:**
- `revisedText` MUST also be **plain text** — no markdown headers (`#`), no formatting markers
- Only correct the text content — do not change paragraph structure or add markup

**Example** — if `getSelectedTextWithFormatting` returned:
```
Dans les années 1960, les grandes puissances [color:#CC0000]se sont lancé[/color] dans une course vers l'espace.
```
The `originalText` MUST be:
```
Dans les années 1960, les grandes puissances se sont lancé dans une course vers l'espace.
```
(All `[color:...]` markers stripped — plain text only)

```json
{
  "revisions": [
    {
      "originalText": "Fort de son expérienc en intégration d'outils IA de pointe dans un environneme sécurisé.",
      "revisedText": "Fort de son expérience en intégration d'outils IA de pointe dans un environnement sécurisé."
    },
    {
      "originalText": "La gestion strict des données est une condition sine qua non.",
      "revisedText": "La gestion stricte des données est une condition sine qua non."
    }
  ],
  "enableTrackChanges": true
}
```

**Only include paragraphs that actually have corrections.** Paragraphs without errors are NOT included.

### Step 4 — If proposeDocumentRevision returns NOT FOUND errors

If some paragraphs are NOT FOUND:
1. Call `getDocumentContent` to read the full document as plain text
2. Find the exact plain-text paragraph in the document (no markers)
3. Retry `proposeDocumentRevision` with the correct plain-text `originalText`
4. **NEVER fall back to `proposeRevision`** — it replaces the entire selection as one giant Track Change, making it impossible to review corrections individually

### Step 5 — Confirm in UI language

After the corrections are applied, respond briefly in the `[UI language: X]` language:
- Number of paragraphs corrected
- List the specific corrections made (original → corrected)
- Remind the user they can accept/reject each change in the Review pane

Example (French UI):
```
✅ 2 paragraphes corrigés :
- "expérienc" → "expérience"
- "environneme" → "environnement"
- "gestion strict" → "gestion stricte"

Les corrections sont visibles dans le suivi des modifications.
```

## Rules

- **Use `proposeDocumentRevision`** — one entry per corrected paragraph, full corrected text
- **NEVER use `proposeRevision`** — it replaces the entire selection as one change, creating a giant unreadable redline block
- **NEVER use `searchAndReplace`** — it does not create Track Changes
- **NEVER rewrite or rephrase** — only fix genuine errors
- **Preserve the author's voice** — do not change sentence structure unless required for correctness
- **originalText must be PLAIN TEXT** — no formatting markers, no color annotations, no markdown
- **revisedText must be PLAIN TEXT** — no markdown headers (`#`), no formatting markup
- **Skip non-text content**: tables, code blocks, images
- **Match language**: French text stays French, English stays English (do not translate)

## What counts as an error

- Typos: `expérienc` → `expérience`, `environneme` → `environnement`
- Agreement: `gestion strict` → `gestion stricte`, `code importante` → `code important`
- Spelling: `recieve` → `receive`, `occured` → `occurred`
- Accents: `recu` → `reçu`, `systeme` → `système`
- Punctuation: missing period at end of sentence, comma splice, wrong apostrophe

## What NOT to change

- Anglicismes volontaires used consistently in the document
- Technical or domain-specific abbreviations
- Formatting structure (paragraph breaks, bullets)
- Intentional repetition or stylistic patterns
