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

Call `proposeDocumentRevision` with one entry per **paragraph that contains corrections**:

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

### Step 4 — Confirm in UI language

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
- **NEVER use `proposeRevision`** — it replaces the entire selection as one change
- **NEVER use `searchAndReplace`** — it does not create Track Changes
- **NEVER rewrite or rephrase** — only fix genuine errors
- **Preserve the author's voice** — do not change sentence structure unless required for correctness
- **Preserve inline formatting** in `revisedText`: keep `**bold**`, `*italic*`, `__underline__` markers exactly as they appear in `originalText` — only correct the text content inside them
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
