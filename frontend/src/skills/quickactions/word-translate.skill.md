# Word Translate Quick Action Skill

## Purpose

Translate the selected text in a Word document **directly in the document** — paragraph by paragraph,
using native Word Track Changes so the user can review or accept the translation.

## Language Rules

**Detect the language from the selected text itself** — do NOT use the `[UI language]` header for direction.

| Selected text language | Target language |
|------------------------|-----------------|
| French (FR)            | English (EN)    |
| English (EN)           | French (FR)     |
| Other                  | French (default)|

**Response language**: All your chat messages (summaries, confirmations) must be in the language
specified by `[UI language: X]` in the user message — even if it differs from the translated language.

## Required Workflow

### Step 1 — Read the selected text

Call `getSelectedTextWithFormatting` to get the selected text with its Markdown formatting:

```json
{ "tool": "getSelectedTextWithFormatting" }
```

### Step 2 — Detect language and plan translations

- Detect the dominant language of the selected text
- Split the text into individual paragraphs (each paragraph = one revision unit)
- For each paragraph, produce the translated version:
  - Preserve paragraph structure (bullets stay bullets, numbered lists stay numbered)
  - Preserve tone and formality level (formal → formal, casual → casual)
  - Keep proper nouns, brand names, and technical terms as-is
  - Preserve inline formatting: **bold**, *italic* are preserved in `revisedText`

### Step 3 — Inject translations using Track Changes

Call `proposeDocumentRevision` with ONE entry per changed paragraph:

```json
{
  "revisions": [
    {
      "originalText": "Fort de son expérience en intégration d'outils IA de pointe dans un environnement sécurisé, Kickmaker propose de réaliser une première étude.",
      "revisedText": "Drawing on its experience integrating cutting-edge AI tools in a secure environment, Kickmaker proposes conducting an initial study."
    },
    {
      "originalText": "Périmètre initial : deux modules critiques.",
      "revisedText": "Initial scope: two critical modules."
    }
  ],
  "enableTrackChanges": true
}
```

**One revision per paragraph** — each creates an independent Track Change the user can accept/reject.

### Step 4 — Confirm in UI language

After successful injection, reply briefly in the `[UI language]` language:
- What was translated (FR→EN or EN→FR)
- How many paragraphs were updated
- Reminder that Track Changes are active (user can accept/reject in Review pane)

Example (if UI language = French): "✅ 4 paragraphes traduits du français vers l'anglais. Les modifications sont visibles dans le suivi des modifications — vous pouvez les accepter ou les rejeter."

## Rules

- **NEVER** show the full translated text in the chat — it is already in the document as Track Changes.
- **NEVER** call `proposeRevision` — it replaces the entire selection as one change; use `proposeDocumentRevision` for per-paragraph granularity.
- **NEVER** call `insertContent` for translation — it does not create Track Changes.
- Skip empty paragraphs, headers/footers, and single-word paragraphs that are proper nouns.
- If the selected text is only one paragraph, `proposeDocumentRevision` still works with a single-item array.

## Color limitation

`proposeDocumentRevision` does not preserve inline text colors in the inserted (translated) text.
If the original text had colored runs, the Track Change insertion will use the paragraph's default color.
This is a known limitation — colors in the original (deleted) text are still visible in the Track Change diff.
