---
name: Réviser le doc
description: "Révise le document Word sélectionné et propose des améliorations sur le fond, la clarté et la cohérence via Track Changes. Fournit aussi un commentaire de synthèse sur la qualité globale."
host: word
executionMode: agent
icon: Eye
actionKey: word-review
---

# Document Review Quick Action Skill (Word)

## Purpose

Read the full Word document and produce a structured review in the chat proposing concrete improvement axes — without modifying the document.

## When to Use

- User clicks "Review Document" Quick Action in Word
- Goal: Provide editorial/strategic feedback on the entire document (structure, clarity, coherence, argumentation, tone, style)
- This is a **read-only, chat-response action** — do NOT use Track Changes, do NOT call `proposeDocumentRevision`

## Input Contract

- **Document content**: Full text retrieved via tools
- **Language**: **ALWAYS respond in the UI language specified at the start of the user message as `[UI language: ...]`.** If it says `[UI language: Français]`, your entire review must be in French. This overrides the document language.
- **Mode**: Agent loop — read the document, then respond in chat

## Workflow

### STEP 1 — Read the full document

Call `getSelectedTextWithFormatting` to retrieve the document content:

```json
{ "includeFormatting": false }
```

If the document is empty, respond with a short message saying there is no content to review.

### STEP 2 — Analyse (internal reasoning)

Before responding, assess the document across these dimensions:

1. **Structure & organisation**: Is there a clear introduction, body, and conclusion? Are sections logically ordered?
2. **Clarity & conciseness**: Are sentences clear and direct? Is there unnecessary repetition or wordiness?
3. **Coherence & flow**: Do ideas connect smoothly? Are transitions between paragraphs logical?
4. **Argumentation & evidence**: Are claims well-supported? Is the reasoning convincing?
5. **Tone & register**: Is the tone appropriate for the document type (formal, professional, academic, etc.)? Is it consistent?
6. **Completeness**: Are there obvious gaps, missing sections, or underdeveloped points?

### STEP 3 — Produce the review in chat

Write a structured review with **4–6 improvement axes**, each with:
- A short title
- 2–3 sentences of specific, actionable feedback referencing actual content from the document
- A concrete suggestion on how to address it

**Format:**

```markdown
## Document Review

### 1. [Title of axis]
[Specific observation from the document + why it's an issue]
**Suggestion**: [Concrete, actionable recommendation]

### 2. [Title of axis]
...
```

End with a one-sentence overall assessment (strengths + priority area).

## Constraints

- **DO NOT modify the document** (no `proposeDocumentRevision`, no `proposeRevision`, no `addComment`)
- **DO NOT** rewrite sections verbatim — describe what to improve and how
- Keep the review focused and actionable (not a comprehensive essay)
- Reference specific passages to make feedback concrete: "In the second paragraph of section X..."
- If the document is short (< 150 words), still apply all axes but scale the depth accordingly

## Quality Check

- ✓ Full document read before responding?
- ✓ Each axis has a concrete suggestion?
- ✓ Response is in the UI language?
- ✓ No document modification attempted?
