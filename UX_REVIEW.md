# KickOffice - Usability & UX Review

**Date**: 2026-02-16
**Scope**: End-user experience, workflow analysis, quick action relevance, settings usability
**Perspective**: First-time user, then daily user of the Office add-in

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [First-Time User Experience](#first-time-user-experience)
3. [Daily User Workflow](#daily-user-workflow)
4. [Quick Actions Audit (per host)](#quick-actions-audit-per-host)
5. [Settings Page Usability](#settings-page-usability)
6. [Chat UX Issues](#chat-ux-issues)
7. [Issues by Priority](#issues-by-priority)
8. [Summary Table](#summary-table)

---

## Executive Summary

KickOffice provides a solid feature set but the UX has several friction points that hurt both discoverability and daily usability. The main issues are:

- **Hardcoded French strings** in the main UI break the i18n system
- **Quick action buttons are icon-only** with no text — poor discoverability
- **Agent mode chat has no streaming** — the user sees no feedback until the full response arrives
- **Tool enable/disable in Settings is dead code** — the toggles save to localStorage but HomePage never reads them
- **Built-in prompts customization exposes developer-level concepts** (`${language}`, `${text}` placeholders) to end users
- **PowerPoint "Visual" action has a broken workflow** — generates a text prompt the user must manually copy to Image mode
- **No way to regenerate, edit messages, or preserve chat history**

---

## First-Time User Experience

### What happens when a user opens KickOffice for the first time

1. Task pane opens → "How can I help you today?" with a sparkle icon
2. Backend status indicator (green/red dot) — **Good**, immediately shows if backend is reachable
3. A row of **icon-only buttons** across the top — no text, no labels. User must hover each one to discover what they do. On touch devices (tablets), tooltips don't work at all.
4. A prompt selector dropdown labeled "Prompt" — unclear what it does
5. A model selector labeled **"Type de tâche :"** (hardcoded French, not translated)
6. A textarea with placeholder text
7. Two tiny checkboxes at the bottom

### Problems

| # | Issue | Impact |
|---|-------|--------|
| 1 | Quick actions are **icon-only** — 5 small icons with no text | User doesn't know what the buttons do without hovering. Very poor discoverability. |
| 2 | **"Type de tâche :"** is hardcoded French (`HomePage.vue:166`) | English users see French label. Breaks i18n. |
| 3 | Model selector shows tier labels like "Nano (rapide)", "Standard", "Raisonnement" | User doesn't understand the difference. No tooltip or description explains when to use which. |
| 4 | No **onboarding** or first-use guidance | User dropped into a blank interface with no tutorial or suggested prompts |
| 5 | Empty state shows only static text, no **clickable suggestions** | Compare to ChatGPT/Copilot which show clickable prompt cards |
| 6 | Checkboxes are **tiny** (h-3.5 = 14px) and labels can be clipped | Hard to click, especially on touch. The checkbox label extends outside the clickable area of the `<label>` due to the h-3.5 constraint on the label itself. |

---

## Daily User Workflow

### Typical workflow: "Polish my selected text"

1. User selects text in Word
2. Clicks the "Polish" icon (but must remember which icon it is)
3. Action runs with streaming — result appears progressively — **Good**
4. User clicks "Replace" icon button below the response
5. Text is replaced in the document — **Good**

**Friction**: Step 2 requires memorizing which tiny icon does what. After first use it's fine, but discoverability is poor.

### Typical workflow: "Chat about my document"

1. User types a question in the textarea
2. Presses Enter
3. A message appears: **"⏳ Analyse de la demande..."** (hardcoded French, `HomePage.vue:954`)
4. Then tool calls appear: **"⚡ Action : getSelectedText..."** (hardcoded French, `HomePage.vue:1004`)
5. After the full agent loop completes, the final response appears all at once — **no streaming**
6. User sees Replace/Append/Copy buttons

**Friction**:
- Steps 3-4 show hardcoded French text regardless of UI language
- Step 5 has no streaming — for complex responses this can take 10-30 seconds with no visual feedback other than the tool action log
- The user cannot edit or regenerate a previous message
- If the user closes the pane, the entire conversation is lost

### Typical workflow: "Generate an image"

1. User switches model selector to "Image" tier
2. Textarea placeholder changes to "Describe the image to generate"
3. User types prompt, presses Enter
4. "Generating image..." message appears — **Good**
5. Image appears inline — **Good**
6. User clicks "Replace" or "Copy" → **base64 text is inserted instead of image** (see DESIGN_REVIEW.md C2)

---

## Quick Actions Audit (per host)

### Word: Translate, Polish, Academic, Summary, Grammar

| Action | Icon | Useful? | Notes |
|--------|------|---------|-------|
| **Translate** | Globe | ✅ Yes | Core use case. Works well. |
| **Polish** | Brush | ✅ Yes | Rewrites text more professionally. Very useful. |
| **Academic** | BookOpen | ⚠️ Niche | Only useful for academic/research users. Could be replaced by a more universal "Formal" action. |
| **Summary** | FileCheck | ✅ Yes | Core use case. Works well. |
| **Grammar** | CheckCheck | ✅ Yes | Essential. Works well. |

**Missing quick actions for Word**:
- **Simplify** — Rewrite text to be simpler/clearer (opposite of Academic)
- **Expand** — Make a short text longer with more details
- **Tone** — Change tone (formal/casual/friendly) — subsumes Academic

### Excel: Analyze, Chart, Formula, Transform, Highlight

| Action | Icon | Type | Useful? | Notes |
|--------|------|------|---------|-------|
| **Analyze** | Eye | Agent | ✅ Yes | Analyzes selected data. Good. |
| **Chart** | FunctionSquare | Agent | ✅ Yes | Creates chart from data. Good. |
| **Formula** | FunctionSquare | Draft | ✅ Yes | Pre-fills input for user to describe formula needed. Good flow. |
| **Transform** | Briefcase | Draft | ⚠️ Unclear | The name "Transform" is vague. What does it transform? Label should be more descriptive (e.g., "Restructure Data"). |
| **Highlight** | Eye | Draft | ⚠️ Unclear | Called "Highlighter" in English but "Révélateur" in French. The purpose isn't clear from the name. |

**Missing quick actions for Excel**:
- **Explain** — Explain what a formula or data pattern does (exists as built-in prompt but not as quick action button)
- **Clean** — Clean/normalize data (exists as built-in prompt "clean" in constants but the quick action mapping is unclear)

**Note**: The built-in prompts settings page shows 5 Excel prompts (Analyze, Chart, Formula, Format, Explain) but the quick action bar shows different ones (Analyze, Chart, Formula, Transform, Highlight). This mismatch is confusing.

### PowerPoint: Bullets, Speaker Notes, Punchify, Shrink, Visual

| Action | Icon | Useful? | Notes |
|--------|------|---------|-------|
| **Bullets** | List | ✅ Yes | Converts text to bullet points. Core PPT use case. |
| **Speaker Notes** | StickyNote | ✅ Yes | Generates speaker notes from slide text. Very useful. |
| **Punchify** | Sparkles | ✅ Yes | Makes text more impactful. Good for presentations. |
| **Shrink** | MinusCircle | ✅ Yes | Shortens text. Good for fitting text in slides. |
| **Visual** | Image | ❌ Broken workflow | See below. |

**"Visual" action broken workflow**:
1. User clicks Visual
2. A prefix text is inserted into the input: "Generate an image prompt for a slide about: "
3. User must type what the slide is about, then send
4. The LLM responds with a **text description of an image prompt** (not an image!)
5. User must then: switch to Image mode, copy the prompt, paste it, send again
6. This is a **5-step manual process** for something that should be 1 click

**Proposed fix**: Visual should either (a) directly generate the image by internally chaining to the image model, or (b) be removed and replaced with a simpler "Generate Image" button that asks the user for a description and generates directly.

### Outlook: Reply, Formalize, Concise, Proofread, Extract Tasks

| Action | Icon | Useful? | Notes |
|--------|------|---------|-------|
| **Smart Reply** | Reply | ✅ Yes | Pre-fills "Draft a reply saying that: " — good flow. |
| **Formalize** | Briefcase | ✅ Yes | Makes email more formal. Useful. |
| **Concise** | MinusCircle | ✅ Yes | Shortens email. Good. |
| **Proofread** | CheckCheck | ✅ Yes | Checks grammar/spelling. Good. |
| **Extract Tasks** | CheckCircle | ✅ Yes | Extracts action items. Very useful for long emails. |

**Outlook quick actions are the best designed** — all relevant, clear purpose, good flows.

**Missing**:
- **Friendly** — Make an email less formal/more approachable (opposite of Formalize)

---

## Settings Page Usability

The Settings page has 4 tabs: General, Prompts, Built-in Prompts, Tools.

### General tab — Mostly good

| Setting | Useful? | Notes |
|---------|---------|-------|
| UI Language | ✅ | Essential |
| Reply Language | ✅ | Essential, 13 languages |
| Excel Formula Language | ✅ | Good, only shown in Excel host |
| First Name / Last Name | ⚠️ | Useful for Outlook (email context), but shown in all hosts. Consider hiding in Word/PPT or adding a note. |
| Gender | ⚠️ | Same — mostly relevant for Outlook (gendered greetings in French). |
| Agent Max Iterations | ❌ | **Too technical** for end users. "Agent Max Iterations" is LLM jargon. Should be hidden or renamed to something like "Response depth" with a simple low/medium/high selector. |
| Backend Status | ✅ | Useful, read-only |
| Configured Models | ⚠️ | Read-only list of model IDs. Meaningless to end users ("gpt-5-nano"?). Consider hiding or showing only labels. |

### Prompts tab — Over-engineered for end users

The prompts system asks users to write a **system prompt** and a **user prompt** — this is developer/LLM jargon. A typical Office user doesn't know what a "system prompt" is.

**Proposed simplification**: Replace with "Custom Instructions" — a single text field where users describe how they want the AI to behave (e.g., "Always reply formally", "Use British English"). The app would inject this as the system prompt behind the scenes.

### Built-in Prompts tab — Developer-facing, not user-facing

This tab exposes raw prompt templates with `${language}` and `${text}` placeholders. Issues:
- Only covers Word and Excel prompts — **PowerPoint and Outlook built-in prompts are not customizable**
- The `${language}` / `${text}` placeholder syntax is developer-level. End users will break it (e.g., deleting `${text}` accidentally)
- No validation — if user removes the `${text}` placeholder, the quick action silently stops working

**Proposed simplification**: Either hide this tab entirely (power users can edit prompts via the regular Prompts tab), or add guardrails (placeholder validation, visual tokens instead of raw syntax).

### Tools tab — Dead code, should be removed or wired up

**Critical finding**: The tool enable/disable toggles save to `localStorage('enabledTools')` but `HomePage.vue` **never reads this value** when building the tool list for the agent. The toggles have zero effect on actual behavior.

Either:
1. **Wire it up**: Read `enabledTools` from localStorage in `HomePage.vue` and filter tools accordingly
2. **Remove the tab**: If the feature isn't ready, don't show it to users — it erodes trust

---

## Chat UX Issues

### U1. No streaming in chat mode (agent loop)

Quick actions use `chatStream()` (streaming, progressive text). But all chat messages go through `chatSync()` (non-streaming agent loop). The user types a question and gets:
- "⏳ Analyse de la demande..." → waits
- "⚡ Action : toolName..." → waits more
- Full response appears all at once

For simple questions that don't need tools, this is a degraded experience compared to streaming. Consider: if the model's response contains no tool calls, stream the response instead of waiting for the full sync response.

### U2. No message regeneration

If the AI gives a poor response, the user's only option is to retype the question. A "Regenerate" button on assistant messages is standard in chat UIs.

### U3. No message editing

The user cannot edit a previously sent message and re-send it. This forces retyping for minor corrections.

### U4. No chat history persistence

Closing the task pane or clicking "New Chat" loses the entire conversation with no confirmation dialog. There's no way to go back to a previous conversation.

### U5. No confirmation on "New Chat"

Clicking the "New Chat" button immediately clears the conversation without asking "Are you sure?". One accidental click destroys all context.

### U6. Hardcoded French strings in chat

Three strings in `HomePage.vue` are hardcoded in French, breaking i18n:
- `"Type de tâche :"` (line 166)
- `"⏳ Analyse de la demande..."` (line 954)
- `"⚡ Action : "` (line 1004)

These should use `t()` / i18n keys.

### U7. "Thought process" label not translated

`HomePage.vue:114` — The `<summary>` tag "Thought process" for reasoning model responses is hardcoded in English, not translated.

### U8. Action buttons (Replace/Append/Copy) always visible

Every assistant message shows 3 small icon buttons. In a long conversation, this creates visual clutter. Consider:
- Showing buttons only on hover
- Or showing them only on the last message

### U9. No typing/loading indicator

During agent mode processing, there's no animated typing indicator or skeleton. Just static text messages updating. A subtle animation would reassure the user that processing is happening.

### U10. Checkbox labels clipped

The checkbox area uses `h-3.5` on the label container, which constrains the height to 14px. The text can be clipped or the clickable area is too small, especially on mobile/touch.

---

## Issues by Priority

### HIGH — Significantly impacts daily usability

| ID | Issue | Where |
|----|-------|-------|
| **U-H1** | **Hardcoded French strings** ("Type de tâche", "Analyse de la demande", "Action :") | `HomePage.vue:166,954,1004` |
| **U-H2** | **Quick actions are icon-only** — no text labels, poor discoverability | `HomePage.vue:41-53` |
| **U-H3** | **Tool toggles in Settings are dead code** — saves to localStorage but never read | `SettingsPage.vue` vs `HomePage.vue` |
| **U-H4** | **No streaming in chat mode** — user sees no progressive output | `HomePage.vue:940-1038` (agent loop) |
| **U-H5** | **PowerPoint "Visual" action broken workflow** — generates text prompt instead of image | `constant.ts` (pptVisualPrefix) |
| **U-H6** | **No confirmation on "New Chat"** — conversation lost on accidental click | `HomePage.vue` |

### MEDIUM — Noticeable friction

| ID | Issue | Where |
|----|-------|-------|
| **U-M1** | **No message regeneration** | `HomePage.vue` |
| **U-M2** | **No chat history persistence** | `HomePage.vue` |
| **U-M3** | **Settings expose developer jargon** (system prompt, ${language}/${text}, agent max iterations) | `SettingsPage.vue` |
| **U-M4** | **Built-in prompts settings incomplete** — only Word/Excel, not PowerPoint/Outlook | `SettingsPage.vue:507-535` |
| **U-M5** | **Quick actions bar mismatch with built-in prompts** — Excel quick actions differ from customizable built-in prompts | `constant.ts` vs `SettingsPage.vue` |
| **U-M6** | **Model selector labels don't explain usage** — "Nano", "Standard", "Raisonnement" mean nothing to users | `HomePage.vue:167-174` |
| **U-M7** | **"Thought process" hardcoded in English** — not in i18n | `HomePage.vue:114` |

### LOW — Polish items

| ID | Issue | Where |
|----|-------|-------|
| **U-L1** | **No empty state clickable suggestions** | `HomePage.vue:74-96` |
| **U-L2** | **Action buttons always visible** — visual clutter on long conversations | `HomePage.vue:129-157` |
| **U-L3** | **No typing/loading animation** during agent processing | `HomePage.vue` |
| **U-L4** | **Checkbox area too small** (h-3.5 constraint on labels) | `HomePage.vue:207-216` |
| **U-L5** | **No message editing** (edit and re-send) | `HomePage.vue` |
| **U-L6** | **First Name/Last Name/Gender shown in all hosts** — mostly relevant for Outlook only | `SettingsPage.vue:92-120` |
| **U-L7** | **Configured models card shows raw model IDs** — meaningless to users | `SettingsPage.vue:149-164` |
| **U-L8** | **Excel "Transform" and "Highlight" labels are vague** | `en.json` / `fr.json` |

---

## Summary Table

| Priority | ID | Issue | Status |
|----------|-----|-------|--------|
| HIGH | U-H1 | Hardcoded French strings in chat UI | ❌ TODO |
| HIGH | U-H2 | Quick actions icon-only — no text labels | ❌ TODO |
| HIGH | U-H3 | Tool toggles in Settings are dead code | ❌ TODO |
| HIGH | U-H4 | No streaming in chat mode (agent loop) | ❌ TODO |
| HIGH | U-H5 | PowerPoint "Visual" broken workflow | ❌ TODO |
| HIGH | U-H6 | No confirmation on "New Chat" | ❌ TODO |
| MEDIUM | U-M1 | No message regeneration | ❌ TODO |
| MEDIUM | U-M2 | No chat history persistence | ❌ TODO |
| MEDIUM | U-M3 | Settings expose developer jargon | ❌ TODO |
| MEDIUM | U-M4 | Built-in prompts only for Word/Excel | ❌ TODO |
| MEDIUM | U-M5 | Quick actions vs built-in prompts mismatch | ❌ TODO |
| MEDIUM | U-M6 | Model selector labels don't explain usage | ❌ TODO |
| MEDIUM | U-M7 | "Thought process" not translated | ❌ TODO |
| LOW | U-L1 | No clickable empty state suggestions | ❌ TODO |
| LOW | U-L2 | Action buttons always visible (clutter) | ❌ TODO |
| LOW | U-L3 | No typing/loading animation | ❌ TODO |
| LOW | U-L4 | Checkbox area too small | ❌ TODO |
| LOW | U-L5 | No message editing | ❌ TODO |
| LOW | U-L6 | Name/Gender shown in all hosts | ❌ TODO |
| LOW | U-L7 | Raw model IDs shown to users | ❌ TODO |
| LOW | U-L8 | Vague Excel action labels | ❌ TODO |
