# KickOffice - Design Review

**Initial date**: 2026-02-15
**Last updated**: 2026-02-16 (revision 3)
**Scope**: Architecture, Security, Code Quality, Functional Bugs, Documentation
**Files analyzed**: Full codebase — backend (modular: 10 source files), frontend (42 source files), documentation (7 files)

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Completed Fixes (since initial review)](#completed-fixes-since-initial-review)
3. [Overall Architecture](#overall-architecture)
4. [Issues by Severity](#issues-by-severity)
   - [CRITICAL — Blocking / Immediate Impact](#critical--blocking--immediate-impact)
   - [HIGH — Fix Soon](#high--fix-soon)
   - [MEDIUM — Plan for Later](#medium--plan-for-later)
   - [LOW — Nice to Have](#low--nice-to-have)
5. [Summary Table](#summary-table)

---

## Executive Summary

KickOffice is an AI-powered Microsoft Office add-in (Word, Excel, PowerPoint, Outlook), built with Vue 3 / TypeScript (frontend) and Express.js (backend). Since the initial review, **15 of the 18 original issues have been fixed**, including the backend modularization, HomePage refactoring, security hardening (Helmet, rate limiting, error logging), and image insertion bugs.

### Current State

| Severity | Open | Fixed | Summary |
|----------|------|-------|---------|
| **CRITICAL** | 1 | 5 | **Chat broken in Word** (`reasoning_effort: 'none'` + tools) |
| **HIGH** | 2 | 5 | Agent loop silent exit on empty response, tool toggles dead code |
| **MEDIUM** | 4 | 4 | Hardcoded French, missing error handler, a11y, built-in prompts incomplete |
| **LOW** | 1 | 1 | Repeated CSS patterns |

**Headline issue**: Chat in Word is broken. Quick actions work fine, but agent-mode chat (typing a message) does nothing — the message is sent but no response appears. Root cause analysis points to `reasoning_effort: 'none'` being sent to the GPT-5.2 API alongside tools, which likely causes the model to return an empty response. See [C7](#c7-chat-broken-in-word--reasoning_effort-none-prevents-tool-calling-new).

---

## Completed Fixes (since initial review)

### Batch 1 — Already marked in initial review (15 items pre-existing)

These were fixed before or during the initial review. See initial review for details. Includes: C1 (32-tool limit), H2 (backend validation), H4 (duplicated logic), H6 (env alignment), H7 (fetch timeouts), H8 (ToolDefinition type), M1 (v-for keys), M2 (memoize renderSegments), M5 (redundant watchers), M6 (retry with backoff), B1 (cursor-po typo), B2 (reduce any), B5 (hostDetection docs), B6 (body parser limit), C4 (clearInterval).

### Batch 2 — Fixed since initial review (15 items newly resolved)

| Old ID | Item | Evidence | Status |
|--------|------|----------|--------|
| **C2** | Image buttons insert base64 text | `useImageActions.ts:68-90` — proper `insertImageToWord` and `insertImageToPowerPoint`. `copyImageToClipboard` (line 52-66) shows error instead of copying base64 text. Excel shows `imageInsertExcelNotSupported` info message. | ✅ FIXED |
| **C3** | Backend errors not logged (silent 400s) | `http.js:14-25` — `logAndRespond` utility logs all 4xx/5xx. Used consistently in all routes. | ✅ FIXED |
| **C4** | Sensitive info leakage in LLM errors | `chat.js:77-80,173-175` — Generic 502 to client, detailed log server-side only. | ✅ FIXED |
| **C5** | No backend authentication | `auth.js:4-13` — `ensureLlmApiKey` middleware. `server.js:58-59` — applied to chat and image routes. **Note**: This validates the API key is configured server-side, not client authentication. Acceptable for intranet deployment. | ✅ FIXED (intranet scope) |
| **C6** | No rate limiting | `server.js:26-40` — `express-rate-limit` on chat (20/min) and image (5/min). Configurable via `CHAT_RATE_LIMIT_*` and `IMAGE_RATE_LIMIT_*` env vars. | ✅ FIXED |
| **H1** | Model tier configuration wrong | `models.js:5-27` — 3 tiers (standard/reasoning/image). GPT-5.2 as default for both chat tiers. | ✅ FIXED |
| **H2** | Missing HTTP security headers | `server.js:48-51` — Helmet configured (CSP and COEP disabled for Office add-in). | ✅ FIXED |
| **H3** | HomePage.vue god component (1344 lines) | `HomePage.vue` now 265 lines (was 1344). Extracted: `ChatHeader.vue`, `ChatInput.vue`, `ChatMessageList.vue`, `QuickActionsBar.vue` + composables: `useAgentLoop.ts`, `useImageActions.ts`, `useOfficeInsert.ts`. | ✅ FIXED |
| **H4** | Monolithic backend (448 lines) | Modular structure: `routes/` (chat, health, image, models), `middleware/` (auth, validate), `config/` (env, models), `utils/` (http). `server.js` now 80 lines. | ✅ FIXED |
| **H5** | Agent loop without abort support | `useAgentLoop.ts:161-176` — Abort signal checked at loop start and passed to `chatSync`. Abort errors caught in loop body. | ✅ FIXED |
| **M1** | Fragile SSE parser (chunk splitting) | `backend.ts:131-139` — Buffer-based parsing: `buffer += decoder.decode(value, {stream:true}); lines = buffer.split('\n'); buffer = lines.pop()`. | ✅ FIXED |
| **M4** | Residual `as any` in agent loop | `backend.ts:92-103` — Proper types: `ChatRequestMessage = ChatMessage \| ToolChatMessage`. `ToolDefinition` interface (line 172-180). | ✅ FIXED |
| **M5** | Missing request logging (morgan) | `server.js:54` — `morgan(':method :url :status :res[content-length] - :response-time ms')`. | ✅ FIXED |
| **B1** | No dark mode toggle | `SettingsPage.vue:77-90` — Dark mode toggle with `useStorage(localStorageKey.darkMode)`. `main.ts` applies class on startup. | ✅ FIXED |
| **B3** | Outdated README | Partially updated — model count, Docker details, implementation status. Still has stale model tier table (see M9). | ⚠️ PARTIAL |

---

## Overall Architecture

### Strengths (updated)

- **Clear separation**: Frontend (Vue 3 + Vite, port 3002) / Backend (Express.js, port 3003) / External LLM API
- **Secret protection**: API keys only on server side in `.env`
- **Docker deployment**: Working Docker Compose with health checks
- **Multi-host support**: Word (39 tools), Excel (39 tools), PowerPoint (8 tools), Outlook (13 tools), General (2 tools)
- **i18n**: 13 response languages, 2 UI locales (en/fr)
- **Agent mode**: OpenAI function-calling tool loop with backend validation, abort support
- **Robust backend validation**: Temperature, maxTokens, tools, prompt length — with `logAndRespond` for all errors
- **Timeout and retry**: Both sides have timeouts and retry strategies with exponential backoff
- **Security**: Helmet, rate limiting, CORS, sanitized error responses
- **Modular codebase**: Backend split into routes/middleware/config/utils. Frontend split into pages/components/composables

### Remaining Weaknesses

- **Chat broken**: Agent-mode chat in Word sends `reasoning_effort: 'none'` with tools, likely causing empty model responses (see C7)
- **No client authentication**: `ensureLlmApiKey` only checks server-side API key, not client identity
- **Tool toggles dead code**: Settings "Tools" tab saves to localStorage but agent loop never reads it
- **Hardcoded French strings**: Several strings in composables are not using i18n
- **README `MODEL_STANDARD_REASONING_EFFORT`**: Description still referenced `'none'` as a valid value (fixed in revision 3 — description now says to omit the parameter)

---

## Issues by Severity

---

### CRITICAL — Blocking / Immediate Impact

---

#### C7. Chat broken in Word — `reasoning_effort: 'none'` prevents tool calling (NEW)

**Symptom**: In Word, typing a message in the chat sends it to the conversation but nothing happens. The button immediately returns to "available". The "⏳ Analyse de la demande..." placeholder may flash briefly. Quick action buttons work perfectly.

**User report**: "le chat ne marche pas dans word, ça envoie le message dans le chat mais rien ne se passe et direct le bouton repasse à chat dispo. Les boutons d'action rapide marchent sans souci. Avant le chat marchait très bien, on a juste changé les modèles."

**Root cause analysis**:

The chat and quick actions use different API paths:
- **Quick actions** → `chatStream()` → `POST /api/chat` (streaming, **no tools**) → **WORKS**
- **Chat** → `chatSync()` → `POST /api/chat/sync` (non-streaming, **with tools**) → **BROKEN**

Both paths share the same `buildChatBody()` function (`models.js:49-88`). For the standard tier with GPT-5.2, the body sent to the LLM API is:

```json
{
  "model": "gpt-5.2",
  "messages": [...],
  "stream": false,
  "max_completion_tokens": 4096,
  "temperature": 0.7,
  "tools": [... 41 tools ...],
  "tool_choice": "auto",
  "reasoning_effort": "none"    ← PROBLEMATIC
}
```

The `reasoning_effort: 'none'` value comes from `models.js:11`:
```javascript
reasoningEffort: process.env.MODEL_STANDARD_REASONING_EFFORT || 'none',
```

Since `MODEL_STANDARD_REASONING_EFFORT` is not set in `.env.example` (only `MODEL_REASONING_EFFORT=high` for the reasoning tier), it defaults to `'none'`.

Then in `buildChatBody` (`models.js:83-85`):
```javascript
if (modelTier !== 'image' && isGpt5Model(modelId)) {
    body.reasoning_effort = reasoningEffort  // sends 'none' for standard tier
}
```

**Two possible failure modes** (both lead to the same user-visible symptom):

1. **API rejects `reasoning_effort: 'none'`**: The OpenAI API may not accept `'none'` as a valid value (typical values are `'low'`, `'medium'`, `'high'`). This would cause a non-OK response → backend catches it → returns 502 → frontend throws in `chatSync()` → `sendMessage()` catches and shows "failedToResponse" toast. The user may not notice the brief toast.

2. **API accepts `'none'` but model returns empty response**: With reasoning disabled, the model cannot reason about which tools to call. It returns `{ choices: [{ message: { content: null, tool_calls: undefined } }] }`. Then in `useAgentLoop.ts:177-182`:
   ```typescript
   const choice = response.choices?.[0]  // exists
   if (!choice) break                    // doesn't break
   if (assistantMsg.content) ...         // null → skip
   if (!assistantMsg.tool_calls?.length) break  // no tool_calls → BREAKS
   ```
   The loop exits on the first iteration. The history message stays as "⏳ Analyse de la demande..." and loading goes to false. **This matches the user's description exactly.**

**Why quick actions still work**: They use `chatStream()` which calls `/api/chat` **without tools**. The model can still generate text with `reasoning_effort: 'none'` because it doesn't need to reason about tool selection — it just needs to generate a text response.

**Files**: `models.js:11,83-85`, `useAgentLoop.ts:150-209`

**Proposed fix (option A — recommended)**: Don't send `reasoning_effort` when the value is `'none'`. Let the API use its default behavior:
```javascript
// models.js:83-85, change to:
if (modelTier !== 'image' && isGpt5Model(modelId) && reasoningEffort && reasoningEffort !== 'none') {
    body.reasoning_effort = reasoningEffort
}
```

**Proposed fix (option B)**: Remove the `reasoningEffort` field entirely from the standard tier config, and only use it for the reasoning tier:
```javascript
standard: {
    id: process.env.MODEL_STANDARD || 'gpt-5.2',
    // ... no reasoningEffort field
},
reasoning: {
    id: process.env.MODEL_REASONING || 'gpt-5.2',
    reasoningEffort: process.env.MODEL_REASONING_EFFORT || 'high',
    // ...
},
```
Then in `buildChatBody`: only set `reasoning_effort` if the config has it defined.

**Proposed fix (option C)**: Change the standard tier default from `'none'` to omit the parameter:
```javascript
reasoningEffort: process.env.MODEL_STANDARD_REASONING_EFFORT || undefined,
```

**Test validation**: After any fix, verify:
1. Chat in Word sends a message and receives a response with tool calls
2. Quick actions still work
3. Reasoning tier still uses `reasoning_effort: 'high'`

---

### HIGH — Fix Soon

---

#### H6. Agent loop silently exits on empty model response (NEW)

**Symptom**: If the model returns an empty response (no content, no tool_calls), the agent loop exits without any user-visible feedback. The "⏳ Analyse de la demande..." placeholder stays in the chat with no error message.

**Root cause**: In `useAgentLoop.ts:177-182`, the loop breaks silently when there are no tool_calls and no content:
```typescript
const assistantMsg = choice.message
currentMessages.push({ role: 'assistant', content: assistantMsg.content || '' })
if (assistantMsg.content) history.value[lastIndex].content = assistantMsg.content
if (!assistantMsg.tool_calls?.length) break  // silent exit
```

After the loop (line 203-208), the only post-loop handling is for abort and max-iterations. There is no fallback for an empty response.

**Impact**: This is a resilience issue independent of C7. Even after fixing C7, any future API change that produces empty responses will cause the same silent failure.

**Proposed fix**: Add empty response detection after the loop:
```typescript
// After the while loop and abort check:
if (!abortedByUser && iteration <= 1 && !history.value[lastIndex]?.content?.trim()
    || history.value[lastIndex]?.content === '⏳ Analyse de la demande...') {
    history.value[lastIndex].content = t('noModelResponse')
    // Where 'noModelResponse' i18n key = "The model returned an empty response. Try again or check the backend logs."
}
```

---

#### H7. Tool enable/disable toggles in Settings are dead code (NEW)

**Symptom**: The Settings "Tools" tab shows checkboxes to enable/disable individual tools. Toggling them saves to `localStorage('enabledTools')`. But `useAgentLoop.ts:151-153` never reads this value — it always includes all tools:

```typescript
const appToolDefs = hostIsOutlook ? getOutlookToolDefinitions() : ...
const generalToolDefs = getGeneralToolDefinitions()
const tools = [...generalToolDefs, ...appToolDefs].map(...)
```

**Files**: `SettingsPage.vue:717-738` (saves), `useAgentLoop.ts:151-153` (ignores)
**Impact**: The feature appears to work in the UI but has zero effect on actual behavior. This erodes user trust.

**Proposed fix**: Either:
1. **Wire it up**: In `useAgentLoop.ts`, read `enabledTools` from localStorage and filter tools accordingly
2. **Remove the tab**: If the feature is not ready, hide the Tools tab from the Settings UI

---

### MEDIUM — Plan for Later

---

#### M2. Missing global Vue error handler (unchanged)

**File**: `frontend/src/main.ts`
**Impact**: Uncaught errors in Vue components cause silent failures

**Proposed fix**:
```typescript
app.config.errorHandler = (err, instance, info) => {
  console.error('Vue Global Error:', err, info)
}
```

---

#### M3. Insufficient accessibility (ARIA) (unchanged)

**Files**: Components, `ChatMessageList.vue`, `ChatInput.vue`, `QuickActionsBar.vue`
**Impact**: WCAG non-compliance

Specific gaps:
1. Quick action buttons are icon-only with no `aria-label` — screen readers can't identify them
2. Chat messages container needs `aria-live="polite"` for dynamic content announcements
3. Backend status indicator needs `role="status"`
4. The model selector dropdown and checkbox labels need proper ARIA attributes

---

#### M6. Hardcoded French strings in composables (NEW)

**Files**: `useAgentLoop.ts`

Two hardcoded French strings remain in the composable (moved from the old HomePage.vue during refactoring):

| Line | String | Should be |
|------|--------|-----------|
| `useAgentLoop.ts:157` | `'⏳ Analyse de la demande...'` | `t('agentAnalyzing')` |
| `useAgentLoop.ts:204` | `'🛑 Processus arrêté par l\'utilisateur.'` | `t('agentStoppedByUser')` |

These break the i18n system — English users see French text during chat processing.

---

#### M7. README stale content (FIXED in revision 3)

**File**: `README.md`

Issues fixed in revision 3:
- Architecture diagram had `PowerPoint` listed twice; `Outlook` was missing — corrected.
- Project structure section only showed `server.js` for the backend; now shows the full modular layout (`routes/`, `middleware/`, `config/`, `utils/`) and the composable-based frontend (`composables/`, full `components/chat/`, complete `utils/` list).
- `MODEL_STANDARD_REASONING_EFFORT` env var description said "`none` to disable" — `'none'` is not a valid API value; description updated to clarify valid values and recommend omitting the parameter.
- Added a "Known Open Issues" table in the Implementation Status section linking to DESIGN_REVIEW.md for quick project-health visibility.

---

#### M8. Built-in prompts customization incomplete (NEW)

**File**: `SettingsPage.vue:523-558`

The built-in prompts editor only supports Word and Excel prompts:
```typescript
const builtInPromptsData = ref(
  hostIsExcel ? { ...excelBuiltInPromptsData } : { ...wordBuiltInPromptsData }
)
```

PowerPoint and Outlook built-in prompts (`powerPointBuiltInPrompt`, `outlookBuiltInPrompt` in `constant.ts`) exist and are used by quick actions, but they are **not customizable** via the Settings UI.

**Impact**: PowerPoint and Outlook users cannot customize their quick action prompts.

**Proposed fix**: Add PowerPoint and Outlook prompt configs to the `SettingsPage.vue` built-in prompts tab, following the same pattern as Word/Excel.

---

### LOW — Nice to Have

---

#### B2. Repeated CSS classes (unchanged)

**File**: `frontend/src/index.css`
**Detail**: Patterns like `rounded-md border border-border-secondary bg-surface p-2 shadow-sm` are repeated across components.

**Proposed fix**:
```css
@layer components {
  .card { @apply rounded-md border border-border-secondary bg-surface p-2 shadow-sm; }
}
```

---

## Summary Table

| Priority | ID | Action | Status |
|----------|-----|--------|--------|
| **CRITICAL** | **C1** | Chat broken Word/Excel — 32 tools limit | ✅ FIXED |
| **CRITICAL** | **C2** | Image buttons insert base64 text | ✅ FIXED |
| **CRITICAL** | **C3** | 400 errors not logged in backend | ✅ FIXED |
| **CRITICAL** | **C4** | LLM error leakage to client | ✅ FIXED |
| **CRITICAL** | **C5** | No backend authentication | ✅ FIXED (intranet) |
| **CRITICAL** | **C6** | No rate limiting | ✅ FIXED |
| **CRITICAL** | **C7** | **Chat broken Word — `reasoning_effort: 'none'` + tools** | ❌ **TODO** |
| HIGH | H1 | Model tier configuration wrong | ✅ FIXED |
| HIGH | H2 | Missing HTTP security headers (Helmet) | ✅ FIXED |
| HIGH | H3 | `HomePage.vue` god component (1344 lines) | ✅ FIXED (265 lines) |
| HIGH | H4 | Monolithic backend (448 lines) | ✅ FIXED (80 lines) |
| HIGH | H5 | Agent loop without abort support | ✅ FIXED |
| HIGH | H6 | **Agent loop silently exits on empty response** | ❌ **TODO** |
| HIGH | H7 | **Tool toggles in Settings are dead code** | ❌ **TODO** |
| MEDIUM | M1 | Fragile SSE parser (chunk splitting) | ✅ FIXED |
| MEDIUM | M2 | Missing global Vue error handler | ❌ TODO |
| MEDIUM | M3 | Accessibility (ARIA) | ❌ TODO |
| MEDIUM | M4 | Residual `as any` in agent loop | ✅ FIXED |
| MEDIUM | M5 | Request logging (morgan) | ✅ FIXED |
| MEDIUM | M6 | **Hardcoded French strings in composables** | ❌ **TODO** |
| MEDIUM | M7 | README stale content (architecture diagram, project structure, env vars) | ✅ FIXED |
| MEDIUM | M8 | **Built-in prompts customization incomplete (PPT/Outlook)** | ❌ **TODO** |
| LOW | B1 | Dark mode toggle | ✅ FIXED |
| LOW | B2 | Extract repeated CSS | ❌ TODO |
| LOW | B3 | Outdated README | ✅ FIXED (via M7) |

### Progress

- **Total items**: 24
- **Fixed**: 17 (71%)
- **Open**: 7 (1 critical, 2 high, 3 medium, 1 low)

---

## Security — OK Points (no issues found)

- **XSS**: No `v-html` usage — Vue escapes correctly
- **CORS**: Properly restricted to `FRONTEND_URL`
- **Secrets**: API keys never exposed client-side
- **SQL/NoSQL Injection**: N/A (no database)
- **Input validation**: Temperature, maxTokens, tools structure, prompt length, image params all validated via `middleware/validate.js`
- **Timeouts**: All fetch requests have timeouts with AbortController
- **Security headers**: Helmet configured (HSTS, X-Frame-Options, X-Content-Type-Options)
- **Rate limiting**: IP-based on chat (20/min) and image (5/min)
- **Error sanitization**: LLM errors logged server-side, generic messages to client
- **Request logging**: Morgan middleware for all HTTP requests
