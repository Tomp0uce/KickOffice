# KickOffice - Design Review

**Initial date**: 2026-02-15
**Last updated**: 2026-02-16
**Scope**: Architecture, Security, Code Quality, Functional Bugs, Documentation
**Files analyzed**: Backend (`server.js` — 448 lines), Frontend (25+ source files), Documentation (README.md, agents.md, manifests)

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Previously Completed Fixes](#previously-completed-fixes)
3. [Overall Architecture](#overall-architecture)
4. [Issues by Severity](#issues-by-severity)
   - [CRITICAL — Blocking / Immediate Impact](#critical--blocking--immediate-impact)
   - [HIGH — Fix Soon](#high--fix-soon)
   - [MEDIUM — Plan for Later](#medium--plan-for-later)
   - [LOW — Nice to Have](#low--nice-to-have)
5. [Summary Table](#summary-table)

---

## Executive Summary

KickOffice is an AI-powered Microsoft Office add-in (Word, Excel, PowerPoint, Outlook), built with Vue 3 / TypeScript (frontend) and Express.js (backend). The architecture is fundamentally sound: the backend acts as a secure LLM proxy, API keys are never exposed client-side, and CORS is properly restricted.

### Current State

| Severity | Count | Summary |
|----------|-------|---------|
| **CRITICAL** | 5 | Images inserted as base64 text, silent errors, security (auth, rate limit, error leakage) |
| **HIGH** | 5 | Wrong model tiers, missing security headers, god component, monolithic backend, agent loop without abort |
| **MEDIUM** | 5 | Fragile SSE parser, missing Vue error handler, accessibility, request logging, residual `as any` |
| **LOW** | 3 | Dark mode toggle, repeated CSS, outdated README |

### Previously Completed Fixes

15 out of 26 initial items were fixed in the first iteration. This document does not revisit them and focuses on remaining issues + newly discovered ones.

---

## Previously Completed Fixes

| Old ID | Item | Evidence |
|--------|------|----------|
| C4 | Cleanup `setInterval` | `HomePage.vue:1338-1342` — `onUnmounted` + `clearInterval` |
| H2 | Backend input validation | `server.js:56-107` — `validateTemperature`, `validateMaxTokens`, `validateTools` |
| C1 | Raise tools limit for Word/Excel chat | `server.js:9,74` — configurable `MAX_TOOLS` (default 128) |
| H4 | Extract duplicated logic | `savedPrompts.ts`, unified `getOfficeSelection()` |
| H6 | Align `.env.example` with defaults | Consistent between `.env.example` and `server.js` |
| H7 | Fetch request timeouts | Backend: `fetchWithTimeout` + AbortController. Frontend: `fetchWithTimeoutAndRetry` |
| H8 | Type `ToolDefinition` | `types/index.d.ts:27-36` — generic type with alias |
| M1 | Unique IDs in `v-for` | `crypto.randomUUID()` in `createDisplayMessage` |
| M2 | Memoize `renderSegments` | `historyWithSegments` computed |
| M5 | Remove redundant watchers | `useStorage` handles persistence |
| M6 | Retry with backoff | `backend.ts:8-75` — 2 retries +10s/+30s |
| B1 | Fix `cursor-po` typo | All classes correct |
| B2 | Reduce `any` types | Interfaces `QuickAction`, `OpenAIChatCompletion` |
| B5 | Document `hostDetection.ts` | Added in README |
| B6 | Reduce body parser limit | 4MB (`server.js:172`) |

---

## Overall Architecture

### Strengths

- **Clear separation**: Frontend (Vue 3 + Vite, port 3002) / Backend (Express.js, port 3003) / External LLM API
- **Secret protection**: API keys only on server side in `.env`
- **Docker deployment**: Working Docker Compose with health checks
- **Multi-host support**: Word (37 tools), Excel (39 tools), PowerPoint (8 tools), Outlook (13 tools)
- **i18n**: 13 response languages, 2 UI locales (en/fr)
- **Agent mode**: OpenAI function-calling tool loop with backend validation
- **Robust backend validation**: Temperature, maxTokens, tools, prompt length
- **Timeout and retry**: Both sides have timeouts and a retry strategy

### Weaknesses

- **Functional bug**: Image insertion inserts base64 text instead of actual images
- **Silent errors**: Backend 400 errors are never logged — impossible to diagnose
- **Security**: No authentication, no rate limiting
- **Maintainability**: `HomePage.vue` = 1344 lines (god component), `server.js` = 448 lines (monolithic)
- **Misconfigured models**: gpt-5.2 (reasoning tier) is much faster than nano/standard — counter-intuitive setup

---

## Issues by Severity

---

### CRITICAL — Blocking / Immediate Impact

---

#### C1. Chat broken in Word and Excel — 32-tool limit exceeded ✅ FIXED (2026-02-16)

**Symptom**: Chat shows a "response error" in the UI under Word and Excel, but no error appears in the backend logs. Chat works in PowerPoint. Quick actions (buttons) still work.

**Root cause**: Backend `validateTools()` (`server.js:73`) rejects requests with more than 32 tools:
```javascript
if (tools.length > 32) return { error: 'tools supports at most 32 entries' }
```

But tools sent by the frontend are:
- **Word**: 37 tools + 2 general = **39 tools** → ❌ rejected (> 32)
- **Excel**: 39 tools + 2 general = **41 tools** → ❌ rejected (> 32)
- **PowerPoint**: 8 tools + 2 general = **10 tools** → ✅ accepted
- **Outlook**: 13 tools + 2 general = **15 tools** → ✅ accepted

The backend returns a 400 error, which is properly caught by the frontend (`backend.ts:192-194`) and displayed as "response error". But on the backend side, this 400 error is not logged (no `console.error`), hence no error in the logs.

**Why quick actions still work**: They use `chatStream()` which calls `/api/chat` (streaming) **without tools**. Only normal chat goes through `chatSync()` → `/api/chat/sync` with tools.

**Files**: `server.js:73`, `HomePage.vue:940-948` (tool construction)
**Impact**: Chat was completely broken for Word and Excel (the 2 primary hosts).

**Resolution implemented**:
1. Added a configurable backend limit:
   ```javascript
   const MAX_TOOLS = parseInt(process.env.MAX_TOOLS || '128', 10)
   if (tools.length > MAX_TOOLS) return { error: `tools supports at most ${MAX_TOOLS} entries` }
   ```
2. Default limit is now **128**, which covers Word/Excel tool sets with margin.
3. Added `MAX_TOOLS` to `backend/.env.example` and backend environment documentation in `README.md`.

**Follow-up (optional)**: Dynamically send only relevant tools to reduce payload size and token usage.

---

#### C2. Image buttons (copy/replace/append) insert base64 text instead of images

**Symptom**: After generating an image, the "Replace", "Append", "Copy" buttons insert the raw base64 data string instead of the image itself, crashing Word/PowerPoint.

**Root cause**: Broken fallback chain in `insertMessageToDocument()` (`HomePage.vue:528-547`).

For **Word** (`insertImageToWord`, line 513-526):
- The function uses `insertInlinePictureFromBase64()` which should work.
- **But** if it fails (e.g., Word context not ready, invalid range), the fallback calls `copyImageToClipboard()`.
- `copyImageToClipboard()` tries `ClipboardItem` (often blocked in the Office WebView iframe), then falls back to `copyToClipboard(imageSrc)` which copies the **full data URL string** (multi-MB text) to clipboard.

For **PowerPoint** and **Excel**:
- No direct image insertion path — the code falls directly to `copyImageToClipboard()` → same text fallback problem.
- However, PowerPoint has an `shapes.addImage(base64)` API (already implemented in `powerpointTools.ts:409-464`), but it's not used by the UI buttons.

**Files**: `HomePage.vue:487-547`
**Impact**: Image action buttons are broken for all hosts

**Proposed fix**:
1. **Word**: Add more granular try-catch in `insertImageToWord()` with explicit error logging. Verify extracted `base64Payload` is valid (length > 0).
2. **PowerPoint**: Add an `insertImageToPowerPoint()` function that uses `PowerPoint.run()` + `slide.shapes.addImage(base64)` (the API already exists in the tools).
3. **Excel**: Image insertion is not supported by the Excel JavaScript API. Document this limitation and show a clear message.
4. **Clipboard fallback**: NEVER copy the raw data URL as text. If `ClipboardItem` fails, show an explicit error message instead of copying the base64 string:
   ```typescript
   async function copyImageToClipboard(imageSrc: string, fallback = false) {
     try {
       const response = await fetch(imageSrc)
       const blob = await response.blob()
       if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
         await navigator.clipboard.write([new ClipboardItem({ [blob.type || 'image/png']: blob })])
         messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))
         return
       }
     } catch (err) {
       console.warn('Image clipboard write failed:', err)
     }
     // Do NOT fall through to copyToClipboard(imageSrc) which copies base64 text
     messageUtil.error(t('imageClipboardNotSupported'))
   }
   ```

---

#### C3. Backend errors not logged (silent 400 errors)

**Symptom**: When the backend rejects a request (tools validation, temperature, etc.), the 400 error is sent to the client but **never logged** server-side. Bug diagnosis is impossible without logs.

**Root cause**: Validation responses return `res.status(400).json({ error })` directly without `console.error` or `console.warn`. Only 500 errors and LLM errors are logged.

**Files**: `server.js` — all `return res.status(400).json(...)` lines (~15 occurrences)
**Impact**: Impossible to diagnose problems without frontend access (bug C1 is direct proof)

**Proposed fix**:
Add systematic logging before each error response:
```javascript
// Create a logging helper
function logAndRespond(res, status, errorObj) {
  if (status >= 400) {
    console.warn(`[${status}] ${errorObj.error}`)
  }
  return res.status(status).json(errorObj)
}
```
Or better: install `morgan` to log all requests (see M5) and add an error logging middleware.

---

#### C4. Sensitive information leakage in LLM errors

**Symptom**: Raw LLM API errors are forwarded to the client via the `details` field.

**Files**: `server.js:254-257`, `349-352`, `421-424`
```javascript
return res.status(response.status).json({
  error: `LLM API error: ${response.status}`,
  details: errorText,  // May contain internal URLs, versions, partial keys
})
```
**Impact**: Infrastructure information leakage. Present in all 3 endpoints.

**Proposed fix**:
```javascript
// Replace in all 3 endpoints:
console.error(`LLM API error ${response.status}:`, errorText)
return res.status(502).json({
  error: 'The AI service returned an error. Please try again later.',
})
```

---

#### C5. No authentication on the backend

**File**: `server.js` (entire file)
**Impact**: Anyone on the network can call the endpoints and consume LLM API credits

**Proposed fix**:
1. `ALLOWED_API_KEYS` variable in `.env` (comma-separated list)
2. `requireAuth` middleware checking the `x-api-key` header
3. Apply to `/api/chat`, `/api/chat/sync`, `/api/image`
4. Keep `/health` and `/api/models` public
5. Frontend: add the `x-api-key` header in `fetchWithTimeoutAndRetry` via a `VITE_API_KEY` variable

---

#### C6. No rate limiting

**File**: `server.js`
**Impact**: DoS possible, unlimited API credit consumption

**Proposed fix**:
```bash
npm install express-rate-limit
```
```javascript
import rateLimit from 'express-rate-limit'
const chatLimiter = rateLimit({ windowMs: 60_000, max: 20 })
const imageLimiter = rateLimit({ windowMs: 60_000, max: 5 })
app.use('/api/chat', chatLimiter)
app.use('/api/image', imageLimiter)
```

---

### HIGH — Fix Soon

---

#### H1. Model tier configuration is wrong

**Symptom**: GPT-5.2 (tier `reasoning`) is much faster than the `nano` (gpt-5-nano) and `standard` (gpt-5-mini) models. Users must manually select "Reasoning" to get the best performance, which is counter-intuitive.

**Files**: `server.js:13-40`, `backend/.env.example`
**Impact**: Degraded UX, users must know internals to choose the right model

**Current configuration**:
| Tier | Model | Intended use |
|------|-------|-------------|
| nano | gpt-5-nano | Fast, basic |
| standard | gpt-5-mini | Normal chat |
| reasoning | gpt-5.2 | Complex → but actually the fastest |
| image | gpt-image-1.5 | Image generation |

**Proposed fix** — Reconfigure to 3 tiers (remove nano):
| Tier | Model | Label | Usage |
|------|-------|-------|-------|
| standard | gpt-5.2 | Standard | Normal chat + agent (fast and performant) |
| reasoning | gpt-5.2 (reasoning mode) | Reasoning | Complex tasks requiring deep reasoning |
| image | gpt-image-1.5 | Image | Image generation |

Changes:
1. `server.js`: Remove `nano` tier, move `gpt-5.2` to `standard`
2. `.env.example`: Update default models
3. Frontend `SettingsPage.vue` / `HomePage.vue`: Model selector shows 3 options instead of 4
4. `getChatTimeoutMs()`: Adjust timeouts accordingly

**Note**: Check whether gpt-5.2 supports a reasoning mode (e.g., `reasoning_effort` parameter) and adapt `buildChatBody()` accordingly.

---

#### H2. Missing HTTP security headers

**File**: `server.js`
**Impact**: Clickjacking, MIME sniffing vulnerabilities, etc.

**Proposed fix**:
```bash
npm install helmet
```
```javascript
import helmet from 'helmet'
app.use(helmet({
  contentSecurityPolicy: false, // Office add-in has its own CSP
  crossOriginEmbedderPolicy: false,
}))
```

---

#### H3. `HomePage.vue` — god component (1344 lines)

**File**: `frontend/src/pages/HomePage.vue`
**Impact**: Maintainability, readability, testability, performance

The component combines: chat UI, agent loop, quick actions, Office API (4 hosts), clipboard, health check polling, system prompts for each host, image insertion.

**Proposed fix** — Extract into 7 pieces:
1. `ChatHeader.vue` — Header with logo, new chat and settings buttons (lines 5-38)
2. `QuickActionsBar.vue` — Quick actions bar with prompt selector (lines 41-67)
3. `ChatMessageList.vue` — Message container with empty state (lines 70-160)
4. `ChatInput.vue` — Input area with mode and model selectors (lines 163-217)
5. Composable `useAgentLoop.ts` — Agent loop + system prompts (lines 653-1039)
6. Composable `useOfficeInsert.ts` — Document insertion + clipboard (lines 1199-1311)
7. Composable `useImageActions.ts` — Image generation and insertion (lines 487-547)

---

#### H4. Monolithic backend (448 lines in 1 file)

**File**: `backend/src/server.js`
**Impact**: Maintainability as the code grows

**Proposed fix**:
```
backend/src/
├── server.js              # Entry point, middleware setup
├── config/
│   └── models.js          # Model configuration
├── middleware/
│   ├── auth.js            # Authentication (C5)
│   └── validate.js        # Input validation (extract existing)
├── routes/
│   ├── health.js          # GET /health
│   ├── models.js          # GET /api/models
│   ├── chat.js            # POST /api/chat, /api/chat/sync
│   └── image.js           # POST /api/image
└── utils/
    └── fetchWithTimeout.js # Fetch helper with timeout
```

---

#### H5. Agent loop without abort support

**Symptom**: When the user clicks "Stop" during agent-mode chat, the `abortController` is triggered but the in-flight `chatSync()` request is not interrupted because `chatSync` doesn't receive the abort signal.

**Files**:
- `backend.ts:183-198`: `chatSync()` doesn't pass a `signal` to `fetchWithTimeoutAndRetry()`
- `HomePage.vue:958-965`: The `while` loop doesn't check `abortController.value?.signal.aborted` between iterations

**Impact**: The "Stop" button doesn't work during agent mode. The request continues in the background and results are ignored when they arrive, wasting LLM tokens.

**Proposed fix**:
1. Add an optional `abortSignal` field to `ChatSyncOptions` and pass it to `fetchWithTimeoutAndRetry()`
2. In `runAgentLoop`, pass `abortController.value?.signal` to `chatSync()`
3. Add a check `if (abortController.value?.signal.aborted) break` at the start of each loop iteration

---

### MEDIUM — Plan for Later

---

#### M1. Fragile SSE parser (chunk splitting)

**Potential symptom**: Truncated responses or random JSON errors during streaming.

**File**: `backend.ts:124-151`

The SSE parser splits by `\n` but doesn't handle the case where a `data: {...}` line is split across two TCP chunks. If a chunk ends in the middle of a JSON line, `JSON.parse()` fails silently (the catch is empty).

**Proposed fix**:
Maintain a residual buffer between chunks:
```typescript
let buffer = ''
while (true) {
  const { done, value } = await reader.read()
  if (done) break
  buffer += decoder.decode(value, { stream: true })
  const lines = buffer.split('\n')
  buffer = lines.pop() || '' // Keep the last incomplete line
  for (const line of lines) {
    if (!line.startsWith('data: ')) continue
    // ... parse as before
  }
}
```

---

#### M2. Missing global Vue error handler

**File**: `frontend/src/main.ts`
**Impact**: Uncaught errors cause silent crashes

**Proposed fix**:
```typescript
app.config.errorHandler = (err, instance, info) => {
  console.error('Vue Global Error:', err, info)
  // Optional: show a toast
}
```

---

#### M3. Insufficient accessibility (a11y)

**Files**: `HomePage.vue`, components
**Impact**: WCAG non-compliance, user exclusion

**Proposed fix**:
1. `aria-label` on all text-less buttons (New Chat, Settings, Stop, Send, Copy, Replace, Append)
2. `aria-live="polite"` on the messages container
3. `role="status"` on the backend online/offline indicator
4. `aria-expanded` on `<details>` elements (think tags)

---

#### M4. Residual `as any` in the agent loop

**File**: `HomePage.vue:1022-1026`
```typescript
currentMessages.push({
  role: 'tool' as any,
  tool_call_id: toolCall.id,
  content: result,
} as any)
```

**File**: `backend.ts:157`
```typescript
tools?: any[]
```

**Proposed fix**:
1. Extend `ChatMessage` to support the `tool` role:
   ```typescript
   export type ChatMessage =
     | { role: 'system' | 'user' | 'assistant'; content: string }
     | { role: 'tool'; tool_call_id: string; content: string }
   ```
2. Type `tools` in `ChatSyncOptions` with the existing `ToolDefinition[]` type.

---

#### M5. Missing request logging

**File**: `server.js`
**Impact**: Impossible to diagnose or audit requests

**Proposed fix**:
```bash
npm install morgan
```
```javascript
import morgan from 'morgan'
app.use(morgan(':method :url :status :response-time ms'))
```

---

### LOW — Nice to Have

---

#### B1. No dark mode toggle in the UI

**File**: `frontend/src/pages/SettingsPage.vue`
**Detail**: Dark mode CSS variables exist in `index.css:162-187` but there's no toggle to activate them.

**Proposed fix**:
```typescript
const darkMode = useStorage(localStorageKey.darkMode, false)
watch(darkMode, (val) => {
  document.documentElement.classList.toggle('dark', val)
}, { immediate: true })
```

---

#### B2. Repeated CSS classes

**File**: `frontend/src/index.css`
**Detail**: Patterns like `rounded-md border border-border-secondary bg-surface p-2 shadow-sm` are repeated across components.

**Proposed fix**:
```css
@layer components {
  .card { @apply rounded-md border border-border-secondary bg-surface p-2 shadow-sm; }
}
```

---

#### B3. Outdated README.md

**File**: `README.md`
**Detail**: Several pieces of information no longer match the code.

**Required corrections**:
1. "23 Word tools" → **37 Word tools**
2. "22 Excel tools" → **39 Excel tools**
3. Add **8 PowerPoint tools** and **13 Outlook tools**
4. Add Quick Actions for Excel, PowerPoint, Outlook
5. Confirm PowerPoint support (marked as not implemented but it is now)
6. Mention the 13 response languages
7. Document the 4 model tiers configuration (soon to be 3)

---

## Summary Table

| Priority | ID | Action | Status |
|----------|-----|--------|--------|
| **CRITICAL** | **C1** | **Chat broken Word/Excel — 32 tools limit** | ✅ FIXED |
| **CRITICAL** | **C2** | **Image buttons insert base64 text** | ❌ TODO |
| **CRITICAL** | **C3** | **400 errors not logged in backend** | ❌ TODO |
| **CRITICAL** | **C4** | **LLM error leakage to client** | ❌ TODO |
| **CRITICAL** | **C5** | **No backend authentication** | ❌ TODO |
| **CRITICAL** | **C6** | **No rate limiting** | ❌ TODO |
| HIGH | H1 | Model tier configuration is wrong | ❌ TODO |
| HIGH | H2 | Missing HTTP security headers (helmet) | ❌ TODO |
| HIGH | H3 | `HomePage.vue` god component (1344 lines) | ❌ TODO |
| HIGH | H4 | Monolithic backend (448 lines) | ❌ TODO |
| HIGH | H5 | Agent loop without abort support | ❌ TODO |
| MEDIUM | M1 | Fragile SSE parser (chunk splitting) | ❌ TODO |
| MEDIUM | M2 | Missing global Vue error handler | ❌ TODO |
| MEDIUM | M3 | Accessibility (ARIA) | ❌ TODO |
| MEDIUM | M4 | Residual `as any` in agent loop | ❌ TODO |
| MEDIUM | M5 | Request logging (morgan) | ❌ TODO |
| LOW | B1 | Dark mode toggle | ❌ TODO |
| LOW | B2 | Extract repeated CSS | ❌ TODO |
| LOW | B3 | Outdated README.md | ❌ TODO |

---

## Security — OK Points (no issues found)

- **XSS**: No `v-html` usage — Vue escapes correctly
- **CORS**: Properly restricted to `FRONTEND_URL`
- **Secrets**: API keys never exposed client-side
- **SQL/NoSQL Injection**: N/A (no database)
- **Input validation**: Temperature, maxTokens, tools structure, prompt length, image params all validated
- **Timeouts**: All fetch requests have timeouts with AbortController
