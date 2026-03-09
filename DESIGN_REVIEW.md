# DESIGN_REVIEW.md — Code Audit v10.1

**Date**: 2026-03-09
**Version**: 10.1
**Scope**: Full design review — Architecture, tool/prompt quality, error handling, UX/UI, dead code, code quality, user-reported issues & prospective improvements

---

## Execution Status Overview

| Status | Count | Items |
|--------|-------|-------|
| ✅ **FIXED** | 4 | TOOL-C1 (partial), TOOL-H1, USR-C1, USR-H1 |
| 🟠 **PARTIALLY FIXED** (deferred actions in Phase 4) | 3 | TOOL-H2, USR-H2, TOOL-C1 remaining |
| ⏳ **IN PROGRESS** | 6 | ERR-H1, ERR-H2, DUP-H1, QUAL-H1, + 5 deferred High items |
| 📋 **BACKLOG** | 9 | Phase 2 Medium items |
| 🎯 **PLANNED** | 5 | Phase 3 Low items |
| 🚀 **DEFERRED IMPROVEMENTS** | 8 | Phase 4 Prospective (PROSP-1/2/3/4/5 + H2 context optimization) |

---

## Health Summary (v10.1)

All previous critical and major items from v9.x have been resolved. This v10.1 review is a comprehensive deep-dive across 8 axes + user-reported issues + prospective improvements, identifying new improvement opportunities after recent large-scale changes (OOXML editing, chart extraction, image registry, session persistence, header auto-detect).

**Latest session (2026-03-09)**: Fixed 4 items (TOOL-H1, USR-H1, USR-C1, TOOL-C1 logging), partially fixed 3 items with deferred actions documented.

| Category | 🔴 Critical | 🟠 High | 🟡 Medium | 🟢 Low |
|----------|----------|------|--------|-----|
| Architecture | 0 | 2 | 3 | 1 |
| Tool/Prompt Quality | 0 | 2 | 4 | 3 |
| Error Handling | 0 | 2 | 2 | 1 |
| UX/UI | 0 | 1 | 3 | 3 |
| Dead Code | 0 | 0 | 2 | 1 |
| Code Duplication | 0 | 1 | 2 | 0 |
| Code Quality | 0 | 1 | 3 | 2 |
| User-Reported Issues | 0 | 2 | 2 | 1 |
| **Total** | **0** | **11** | **21** | **12** |
| **Status** | ✅ All critical items fixed or deferred | 6 active, 5 deferred | 21 items | 12 items |

---

## 1. ARCHITECTURE

### ARCH-H1 — `useAgentLoop.ts` is a monolith (1 145 lines) [HIGH]

**File**: `frontend/src/composables/useAgentLoop.ts`

The largest composable handles too many concerns: message orchestration, stream processing, tool execution coordination, loop detection, session management, document context injection, quick actions, and scroll management. It imports from 12+ utility files, creating a star dependency pattern.

**Impact**: Hard to test, hard to extend (adding a new Office host requires modifying imports), hard to reason about state.

**Recommendation**: Extract into focused composables:
- `useMessageOrchestration.ts` — message building, context injection
- `useQuickActions.ts` — quick action dispatch
- `useSessionFiles.ts` — uploaded file management
- Keep `useAgentLoop.ts` as a thin orchestrator

---

### ARCH-H2 — HomePage.vue prop drilling (44+ bindings) [HIGH]

**File**: `frontend/src/pages/HomePage.vue`

HomePage passes 44+ props and event bindings down to child components (ChatHeader: 7, ChatMessageList: 17, ChatInput: 13, QuickActionsBar: 6). This creates tight coupling between the page and its children.

**Impact**: Every state change requires updating prop chains. Adding a new feature touches multiple components.

**Recommendation**: Use `provide/inject` or a page-level composable (`useHomePageState`) to share reactive state directly, reducing prop drilling by ~60%.

---

### ARCH-M1 — No abstraction layer for tool providers [MEDIUM]

**Files**: `useAgentLoop.ts:1-30`, `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`

Tool definitions are imported directly with host-specific imports. Adding support for a new Office host (e.g., OneNote) requires modifying the agent loop imports and switch logic.

**Recommendation**: Create a `ToolProviderRegistry` that dynamically registers tool definitions by host, making the agent loop host-agnostic.

---

### ARCH-M2 — Backend validation in single 236-line file [MEDIUM]

**File**: `backend/src/middleware/validate.js`

All request validation is in one file. `validateTools()` has 8 error paths with deep nesting. Changes to one endpoint's validation can inadvertently affect others.

**Recommendation**: Extract domain-specific validators (`chatValidator.js`, `imageValidator.js`, `fileValidator.js`).

---

### ARCH-M3 — Credential storage migration complexity [MEDIUM]

**File**: `frontend/src/utils/credentialStorage.ts:34-91`

Dual-storage migration pattern (localStorage ↔ sessionStorage) with 6 fallback paths. If migration fails mid-process, credentials could be lost. No atomic transaction semantics.

**Recommendation**: Simplify to a single storage strategy with explicit migration on app startup (not on every read).

---

### ARCH-L1 — Frontend Dockerfile uses `npm install` instead of `npm ci` [LOW]

**File**: `frontend/Dockerfile:12-13`

`npm install` allows version range violations. Comment says "for better compatibility with local file dependencies" (`office-word-diff`), but `npm ci` works with local deps if the lockfile is correct.

**Recommendation**: Switch to `npm ci --no-audit --no-fund` after verifying lockfile integrity.

---

### ARCH-L2 — Generated manifests served from root instead of `frontend/public/assets/` [LOW]

**File**: `scripts/generate-manifests.js:44` — `OUTPUT_DIR = path.join(ROOT_DIR, 'generated-manifests')`

The manifest generation script outputs to a `generated-manifests/` directory at the project root. These files are currently only accessible from within the server environment (e.g., via `localhost:3000/manifests/`), not served as static assets from the frontend.

**Current setup**: Manifests are served via an Express route that reads the filesystem. External access requires tunneling (ngrok, Cloudflare Tunnel, etc.).

**Proposed alternative**: Output manifests to `frontend/public/assets/manifests/` so they are bundled and served directly by the Vite/Nginx static file server.

**Benefits**:
- Directly accessible via the frontend URL (same origin, no separate Express route)
- Works out-of-the-box in static hosting scenarios
- Simplified distribution: one URL serves both the add-in UI and the manifest

**Security considerations**:
- Manifests contain the add-in's internal hostname/URL — exposing them publicly means revealing server URLs
- If the frontend is on a public CDN, manifests become publicly discoverable
- Current approach (behind Express with optional auth) is more defensible
- Mitigation: strip internal hostnames from the manifest (use relative paths where possible), or serve manifests only at a non-obvious path

**Recommendation**: Keep the current approach for self-hosted deployments. If/when a SaaS distribution model is desired (users install from a public URL), move manifests to `frontend/public/assets/manifests/` but implement a route-level allowlist for which add-in configurations can be publicly served.

---

## 2. TOOL/PROMPT QUALITY — Full Potential Usage

### TOOL-C1 — Uploaded files sent inline instead of using /v1/files references [CRITICAL — PARTIALLY FIXED ✅ — REMAINING ITEMS DEFERRED]

**Fix applied**: `/v1/files` failure is now logged via `logService.warn` instead of silently swallowed. Token budget now counts `type: 'file'` content parts (200 token fixed cost). Architecture is sound — the inline fallback is correct behavior when the provider doesn't support `/v1/files`.

**Deferred (intentionally not fixed now)**:
- Images still always sent inline as base64 (never use `/v1/files`) — acceptable until image context costs become a bottleneck
- No UI indicator when `/v1/files` upload fails and falls back to inline — low visibility bug moved to USR-L1
- Full document content re-sent on every iteration — blocked on PROSP-H2 (context optimization)

**Files**:
- `frontend/src/composables/useAgentLoop.ts:590-613` — file inclusion in messages
- `frontend/src/composables/useAgentLoop.ts:817-822` — /v1/files upload attempt (silent fallback)
- `frontend/src/composables/useAgentLoop.ts:628-647` — images always inline as base64
- `frontend/src/utils/tokenManager.ts:56-59` — `type: 'file'` not counted in token budget

**Problem**: While the `/api/files` proxy endpoint exists and `uploadFileToPlatform()` attempts to upload files to the LLM provider's `/v1/files` API, the integration has critical gaps:

1. **Silent fallback**: If `/v1/files` upload fails (line 821-822), the error is silently caught and the file falls back to inline content — the user never knows
2. **Images never use /v1/files**: All uploaded images (PNG, JPG) are ALWAYS sent as base64 data-URIs inline (lines 641-644), never as file references
3. **Full content re-sent every iteration**: When the agent loop iterates (tool calls), the entire file content is re-sent in every LLM request as part of the last user message
4. **Token budget blind spot**: `getMessageContentLength()` (tokenManager.ts:47-69) does not account for `type: 'file'` parts — only `text` and `image_url`
5. **Bandwidth waste**: A 5MB PDF's extracted text (~50k chars) is sent inline on every agent iteration instead of being referenced by file_id once

**Impact**: Increased latency, higher token costs, potential context overflow on large documents, unnecessary bandwidth consumption.

**Action**:
1. Make `/v1/files` upload failure visible to the user (warning toast)
2. When `fileId` is available, use `{ type: 'file', file: { file_id: fileId } }` consistently
3. Add `type: 'file'` handling in `getMessageContentLength()`
4. Consider uploading images to `/v1/files` too, not just text files
5. Only inject inline content as a last resort when `/v1/files` is unavailable

---

### TOOL-H1 — Skill doc references non-existent tools [HIGH — FIXED ✅]

**Files**: `frontend/src/composables/useAgentPrompts.ts:101`

`useAgentPrompts.ts` referenced `insertBookmark` and `goToBookmark` tools in the Word agent prompt under **STRUCTURE & ANALYTICS**, but these tools are not defined in `wordTools.ts`. The agent could attempt to call them, resulting in a "tool not found" error.

**Fix applied**: Removed the `insertBookmark` / `goToBookmark` line from the Word agent prompt.

---

### TOOL-H2 — Screenshots underutilized: no auto-verification, not visible to user [HIGH — PARTIALLY FIXED ✅]

**Files**:
- `frontend/src/utils/excelTools.ts:1603-1623` — `screenshotRange` tool
- `frontend/src/utils/powerpointTools.ts:1118-1146` — `screenshotSlide` tool
- `frontend/src/composables/useToolExecutor.ts:89-105` — `__screenshot__` detection
- `frontend/src/types/chat.ts:3-9` — `ToolCallPart.screenshotSrc` field (added)
- `frontend/src/components/chat/ToolCallBlock.vue` — screenshot display (added)

**Fix applied**:
- Added `screenshotSrc?: string` to `ToolCallPart` type
- `useToolExecutor.ts` now stores the screenshot as a data URI on the tool call object when `__screenshot__: true` is detected
- `ToolCallBlock.vue` now displays the screenshot image inline in the chat when `screenshotSrc` is present

**Remaining gaps (not fixed)**:
1. **No auto-verification prompting**: Agent prompts still do NOT instruct the LLM to screenshot after creating charts or modifying slides
2. **No Word screenshot**: Word has no screenshot tool at all
3. **PowerPoint explicitly blocks verification**: `powerpoint.skill.md` says "Do NOT call getAllSlidesOverview to verify" — prevents legitimate verification

---

### TOOL-M1 — Excel `values` parameter typed as `string` but accepts mixed types [MEDIUM]

**File**: `frontend/src/utils/excelTools.ts:182`

The `values` parameter description says items are "string", but Excel cells accept numbers, booleans, dates, and nulls. This can mislead the LLM into always quoting numeric values.

**Action**: Update parameter description to document supported types.

---

### TOOL-M2 — Overlapping Excel read tools [MEDIUM]

**Files**: `frontend/src/utils/excelTools.ts`

`getWorksheetData` (reads active sheet) and `getDataFromSheet` (reads any sheet by name) overlap. Both return CSV data from a worksheet. Could be unified into a single tool with an optional `sheetName` parameter.

**Impact**: The agent may use the wrong tool or call both unnecessarily, wasting tool calls.

---

### TOOL-M3 — No PowerPoint equivalent to Word's `searchAndFormat` [MEDIUM]

**File**: `frontend/src/utils/powerpointTools.ts`

PowerPoint has no tool for applying formatting to specific text within existing shapes (unlike Word's `searchAndFormat`). The only workaround is `eval_powerpointjs`, which is error-prone for the LLM.

**Impact**: Agent cannot reliably bold, color, or resize specific words in PowerPoint slides.

---

### TOOL-M4 — Inconsistent formula locale support [MEDIUM]

**Files**: `frontend/src/composables/useAgentPrompts.ts:28-30`, `frontend/src/utils/constant.ts:2-16`

Agent prompt only handles English/French formula locales, but the language map in `constant.ts` lists 10 languages. German, Spanish, Italian, etc. Excel users won't get correct formula separator guidance (`;` vs `,`).

**Action**: Extend locale detection in `useAgentPrompts.ts` to cover all languages in the language map.

---

### TOOL-L1 — `getRangeAsCsv` missing format documentation [LOW]

**File**: `frontend/src/utils/excelTools.ts:174-176`

No description of the CSV format returned (delimiter, quoting, header handling). The LLM may parse incorrectly.

---

### TOOL-L2 — PowerPoint `slideNumber` should clarify 1-based indexing [LOW]

**File**: `frontend/src/utils/powerpointTools.ts:769-770`

Parameter says "1 = first slide" but could be clearer: "1-based index (1 = first slide, not 0-based)."

---

### TOOL-L3 — Style rules ban em-dashes globally [LOW]

**File**: `frontend/src/utils/constant.ts:20-22`

Global style instructions prohibit em-dashes and semicolons, but these are standard in professional English typography. This may produce unnatural output in formal documents.

**Recommendation**: Restrict this rule to PowerPoint/bullet contexts only, not all content generation.

---

## 3. ERROR HANDLING & DEBUGGABILITY

### ERR-H1 — 4 backend routes bypass `logAndRespond()` and ErrorCodes [HIGH]

**Files**:
- `backend/src/routes/files.js:31, 64, 72, 79` — returns `{ error: '...' }` without code
- `backend/src/routes/feedback.js:23, 46` — returns `{ error: '...' }` without code
- `backend/src/routes/logs.js:25, 29, 56` — returns `{ error: '...' }` without code
- `backend/src/routes/icons.js:13, 25, 47` — returns `{ error: '...', details: '...' }` without code

All other routes use `logAndRespond()` from `utils/http.js` with structured `ErrorCodes`. These 4 routes break the pattern, meaning:
1. Frontend's `categorizeError()` (`backend.ts:101-122`) cannot map error codes — falls back to fragile string inspection
2. Errors are logged without req.logger context enrichment (userId, host, session)
3. The `files.js:79` handler leaks raw error messages to the client

**Action**: Standardize all routes to use `logAndRespond()` with appropriate `ErrorCodes`.

---

### ERR-H2 — Frontend uses `console.warn/error` instead of `logService` (27 instances) [HIGH]

**Files** (27 occurrences across 6 composables):
- `useAgentLoop.ts`: 12 instances
- `useImageActions.ts`: 5 instances
- `useOfficeInsert.ts`: 5 instances
- `useSessionDB.ts`: 3 instances
- `useHealthCheck.ts`: 1 instance
- `useOfficeSelection.ts`: 1 instance

These bypass the centralized `logService` (`logger.ts`), meaning:
1. Errors are not captured in session logs
2. Cannot be submitted to backend via `/api/logs`
3. Cannot be reviewed in the feedback report
4. Debugging requires accessing browser DevTools

**Action**: Replace all `console.warn/error` with `logService.warn/error` calls.

---

### ERR-M1 — Chat route duplicate error handling [MEDIUM]

**File**: `backend/src/routes/chat.js`

`/api/chat` (streaming, lines 12-186) and `/api/chat/sync` (synchronous, lines 188-306) contain ~80% identical error handling code (validation, upstream errors, AbortError/RateLimitError branching). Changes must be applied twice.

**Recommendation**: Extract shared error handler: `handleChatError(res, err, endpoint)`.

---

### ERR-M2 — `files.js:79` leaks raw error message to client [MEDIUM]

**File**: `backend/src/routes/files.js:79`

```javascript
return res.status(500).json({ error: `File upload failed: ${err.message}` })
```

Raw `err.message` could contain internal paths, stack traces, or upstream provider details.

**Action**: Use `sanitizeErrorText()` before including in response, or return a generic message.

---

### ERR-L1 — Silent failures in empty catch blocks [LOW]

**Files**:
- `frontend/src/composables/useAgentLoop.ts` — multiple `try { ... } catch {}` blocks that silently swallow errors
- `frontend/src/utils/powerpointTools.ts:1375-1380` — empty catch in slide iteration loop

**Impact**: Masks API errors that could indicate real problems.

**Recommendation**: At minimum, log a warning in catch blocks.

---

## 4. UX & UI

### UX-M1 — Missing focus indicators (accessibility) [MEDIUM]

**File**: `frontend/src/components/chat/ChatInput.vue:21`

`focus:outline-none` removes the visual focus indicator on the main textarea. Only 8 `focus:ring` instances exist across the entire frontend. Keyboard-only users cannot see which element is focused.

**Action**: Add `focus:ring-2 focus:ring-primary/50` to all interactive elements (input, buttons, select).

---

### UX-H1 — Screenshot images not visible in chat [HIGH]

**File**: `frontend/src/components/chat/ChatMessageList.vue:91-96`

When a screenshot tool executes, the image is injected into the LLM's vision context but **never displayed** to the user. The `imageSrc` field on messages is only populated for DALL-E generated images. Screenshots are invisible — the user only sees "Screenshot captured."

**Action**: When a tool result contains `__screenshot__: true`, render the base64 image inline in the tool call result block. This gives users visual feedback and helps them understand what the agent "sees."

---

### UX-M2 — Hardcoded tooltip strings (i18n gap) [MEDIUM]

**File**: `frontend/src/components/chat/StatsBar.vue:9, 12, 18`

Tooltip texts "Input tokens:", "Output tokens:", "Context usage:" are hardcoded in English. Non-English users see untranslated tooltips.

**Also**: `frontend/src/components/chat/ToolCallBlock.vue:20, 25` — "args", "error", "result" labels are hardcoded.

**Action**: Wrap in `t()` with i18n keys.

---

### UX-L1 — Inline animation styles in ChatInput.vue [LOW]

**File**: `frontend/src/components/chat/ChatInput.vue:54`

Uses `:style="isDraftFocusGlowing ? 'animation-iteration-count: 3; ...' : ''"` inline. Should be in `<style scoped>` with a conditional class.

---

### UX-L2 — Bare URL as link text in AccountTab [LOW]

**File**: `frontend/src/components/settings/AccountTab.vue:61-65`

The link text is a raw URL (`https://getkey.ai.kickmaker.net/`) instead of descriptive text. Poor accessibility for screen readers.

---

### UX-L3 — ChatMessageList max width on mobile [LOW]

**File**: `frontend/src/components/chat/ChatMessageList.vue:47`

`max-w-[95%]` on message bubbles may reduce usable space on small task pane widths (300-450px).

---

## 5. DEAD CODE

### DEAD-M1 — Duplicate tool export aliases in all 4 tool files [MEDIUM]

**Files**:
- `wordTools.ts:1562-1568` — exports both `getToolDefinitions()` and `getWordToolDefinitions`
- `excelTools.ts:1928-1934` — exports both `getToolDefinitions()` and `getExcelToolDefinitions`
- `powerpointTools.ts:1397-1403` — exports both `getToolDefinitions()` and `getPowerPointToolDefinitions`
- `outlookTools.ts:516-522` — exports both `getToolDefinitions()` and `getOutlookToolDefinitions`

Each file exports a generic `getToolDefinitions()` AND a host-specific alias. Only the host-specific names are used in `useAgentLoop.ts`. The generic names are dead code.

**Action**: Remove the redundant `getToolDefinitions()` exports.

---

### DEAD-M2 — `formatRange` redundant with `setCellRange` [MEDIUM]

**File**: `frontend/src/utils/excelTools.ts`

`formatRange` (lines 525-737) is functionally redundant with `setCellRange`'s formatting parameter (lines 189-239). Both apply formatting to Excel ranges. The agent prompt already marks `setCellRange` as PREFERRED.

**Impact**: Occupies a tool slot (139 tools total, max 128 per host), confuses the LLM about which to use.

**Action**: Deprecate `formatRange` or merge its unique features into `setCellRange`.

---

### DEAD-L1 — Unused tool signature for deduplication [LOW]

**File**: `frontend/src/composables/useToolExecutor.ts:78`

`safeStringify(toolArgs)` creates a call signature, but no deduplication logic uses it. Appears to be a remnant of an incomplete feature.

---

## 6. CODE DUPLICATION & GENERALIZATION

### DUP-H1 — Identical tool wrapper pattern repeated 4 times [HIGH]

**Files**: `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`

Each file independently defines:
1. A host-specific type (`WordToolTemplate`, `ExcelToolTemplate`, etc.) — all follow `Omit<ToolDefinition, 'execute'> & { executeXXX: ... }`
2. A host runner (`runWord`, `runExcel`, etc.) — all are `<T>(action) => executeOfficeAction(action)`
3. An error wrapper in `buildExecute` — identical try/catch with `JSON.stringify({ error: true, message, tool, suggestion })`
4. A `getToolDefinitions()` + `getXxxToolDefinitions` alias pair

The shared factory `createOfficeTools()` in `common.ts:48-58` already exists but the individual wrapper functions and types are still duplicated.

**Action**: Create a generic `OfficeToolTemplate<THost>` type and a shared `buildExecuteWrapper(runner)` factory in `common.ts`. Each tool file would only define its tool definitions, not boilerplate.

---

### DUP-M1 — String truncation pattern repeated 4 times [MEDIUM]

**Files**:
- `wordTools.ts:1511` — `code.slice(0, 300) + (code.length > 300 ? '...' : '')`
- `wordTools.ts:1543` — `code.slice(0, 200) + '...'`
- `outlookTools.ts:463` — `code.slice(0, 300) + (code.length > 300 ? '...' : '')`
- `outlookTools.ts:494` — `code.slice(0, 200) + '...'`

**Action**: Extract to `truncateString(str: string, maxLen: number): string` in `common.ts`.

---

### DUP-M2 — Inconsistent error response format across tools [MEDIUM]

Tool implementations return errors in multiple formats:
- `JSON.stringify({ error: true, message, tool, suggestion })` (most tools)
- `JSON.stringify({ success: false, error })` (some Excel tools)
- Plain string `"Error: ..."` (some edge cases)

**Action**: Standardize on a single error format. The `buildExecute` wrapper already handles most cases — ensure all tools go through it.

---

## 7. CODE QUALITY & MAINTAINABILITY

### QUAL-H1 — 128 instances of `: any` across tool utilities [HIGH]

**Files** (top offenders):
- `powerpointTools.ts`: 50 instances
- `outlookTools.ts`: 21 instances
- `excelTools.ts`: 20 instances
- `officeDocumentContext.ts`: 12 instances

Office.js types are available via `@types/office-js`. The `declare const Office: any` pattern (e.g., `powerpointTools.ts:18-19`) bypasses all type checking.

**Impact**: No compile-time safety for Office API calls. Typos in property names or method signatures go undetected.

**Action**: Install `@types/office-js` (if not already) and replace `any` with proper types, at least for the most-used APIs (`Excel.run`, `PowerPoint.createPresentation`, `Office.context.mailbox`).

---

### QUAL-M1 — Magic numbers not in constants [MEDIUM]

Scattered values that should be in `frontend/src/constants/limits.ts` or `backend/src/config/limits.js`:

| Value | Location | Meaning |
|-------|----------|---------|
| `255` | `wordTools.ts:338, 429, 532` | Word search text max length |
| `20`, `15`, `12.5` | `wordTools.ts:128-130` | Font size thresholds for heading detection |
| `300`, `200` | `wordTools.ts:1511, 1543` | Code truncation lengths |
| `20_000` | `outlookTools.ts:39, 504` | Outlook action timeout |
| `1000, 2000` | `officeAction.ts:20` | Retry backoff delays |
| `50 * 1024 * 1024` | `files.js:20` | Files API max size |

---

### QUAL-M2 — Frontend console.log in production code [MEDIUM]

**27 instances** in composables (see ERR-H2) plus additional instances in utility files:
- `credentialCrypto.ts`: 7 instances
- `credentialStorage.ts`: 3 instances
- `cryptoPolyfill.ts`: 2 instances

These should use `logService` for structured logging.

---

### QUAL-M3 — Large Vue components exceeding 300 lines [MEDIUM]

| Component | Lines | Responsibilities |
|-----------|-------|-----------------|
| `HomePage.vue` | 592 | Layout, routing, state orchestration, confirmation dialogs |
| `ChatMessageList.vue` | 336 | Message rendering, tool call display, actions, markdown |
| `ChatInput.vue` | 307 | Input, file upload, model selection, send/stop |

**Recommendation**: Extract focused sub-components:
- `AttachedFilesList.vue` from ChatInput
- `MessageItem.vue` from ChatMessageList
- `ConfirmationDialogs.vue` from HomePage

---

### QUAL-L1 — Boolean parameter overloading [LOW]

**Files**:
- `powerpointTools.ts:256` — `insertIntoPowerPoint(text, useHtml = true)`
- `powerpointTools.ts:301` — `insertMarkdownIntoTextRange(..., forceStripBullets = false)`

Boolean parameters are unclear at call sites. Prefer options objects or enums.

---

### QUAL-L2 — Async/Promise pattern inconsistency [LOW]

**File**: `frontend/src/utils/outlookTools.ts`

Outlook tools mix `async/await` with callback-based `Office.AsyncResult` patterns (due to Outlook API limitations). While necessary, the wrapping in `resolveAsyncResult()` could be documented more clearly.

---

## 8. CROSS-CUTTING: SECURITY

### SEC-M1 — Rate limiting is IP-based only [MEDIUM — INFO]

**File**: `backend/src/server.js:42-83`

Rate limits use IP-based tracking (`express-rate-limit`). Behind a shared proxy (e.g., corporate network), all users share the same limit. Per-user rate limiting (via `X-User-Key`) would be more accurate.

**Note**: This is documented as a known limitation, not a regression.

---

### SEC-L1 — CSP allows `unsafe-inline` and `unsafe-eval` [LOW — ACCEPTED]

**File**: `frontend/nginx.conf:25`

Required for Office add-in compatibility. Cannot be removed without breaking Office.js runtime. Accepted risk.

---

## 9. USER-REPORTED ISSUES

### USR-C1 — Feedback system does not save complete debug bundle [CRITICAL — FIXED ✅]

**Fix applied**: `FeedbackDialog.vue` now includes: full chat history (stripped of base64 data) from IndexedDB, and `systemContext` (`host`, `appVersion`, `modelTier`, `userAgent`). Backend `feedback.js` now saves all fields and logs a summary (logCount, chatMessageCount, hasSystemContext). Payload limit raised to 20MB.

**Files**:
- `frontend/src/components/settings/FeedbackDialog.vue:89-116` — feedback submission UI
- `frontend/src/api/backend.ts:562-578` — `submitFeedback()` API call
- `backend/src/routes/feedback.js:17-48` — feedback storage
- `frontend/src/utils/logger.ts:131-133` — `getSessionLogs()` only returns frontend buffer

**Problem**: The feedback form collects a user comment, a category, and the current frontend log buffer — but does NOT create a complete debug bundle:

1. **No chat history**: The full conversation (messages, tool calls, tool results) is NOT included in the feedback payload
2. **No backend logs**: Only frontend `logService` buffer is sent — backend request logs (with correlation IDs) are not included
3. **No system context**: Browser version, Office host version, add-in version, model used — none of this is captured
4. **Frontend-only logs**: `logService.getSessionLogs(sessionId)` returns only what was logged via `logService.*()` — and since 27 console.warn/error calls bypass logService (see ERR-H2), many errors are missing

**Impact**: When a user reports a bug, the developer has no way to reconstruct the session without asking the user for more details.

**Action**:
1. Include full chat `history` (messages array) in the feedback payload
2. Add system context: `{ officeHost, officeVersion, browserUA, addinVersion, modelTier, sessionId }`
3. Optionally correlate backend logs by `x-request-id` or `sessionId`
4. Save the complete bundle as a structured JSON in `logs/feedback/`

---

### USR-H1 — Double bullets generated in PowerPoint [HIGH — FIXED ✅]

**Files**:
- `frontend/src/utils/powerpointTools.ts:344` — `findShapeOnSlide()` shapes.load call (fixed)
- `frontend/src/utils/powerpointTools.ts:301-330` — `insertMarkdownIntoTextRange()`
- `frontend/src/utils/powerpointTools.ts:387-404` — `hasNativeBullets()` detection

**Root cause**: `findShapeOnSlide()` loaded `placeholderFormat` but NOT `placeholderFormat/type`. This meant `placeholderFormat.type` was always `undefined`, so `isBodyPlaceholder` never became `true` for placeholder shapes — causing the native bullet strip logic to be skipped, resulting in double bullets.

**Fix applied**: Changed `shapes.load('items,items/id,items/name,items/placeholderFormat')` to `shapes.load('items,items/id,items/name,items/placeholderFormat,items/placeholderFormat/type')`. Now `placeholderFormat.type` is correctly loaded, `isBodyPlaceholder` returns `true` for body/content placeholders, and `forceStripBullets` is set — preventing double bullets.

**Remaining gaps (not fixed)**:
1. `hasNativeBullets()` only checks EXISTING paragraphs — empty shapes with bullet defaults in XML still may double-bullet on fresh insert
2. No stronger prompt guidance added yet

---

### USR-H2 — Long latency (1-2 min) between successive tool calls [HIGH — PARTIALLY FIXED ✅]

**Files**:
- `frontend/src/composables/useAgentLoop.ts:272-305` — agent loop iteration (timer added)
- `frontend/src/utils/tokenManager.ts:71-160` — `prepareMessagesForContext()`
- `backend/src/config/models.js:147-149` — `reasoning_effort` parameter

**Fix applied**: Added an elapsed time counter to the `currentAction` status label during LLM inference. The status now updates every second: "Analyzing... (5s)", "Waiting for AI... (12s)", etc. Users can now see that the system is working and how long it has been waiting — reducing perceived latency and anxiety about hung states.

**Root causes** (not fixed — structural):

1. **Context bloat**: Each iteration re-sends the full message history (up to 1.2M chars / ~400k tokens). After several tool calls that return large results (full spreadsheet data, document content), the context sent to the LLM grows significantly, increasing inference time.

2. **`reasoning_effort` parameter**: If set to `"high"` (backend config/models.js:147), GPT-5 models spend extra time in the "thinking" phase. This can add 30-90 seconds per call with complex tool history.

3. **Tool result accumulation**: Tool results are pushed to `currentMessages` and never summarized. After 5-6 tool calls, potentially 500k+ chars in context.

**Remaining actions**:
1. Add aggressive truncation for tool results older than N iterations (see PROSP-H2)
2. Add a visible context window % indicator in the status bar
3. Log context size per iteration to help diagnose

---

### USR-M1 — Scroll behavior doesn't match user expectations [MEDIUM]

**Files**:
- `frontend/src/composables/useHomePage.ts:71-107` — scroll helpers
- `frontend/src/composables/useAgentLoop.ts:254-255, 429-430` — scroll trigger points
- `frontend/src/composables/useAgentStream.ts:51-56` — no auto-scroll during streaming

**Current behavior**:
- Session load: scrolls to **top of last message** (not top of conversation)
- Message send: scrolls to **bottom** (shows user input at bottom)
- During streaming: **no auto-scroll** (user can scroll freely) — correct
- Response complete: scrolls to **bottom** of response

**User expectations** (from feedback):
1. On session load / app start: scroll to **top of conversation** (first message)
2. On message send: scroll to **top of user's message** (to see their request)
3. During LLM streaming: free scroll — correct, keep as-is
4. On response complete: scroll to **top of assistant response** (to start reading)

**Gaps**:
- Point 1: Currently scrolls to last message, not conversation top. For old sessions with long history, user sees the end.
- Point 2: Currently scrolls to bottom, pushing user message out of view on long contexts.
- Point 4: Currently scrolls to bottom of response. If response is long, user sees the end first and must scroll up.

**Action**:
1. Session load: `container.scrollTo({ top: 0 })` to start at conversation top
2. Message send: scroll to the newly created user message element (`scrollToMessageTop()` targeting user message, not assistant)
3. Response complete: scroll to top of assistant response message element (current `scrollToMessageTop()` logic, but ensure it targets response start)

---

### USR-M2 — Context window percentage already visible but not prominent enough [MEDIUM]

**File**: `frontend/src/components/chat/StatsBar.vue`

The stats bar already shows context usage with color-coded warnings (green <70%, orange 70-89%, red >=90%), but users don't understand WHY the agent is slow. The context % is visible but not prominent enough during long agent sessions.

**Action**: Consider adding a tooltip or notification when context exceeds 80%: "Response may be slower — large conversation context."

---

### USR-L1 — No visual feedback when /v1/files upload silently fails [LOW]

**File**: `frontend/src/composables/useAgentLoop.ts:821-822`

When `uploadFileToPlatform()` fails, the error is caught silently and the file falls back to inline content. The user has no idea their file was not uploaded efficiently.

**Action**: Show a subtle warning: "File uploaded inline (provider file API unavailable)."

---

## 10. PROSPECTIVE IMPROVEMENTS (DEFERRED)

### PROSP-1 — Dynamic tool loading: lazy tool categories instead of full tool set [DEFERRED]

**Current state**: All tools for the active Office host (up to 49 for Excel, 41 for Word) are sent in every LLM request in the `tools` parameter. This consumes significant context window space (tool schemas are verbose JSON).

**Proposal**: Instead of sending all tools upfront, the system prompt would list available tool categories (e.g., "Reading", "Writing", "Formatting", "Charts", "Tables") and the agent would request a specific category when needed. The frontend would then inject only those tools for the next iteration.

**Analysis**:

| Aspect | Assessment |
|--------|-----------|
| **Context savings** | HIGH — Tool schemas can consume 20-40k tokens. Loading only relevant categories could reduce this by 50-70%. |
| **Latency improvement** | MEDIUM — Fewer tools = faster LLM inference per iteration. BUT adds 1 extra round-trip to "request" the category. |
| **Accuracy impact** | MIXED — LLMs perform better with fewer, more focused tools. But the agent may not know which category it needs upfront, leading to wrong category requests and wasted iterations. |
| **Quick action impact** | NEGATIVE — Quick actions need specific tools immediately (e.g., "Bullets" needs `insertContent`). Adding a category-request step would slow down quick actions significantly. Quick actions should bypass this mechanism and always include their required tools. |
| **Implementation complexity** | HIGH — Requires: tool categorization metadata, category request/response protocol in agent loop, bypass for quick actions, prompt engineering for category selection. |

**My recommendation**: **NOT recommended as described.** The category-request pattern adds latency (extra LLM round-trip) and uncertainty (wrong category). Better alternatives:

1. **Static tool profiles per intent**: Detect user intent from the first message (e.g., "make a chart" → chart tools + write tools) and pre-select relevant tools. No extra round-trip needed.
2. **Two-tier tools**: Always include a small "core" tool set (read, write, format) + dynamically add specialized tools (charts, conditional formatting, pivot tables) based on keywords in the user message.
3. **Tool description compression**: Shorten tool descriptions in the JSON schema (move detailed guidance to skill docs which are in the system prompt). This reduces token cost without changing the protocol.

**Criticality**: LOW — Current tool count (max 49) is within LLM comfort zone. GPT-5.2 handles 128+ tools. Optimize only if latency measurements confirm tools are the bottleneck.

---

### PROSP-2 — Conversation history optimization and document re-accessibility [DEFERRED]

**Current state**: `prepareMessagesForContext()` in `tokenManager.ts` does backwards iteration, keeping recent messages within a 1.2M char budget. Document content injected via `<attached_files>` tags is only in the user message where it was attached. Tool results accumulate without summarization.

**Questions raised**:

1. **Can the LLM re-access an uploaded document in a later message?**
   Answer: YES, if the message containing the document is still within the context window. But if truncated by `prepareMessagesForContext()`, it's lost. There is no persistent document store accessible by the LLM.

2. **Is too much irrelevant history sent?**
   Likely YES. Users often switch topics within a session (e.g., "format this table" then "write a formula" then "create a chart"). Old tool results from unrelated actions waste context space.

3. **Are `<doc_context>` and `<attached_files>` tags optimal?**
   The separation is good (document content vs. uploaded files), but both are injected into the last user message on every iteration, meaning the LLM re-processes them each time.

**Analysis**:

| Strategy | Pros | Cons |
|----------|------|------|
| **Current (full history, backward truncation)** | Simple, preserves recent context | Wastes tokens on old irrelevant tool results |
| **Summarize old turns** | Reduces token usage, keeps key context | Requires an extra LLM call to summarize, adds latency |
| **Sliding window (last N turns only)** | Predictable context size | Loses document context if attached early |
| **Topic-based segmentation** | Only relevant history sent | Very hard to implement reliably |
| **Document pinning** | Uploaded documents always in context | Uses tokens even when doc is irrelevant |

**My recommendation**: Implement **aggressive tool result summarization** (replace old tool results with 1-line summaries after 3 iterations) + **document pinning** (uploaded files stay in context until explicitly dismissed). This is the best balance of simplicity and effectiveness.

**Criticality**: MEDIUM — Directly impacts latency (USR-H2) and document accessibility. Should be implemented alongside tool result truncation.

---

### PROSP-3 — Split PRD into domain-specific sub-documents [DEFERRED]

**Current state**: `PRD.md` is a single 550+ line document covering all features across all Office hosts + infrastructure + chat UX.

**Proposal**: Split into:
- `PRD.md` — Overview, deployment, cross-app features, links to sub-PRDs
- `docs/PRD-infrastructure.md` — Backend, deployment, Docker, security
- `docs/PRD-chat.md` — Chat UI, sessions, stats, general features
- `docs/PRD-word.md` — Word-specific features
- `docs/PRD-excel.md` — Excel-specific features
- `docs/PRD-powerpoint.md` — PowerPoint-specific features
- `docs/PRD-outlook.md` — Outlook-specific features

**Analysis**:

| Aspect | Assessment |
|--------|-----------|
| **Maintainability** | HIGH benefit — Each sub-PRD is focused and smaller, easier to update |
| **Discoverability** | MEDIUM — Requires proper cross-linking; risk of stale links |
| **Agent context** | HIGH benefit — Agent can load only the relevant sub-PRD for the current Office host instead of the entire 550-line document |
| **Claude.md integration** | EASY — Add rule: "When working on {host}, read `docs/PRD-{host}.md` before implementing" |
| **Migration effort** | LOW — Content already organized by host in current PRD, just needs extraction |

**My recommendation**: **Recommended.** The current PRD is too large for efficient agent consumption. Split it and add a routing rule in `Claude.md`. Keep the root `PRD.md` as an index with links.

**Criticality**: LOW — Nice to have for DX, not blocking.

---

### PROSP-4 — Templates for Design Review, Commits, and PRs [DEFERRED]

**Proposal**: Create reusable templates in `docs/templates/` or directly in `Claude.md`:
- **Design Review template**: Standard axes (Architecture, Tool/Prompt Quality, Error Handling, UX/UI, Dead Code, Code Duplication, Code Quality) with severity levels
- **Commit message template**: Type prefix + scope + description format
- **PR template**: Summary bullets + test plan + compatibility notes

**Analysis**:

The Design Review template is already effectively defined by this v10.1 document's structure. Formalizing it would:

1. **Ensure consistency** across reviews — each review covers the same axes
2. **Speed up reviews** — agent knows exactly what to analyze
3. **Enable diff tracking** — compare v10 to v11 systematically

For commits and PRs, `Claude.md` sections 12-13 already define expectations. A `.github/pull_request_template.md` file would enforce PR structure automatically.

**My recommendation**:
1. Add a DR template section in `Claude.md` with the 8 standard axes
2. Create `.github/pull_request_template.md` with the Summary/Test Plan/Compatibility format
3. Commit templates are already well-defined in `Claude.md` section 12 — no change needed

**Criticality**: LOW — Process improvement, not functional.

---

### PROSP-5 — Claude.md overhaul: is it actually used effectively? [DEFERRED]

**Current state**: `Claude.md` is 302 lines covering 15 sections: scope, architecture, working principles, API contracts, frontend/backend guidelines, docs, PRD, PowerPoint agent, known issues, validation, commit/PR, strict agent rules, vibe coding rules.

**Honest assessment**:

| Section | Actually Used? | Value |
|---------|---------------|-------|
| §1 Scope + companion docs | YES — agents check this | HIGH |
| §2 Architecture snapshot | YES — critical reference | HIGH |
| §3 Working principles | PARTIALLY — too generic | MEDIUM |
| §4 API contract rules | YES — prevents regressions | HIGH |
| §5 Frontend guidelines | PARTIALLY — tool counts may be stale | MEDIUM |
| §6 Backend guidelines | YES — actively used | HIGH |
| §7 Docs guidelines | RARELY — agents often skip | LOW |
| §8 PRD guidelines | RARELY — too detailed for daily use | LOW |
| §9 PowerPoint agent | YES — used for prompt construction | HIGH |
| §10 Known issues | REDIRECT only — good pattern | HIGH |
| §11 Validation checklist | SOMETIMES — not enforced | MEDIUM |
| §12-13 Commit/PR | YES — actively followed | HIGH |
| §14 Strict agent rules | YES — security boundary | HIGH |
| §15 Vibe coding rules | YES — prevents PowerShell errors | HIGH |

**Issues identified**:
1. **Tool counts in §5 (line 150-155) get stale quickly** — should be auto-generated or reference code, not hardcoded
2. **§7-8 (Docs/PRD guidelines) are rarely consulted** — too verbose, could be simplified
3. **No host-specific routing** — agent reads entire 302 lines regardless of whether task is Word, Excel, or Outlook
4. **Missing rules**: No guidance on screenshot verification, file upload strategy, or context optimization
5. **Language inconsistency**: Some sections reference French terms despite §7 requiring English docs

**My recommendation**:
1. **Trim §7-8** to 2-3 key rules each (currently 40+ lines combined)
2. **Add missing rules**: screenshot verification after visual modifications, /v1/files usage preference, tool result size awareness
3. **Make tool counts dynamic**: Replace hardcoded counts with "see each `*Tools.ts` file"
4. **Add routing rule**: "When task involves a specific Office host, prioritize reading the corresponding skill doc"
5. **Add DR/PR templates** as proposed in PROSP-4
6. **Full rewrite is NOT recommended** — the structure is sound, just needs targeted updates

**Criticality**: MEDIUM — An outdated Claude.md causes agent drift and inconsistency. Targeted updates would have high ROI.

---

## PREVIOUSLY FIXED ITEMS (v9.x — All Verified OK)

| ID | Description | Status |
|----|-------------|--------|
| GEN-C1 | File attachment race condition | FIXED |
| PPT-C2 | Infinite loop on image slide creation | FIXED |
| GEN-C3 | Slowdown after 5+ tool calls | FIXED (context pruning) |
| WD-C4 | Crash on PDF insert without selection | FIXED |
| XL-M1 | Chart X-axis treated as data series | FIXED |
| PPT-M2 | Agent ignores template placeholders | FIXED |
| GEN-M3 | Task pane width stuck at 300px | FIXED |
| GEN-M4 | Premature timestamp display | FIXED |
| PPT-M5 | Speaker notes not self-inserting | FIXED |
| PPT-L1 | Impact action not suited for PowerPoint | FIXED |
| PPT-L2 | Generated image cropped/square | FIXED |
| GEN-L3 | Formatting checkboxes utility | FIXED (Phantom Context) |

---

## SUMMARY & PRIORITY MATRIX

### Phase 0 — 🔴 CRITICAL (User-facing bugs & data inefficiency) — ✅ COMPLETE
1. **TOOL-C1**: ~~Fix /v1/files integration~~ — **PARTIALLY FIXED** ✅ (silent failure logged, token budget fixed; remaining items → Phase 4 deferred)
2. **USR-C1**: ~~Complete the feedback debug bundle~~ — **FIXED** ✅

### Phase 1 — 🟠 HIGH (Reliability & User Experience) — 6 Active / 5 Deferred
**FIXED** (4 items):
3. **USR-H1**: ~~Fix double bullets in PowerPoint~~ — **FIXED** ✅ (`placeholderFormat/type` now loaded)
4. **USR-H2**: ~~Reduce latency between tool calls~~ — **PARTIALLY FIXED** ✅ (elapsed timer added; structural context optimization → Phase 4)
5. **TOOL-H2**: ~~Display screenshots in chat~~ — **PARTIALLY FIXED** ✅ (screenshots now visible; auto-verification → Phase 4)
9. **TOOL-H1**: ~~Fix skill doc referencing non-existent tools~~ — **FIXED** ✅

**Still Active** (5 items):
6. **ERR-H1**: Standardize all backend routes to use `logAndRespond()` + ErrorCodes
7. **ERR-H2**: Replace all `console.warn/error` with `logService` (27 instances)
8. **DUP-H1**: Extract shared tool wrapper boilerplate to `common.ts`
10. **QUAL-H1**: Replace critical `any` types with proper Office.js types
— **PROSP-H2**: Conversation history optimization (blocking 3 deferred items) → Phase 4

### Phase 2 — 🟡 MEDIUM (Maintainability & DX) — 9 Active
11. **USR-M1**: Fix scroll behavior (session load → top, send → user msg, complete → response top)
12. **ARCH-H1**: Split `useAgentLoop.ts` into focused composables
13. **ARCH-H2**: Reduce prop drilling in HomePage with provide/inject
14. **ERR-M1**: Extract shared chat error handler
15. **ERR-M2**: Sanitize error message in files.js:79
16. **TOOL-M1-M4**: Fix parameter docs, merge overlapping tools, extend locale support
17. **DEAD-M1-M2**: Remove dead exports, deprecate redundant `formatRange`
18. **DUP-M1-M2**: Extract `truncateString`, standardize error format
19. **QUAL-M1-M3**: Consolidate magic numbers, fix console logging, split large components
20. **UX-M1-M3**: Restore focus indicators, translate hardcoded strings, context % warning
— **PROSP-2**: Claude.md overhaul (missing rules, stale counts) → Phase 4

### Phase 3 — 🟢 LOW (Polish) — 5 Active
21. **UX-L1-L3**: Inline styles, link text, mobile width
22. **ARCH-L1**: Switch to `npm ci` in Dockerfile
23. **ARCH-L2**: Evaluate manifest accessibility — move `generated-manifests/` to `frontend/public/assets/` for SaaS distribution
24. **QUAL-L1-L2**: Boolean params, async pattern docs
25. **USR-L1**: Show warning when /v1/files upload silently fails
— **PROSP-1/3/4/5**: Dynamic tool loading, PRD split, templates, intent profiles → Phase 4

### Phase 4 — Deferred Items (Not Yet Addressed)

**Deferred actions from partially-fixed Phase 0–1 items** (actionable, blocked on design decisions or dependencies):

#### 🟠 TOOL-C1 Remaining Items (HIGH)
- **Images never use /v1/files**: All uploaded images are always sent inline as base64, never as file references. Consider uploading images to `/v1/files` too.
- **No UI indicator for /v1/files fallback**: When `/v1/files` upload fails and falls back to inline, user has no visual feedback. Add a warning toast.
- **Full document re-sent on every iteration**: Files injected in iteration 1 are re-sent in full on iterations 2+. Blocked on PROSP-H2 (context optimization).

#### 🟠 TOOL-H2 Remaining Items (HIGH)
- **No auto-verification prompting**: Agent prompts do NOT instruct the LLM to screenshot after creating charts or modifying slides for self-verification. Add screenshot guidance to Excel and PowerPoint prompts.
- **PowerPoint explicitly blocks verification**: `powerpoint.skill.md` line 224 says "Do NOT call getAllSlidesOverview to verify" — defensive rule prevents legitimate verification workflows.
- **No Word screenshot tool**: Word has no screenshot capability at all, preventing visual verification of formatting changes.

#### 🟠 USR-H1 Remaining Items (HIGH)
- **Empty shapes with default bullets**: `hasNativeBullets()` only checks EXISTING paragraphs — empty shapes with bullet XML defaults still may double-bullet on first insert.
- **Stronger prompt guidance needed**: Add explicit rule to Word agent prompt: "When inserting into PowerPoint body placeholders, NEVER use markdown bullet syntax (`- `). The shape already has native bullets."

#### 🟠 USR-H2 Remaining Items (HIGH)
- **Context bloat structural issue**: Each iteration re-sends full message history (up to 1.2M chars / ~400k tokens) without summarization. Blocked on PROSP-H2.
- **Tool result accumulation**: Tool results pushed to `currentMessages` never summarized. After 5–6 tool calls, context can exceed 500k chars.
- **No context window % indicator**: Add visible indicator in status bar so users understand why LLM is slow.

**Prospective improvements** (architectural enhancements, not blocking but high-value):

#### PROSP-H2: Conversation History Optimization & Context Management 🟠 HIGH
- **Tool result summarization**: After N iterations, replace detailed tool results with brief "Tool X: [1-line summary]"
- **Document pinning**: Keep recently-uploaded files pinned in context window instead of re-injecting on every iteration
- **Backwards iteration improvements**: Smarter message selection that prioritizes tool calls/responses over old chat history
- **Root blocker for**: TOOL-C1 remaining items, USR-H2 latency, context overflow on large projects

---

#### PROSP-1: Dynamic Tool Loading — Intent-Based Tool Sets 🟢 LOW
**Current**: All tools (up to 49 for Excel, 41 for Word) sent in every LLM request.
**Problem**: Verbose JSON schemas consume significant context window (~50k chars per tool set).
**Not recommended as-is** — Rather than lazy loading, consider static intent profiles:
- `excel-chart-creation`: chart, data, analysis tools only
- `excel-data-entry`: data manipulation, cell formatting tools only
- `word-formatting`: text, style, formatting tools only
- Agent selects profile based on user instruction or first message

#### PROSP-2: Claude.md Targeted Overhaul 🟡 MEDIUM
**Current state**: 302 lines, 15 sections. Some sections never consulted (§7–8 Docs/PRD guidelines).
**Issues**:
1. Tool counts in §5 get stale quickly — should auto-reference code
2. §7–8 (Docs/PRD) are verbose and rarely used — trim to 3–5 key rules each
3. Missing rules on screenshot verification, /v1/files strategy, context management
4. No host-specific routing — agent reads all 302 lines regardless of task

**Recommended actions**:
- Trim §7–8 from 40+ lines to 5–10 lines combined
- Add screenshot verification guidance: "After creating visuals, call screenshot tools and compare with originals"
- Add /v1/files guidance: "Prefer file references for documents >10KB; upload as JSON multipart if provider supports /v1/files"
- Replace hardcoded tool counts with "See each `*Tools.ts` file for the complete list"
- Add routing rule: "When task is {Host}-specific, prioritize reading `docs/PRD-{Host}.md` and `frontend/src/skills/{host}.skill.md`"

#### PROSP-3: Split PRD into Domain-Specific Sub-Documents 🟢 LOW
**Current**: Single 550+ line `PRD.md` covering all Office hosts + infrastructure + UX.
**Proposal**:
- `docs/PRD-index.md` — Top-level overview and cross-links
- `docs/PRD-word.md` — Word-specific features, constraints, workflows
- `docs/PRD-excel.md` — Excel-specific features, constraints, workflows
- `docs/PRD-powerpoint.md` — PowerPoint-specific features, constraints, workflows
- `docs/PRD-outlook.md` — Outlook-specific features, constraints, workflows

**Benefits**: Smaller docs, more agent-friendly, direct host-specific context without bloat. Easy migration (content already organized by host). Add routing rule in `Claude.md`.

#### PROSP-4: Templates for Design Review, Commits, and PRs 🟢 LOW
**Current**: Design Review structure is well-defined by v10.1. Commit/PR expectations are in `Claude.md` §12–13.
**Proposal**:
1. Formalize DR template in `Claude.md` with the 8 standard axes + severity levels
2. Create `.github/pull_request_template.md` with Summary / Test Plan / Breaking Changes format
3. No commit template needed — `Claude.md` §12 already sufficient

**Value**: Consistency across reviews, easier diff tracking v10 → v11, automatic PR structure enforcement.

#### PROSP-5: Consider Static Intent Profiles Instead of Full Dynamic Loading 🟡 MEDIUM
**Alternative to PROSP-1**: Rather than lazy-load tools by category, define static profiles:
- `chart-workflows`: Excel + PowerPoint tools for chart creation from data or images
- `data-cleanup`: Excel data validation, dedupe, format tools
- `document-assembly`: Word + Excel + PowerPoint tools for multi-document generation

**Benefit**: Predictable, testable, avoids context thrashing from tool set switching.
**Drawback**: Requires upfront profiling of common user workflows.

---

## Deferred Items Summary by Severity

| Severity | Count | Status | Action Required |
|----------|-------|--------|-----------------|
| 🔴 **Critical** | 0 | ✅ All fixed or deferred | None — Phase 0 complete |
| 🟠 **High** | 5 deferred | ⏳ Pending | 4 items from partial fixes; 1 prospective (PROSP-H2) |
| 🟡 **Medium** | — | — | 1 prospective: Claude.md overhaul |
| 🟢 **Low** | 3 | — | 3 prospective: tool loading, PRD split, templates |

---

*This review covers the full codebase as of 2026-03-09. Line numbers reference the current state of the repository on the `review/design-review-v10` branch.*
