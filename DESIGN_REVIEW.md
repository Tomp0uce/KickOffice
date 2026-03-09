# DESIGN_REVIEW.md — Code Audit v10.0

**Date**: 2026-03-09
**Version**: 10.0
**Scope**: Full design review — Architecture, tool/prompt quality, error handling, UX/UI, dead code, code quality & maintainability

---

## Health Summary (v10.0)

All previous critical and major items from v9.x have been resolved. This v10.0 review is a comprehensive deep-dive across 8 axes, identifying new improvement opportunities after recent large-scale changes (OOXML editing, chart extraction, image registry, session persistence, header auto-detect).

| Category | Critical | High | Medium | Low |
|----------|----------|------|--------|-----|
| Architecture | 0 | 2 | 3 | 1 |
| Tool/Prompt Quality | 0 | 1 | 4 | 3 |
| Error Handling | 0 | 2 | 2 | 1 |
| UX/UI | 0 | 0 | 2 | 3 |
| Dead Code | 0 | 0 | 2 | 1 |
| Code Duplication | 0 | 1 | 2 | 0 |
| Code Quality | 0 | 1 | 3 | 2 |
| **Total** | **0** | **7** | **18** | **11** |

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

## 2. TOOL/PROMPT QUALITY — Full Potential Usage

### TOOL-H1 — Skill doc references non-existent tools [HIGH]

**Files**: `frontend/src/skills/word.skill.md` (line 101), `frontend/src/composables/useAgentPrompts.ts`

`word.skill.md` references `insertBookmark` and `goToBookmark` tools, but these are not defined in `wordTools.ts`. The agent may attempt to call them, resulting in a "tool not found" error.

**Action**: Either implement these tools or remove references from the skill document.

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

### Phase 1 — High Priority (Reliability & Debuggability)
1. **ERR-H1**: Standardize all backend routes to use `logAndRespond()` + ErrorCodes
2. **ERR-H2**: Replace all `console.warn/error` with `logService` (27 instances)
3. **DUP-H1**: Extract shared tool wrapper boilerplate to `common.ts`
4. **TOOL-H1**: Fix skill doc referencing non-existent tools
5. **QUAL-H1**: Replace critical `any` types with proper Office.js types

### Phase 2 — Medium Priority (Maintainability & DX)
6. **ARCH-H1**: Split `useAgentLoop.ts` into focused composables
7. **ARCH-H2**: Reduce prop drilling in HomePage with provide/inject
8. **ERR-M1**: Extract shared chat error handler
9. **ERR-M2**: Sanitize error message in files.js:79
10. **TOOL-M1-M4**: Fix parameter docs, merge overlapping tools, extend locale support
11. **DEAD-M1-M2**: Remove dead exports, deprecate redundant `formatRange`
12. **DUP-M1-M2**: Extract `truncateString`, standardize error format
13. **QUAL-M1-M3**: Consolidate magic numbers, fix console logging, split large components

### Phase 3 — Low Priority (Polish)
14. **UX-M1-M2**: Restore focus indicators, translate hardcoded strings
15. **UX-L1-L3**: Inline styles, link text, mobile width
16. **ARCH-L1**: Switch to `npm ci` in Dockerfile
17. **QUAL-L1-L2**: Boolean params, async pattern docs
18. Remaining LOW items

---

*This review covers the full codebase as of 2026-03-09. Line numbers reference the current state of the repository on the `fix/speaker-notes-and-related-issues` branch.*
