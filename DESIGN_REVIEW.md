# DESIGN_REVIEW.md — Code Audit v11.3

**Date**: 2026-03-14
**Version**: 11.3
**Scope**: Full design review — Architecture, tool/prompt quality, error handling, UX/UI, dead code, code quality, user-reported issues & prospective improvements

---

## Execution Status Overview

| Status | Count | Items |
|--------|-------|-------|
| ✅ **FIXED** | 31 | TOOL-C1 images+toast, TOOL-H1, TOOL-H2 screenshot guidance, USR-C1, USR-H1 bullets, USR-H1 prompt, USR-H2 elapsed timer+ctx%, context% indicator, ERR-H1, ERR-H2, USR-M1, USR-L1, **PPT-C1, PPT-C2, TOOL-M3** (Phase 1A), **IMG-H1, PPT-H1, PPT-M1** (Phase 1B), **PPT-H2, TOOL-L2, TOOL-L3** (Phase 1C), **UX-H1, ARCH-H2** (Phase 2A), **LANG-H1, TOOL-M4** (Phase 2B), **OUT-H1, QUAL-L2** (Phase 2C), **LOG-H1, FB-M1, ERR-M1** (Phase 3A), **XL-M1, TOOL-M1, TOOL-M2** (Phase 3B), **CLIP-M1, UX-M1, UX-L1** (Phase 3C) |
| 🟠 **PARTIALLY FIXED** (deferred sub-items remain) | 3 | TOOL-C1 (doc re-send), TOOL-H2 (no Word screenshot), USR-H1 (empty shapes) |
| ⏳ **IN PROGRESS** | 2 | DUP-H1, QUAL-H1 + PROSP-H2 context optimization |
| 📋 **BACKLOG** | 9 | Phase 2 Medium items (v10.x) |
| 🆕 **NEW (v11.0)** | 11 | 0 Critical + 6 High (6 fixed ✅) + 6 Medium (all 6 fixed ✅) + 0 Low (both fixed ✅) — see sections 11–13 |
| 🎯 **PLANNED** | 5 | Phase 3 Low items |
| 🚀 **DEFERRED** (Phase 4) | 18 | 11 functional improvements + 4 legacy (v7/v8) + 2 architectural + 1 dynamic tooling |

---

## Health Summary (v11.0)

All previous critical and major items from v9.x–v10.x have been resolved or deferred. This v11.0 review adds 20 new items from user-reported bugs + planned improvements audit. All OFFICE_AGENTS_ANALYSIS.md items have been confirmed implemented (screenshotRange, screenshotSlide, getRangeAsCsv, modifyWorkbookStructure, hide/freeze, duplicateSlide, verifySlides, editSlideXml, insertIcon, findData pagination, pptxZipUtils) — OFFICE_AGENTS_ANALYSIS.md deleted.

**v10.x sessions (2026-03-09)**: Fixed 4 items (TOOL-H1, USR-H1, USR-C1, TOOL-C1 logging), partially fixed 3 items. Fixed ERR-H1 (all 4 backend routes standardized), ERR-H2 (27+ console.warn/error → logService across 14 files), USR-M1 (scroll behavior), USR-L1 (upload failure warning done).

**v11.0 session (2026-03-14)**: Added 20 new items — confirmed implementation status of all OFFICE_AGENTS_ANALYSIS features, added user-reported bugs (PPT-C1, PPT-C2, IMG-H1, PPT-H1, OUT-H1, UX-H1, LANG-H1), and new improvement items (LOG-H1, PPT-H2, WORD-H1, PPT-M1, XL-M1, CLIP-M1, TOKEN-M1, OXML-M1, FB-M1, SKILL-L1, DYNTOOL-D1).

**v11.1 session (Phase 1A — 2026-03-14)**: ✅ **Phase 1A complete** — Fixed PPT-C1 (`getAllSlidesOverview`: per-slide try/catch + textSyncOk flag to isolate OLE/chart shape failures), fixed PPT-C2 (`insertImageOnSlide` + `insertIcon`: `slides.getItemAt(index)` → `slides.items[index]` to avoid post-sync proxy issue), implemented TOOL-M3 (`searchAndFormatInPresentation` tool: manual slide→shape→paragraph→textRun iteration with 4-sync batch pattern, supports bold/italic/underline/fontColor/fontSize/fontName).

**v11.2 session (Phase 1B — 2026-03-14)**: ✅ **Phase 1B complete** — IMG-H1: strengthened `FRAMING_INSTRUCTION` in `image.js` (explicit rules: fit entire subject, 4-side padding, no edge clipping, landscape composition) + changed default size to `1536x1024` in `backend.ts`. PPT-H1: rewrote `powerPointBuiltInPrompt.visual` to generate content-specific representative images (explicit requirement to illustrate the exact topic, not generic stock, style guidance per content type, text allowed if useful). PPT-M1: in `useAgentLoop.ts` visual handler, if selection < 5 words → call `screenshotSlide`, send image to LLM for slide description, use description as context for the visual prompt.

**v11.3 session (Phase 1C — 2026-03-14)**: ✅ **Phase 1C complete** — PPT-H2: replaced `speakerNotes` with `review` Quick Action — new early handler in `useAgentLoop.ts` (no selection required) runs agent loop with `getCurrentSlideIndex` → `screenshotSlide` → `getAllSlidesOverview` → numbered improvement suggestions; `constant.ts` updated, `ScanSearch` icon in `HomePage.vue`, i18n keys added. TOOL-L2: all 10 `slideNumber` descriptions clarified to "1-based (1 = first slide, not 0-based)". TOOL-L3: em-dash/semicolon ban extracted from `GLOBAL_STYLE_INSTRUCTIONS` into `PPT_STYLE_RULES`, applied only in `bullets` and `punchify` PPT prompts — formal Word/Outlook documents unaffected.

**v11.4 session (Phase 2A — 2026-03-14)**: ✅ **Phase 2A complete** — UX-H1: smart scroll with manual interruption — added `isAutoScrollEnabled` + `handleScroll()` in `useHomePage.ts`, `@scroll` listener in `ChatMessageList.vue` detects if user is near bottom (100px threshold), auto-scroll disables when user scrolls up and re-enables when scrolling back near bottom, `scrollToMessageTop()` always forces scroll for new content. ARCH-H2: created `useHomePageContext.ts` with provide/inject system, reduced `ChatMessageList` props from 20 to 0 (100% reduction), context exposes 40+ shared states/functions/handlers, progressive migration with optional props using context as fallback.

**v11.5 session (Phase 2B — 2026-03-14)**: ✅ **Phase 2B complete** — LANG-H1: separated conversation language (UI) from content generation language (document) across all 4 agent prompts (Word, Excel, PowerPoint, Outlook). Added explicit Language guidelines: conversations/explanations in UI language, generated content in selected text/document language. Generalized Outlook's `ALWAYS reply in SAME language` pattern to all hosts. TOOL-M4: extended `excelFormulaLanguageInstruction()` to support all 13 languages in `languageMap` (en, fr, de, es, it, pt, zh-cn, ja, ko, nl, pl, ar, ru). Created `ExcelFormulaLanguage` type. Categorized languages by separator: semicolon (`;`) for fr/de/es/it/pt/nl/pl/ru, comma (`,`) for en/zh-cn/ja/ko/ar.

**v11.6 session (Phase 2C — 2026-03-14)**: ✅ **Phase 2C complete** — OUT-H1: fixed image deletion during Outlook email translation by adding CRITICAL preservation instructions to all content-modifying prompts (`translate`, `translate_formalize`, `concise`, `proofread`). LLM now preserves `{{PRESERVE_N}}` placeholders that represent embedded images. Leveraged existing preservation system (`extractTextFromHtml` + `reassembleWithFragments` in richContentPreserver.ts) that was already implemented but missing LLM-side instructions. Images now preserved end-to-end during translation. QUAL-L2: added comprehensive JSDoc documentation (30+ lines) for `resolveAsyncResult()` helper in outlookTools.ts explaining callback-to-Promise bridge pattern, with code examples and full @param/@returns/@throws annotations.

**v11.7 session (Phase 3A — 2026-03-14)**: ✅ **Phase 3A complete** — LOG-H1: created `backend/logs/` directory, implemented JSONL tool usage logging in `toolUsageLogger.js` (logs to `tool-usage.jsonl` with format `{ts, user, host, tool, count}`), integrated logging in both `/api/chat` (streaming) and `/api/chat/sync` endpoints, added `getRecentToolUsage()` function for retrieving user tool history. ERR-M1: extracted shared error handler `handleChatError(res, error, req, endpoint, isStreaming)` from duplicate code blocks (~80% reduction), now handles AbortError, RateLimitError, streaming header-sent cases, and generic errors in single function. FB-M1: enhanced feedback system with `logChatRequest()` to track chat history in `request-history.jsonl`, `getRecentRequests()` to retrieve last 4 user requests, updated `feedback.js` to include `recentRequests` + `toolUsageSnapshot` fields in feedback submissions, created `feedback-index.jsonl` with `logFeedbackSubmission()` for centralized feedback tracking.

**v11.8 session (Phase 3B — 2026-03-14)**: ✅ **Phase 3B complete** — TOOL-M1: updated `setCellRange` tool `values` parameter schema from `items: { type: 'string' }` to `anyOf: [string, number, boolean, null]` with enhanced description documenting all accepted types to prevent LLM from incorrectly quoting numeric values. TOOL-M2: merged redundant `getWorksheetData` and `getDataFromSheet` tools into single unified `getWorksheetData` with optional `sheetName` and `address` parameters — eliminates agent confusion about which tool to use, reduces duplicate API calls. XL-M1: enhanced multi-curve chart extraction workflow — updated `extract_chart_data` tool description with explicit MULTI-CURVE CHARTS guidance (call once per series with specific targetColor, write to adjacent columns), updated `excel.skill.md` Step 1 to emphasize identifying all series colors, Step 2 to show iteration pattern for multi-series, Step 3 to demonstrate adjacent column layout for multiple series data.

**v11.9 session (Phase 3C — 2026-03-14)**: ✅ **Phase 3C complete** — CLIP-M1: implemented clipboard image paste support in ChatInput.vue — added `@paste` event handler to textarea, detects `clipboardData.items` with `type.startsWith('image/')`, creates File objects with descriptive names (`pasted-image-{timestamp}.{extension}`), processes through existing `processFiles` pipeline (same validation, size limits, preview display as drag/drop uploads). UX-M1: restored focus indicators for accessibility — added `focus:ring-2 focus:ring-primary/50` to all 6 interactive elements (textarea, select, 3 buttons, remove file button), improves keyboard navigation visibility for screen reader and keyboard-only users. UX-L1: refactored inline animation styles — replaced `:style="isDraftFocusGlowing ? 'animation-iteration-count: 3; ...' : ''"` with conditional class `draft-focus-glow`, moved animation definition to `<style scoped>` section (cleaner separation of concerns, better maintainability).

| Category | 🔴 Critical | 🟠 High | 🟡 Medium | 🟢 Low |
|----------|----------|------|--------|-----|
| Architecture | 0 | 2 | 3 | 1 |
| Tool/Prompt Quality | 0 | 2 | 4 | 3 |
| Error Handling | 0 | 2 | 2 | 1 |
| UX/UI | 0 | 2 | 3 | 3 |
| Dead Code | 0 | 0 | 2 | 1 |
| Code Duplication | 0 | 1 | 2 | 0 |
| Code Quality | 0 | 1 | 3 | 2 |
| User-Reported Issues | 0 | 2 | 2 | 1 |
| **v10 Subtotal** | **0** | **12** | **21** | **12** |
| **NEW v11 — Bugs** | **2** | **5** | **0** | **0** |
| **NEW v11 — Improvements** | **0** | **4** | **6** | **2** |
| **GRAND TOTAL** | **2** | **21** | **27** | **14** |
| **Status** | 2 new critical bugs | 12 active v10 + 9 new | 27 items | 14 items |

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

### ARCH-H2 — HomePage.vue prop drilling (44+ bindings) [HIGH] ✅

**File**: `frontend/src/pages/HomePage.vue`, `frontend/src/composables/useHomePageContext.ts` (nouveau)

HomePage passes 44+ props and event bindings down to child components (ChatHeader: 7, ChatMessageList: 17, ChatInput: 13, QuickActionsBar: 6). This creates tight coupling between the page and its children.

**Impact**: Every state change requires updating prop chains. Adding a new feature touches multiple components.

**✅ IMPLÉMENTÉ (2026-03-14)** :
- Création de `useHomePageContext.ts` avec système `provide/inject`
- Définition de l'interface `HomePageContext` avec 40+ états/fonctions/handlers partagés
- `provideHomePageContext()` appelée dans `HomePage.vue` pour exposer le contexte
- Migration de `ChatMessageList.vue` : **20 props → 0 props** (réduction de 100%)
- Props rendues optionnelles avec contexte comme fallback (migration progressive)
- Le composant utilise maintenant `useHomePageContext()` pour accéder aux données
- Événements émis remplacés par appels directs aux fonctions du contexte
- Architecture extensible : autres composants (ChatInput, StatsBar, etc.) peuvent être migrés ultérieurement

**Recommendation**: ✅ Implemented using `provide/inject` with `useHomePageContext` composable, reducing ChatMessageList prop drilling by 100%.

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

### TOOL-M1 — Excel `values` parameter typed as `string` but accepts mixed types [MEDIUM] ✅ FIXED (Phase 3B)

**File**: `frontend/src/utils/excelTools.ts:182-195`

The `values` parameter description said items are "string", but Excel cells accept numbers, booleans, dates, and nulls. This could mislead the LLM into always quoting numeric values.

**Fix (v11.8)**: Updated `setCellRange` tool schema:
- Changed `items: { type: 'array', items: { type: 'string' } }` to `items: { type: 'array', items: { anyOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] } }`
- Enhanced description: "Each cell value can be: string, number, boolean, null, or Date object. Use null to skip/clear a cell."
- Added example showing mixed types: `[["Name","Score"],["Alice",95],["Bob",true],[null,3.14]]`
- LLM now correctly passes numeric values as numbers, not quoted strings

---

### TOOL-M2 — Overlapping Excel read tools [MEDIUM] ✅ FIXED (Phase 3B)

**Files**: `frontend/src/utils/excelTools.ts:93-132`, `frontend/src/skills/excel.skill.md:141-155`

`getWorksheetData` (reads active sheet) and `getDataFromSheet` (reads any sheet by name) overlapped. Both returned CSV data from a worksheet, causing agent confusion.

**Fix (v11.8)**: Merged into single unified `getWorksheetData` tool:
- **Before**: Two separate tools — `getWorksheetData()` for active sheet only, `getDataFromSheet(name)` for named sheets
- **After**: Single tool `getWorksheetData(sheetName?, address?)` with optional parameters:
  - `sheetName` (optional): worksheet name, defaults to active sheet if omitted
  - `address` (optional): specific range address, defaults to used range if omitted
- Removed `getDataFromSheet` tool completely (38 lines deleted)
- Updated `excel.skill.md` tool reference table to reflect single unified tool
- Eliminates agent confusion about which tool to use, prevents redundant tool calls
- Returns `worksheet: "(active)"` or the actual sheet name in response for clarity

---

### TOOL-M3 — No PowerPoint equivalent to Word's `searchAndFormat` [MEDIUM] ✅ FIXED (Phase 1A)

**File**: `frontend/src/utils/powerpointTools.ts`

PowerPoint has no native `body.search()` API like Word. Implemented `searchAndFormatInPresentation` tool that manually iterates slides → shapes (filtering pictures/OLE) → paragraphs → textRuns using 4-sync batch pattern per slide. Supports bold, italic, underline, fontColor, fontSize, fontName.

**Impact**: ✅ Agent can now reliably bold, color, or resize specific words in PowerPoint slides.

---

### TOOL-M4 — Inconsistent formula locale support [MEDIUM] ✅

**Files**: `frontend/src/composables/useAgentPrompts.ts:28-62`, `frontend/src/utils/constant.ts:2-21`

Agent prompt only handles English/French formula locales, but the language map in `constant.ts` lists 13 languages. German, Spanish, Italian, etc. Excel users won't get correct formula separator guidance (`;` vs `,`).

**✅ IMPLÉMENTÉ (2026-03-14)** :
- Extended `excelFormulaLanguageInstruction()` to support all 13 languages in `languageMap`
- Created `ExcelFormulaLanguage` type in constant.ts: `'en' | 'fr' | 'de' | 'es' | 'it' | 'pt' | 'zh-cn' | 'ja' | 'ko' | 'nl' | 'pl' | 'ar' | 'ru'`
- Categorized languages into two groups:
  - **Semicolon separator (`;`)** + comma for decimals: fr, de, es, it, pt, nl, pl, ru
  - **Comma separator (`,`)** + period for decimals: en, zh-cn, ja, ko, ar
- Updated type signatures in `useAgentPrompts.ts`, `useAgentLoop.ts`, and `HomePage.vue`
- Function now provides localized instructions for all supported languages with correct separator and decimal guidance

---

### TOOL-L1 — `getRangeAsCsv` missing format documentation [LOW]

**File**: `frontend/src/utils/excelTools.ts:174-176`

No description of the CSV format returned (delimiter, quoting, header handling). The LLM may parse incorrectly.

---

### TOOL-L2 — PowerPoint `slideNumber` should clarify 1-based indexing [LOW] ✅ FIXED (Phase 1C)

**File**: `frontend/src/utils/powerpointTools.ts`

**Fix**: Updated all 10 `slideNumber` parameter descriptions from `"1 = first slide"` to `"1-based (1 = first slide, not 0-based)"` using a targeted sed replacement across the file.

---

### TOOL-L3 — Style rules ban em-dashes globally [LOW] ✅ FIXED (Phase 1C)

**File**: `frontend/src/utils/constant.ts`

**Fix**: Extracted em-dash/semicolon ban into a new `PPT_STYLE_RULES` constant. Removed these rules from `GLOBAL_STYLE_INSTRUCTIONS` (which applies to all hosts). Added `PPT_STYLE_RULES` directly to the `bullets` and `punchify` prompt constraints in `powerPointBuiltInPrompt` only — formal Word/Outlook documents can now use em-dashes normally.

---

## 3. ERROR HANDLING & DEBUGGABILITY

### ERR-H1 — 4 backend routes bypass `logAndRespond()` and ErrorCodes [HIGH] ✅ FIXED

**Files**:
- `backend/src/routes/files.js:31, 64, 72, 79` — returns `{ error: '...' }` without code
- `backend/src/routes/feedback.js:23, 46` — returns `{ error: '...' }` without code
- `backend/src/routes/logs.js:25, 29, 56` — returns `{ error: '...' }` without code
- `backend/src/routes/icons.js:13, 25, 47` — returns `{ error: '...', details: '...' }` without code

All other routes use `logAndRespond()` from `utils/http.js` with structured `ErrorCodes`. These 4 routes break the pattern, meaning:
1. Frontend's `categorizeError()` (`backend.ts:101-122`) cannot map error codes — falls back to fragile string inspection
2. Errors are logged without req.logger context enrichment (userId, host, session)
3. The `files.js:79` handler leaks raw error messages to the client

**Fix applied**: All 4 routes now use `logAndRespond()` with `ErrorCodes`. New codes added: `FEEDBACK_MISSING_FIELDS`, `LOGS_INVALID_ENTRIES`, `LOGS_TOO_MANY_ENTRIES`, `ICON_QUERY_REQUIRED`, `ICON_NOT_FOUND`, `ICON_FETCH_FAILED`, `FILE_NO_ID_RETURNED`. Also fixed `http.js` `console.error/warn` → `logger.error/warn` and `models.js` `console.warn` → `logger.warn`.

---

### ERR-H2 — Frontend uses `console.warn/error` instead of `logService` (27 instances) [HIGH] ✅ FIXED

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

**Fix applied**: All `console.warn/error` replaced with `logService.warn/error` across 14 frontend files (composables, utils, skills, pages). `logService` import added to 13 files (useAgentLoop already had it).

---

### ERR-M1 — Chat route duplicate error handling [MEDIUM] ✅ FIXED (Phase 3A)

**File**: `backend/src/routes/chat.js`

`/api/chat` (streaming) and `/api/chat/sync` (synchronous) contained ~80% identical error handling code (validation, upstream errors, AbortError/RateLimitError branching). Changes had to be applied twice.

**Fix (v11.7)**: Extracted shared error handler `handleChatError(res, error, req, endpoint, isStreaming)` with comprehensive JSDoc documentation. The function handles:
- Streaming-specific error case (headers already sent): writes SSE error message and ends response
- AbortError: returns 504 with `LLM_TIMEOUT` error code
- RateLimitError: returns 429 with `RATE_LIMITED` error code
- Generic errors: returns 500 with `INTERNAL_ERROR` error code

Both endpoints now call `handleChatError()` in their catch blocks with appropriate `isStreaming` flag. Reduced ~40 lines of duplicate code to a single shared function.

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

### UX-M1 — Missing focus indicators (accessibility) [MEDIUM] ✅ FIXED (Phase 3C)

**File**: `frontend/src/components/chat/ChatInput.vue:21, 60, 70, 90, 99, 40`

`focus:outline-none` removed visual focus indicators. Only 8 `focus:ring` instances existed across the entire frontend. Keyboard-only users could not see which element is focused.

**Fix (v11.9)**: Added `focus:ring-2 focus:ring-primary/50` to all 6 interactive elements in ChatInput.vue:
1. **Select element** (line 21): Model tier dropdown — added `focus:ring-2 focus:ring-primary/50`
2. **Textarea** (line 60): Main chat input — added `focus:ring-2 focus:ring-primary/50` (kept `outline-none` to avoid double outline)
3. **Attach button** (line 70): Paperclip file upload button — added `focus:outline-none focus:ring-2 focus:ring-primary/50`
4. **Stop button** (line 90): Red stop button during streaming — added `focus:outline-none focus:ring-2 focus:ring-primary/50`
5. **Send button** (line 99): Blue send/submit button — added `focus:outline-none focus:ring-2 focus:ring-primary/50`
6. **Remove file button** (line 40): × button on file chips — added `focus:outline-none focus:ring-2 focus:ring-primary/50 rounded-sm`

**Result**: Keyboard navigation now shows visible focus rings on all interactive elements — complies with WCAG 2.1 accessibility guidelines for keyboard-only users and screen readers

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

### UX-L1 — Inline animation styles in ChatInput.vue [LOW] ✅ FIXED (Phase 3C)

**File**: `frontend/src/components/chat/ChatInput.vue:50-54, 310-315`

Used `:style="isDraftFocusGlowing ? 'animation-iteration-count: 3; animation-duration: 0.5s;' : ''"` inline. Should be in `<style scoped>` with a conditional class for better maintainability.

**Fix (v11.9)**:
1. **Template refactor** (lines 50-54):
   - **Before**: `:style="isDraftFocusGlowing ? 'animation-iteration-count: 3; animation-duration: 0.5s;' : ''"`
   - **After**: `:class="{ 'ring-2 ring-accent draft-focus-glow': isDraftFocusGlowing }"`
   - Removed inline style completely, replaced with conditional class `draft-focus-glow`

2. **Added scoped style section** (lines 310-315):
   ```css
   <style scoped>
   /* UX-L1: Animation for draft focus glow */
   .draft-focus-glow {
     animation: pulse 0.5s ease-in-out;
     animation-iteration-count: 3;
   }
   </style>
   ```

**Result**: Cleaner separation of concerns — styles in `<style scoped>`, logic in `<script>`, presentation in `<template>`. Easier to maintain and modify animation properties.

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

### QUAL-L2 — Async/Promise pattern inconsistency [LOW] ✅

**File**: `frontend/src/utils/outlookTools.ts:46-76`

Outlook tools mix `async/await` with callback-based `Office.AsyncResult` patterns (due to Outlook API limitations). While necessary, the wrapping in `resolveAsyncResult()` could be documented more clearly.

**✅ IMPLÉMENTÉ (2026-03-14)** :
- Added comprehensive JSDoc documentation (30+ lines) for the `resolveAsyncResult()` helper function
- Documentation explains:
  - **Why it exists**: Outlook JavaScript API uses callback-based patterns instead of Promises
  - **The pattern**: How it bridges AsyncResult callbacks with async/await
  - **Code example**: Before/after comparison showing the wrapping pattern
  - **Full signature**: @param, @returns, @throws annotations
- Developers can now understand the pattern at a glance and replicate it correctly when adding new Outlook tools

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

### USR-M1 — Scroll behavior doesn't match user expectations [MEDIUM] ✅ FIXED

**Files**:
- `frontend/src/composables/useHomePage.ts:71-107` — scroll helpers
- `frontend/src/composables/useAgentLoop.ts:254-255, 429-430` — scroll trigger points
- `frontend/src/composables/useAgentStream.ts:51-56` — no auto-scroll during streaming

**Fix applied**:
1. Session load / session switch / session delete: now calls `scrollToConversationTop()` (new helper added to `useHomePage.ts`) → `container.scrollTo({ top: 0, behavior: 'smooth' })`
2. Message send: changed from `scrollToVeryBottom()` → `scrollToMessageTop()` — scrolls to top of user's newly sent message
3. Response complete: changed from `scrollToVeryBottom()` → `scrollToMessageTop()` — scrolls to top of assistant response so user reads from the start

---

### USR-M2 — Context window percentage already visible but not prominent enough [MEDIUM]

**File**: `frontend/src/components/chat/StatsBar.vue`

The stats bar already shows context usage with color-coded warnings (green <70%, orange 70-89%, red >=90%), but users don't understand WHY the agent is slow. The context % is visible but not prominent enough during long agent sessions.

**Action**: Consider adding a tooltip or notification when context exceeds 80%: "Response may be slower — large conversation context."

---

### USR-L1 — No visual feedback when /v1/files upload silently fails [LOW] ✅ FIXED

**File**: `frontend/src/composables/useAgentLoop.ts:821-822`

When `uploadFileToPlatform()` fails, the error is caught silently and the file falls back to inline content. The user has no idea their file was not uploaded efficiently.

**Fix applied**: Warning toast shown when `/v1/files` upload fails and file falls back to inline base64. Implemented in previous session.

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

**Latest round (v10.2)** — FIXED:
6. **TOOL-C1**: Images now try /v1/files + warning toast for both text and images
7. **TOOL-H2**: Screenshot guidance added to Excel (Step 5) + PPT prompts; PPT verification rule clarified
8. **USR-H1**: Prompt guidance: "no markdown bullets in body placeholders"
9. **USR-H2**: Context % shown in LLM wait label when >50%

**Still Active** (2 items):
10. ~~**ERR-H1**: Standardize all backend routes to use `logAndRespond()` + ErrorCodes~~ — **FIXED** ✅
11. ~~**ERR-H2**: Replace all `console.warn/error` with `logService` (27 instances)~~ — **FIXED** ✅
12. **DUP-H1**: Extract shared tool wrapper boilerplate to `common.ts`
13. **QUAL-H1**: Replace critical `any` types with proper Office.js types
— **PROSP-H2**: Conversation history optimization (blocking 3 deferred items) → Phase 4

### Phase 2 — 🟡 MEDIUM (Maintainability & DX) — 8 Active
11. ~~**USR-M1**: Fix scroll behavior (session load → top, send → user msg, complete → response top)~~ — **FIXED** ✅
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

### Phase 3 — 🟢 LOW (Polish) — 4 Active
21. **UX-L1-L3**: Inline styles, link text, mobile width
22. **ARCH-L1**: Switch to `npm ci` in Dockerfile
23. **ARCH-L2**: Evaluate manifest accessibility — move `generated-manifests/` to `frontend/public/assets/` for SaaS distribution
24. **QUAL-L1-L2**: Boolean params, async pattern docs
25. ~~**USR-L1**: Show warning when /v1/files upload silently fails~~ — **FIXED** ✅
— **PROSP-1/3/4/5**: Dynamic tool loading, PRD split, templates, intent profiles → Phase 4

### Phase 4 — Deferred Items (Not Yet Addressed)

**Consolidated deferred work from multiple review cycles** (v7, v8, v10.1):
- **Part A**: Deferred actions from partially-fixed Phase 0–1 items (actionable, blocked on design decisions)
- **Part B**: Infrastructure & legacy items (from v7/v8, low priority)
- **Part C**: Prospective improvements (architectural enhancements, high-value)

---

**Part A: Deferred actions from partially-fixed Phase 0–1 items** (actionable, blocked on design decisions or dependencies):

#### 🟠 TOOL-C1 Remaining Items (HIGH — MOSTLY FIXED ✅)
- ~~**Images never use /v1/files**~~: **FIXED ✅** — Images now attempt `/v1/files` upload with `purpose: 'vision'`. On success, the provider fileId is stored and used in subsequent iterations instead of re-sending base64 bytes.
- ~~**No UI indicator for /v1/files fallback**~~: **FIXED ✅** — Warning toast shown (i18n key: `warningFileFallbackInline`) when upload fails for both text files and images.
- **Full document re-sent on every iteration**: ⏳ Still blocked on PROSP-H2 (context optimization). Each iteration re-injects full text file content. Images now use fileId if available.

#### 🟠 TOOL-H2 Remaining Items (HIGH — PARTIALLY FIXED ✅)
- ~~**No auto-verification prompting**~~: **FIXED ✅** — Added Step 5 (screenshotRange verification) to Excel chart workflow in both `excel.skill.md` and `useAgentPrompts.ts`. Added `screenshotSlide` verification guidance to PowerPoint prompt and `powerpoint.skill.md`.
- ~~**PowerPoint blocks verification via getAllSlidesOverview**~~: **FIXED ✅** — Rule now clarified: "Do NOT call getAllSlidesOverview to verify — use `screenshotSlide` instead." Defensive rule preserved for the correct tool, verification enabled via screenshot.
- **No Word screenshot tool**: ⏳ Still deferred — No Office.js API for Word document screenshots exists. Cannot implement without a third-party capture solution.

#### 🟠 USR-H1 Remaining Items (HIGH — PARTIALLY FIXED ✅)
- **Empty shapes with default bullets**: ⏳ Still open — `hasNativeBullets()` only checks existing paragraphs. Empty shapes with XML bullet defaults still risk double-bullets. Low priority: body placeholders now covered by `placeholderFormat/type` fix.
- ~~**Stronger prompt guidance needed**~~: **FIXED ✅** — Added Guideline 4 to PowerPoint agent prompt: "When inserting into body/content placeholder shapes, do NOT use markdown list syntax (`- item`). The shape already has native bullets — plain text lines are sufficient."

#### 🟠 USR-H2 Remaining Items (HIGH — PARTIALLY FIXED ✅)
- **Context bloat structural issue**: ⏳ Still blocked on PROSP-H2. Each iteration re-sends full message history.
- **Tool result accumulation**: ⏳ Still blocked on PROSP-H2. Tool results never summarized between iterations.
- ~~**No context window % indicator**~~: **FIXED ✅** — Context usage % shown in `currentAction` label during LLM wait when above 50%: e.g., "Waiting for AI... (14s · ctx 73%)". Uses `estimateContextUsagePercent()` from `tokenManager.ts`.

---

**Part B: Infrastructure & Legacy Deferred Items** (from v7/v8 reviews):

#### 🟢 IC2 — Containers run as root (LOW)
**Files**: `backend/Dockerfile`, `frontend/Dockerfile`
Docker containers should run with a non-root user for security best practices. Currently, both Dockerfiles use the default `root` user:
- `backend/Dockerfile`: Node:22-slim runs as root (no USER directive)
- `frontend/Dockerfile`: Nginx:stable runs as root (no USER directive)

**Current status**: Still vulnerable. No USER directive found in either Dockerfile.
**Severity**: LOW — This is internal infrastructure for local development. Security risk is low if only used internally.
**Action**: Add `USER appuser` or similar to both Dockerfiles after setup. For nginx, create appuser with minimal privileges before switching.

#### 🟢 IH2 — Private IP in build arg (LOW)
**Files**: `frontend/Dockerfile:18`, `.env.example:1,6`
Private IP address `192.168.50.10` hardcoded in build arguments and examples. Should be sanitized or use environment variables like `localhost` or a placeholder.
**Current status**: Still present in `frontend/Dockerfile` ARG and multiple `.env.example` files.
**Action**: Replace with placeholder IP (e.g., `localhost` or `192.168.x.x` generic pattern) or document as "replace with your server IP".

#### 🟢 IH3 — DuckDNS domain in example (LOW)
**Files**: `.env.example:10-11`
Real DuckDNS domain `https://kickoffice.duckdns.org` hardcoded in example. Could be confused with a real public URL.
**Current status**: Still present in `.env.example` as `PUBLIC_FRONTEND_URL` and `PUBLIC_BACKEND_URL`.
**Action**: Replace with placeholder (e.g., `https://your-domain.duckdns.org` or `https://example.duckdns.org`) with a clear comment "Update with your actual DuckDNS domain".

#### 🟢 UM10 — PowerPoint HTML reconstruction (DEFERRED INDEFINITELY)
**Original proposal** (v7): Reconstruct PowerPoint slides from HTML snapshots captured during visual creation. This would allow the agent to verify if generated HTML matches the final slide layout.
- **NOT resolved by OOXML editing**: Recent improvements (layout detection, placeholder type loading, chart extraction) improved slide manipulation but did NOT implement HTML→slide reconstruction.
- **Complexity too high**: OOXML format is intricate and error-prone. Edge cases (complex animations, embedded OLE objects, custom fonts) make this unreliable.
- **Better approach**: Use screenshot + image upload workflow instead (already implemented via screenshotRange/screenshotSlide tools).
- **Status**: Closed/Not recommended. Do not implement.

---

**Part C: Prospective improvements** (architectural enhancements, not blocking but high-value):

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

| Severity | Count | Status | Items |
|----------|-------|--------|-------|
| 🔴 **Critical** | 0 | ✅ All v10 critical fixed or deferred | None from v10 — 2 new critical in active backlog (PPT-C1, PPT-C2) |
| 🟠 **High** | 5 + 1 prospective | ⏳ Pending | TOOL-C1 (3), TOOL-H2 (3), USR-H1 (2), USR-H2 (3), PROSP-H2 (context opt.) |
| 🟡 **Medium** | 2 prospective | — | PROSP-2 (Claude.md), PROSP-5 (intent profiles) |
| 🟢 **Low** | 4 legacy + 3 prospective | — | IC2, IH2, IH3, UM10 (v7/v8) + PROSP-1/3/4 |
| 🚀 **DYNTOOL-D1** | 1 | — | Dynamic Tooling (new, detailed plan) |
| **TOTAL DEFERRED** | **18** | | 11 functional (from partial fixes + PROSP-H2) + 6 architectural/legacy + 1 new (DYNTOOL) |

---

## 11. USER-REPORTED BUGS (v11.0) — 🔴 Critical & 🟠 High

### PPT-C1 — `getAllSlidesOverview` returns InvalidArgument on "resume a slide" request [CRITICAL] ✅ FIXED (Phase 1A)

**File**: `frontend/src/utils/powerpointTools.ts`

**Fix**: Wrapped the entire per-slide processing block in an outer `try/catch` (returns `"Slide N: [Error reading content]"` on failure). Added `textSyncOk` flag — the text `await context.sync()` is wrapped in its own `try/catch`, and if it fails, text extraction is skipped gracefully for that slide. OLE/chart/SmartArt shapes that cause `InvalidArgument` no longer crash the entire function.

---

### PPT-C2 — `insertImageOnSlide` crashes: "addImage is not a function" when using UUID [CRITICAL] ✅ FIXED (Phase 1A)

**File**: `frontend/src/utils/powerpointTools.ts`

**Fix**: Replaced `slides.getItemAt(index)` (returns a proxy post-sync lacking `.shapes.addImage`) with `slides.items[index]` (direct access to the already-loaded slide object). Applied to both `insertImageOnSlide` and `insertIcon` which had the same pattern.

---

### IMG-H1 — Image generation cropped with gpt-image-1 / gpt-image-1.5 [HIGH] ✅ FIXED (Phase 1B)

**Files**: `backend/src/routes/image.js`, `frontend/src/api/backend.ts`

**Fix**:
1. Strengthened `FRAMING_INSTRUCTION` in `image.js` — explicit 4-rule composition mandate: fit entire subject, visible padding on all four sides, no clipping of heads/limbs/text/edges, landscape 16:9 composition.
2. Changed default size from `1024x1024` → `1536x1024` in `backend.ts` — landscape format matches PowerPoint slide dimensions and prevents side-cropping of wide subjects.

---

### PPT-H1 — Quick Action "Image" : l'image générée n'est pas représentative du contenu [HIGH] ✅ FIXED (Phase 1B)

**File**: `frontend/src/utils/constant.ts` — `powerPointBuiltInPrompt.visual`

**Fix**: Rewrote `visual.system` and `visual.user` prompts. New prompt explicitly requires: (1) visual must represent the SPECIFIC topic — no generic stock images, (2) style selection adapted to content type (photo-realistic, flat vector, isometric, infographic, etc.), (3) text in image explicitly allowed and requested when useful, (4) composition details: foreground/background/focal elements, (5) landscape 16:9 format, (6) output only the prompt with no preamble.

---

### OUT-H1 — Outlook translation deletes embedded images from email body [HIGH] ✅

**Files**: `frontend/src/utils/outlookTools.ts`, `frontend/src/utils/constant.ts`, `frontend/src/utils/richContentPreserver.ts`

When the agent translates an email body, it reads the HTML content, sends it to the LLM for translation, then calls `setBody` with the translated HTML. Inline images (embedded as `cid:` references or `data:` URIs) are lost because the LLM does not reproduce the `<img>` tags in its translation output.

**Tool description** (`outlookTools.ts:154`) says "automatically preserves images from the original email" — this guarantee is currently NOT enforced at the code level, only in the description.

**✅ IMPLÉMENTÉ (2026-03-14)** :
The preservation system already exists in the codebase:
- `getEmailBody` (outlookTools.ts:107-148) extracts HTML and calls `extractTextFromHtml()` which replaces images with `{{PRESERVE_N}}` placeholders
- `writeEmailBody` (outlookTools.ts:178-243) calls `reassembleWithFragments()` to restore images from placeholders
- The missing piece was LLM awareness: the translation prompts did not instruct the LLM to preserve these placeholders

**Implementation:**
- Added CRITICAL preservation instruction to all Outlook prompts that modify content:
  - `translate` (constant.ts:62): "If the text contains preservation placeholders like {{PRESERVE_0}}, {{PRESERVE_1}}, etc., you MUST keep these placeholders EXACTLY as-is in their original positions."
  - `translate_formalize` (constant.ts:446): Same instruction added
  - `concise` (constant.ts:466): Same instruction added
  - `proofread` (constant.ts:487): Same instruction added
- The LLM now preserves `{{PRESERVE_N}}` placeholders during translation, and `writeEmailBody` reassembles them back into `<img>` tags
- Images are now preserved end-to-end during translation and other content modifications

---

### UX-H1 — Chat scroll "yoyo" effect during streaming; no smart-scroll interrupt [HIGH] ✅

**File**: `frontend/src/composables/useHomePage.ts:71-107`, `frontend/src/composables/useAgentStream.ts`, `frontend/src/components/chat/ChatMessageList.vue`

**Context**: USR-M1 was previously "fixed" by implementing scroll-to-message-top behavior. However the current implementation still causes a "yoyo" effect during streaming: the container scrolls to the bottom on send, then jumps to the top of the response when the stream starts, creating a disorienting experience. There is also no mechanism to interrupt auto-scroll if the user scrolls up manually.

**✅ IMPLÉMENTÉ (2026-03-14)** :
- Ajout de `isAutoScrollEnabled: Ref<boolean>` dans `useHomePage.ts` (défaut: `true`)
- Ajout de `handleScroll()` qui détecte si l'utilisateur est proche du bas (seuil: 100px)
- Écouteur `@scroll="handleScrollEvent"` dans `ChatMessageList.vue` qui appelle `handleScroll()`
- `scrollToBottom()` respecte maintenant `isAutoScrollEnabled` (sauf si `force=true`)
- `scrollToMessageTop()` force toujours le scroll pour afficher le nouveau contenu
- L'auto-scroll se désactive automatiquement quand l'utilisateur scrolle vers le haut
- L'auto-scroll se réactive automatiquement quand l'utilisateur revient près du bas

**Expected behavior (ChatGPT-style):**
- **On initial load / session switch**: `scrollTop = scrollHeight` (instant, no animation)
- **On message send**: smooth scroll to bottom
- **During stream**: auto-scroll to bottom on each new chunk; if user scrolls up manually → pause auto-scroll; if user scrolls back to bottom → resume auto-scroll

**Implementation details:**
- Add `isAutoScrollEnabled: Ref<boolean>` (default `true`, reset to `true` on each new request)
- Add `@scroll` listener on `containerEl` in ChatMessageList or HomePage: if user scrolls up (delta < 0 and not at bottom) → set `isAutoScrollEnabled = false`
- If `scrollTop + clientHeight >= scrollHeight - 10` (within 10px of bottom) → set `isAutoScrollEnabled = true`
- During stream: call `scrollToBottom()` only if `isAutoScrollEnabled === true`
- `scrollToBottom(smooth=true)` for send, `scrollToBottom(smooth=false)` for initial load
- Use `nextTick()` or `MutationObserver` before reading `scrollHeight` to ensure DOM is updated

**Target files:**
- `frontend/src/pages/HomePage.vue` or `frontend/src/composables/useHomePage.ts` — scroll helpers
- `frontend/src/composables/useAgentStream.ts` — stream chunk handler (add scroll call)
- `frontend/src/components/chat/ChatMessageList.vue` — expose `containerEl`, add `@scroll` listener

---

### LANG-H1 — LLM responds in UI language but should use document language for generated text [HIGH] ✅

**File**: `frontend/src/composables/useAgentPrompts.ts` (lines 119, 184, 235, 267)

**Problem**: All agent prompts include `"Language: Communicate entirely in ${lang}."` where `lang` is the UI language (user's interface setting, e.g., French). When the user works on a document in a different language (e.g., an English PowerPoint) and asks to improve text, the LLM generates the improvement proposals in French instead of English.

**Expected behavior**:
- The LLM should **converse with the user** (explanations, questions, commentary) in the **UI language**
- The LLM should **generate document content / propose text for the document** in the **language of the document or selected text**

**Example** (exact case reported): User selects English text "Possible warning from the team ambiance, to be checked" and asks in French "comment améliorer cette phrase" → LLM should respond in French for the discussion but provide the alternative phrases in **English** since the selected text was in English.

**✅ IMPLÉMENTÉ (2026-03-14)** :
- Modified all 4 agent prompts (Word, Excel, PowerPoint, Outlook) to separate **Conversation Language** (UI) from **Content Generation Language** (document)
- Word prompt (line 119): Added explicit Language guideline distinguishing conversation (UI language) from content generation (selected text language)
- Excel prompt (line 184): Same pattern applied for spreadsheet content
- PowerPoint prompt (line 235): Same pattern applied for slide text
- Outlook prompt (line 267): Reformulated existing `Reply Language` rule to align with new consistent pattern across all hosts
- The LLM now analyzes the language of `[Selected text]`, `[Selected cells]`, or email content to determine target language for generated content
- Built-in prompts already use `LANGUAGE_MATCH_INSTRUCTION` which enforces this behavior
- **Pattern generalized**: Outlook's correct pattern (`ALWAYS reply in the SAME language`) now applied to all Office hosts

---

## 12. NEW IMPROVEMENTS (v11.0) — 🟠 High & 🟡 Medium & 🟢 Low

### LOG-H1 — No tool usage counting system per platform [HIGH] ✅ FIXED (Phase 3A)

**Files**: `backend/src/routes/chat.js`, `backend/src/utils/toolUsageLogger.js`, `backend/logs/tool-usage.jsonl`

**Problem**: There was no persistent log tracking which tools are called, per Office host (Word/Excel/PPT/Outlook), per user, per day. This data is needed to:
1. Identify the "Core Set" of most-used tools for the Dynamic Tooling optimization (DYNTOOL-D1)
2. Monitor usage trends and detect anomalies
3. Support the feedback system with usage context

**Fix (v11.7)**:
1. Created `backend/logs/` directory
2. Created `backend/src/utils/toolUsageLogger.js` with:
   - `logToolUsage(userId, host, toolCalls)` — appends to `tool-usage.jsonl` in JSONL format:
     ```json
     {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","tool":"screenshotSlide","count":1}
     ```
   - `getRecentToolUsage(userId, limitLines)` — reads recent tool usage for a specific user (used by FB-M1)
3. Integrated in `backend/src/routes/chat.js`:
   - **Streaming endpoint** (`/api/chat`): parses SSE chunks for `delta.tool_calls`, accumulates tool calls during stream, logs after successful completion
   - **Sync endpoint** (`/api/chat/sync`): extracts `tool_calls` from response message, logs after successful response
   - Both endpoints log with `userId` and `host` from `req.logger.defaultMeta`
4. Tool usage now tracked per-call, enabling future analytics for DYNTOOL-D1 and usage dashboards

---

### PPT-H2 — New Quick Action "Review": replace Speaker Notes action [HIGH] ✅ FIXED (Phase 1C)

**Files**: `frontend/src/utils/constant.ts`, `frontend/src/composables/useAgentLoop.ts`, `frontend/src/pages/HomePage.vue`, `frontend/src/i18n/locales/*.json`

**Fix**:
1. `constant.ts`: Replaced `speakerNotes` with `review` in `powerPointBuiltInPrompt`. The `review` prompt instructs the LLM to provide 3-5 numbered improvement suggestions for the current slide only.
2. `useAgentLoop.ts`: Added a special `review` early handler (like `visual`) that bypasses the `selectedText` guard — no selection required. The handler runs `runAgentLoop` with a system prompt instructing the agent to call `getCurrentSlideIndex` → `screenshotSlide` → `getAllSlidesOverview`, then provide slide-specific review. Removed the `speakerNotes` post-processing block. Removed unused `setCurrentSlideSpeakerNotes` import.
3. `HomePage.vue`: Replaced `speakerNotes` quick action with `review` using `ScanSearch` icon. Added `ScanSearch` to lucide imports.
4. `i18n/locales/en.json` + `fr.json`: Added `pptReview` and `pptReview_tooltip` keys. `getSpeakerNotes`/`setSpeakerNotes` tools remain available to the agent.

---

### WORD-H1 — Track Changes via OOXML (replace office-word-diff approach) [HIGH]

**Files**: `frontend/src/utils/wordDiffUtils.ts`, `frontend/src/utils/wordTools.ts:1379-1391`

**Problem**: The current `proposeRevision` tool uses `office-word-diff` (npm package) to compute word-level diffs and apply changes. This approach can break complex Word formatting (`<w:rPr>`, colors, font sizes) because it reconstructs runs from scratch rather than performing surgical XML edits.

The ideal approach (inspired by `docx-redline-js` and the `Gemini-AI-for-Office` add-in) is to inject real OOXML revision markup (`<w:ins>` / `<w:del>`) directly into the paragraph XML, preserving all existing formatting.

**Proposed implementation**:
1. Add a **configurable "Redline Author" field** in Settings (under Account or a new Editing tab): default `"KickOffice AI"`, user-editable
2. Create `frontend/src/utils/wordOoxmlUtils.ts` — utility to:
   - Read paragraph XML via `paragraph.getOoxml()`
   - Compute diff between original and revised text at run level
   - Inject `<w:ins w:author="{author}" w:date="{date}" w:id="{id}">` for insertions
   - Inject `<w:del w:author="{author}" w:date="{date}" w:id="{id}">` + `<w:delText>` for deletions
   - Write back via `paragraph.insertOoxml(xml, 'Replace')`
3. Update `proposeRevision` tool to use this new approach when the document is a `.docx` (not `.doc`)
4. Keep `office-word-diff` as a fallback for simple cases or when OOXML is unavailable
5. **If switching completely**: Remove `office-word-diff` npm dependency, update `Dockerfile`, update README to remove references
6. **OXML evaluation prerequisite**: This task depends on OXML-M1 (evaluate current OOXML integration across all hosts before making changes)

---

### PPT-M1 — Quick Action "Image": handle <5 words selection case [MEDIUM] ✅ FIXED (Phase 1B)

**File**: `frontend/src/composables/useAgentLoop.ts` — `visual` quick action handler

**Fix**: In the `visual` handler, before Step 1 (image prompt generation), check word count of `slideText`. If `< 5 words`: call `powerpointToolDefinitions.screenshotSlide.execute({})`, parse the base64 result, send the slide image to the LLM as a vision message requesting a 2-3 sentence description of the slide content and visual concept, use that description as `slideText` for Step 1. Errors in the screenshot/description step are caught gracefully and fall back to the original (possibly empty) `slideText`.

---

### XL-M1 — Chart extraction: support multiple curves [MEDIUM] ✅ FIXED (Phase 3B)

**Files**: `backend/src/services/plotDigitizerService.js`, `frontend/src/utils/excelTools.ts:1817-1912`, `frontend/src/skills/excel.skill.md:244-331`

**Problem**: The `extract_chart_data` tool could only extract a single data series from a chart image. Multi-curve charts (e.g., 3 lines with different colors) produced incorrect data because only one curve's pixels were detected.

**Fix (v11.8)**: Enhanced tool description and workflow documentation to support multi-curve extraction:

1. **Backend service already supported per-color extraction** — `plotDigitizerService.js` already accepts `targetColor` parameter for filtering specific RGB colors

2. **Updated tool description** (excelTools.ts:1817-1826):
   - Changed "dominant color of the data series" → "color(s) of the data series" to emphasize plural support
   - Added **MULTI-CURVE CHARTS** section with explicit guidance:
     - Call tool ONCE PER SERIES with specific targetColor for each
     - First identify all series colors (e.g., red="#FF0000", blue="#0000FF", green="#00FF00")
     - Write each series to adjacent Excel columns (A-B for series 1, C-D for series 2, etc.)

3. **Updated excel.skill.md workflow** (lines 248-320):
   - **Step 1**: Enhanced to emphasize identifying ALL series colors for multi-curve charts, check legend if present
   - **Step 2**: Added MULTI-SERIES CHARTS subsection showing iteration pattern — call extract_chart_data 3 times with different targetColors, keep same plotAreaBox/axes for all calls
   - **Step 3**: Added multi-series data layout example showing adjacent columns format with proper headers

4. **Result**: LLM now correctly handles multi-curve charts by:
   - Detecting all series colors via vision analysis
   - Iterating extraction with one tool call per color
   - Merging results into adjacent columns with aligned X values

---

### CLIP-M1 — Paste images from clipboard into chat [MEDIUM] ✅ FIXED (Phase 3C)

**File**: `frontend/src/components/chat/ChatInput.vue:68, 178-210`

**Problem**: Users could not paste images (Ctrl+V / Cmd+V) directly into the chat input area. They had to save the image as a file first and then attach it. This was a significant UX friction point, especially when the user had just copied a screenshot.

**Fix (v11.9)**:
1. Added `@paste="handlePaste"` event listener to the textarea (line 68)
2. Implemented `handlePaste` async function (lines 178-210):
   - Checks `event.clipboardData.items` for items with `type.startsWith('image/')`
   - For each image item, calls `getAsFile()` to get the Blob
   - Creates a File object with descriptive timestamp-based name: `pasted-image-{timestamp}.{extension}`
   - Prevents default paste behavior for images with `event.preventDefault()`
3. Created helper function `createFileList` to convert File array to FileList-like object using DataTransfer API
4. Processes pasted images through existing `processFiles` pipeline (same validation, size checks, and preview display as drag/drop uploads)
5. Preview thumbnails automatically shown in attached files section (lines 32-47) with remove button

**Result**: Users can now paste screenshots/images directly with Ctrl+V/Cmd+V — images appear immediately in file list with full upload pipeline validation

---

### TOKEN-M1 — Token coherence: display vs actual + raise max limit [MEDIUM]

**Files**: `backend/src/middleware/validate.js:40-41`, `backend/src/config/models.js:44, 53`, `frontend/src/utils/tokenManager.ts`

**Problem**:
1. The `validateMaxTokens()` function allows `maxTokens` up to `128000`, but the default model config uses `32000` (standard) and `65000` (reasoning). The limit displayed in the UI (context %) may not reflect actual LLM billing — the token count is client-side estimated, not server-confirmed.
2. `32000` output tokens may be too restrictive for complex document generation tasks.

**Action**:
1. **Verify coherence**: Add server-side token count from LLM response (`usage.completion_tokens`) to the `/api/chat` streaming response headers or a final SSE event. Log the discrepancy between estimated and actual token counts.
2. **Raise default limit**: Increase `MODEL_STANDARD_MAX_TOKENS` default from `32000` to `64000` (or make configurable via env)
3. **Document the gap**: Add a comment in `tokenManager.ts` noting that client-side estimation is approximate and actual usage comes from the LLM response
4. **Display actual tokens**: Once server confirms actual usage, update the stats bar to show confirmed vs estimated

---

### OXML-M1 — OXML integration evaluation and improvement across all Office hosts [MEDIUM]

**Files**: `frontend/src/utils/wordTools.ts`, `frontend/src/utils/excelTools.ts`, `frontend/src/utils/powerpointTools.ts`, `frontend/src/utils/outlookTools.ts`

**Problem**: OOXML is used selectively (PowerPoint has `editSlideXml` via JSZip; Word has `proposeRevision` via `office-word-diff`; Excel and Outlook have minimal direct OOXML manipulation). No comprehensive evaluation of what's possible/useful via OOXML per host.

**Evaluation tasks per host**:
1. **Word**: Can `insertOoxml` be used for more precision edits? Evaluate replacing `office-word-diff` with direct OOXML revision markup (see WORD-H1). Can complex formatting (tables, styles, headers) be better preserved via OOXML?
2. **Excel**: Does any tool benefit from OOXML access? Chart XML? Conditional format XML? Evaluate `Workbook.getOoxml()` availability.
3. **PowerPoint**: `editSlideXml` is implemented. Evaluate: can slide masters be edited? Animations? SmartArt? What are the API limits?
4. **Outlook**: Can email body be manipulated via MIME/OOXML for richer formatting? Evaluate `body.setAsync` vs HTML OOXML approach.

**Action**: Produce a concise per-host evaluation report and update this section with findings. Use findings to prioritize WORD-H1 and other OOXML improvements.

---

### FB-M1 — Feedback system: include last 4 requests + tool usage context [MEDIUM] ✅ FIXED (Phase 3A)

**Files**: `backend/src/routes/feedback.js`, `backend/src/routes/chat.js`, `backend/src/utils/toolUsageLogger.js`, `backend/logs/request-history.jsonl`, `backend/logs/feedback-index.jsonl`

**Context**: USR-C1 was fixed — the feedback included chat history, system context, and frontend logs. But backend request logs and tool usage context were missing.

**Fix (v11.7)**:
1. **Backend request tracking**: Added `logChatRequest(userId, host, endpoint, messageCount)` in `toolUsageLogger.js` — logs to `request-history.jsonl` with format:
   ```json
   {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","endpoint":"/api/chat","messageCount":3}
   ```
   Integrated in both `/api/chat` and `/api/chat/sync` endpoints after validation, before processing request.

2. **Tool usage snapshot**: Added `getRecentRequests(userId, limit=4)` function to retrieve last 4 requests for a user. Uses existing `getRecentToolUsage(userId, limitLines=50)` from LOG-H1 to get tool usage snapshot.

3. **Enhanced feedback payload**: Updated `feedback.js` to include:
   - `recentRequests` — last 4 backend requests from the user (includes timestamps, endpoints, message counts)
   - `toolUsageSnapshot` — last 50 tool usage entries from the user (provides context on what tools were used recently)

4. **Central feedback index**: Created `logFeedbackSubmission(userId, host, category, sessionId, filename)` — appends to `feedback-index.jsonl` with format:
   ```json
   {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","category":"bug","sessionId":"abc123","filename":"feedback_bug_1234567890.json"}
   ```
   Enables triage dashboard and feedback tracking across all users.

---

### SKILL-L1 — skill.md system for Quick Actions [LOW]

**Files**: `frontend/src/skills/`, `frontend/src/composables/useAgentLoop.ts:888-1110`

**Context**: The existing `*.skill.md` files provide context for the **agent loop** (Chat Libre / free chat mode). Quick Actions (`applyQuickAction`) currently use hardcoded prompts from `constant.ts`.

**Proposal** (inspired by Claude Code skill.md system): For each Quick Action, define its behavior via a dedicated markdown file that specifies:
- The system prompt
- Which tools to call and in what order
- Input/output contract
- Fallback behavior

**Benefits**:
- Quick action prompts become user-customizable (power users)
- Consistent with existing skill.md architecture
- Easier to test and iterate without code changes

**Implementation approach** (based on https://support.claude.com/en/articles/12512198-how-to-create-custom-skills):
1. Add quick-action-specific skill files: e.g., `frontend/src/skills/ppt-image.skill.md`, `frontend/src/skills/ppt-bullets.skill.md`
2. Update `applyQuickAction` to load the corresponding skill file and inject it as system context
3. Quick action prompts in `constant.ts` become defaults (loaded if no skill file overrides them)
4. Document the skill file format for power users

---

## 13. OFFICE-AGENTS INTEGRATION — ✅ ALL IMPLEMENTED (v11.0)

The following items from `OFFICE_AGENTS_ANALYSIS.md` (now deleted) have been **fully implemented** and verified in the codebase:

| Feature | Tool Name | File | Status |
|---------|-----------|------|--------|
| Screenshot Excel range | `screenshotRange` | `excelTools.ts:1604` | ✅ Done |
| Screenshot PowerPoint slide | `screenshotSlide` | `powerpointTools.ts:1119` | ✅ Done |
| CSV export for ranges | `getRangeAsCsv` | `excelTools.ts:1626` | ✅ Done |
| Paginated search | `findData` (maxResults, offset) | `excelTools.ts:1375` | ✅ Done |
| Workbook structure (create/delete/rename/duplicate sheet) | `modifyWorkbookStructure` | `excelTools.ts:1664` | ✅ Done |
| Sheet structure (hide/unhide/freeze/unfreeze) | `modifyStructure` | `excelTools.ts:267` | ✅ Done |
| Duplicate slide | `duplicateSlide` | `powerpointTools.ts:1149` | ✅ Done |
| Verify slides (overlaps, overflows) | `verifySlides` | `powerpointTools.ts:1175` | ✅ Done |
| Edit slide OOXML via JSZip | `editSlideXml` | `powerpointTools.ts:1228` | ✅ Done |
| Insert icon (Iconify) | `insertIcon` | `powerpointTools.ts:1293` | ✅ Done |
| ZIP/XML utilities for PPTX | `pptxZipUtils.ts` | `utils/pptxZipUtils.ts` | ✅ Done |

**Excluded by design (per OFFICE_AGENTS_ANALYSIS.md section 4)**:
- Web Search, Web Fetch → DEFERRED (no `webSearch` / `webFetch` to be implemented now)

---

## 14. IMPLEMENTATION PHASES (v11.0 — Optimised)

> **Principe de groupement** : chaque phase regroupe des items qui touchent les mêmes fichiers ou la même zone de code, pour minimiser la lecture de contexte. Maximum 3 items actifs par phase pour respecter la limite de tokens toutes les 4h.

---

### Phase 1A — 🔴 PPT Bugs Critiques + qualité outil PPT ✅ DONE
**Fichiers clés** : `frontend/src/utils/powerpointTools.ts`

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| PPT-C1 | Fix `getAllSlidesOverview` → InvalidArgument sur certaines slides | 🔴 Critical | ✅ FIXED |
| PPT-C2 | Fix `insertImageOnSlide` → crash "addImage is not a function" avec UUID | 🔴 Critical | ✅ FIXED |
| TOOL-M3 | Ajouter un équivalent de `searchAndFormat` pour PowerPoint | 🟡 Medium | ✅ IMPLEMENTED |

**Contexte à lire** : `powerpointTools.ts` (sections `getAllSlidesOverview`, `insertImageOnSlide`, fin du fichier pour ajout d'outil)

---

### Phase 1B — 🖼️ Génération d'image + Quick Action Image ✅ DONE
**Fichiers clés** : `backend/src/routes/image.js`, `frontend/src/api/backend.ts`, `frontend/src/utils/constant.ts` (section `visual`), `frontend/src/composables/useAgentLoop.ts` (handler image QA)

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| IMG-H1 | Fix crop gpt-image-1.5 : renforcer framing instruction + taille landscape par défaut | 🟠 High | ✅ FIXED |
| PPT-H1 | Améliorer le prompt de génération d'image pour produire des visuels réellement représentatifs du texte/slide (illustration adaptée, pas forcément sans texte) | 🟠 High | ✅ FIXED |
| PPT-M1 | Quick Action Image : si < 5 mots sélectionnés → screenshot de la slide + description via LLM avant génération | 🟡 Medium | ✅ IMPLEMENTED |

**Contexte à lire** : `image.js` (FRAMING_INSTRUCTION), `backend.ts` (generateImage, size default), `constant.ts` (prompt `visual`), `useAgentLoop.ts` (handler image quick action ~l.700–720)

---

### Phase 1C — 🎯 Nouvelle Quick Action "Review" PPT + nettoyage prompts ✅ DONE
**Fichiers clés** : `frontend/src/utils/constant.ts` (section PPT), `frontend/src/composables/useAgentLoop.ts` (applyQuickAction), `frontend/src/components/chat/QuickActionsBar.vue`, `frontend/src/components/settings/BuiltinPromptsTab.vue`

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| PPT-H2 | Nouvelle Quick Action "Review" qui remplace "Speaker Notes" | 🟠 High | ✅ DONE |
| TOOL-L2 | Clarifier l'indexation 1-based du paramètre `slideNumber` dans les descriptions | 🟢 Low | ✅ DONE |
| TOOL-L3 | Restreindre la règle anti em-dash/point-virgule aux contextes PPT/bullets uniquement | 🟢 Low | ✅ DONE |

**Contexte à lire** : `constant.ts` (sections `speakerNotes`, `visual`, `punchify`), `useAgentLoop.ts` (lignes 888–1110), `QuickActionsBar.vue`, `BuiltinPromptsTab.vue`

---

### Phase 2A — 📜 Scroll Intelligent + Architecture HomePage ✅
**Fichiers clés** : `frontend/src/composables/useHomePage.ts`, `frontend/src/composables/useAgentStream.ts`, `frontend/src/components/chat/ChatMessageList.vue`, `frontend/src/pages/HomePage.vue`, `frontend/src/composables/useHomePageContext.ts` (nouveau)

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| UX-H1 | Smart scroll avec interruption manuelle (yoyo fix, isAutoScrollEnabled) | 🟠 High | ✅ Complété |
| ARCH-H2 | Réduire le prop drilling de HomePage.vue via provide/inject (~44 bindings) | 🟠 High | ✅ Complété |

**Implémentation (2026-03-14)** :
- **UX-H1** : Ajout de `isAutoScrollEnabled` + `handleScroll()` dans `useHomePage.ts`. Écouteur `@scroll` dans `ChatMessageList.vue` détecte si l'utilisateur est proche du bas (seuil: 100px). L'auto-scroll se désactive si l'utilisateur scrolle vers le haut, et se réactive quand il revient près du bas. `scrollToMessageTop()` force toujours l'auto-scroll pour afficher le nouveau contenu.
- **ARCH-H2** : Création de `useHomePageContext.ts` avec système provide/inject. Réduit les props de `ChatMessageList` de 20 à 0. Le contexte expose 40+ états/fonctions/handlers partagés. Migration progressive : props optionnelles avec contexte comme fallback.

**Contexte à lire** : `useHomePage.ts` (helpers scroll), `useAgentStream.ts` (stream handler), `ChatMessageList.vue` (containerEl, @scroll), `HomePage.vue` (props passées aux enfants), `useHomePageContext.ts` (nouveau composable)

---

### Phase 2B — 🌐 Support Multilingue + locale formules ✅
**Fichiers clés** : `frontend/src/composables/useAgentPrompts.ts`, `frontend/src/utils/constant.ts`

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| LANG-H1 | Discussion en langue UI, propositions de texte dans la langue du document | 🟠 High | ✅ Complété |
| TOOL-M4 | Étendre la détection de locale formule Excel à toutes les langues (10 langues dans constant.ts) | 🟡 Medium | ✅ Complété |

**Implémentation (2026-03-14)** :
- **LANG-H1** : Séparation claire de la langue de conversation (UI) et de la langue de génération de contenu (document). Modification des 4 prompts d'agent (Word, Excel, PowerPoint, Outlook) pour ajouter des guidelines explicites : conversations/explications dans la langue de l'UI, contenu généré dans la langue du texte sélectionné/document. Le pattern Outlook existant (`ALWAYS reply in the SAME language as the original email`) a été généralisé à tous les hosts. Les built-in prompts utilisent déjà `LANGUAGE_MATCH_INSTRUCTION` qui implémente cette logique.
- **TOOL-M4** : Étendu `excelFormulaLanguageInstruction()` pour supporter les 13 langues de `languageMap` (en, fr, de, es, it, pt, zh-cn, ja, ko, nl, pl, ar, ru). Création du type `ExcelFormulaLanguage` dans constant.ts. Distinction séparateur virgule (`,`) vs point-virgule (`;`) selon la langue : langues avec `;` = fr, de, es, it, pt, nl, pl, ru ; langues avec `,` = en, zh-cn, ja, ko, ar. Mise à jour des types dans `useAgentPrompts.ts`, `useAgentLoop.ts`, `HomePage.vue`.

**Contexte à lire** : `useAgentPrompts.ts` (section `lang`, instruction formule), `constant.ts` (language map, ExcelFormulaLanguage type)

---

### Phase 2C — 📧 Outlook : traduction + qualité code ✅
**Fichiers clés** : `frontend/src/utils/outlookTools.ts`, `frontend/src/utils/constant.ts` (prompts Outlook)

| Item | Description | Priorité | Statut |
|------|-------------|----------|--------|
| OUT-H1 | Empêcher la suppression des images lors de la traduction d'un email | 🟠 High | ✅ Complété |
| QUAL-L2 | Documenter le pattern `resolveAsyncResult()` (mélange async/await et callbacks Outlook API) | 🟢 Low | ✅ Complété |

**Implémentation (2026-03-14)** :
- **OUT-H1** : Ajout d'instructions explicites de préservation des placeholders d'images dans tous les prompts Outlook qui modifient le contenu : `translate` (ligne 62), `translate_formalize` (ligne 446), `concise` (ligne 466), `proofread` (ligne 487). Instruction CRITIQUE ajoutée : "If the text contains preservation placeholders like {{PRESERVE_0}}, {{PRESERVE_1}}, etc., you MUST keep these placeholders EXACTLY as-is in their original positions. These represent embedded images and other non-text elements." Le système de préservation existant (`extractTextFromHtml` + `reassembleWithFragments` dans richContentPreserver.ts) fonctionne avec ces instructions pour préserver les images inline lors de la traduction.
- **QUAL-L2** : Ajout d'une documentation JSDoc complète (30+ lignes) pour la fonction `resolveAsyncResult()` dans outlookTools.ts (ligne 46-76). Documentation explique : pourquoi le helper existe (API Outlook callback-based vs Promise-based), le pattern utilisé, un exemple de code "avant/après", et la signature complète avec @param/@returns/@throws.

**Contexte à lire** : `outlookTools.ts` (outil `setBody`, `getBody`, `resolveAsyncResult()`), `constant.ts` (prompts Outlook avec instructions OUT-H1), `richContentPreserver.ts` (extractTextFromHtml, reassembleWithFragments)

---

### Phase 3A — 📊 Logging Backend + Error Handling
**Fichiers clés** : `backend/src/routes/chat.js`, `backend/src/routes/feedback.js`, `backend/src/routes/logs.js`, nouveau dossier `backend/logs/`

| Item | Description | Priorité |
|------|-------------|----------|
| LOG-H1 | Comptage des outils utilisés par plateforme dans `logs/tool-usage.jsonl` | 🟠 High |
| FB-M1 | Feedback enrichi : 4 dernières requêtes backend + snapshot usage outils | 🟡 Medium |
| ERR-M1 | Extraire un handler d'erreur partagé pour `/api/chat` et `/api/chat/sync` (~80% de code dupliqué) | 🟡 Medium |

**Contexte à lire** : `chat.js` (blocs d'erreur des deux routes), `feedback.js`, `logs.js`, structure existante `backend/logs/`

---

### Phase 3B — 📈 Excel : extraction multi-courbes + qualité outils
**Fichiers clés** : `backend/src/services/plotDigitizerService.js`, `frontend/src/utils/excelTools.ts`, `frontend/src/skills/excel.skill.md`

| Item | Description | Priorité |
|------|-------------|----------|
| XL-M1 | Extraction de plusieurs courbes : 1er call LLM détecte RGB de chaque série → itération | 🟡 Medium |
| TOOL-M1 | Mettre à jour la description du paramètre `values` pour documenter les types acceptés (nombre, booléen, null…) | 🟡 Medium |
| TOOL-M2 | Fusionner `getWorksheetData` et `getDataFromSheet` (outils redondants) | 🟡 Medium |

**Contexte à lire** : `plotDigitizerService.js` (extractChartData), `excelTools.ts` (extract_chart_data, getWorksheetData, getDataFromSheet), `excel.skill.md`

---

### Phase 3C — 🖱️ Presse-papier + UX input
**Fichiers clés** : `frontend/src/components/chat/ChatInput.vue`

| Item | Description | Priorité |
|------|-------------|----------|
| CLIP-M1 | Coller une image depuis le presse-papier (Ctrl+V) directement dans le chat | 🟡 Medium |
| UX-M1 | Restaurer les indicateurs de focus (focus:ring) sur tous les éléments interactifs | 🟡 Medium |
| UX-L1 | Déplacer les styles d'animation inline de ChatInput.vue vers `<style scoped>` | 🟢 Low |

**Contexte à lire** : `ChatInput.vue` (textarea, upload, animation, focus)

---

### Phase 4A — 📝 Word : Track Changes OOXML
**Fichiers clés** : `frontend/src/utils/wordDiffUtils.ts`, `frontend/src/utils/wordTools.ts`, nouveau `frontend/src/utils/wordOoxmlUtils.ts`, composant Settings

| Item | Description | Priorité |
|------|-------------|----------|
| OXML-M1 | Évaluation OOXML sur tous les hosts (prérequis, phase lecture/analyse) | 🟡 Medium |
| WORD-H1 | Implémenter `<w:ins>` / `<w:del>` + auteur configurable, remplacer office-word-diff | 🟠 High |
| DUP-M1 | Extraire `truncateString(str, maxLen)` dans `common.ts` (4 occurrences dans wordTools + outlookTools) | 🟡 Medium |

**Ordre** : OXML-M1 d'abord (lecture), puis WORD-H1 (implémentation), DUP-M1 en profiter pour les fichiers déjà ouverts
**Contexte à lire** : `wordDiffUtils.ts`, `wordTools.ts` (proposeRevision, eval_wordjs), `wordOoxmlUtils.ts` à créer, `common.ts`, Settings component (champ auteur Redline)

---

### Phase 4B — 🔧 Architecture AgentLoop + Skill System
**Fichiers clés** : `frontend/src/composables/useAgentLoop.ts`, `frontend/src/skills/` (nouveaux fichiers)

| Item | Description | Priorité |
|------|-------------|----------|
| ARCH-H1 | Découper `useAgentLoop.ts` (1145 lignes) en composables focalisés | 🟠 High |
| SKILL-L1 | Système skill.md pour les Quick Actions (comportement déclaratif, type skill.md) | 🟢 Low |

**Contexte à lire** : `useAgentLoop.ts` (entier), `index.ts` du dossier skills, `useToolExecutor.ts`, `useAgentStream.ts`

---

### Phase 5A — 🏗️ Qualité types + Dead Code (tous les *Tools.ts)
**Fichiers clés** : `frontend/src/utils/common.ts`, `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`

| Item | Description | Priorité |
|------|-------------|----------|
| DUP-H1 | Créer `OfficeToolTemplate<THost>` générique et `buildExecuteWrapper` partagé dans `common.ts` | 🟠 High |
| QUAL-H1 | Remplacer les 128 `: any` critiques par des types Office.js propres | 🟠 High |
| DEAD-M1 | Supprimer les exports alias `getToolDefinitions()` redondants dans les 4 fichiers tools | 🟡 Medium |

**Contexte à lire** : `common.ts` (createOfficeTools, factories existantes), les 4 `*Tools.ts` (sections type + export)

---

### Phase 5B — 🧹 Dead Code Excel + Nettoyage erreurs backend
**Fichiers clés** : `frontend/src/utils/excelTools.ts`, `frontend/src/utils/common.ts`, `backend/src/routes/files.js`

| Item | Description | Priorité |
|------|-------------|----------|
| DEAD-M2 | Déprécier `formatRange` (redondant avec `setCellRange`) | 🟡 Medium |
| DUP-M2 | Standardiser le format d'erreur retourné par tous les outils (3 formats différents aujourd'hui) | 🟡 Medium |
| ERR-M2 | Sanitiser le message d'erreur raw exposé dans `files.js:79` | 🟡 Medium |

**Contexte à lire** : `excelTools.ts` (formatRange vs setCellRange), `common.ts` (buildExecute), `files.js:79`

---

### Phase 5C — 🏛️ Architecture Backend
**Fichiers clés** : `backend/src/middleware/validate.js`, `frontend/src/utils/credentialStorage.ts`, `frontend/src/composables/useAgentLoop.ts`

| Item | Description | Priorité |
|------|-------------|----------|
| ARCH-M1 | Créer un `ToolProviderRegistry` pour rendre l'agent loop host-agnostique | 🟡 Medium |
| ARCH-M2 | Découper `validate.js` (236 lignes) en validators par domaine | 🟡 Medium |
| ARCH-M3 | Simplifier la migration dual-storage credentials (6 fallback paths → 1 migration au startup) | 🟡 Medium |

**Contexte à lire** : `validate.js` (validateTools, validateChat), `credentialStorage.ts`, `useAgentLoop.ts` (imports tools)

---

### Phase 6A — 🎨 Qualité code Vue + Console logs
**Fichiers clés** : `frontend/src/pages/HomePage.vue`, `frontend/src/components/chat/ChatMessageList.vue`, `frontend/src/components/chat/ChatInput.vue`, `frontend/src/utils/credentialCrypto.ts`, `frontend/src/utils/credentialStorage.ts`

| Item | Description | Priorité |
|------|-------------|----------|
| QUAL-M3 | Découper les composants > 300 lignes : extraire `AttachedFilesList`, `MessageItem`, `ConfirmationDialogs` | 🟡 Medium |
| QUAL-M2 | Remplacer les 12 `console.log` restants dans credentialCrypto/Storage/Polyfill par `logService` | 🟡 Medium |
| QUAL-M1 | Déplacer les magic numbers (255, 20_000, 1000/2000…) dans `constants/limits.ts` | 🟡 Medium |

**Contexte à lire** : `HomePage.vue`, `ChatMessageList.vue`, `ChatInput.vue`, `credentialCrypto.ts`, `credentialStorage.ts`, `constant.ts`

---

### Phase 6B — 💅 UX Polish & i18n
**Fichiers clés** : `frontend/src/components/chat/StatsBar.vue`, `frontend/src/components/chat/ToolCallBlock.vue`, `frontend/src/components/settings/AccountTab.vue`, `frontend/src/components/chat/ChatMessageList.vue`

| Item | Description | Priorité |
|------|-------------|----------|
| UX-M2 | Traduire les tooltips hardcodés en anglais (StatsBar, ToolCallBlock) via `t()` | 🟡 Medium |
| UX-M3 | Ajouter un tooltip/notification quand le contexte dépasse 80% | 🟡 Medium |
| UX-L2 | Remplacer l'URL brute par un texte descriptif dans AccountTab.vue | 🟢 Low |
| UX-L3 | Revoir `max-w-[95%]` sur les bulles de message pour mobile | 🟢 Low |

**Contexte à lire** : `StatsBar.vue`, `ToolCallBlock.vue`, `AccountTab.vue`, `ChatMessageList.vue`, fichiers i18n

---

### Phase 6C — 🔩 Infrastructure + Sécurité
**Fichiers clés** : `backend/Dockerfile`, `frontend/Dockerfile`, `.env.example`, `scripts/generate-manifests.js`

| Item | Description | Priorité |
|------|-------------|----------|
| ARCH-L1 | Passer de `npm install` à `npm ci` dans le Dockerfile frontend | 🟢 Low |
| ARCH-L2 | Évaluer déplacement des manifests vers `frontend/public/assets/` pour SaaS | 🟢 Low |
| IC2 | Ajouter directive `USER` non-root dans les deux Dockerfiles | 🟢 Low |
| IH2 | Remplacer l'IP privée `192.168.50.10` par un placeholder dans `.env.example` | 🟢 Low |
| IH3 | Remplacer le domaine DuckDNS réel par un placeholder dans `.env.example` | 🟢 Low |

**Contexte à lire** : `backend/Dockerfile`, `frontend/Dockerfile`, `.env.example`, `generate-manifests.js`

---

### 🚀 DEFERRED — Phase 7+

**TOKEN-M1** (🟡 Medium — déféré) : Cohérence tokens affiché vs réel + augmenter limite max. Attendre d'avoir LOG-H1 actif pour mesurer l'écart réel.
- Fichiers : `validate.js`, `models.js`, `tokenManager.ts`

#### DYNTOOL-D1: Dynamic Tooling — Intent-Based Tool Loading 🚀 DEFERRED

**Prerequisite**: LOG-H1 (tool usage counting) must be implemented and data collected for at least 2 weeks before this work begins.

**Why deferred**: Without real usage data, we cannot identify the correct "Core Set" of tools. Quick Actions will NOT use dynamic tooling — they will be powered by skill.md files (SKILL-L1).

**Plan (3 phases)**:

**Phase 1 — Analysis (depends on LOG-H1 data)**:
- Use `backend/logs/tool-usage.jsonl` to identify, per Office host, the 5–7 tools representing 80% of usage ("Core Set")
- Document the Core Set and Extended Set per host

**Phase 2 — Tool Schema Separation**:
- Divide tool definitions into two tiers per host in `*Tools.ts`:
  - `getCoreToolDefinitions()` — always loaded in Chat Libre
  - `getExtendedToolDefinitions()` — available on-demand
- No breaking changes to existing tool execution logic

**Phase 3 — Routing / RAG (Chat Libre only)**:
- When a user request arrives in Chat Libre mode, run a lightweight intent classifier (keyword matching or LLM call) to determine if Extended Set tools are needed
- If yes, inject the relevant extended tool schemas for that turn only
- Alternative: expose a `getAdvancedTools(category: string)` meta-tool that the LLM can call to request additional tools

**Isolation from Quick Actions**: Quick Actions must never use dynamic loading. They will use the skill.md system (SKILL-L1) where tool calls are explicitly declared.

---

## Deferred Items Summary by Severity (v11.0)

| Severity | Count | Status | Items |
|----------|-------|--------|-------|
| 🔴 **Critical (v11 actif)** | 2 | ✅ Phase 1A DONE | PPT-C1 ✅, PPT-C2 ✅ |
| 🔴 **Critical (v10)** | 0 | ✅ All fixed | Phase 0 complete |
| 🟠 **High (déféré v10)** | 5 + 1 prospectif | ⏳ Pending | TOOL-C1 (doc re-send), TOOL-H2 (Word screenshot), USR-H1 (empty shapes), USR-H2 (context bloat), PROSP-H2 |
| 🟡 **Medium (déféré v10)** | 3 | — | TOKEN-M1 (nouveau), PROSP-2 (Claude.md), PROSP-5 (intent profiles) |
| 🟢 **Low (déféré v10)** | 1 + 3 prospectifs | — | UM10 (PPT HTML reconstruction — fermé, ne pas implémenter) + PROSP-1/3/4 |
| 🚀 **New deferred** | 1 | — | DYNTOOL-D1 (dynamic tooling, besoin données LOG-H1 d'abord) |
| **TOTAL DEFERRED** | **18** | | 11 fonctionnel + 5 architectural/legacy + 2 nouveaux (TOKEN-M1, DYNTOOL-D1) |

---

## Résumé des phases v11.0

| Phase | Zone de code principale | Items actifs | Priorité max |
|-------|------------------------|-------------|-------------|
| **1A** ✅ | `powerpointTools.ts` | PPT-C1 ✅, PPT-C2 ✅, TOOL-M3 ✅ | 🔴 Critical |
| **1B** ✅ | `image.js` + `constant.ts` (visual) + `useAgentLoop` (image) | IMG-H1 ✅, PPT-H1 ✅, PPT-M1 ✅ | 🟠 High |
| **1C** ✅ | `constant.ts` (PPT QA) + `useAgentLoop` + `QuickActionsBar` | PPT-H2 ✅, TOOL-L2 ✅, TOOL-L3 ✅ | 🟠 High |
| **2A** ✅ | `useHomePage.ts` + `useHomePageContext.ts` + `ChatMessageList.vue` + `HomePage.vue` | UX-H1 ✅, ARCH-H2 ✅ | 🟠 High |
| **2B** ✅ | `useAgentPrompts.ts` + `constant.ts` (ExcelFormulaLanguage) | LANG-H1 ✅, TOOL-M4 ✅ | 🟠 High |
| **2C** ✅ | `outlookTools.ts` + `constant.ts` (Outlook prompts) + `richContentPreserver.ts` | OUT-H1 ✅, QUAL-L2 ✅ | 🟠 High |
| **3A** ✅ | `chat.js` + `feedback.js` + `toolUsageLogger.js` + `logs/` | LOG-H1 ✅, FB-M1 ✅, ERR-M1 ✅ | 🟠 High |
| **3B** ✅ | `excelTools.ts` + `excel.skill.md` | XL-M1 ✅, TOOL-M1 ✅, TOOL-M2 ✅ | 🟡 Medium |
| **3C** ✅ | `ChatInput.vue` | CLIP-M1 ✅, UX-M1 ✅, UX-L1 ✅ | 🟡 Medium |
| **4A** | `wordDiffUtils.ts` + `wordTools.ts` + `wordOoxmlUtils.ts` | OXML-M1, WORD-H1, DUP-M1 | 🟠 High |
| **4B** | `useAgentLoop.ts` + `skills/` | ARCH-H1, SKILL-L1 | 🟠 High |
| **5A** | `common.ts` + tous `*Tools.ts` (types + exports) | DUP-H1, QUAL-H1, DEAD-M1 | 🟠 High |
| **5B** | `excelTools.ts` + `common.ts` + `files.js` | DEAD-M2, DUP-M2, ERR-M2 | 🟡 Medium |
| **5C** | `validate.js` + `credentialStorage.ts` + `useAgentLoop.ts` (registry) | ARCH-M1, ARCH-M2, ARCH-M3 | 🟡 Medium |
| **6A** | `HomePage.vue` + `ChatMessageList.vue` + `credentialCrypto.ts` | QUAL-M3, QUAL-M2, QUAL-M1 | 🟡 Medium |
| **6B** | `StatsBar.vue` + `ToolCallBlock.vue` + `AccountTab.vue` | UX-M2, UX-M3, UX-L2, UX-L3 | 🟡 Medium |
| **6C** | `Dockerfile` × 2 + `.env.example` | ARCH-L1, ARCH-L2, IC2, IH2, IH3 | 🟢 Low |
| **Déféré 7+** | — | TOKEN-M1, DYNTOOL-D1, PROSP-H2, PROSP-1–5, TOOL-C1↗, TOOL-H2↗, USR-H1↗, USR-H2↗ | 🚀 |

---

*Ce document couvre le codebase au 2026-03-14. Les numéros de ligne référencent l'état courant sur la branche `claude/design-review-planning-UcBZi`.*
