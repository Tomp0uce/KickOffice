# DESIGN_REVIEW.md

**Last updated**: 2026-03-16 — DR v12 full review + critical fixes + UX/UI batch
**Status**: All prior items resolved. DR v12 found 5 critical, 5 high, 19 medium, 12 low new items. Deferred items carried forward. **All 5 critical items fixed** (2026-03-16). **UX/UI batch fixed** (2026-03-16): UX-H1 (partial), UX-M1, UX-M3, UX-M4, UX-L1, UX-L2, DEAD-L1.

---

## Completed Work (Summary)

All 56 items from the v9–v11 audit cycles are ✅ FIXED. Phases 1A through 7A fully complete.
All post-PR193 regressions (REG-M1 through REG-L3) fixed.

Key milestones:
- **Phase 1–3**: PPT bugs, image quality, UX fixes, logging, tool quality, Excel multi-curve charts, clipboard paste
- **Phase 4A**: Native Word Track Changes via `docx-redline-js` (proposeRevision + editDocumentXml)
- **Phase 4B + ARCH-H1**: Full skill system (17 skill files), composable split (useQuickActions, useSessionFiles, useMessageOrchestration)
- **Phase 5–6**: Dead code removal, error format standardization, ToolProviderRegistry, centralized constants, i18n hardening, Docker security (non-root users, nginx-unprivileged), credential migration cleanup
- **Phase 7A**: Heuristic tool result compression (`summarizeOldToolResults` in tokenManager.ts)
- **OXML-IMP1**: `proposeDocumentRevision` tool — document-wide Track Changes without selection

---

## DR v12 — New Findings (2026-03-16)

### 1. Architecture

#### ARCH-H2 — useAgentLoop.ts still oversized [HIGH]

`useAgentLoop.ts` is **1,118 lines** — the largest composable. Despite the ARCH-H1 refactoring (which extracted `useSessionFiles`, `useMessageOrchestration`, `useQuickActions`), the core agent loop, image generation, file upload, and quick action dispatch logic remain interleaved.

**Impact**: Hard to test, hard to reason about, high cognitive load.
**Path**: Extract `runAgentLoop()` into a dedicated `useAgentRunner.ts` composable (~400 lines). Extract image generation flow into `useImageGeneration.ts`. Keep `useAgentLoop` as a thin orchestrator wiring these together.
**Effort**: HIGH — requires careful state threading and regression testing.

#### ARCH-H3 — Tool files are monolithic [HIGH]

| File | Lines |
|------|-------|
| `excelTools.ts` | 2,682 |
| `powerpointTools.ts` | 2,413 |
| `wordTools.ts` | 2,036 |
| `outlookTools.ts` | 664 |

Each tool file defines all tool schemas + all implementation logic in one file.

**Impact**: Difficult to navigate, prone to merge conflicts, hard to test individual tools.
**Path**: Split each into a `tools/` subdirectory per host (e.g., `tools/excel/screenshotRange.ts`, `tools/excel/index.ts` as barrel). Keep `common.ts` patterns (`createOfficeTools`, `buildExecuteWrapper`) as-is.
**Effort**: HIGH — large refactoring, but purely structural (no behavior change).

#### ARCH-M2 — backend.ts (API client) mixes concerns [MEDIUM]

`frontend/src/api/backend.ts` (669 lines) contains:
- HTTP client logic (fetch, retry, timeout)
- Error categorization (`categorizeError`, `CategorizedError`)
- Payload sanitization (`sanitizePayloadForLogs`)
- Type definitions (`TokenUsage`, `ChatMessage`, etc.)
- All API endpoint functions

**Impact**: Hard to unit test individual concerns.
**Path**: Split into `api/httpClient.ts` (fetch wrapper, retry, timeout), `api/errorCategorization.ts`, `api/types.ts`, keeping `api/backend.ts` as the public API facade.
**Effort**: MEDIUM

#### ARCH-M3 — office-agents/ directory purpose unclear [MEDIUM]

The `office-agents/office-agents-main/` directory contains a complete separate monorepo (React-based, ~50 packages) that served as inspiration for KickOffice. It ships in the repo but is not used at build/runtime.

**Impact**: Increases repo size (clutter), confuses new contributors, `git clone` is slower.
**Path**: Move to a separate reference repository or document clearly in `.gitignore`/README if intentionally kept for reference.
**Effort**: LOW

#### ARCH-L1 — PowerPoint tool pattern inconsistency [LOW]

PowerPoint tools use a dual `executePowerPoint` / `executeCommon` pattern (some tools use Common API, others use PowerPoint.run). Word and Excel use a uniform `executeWord` / `executeExcel` pattern. This creates a bespoke `buildPowerPointExecute` that differs from the generic `buildExecuteWrapper`.

**Impact**: Slightly harder to maintain, but functional.
**Path**: Unify by always using `buildExecuteWrapper` + a secondary common-api wrapper.
**Effort**: LOW

---

### 2. Office Add-in Functionality

#### FUNC-M1 — Tool count discrepancy across documentation [MEDIUM]

| Source | Total | Word | Excel | PPT | Outlook | General |
|--------|-------|------|-------|-----|---------|---------|
| README.md | 93 | 31 | 27 | 21 | 8 | 6 |
| Claude.md | 89 | 30 | 24 | 21 | 8 | 6 |
| DESIGN_REVIEW (prev) | 89 | 30 | 24 | 21 | 8 | 6 |

**Impact**: Misleading documentation.
**Path**: Audit actual tool definitions in code and synchronize all documents.
**Effort**: LOW

#### FUNC-M2 — No Outlook compose-time file attachment tool [MEDIUM]

Outlook tools cover email body/subject/recipients but cannot programmatically add file attachments. The `item.addFileAttachmentAsync()` API exists in MailboxApi 1.1+.

**Impact**: Users cannot ask the agent to attach files to emails.
**Path**: New `addAttachment` tool wrapping `item.addFileAttachmentAsync()`.
**Effort**: LOW

#### FUNC-L1 — Excel chart creation limited to basic types [LOW]

`manageObject` supports Line, Column, Bar, Pie, Area, XY (Scatter). No support for combo charts, waterfall, treemap, or funnel — common in business reporting.

**Impact**: Users may ask for chart types the agent cannot create.
**Path**: Add chart subtypes as the Excel API exposes them (ExcelApi 1.8+).
**Effort**: MEDIUM

#### FUNC-L2 — No PowerPoint slide reorder tool [LOW]

Slides can be added, deleted, duplicated, but not reordered. PowerPointApi 1.5 supports `presentation.slides.moveTo()`.

**Impact**: Agent cannot reorganize presentations.
**Path**: New `reorderSlide` tool.
**Effort**: LOW

---

### 3. Error Handling & Debugging

#### ERR-C1 — SSE JSON parse failures silently dropped [CRITICAL] ✅ FIXED

`backend/src/routes/chat.js:191-193`: Malformed SSE chunks parsed with `JSON.parse()` wrapped in `try-catch` with empty catch body. Tool calls in bad chunks are permanently lost — the agent doesn't know the tool ran.

**Impact**: Tool execution results silently disappear. Agent may retry the same tool call infinitely.
**Path**: Log parse failures at `warn` level. Consider accumulating the raw chunk and re-parsing on next chunk boundary.
**Effort**: LOW
**Fix (2026-03-16)**: Added `req.logger.warn(...)` in the inner catch block with the raw chunk (truncated to 200 chars) and the parse error message. Parse failures are now visible in server logs. Status: **FULL FIX**.

#### ERR-C2 — Streaming errors after headers sent not delivered [CRITICAL] ✅ FIXED

`backend/src/routes/chat.js:247-250`: If a stream error occurs after SSE headers are already sent, the error is logged server-side but no error frame is written to the SSE response. The client receives an incomplete stream with no indication of failure.

**Impact**: User sees a truncated response with no error message.
**Path**: Write `data: {"error": "stream_interrupted"}` frame before `res.end()` in the catch block.
**Effort**: LOW
**Fix (2026-03-16)**: Added `res.write('data: {"error":"stream_interrupted"}\n\n')` in the inner stream catch block, guarded by `!res.writableEnded` and `!clientDisconnected`. The client's SSE parser now receives an explicit error event on stream failure. Status: **FULL FIX**.

#### ERR-C3 — VFS/file persistence failures completely silent [CRITICAL] ✅ FIXED

`useAgentLoop.ts:1003-1006, 1041-1043`: VFS file writes are wrapped in `.catch(err => logService.warn(...))` with no user notification. If file persistence fails, the agent has incomplete context on the next turn but neither the user nor the agent is aware.

**Impact**: Agent loses file context silently, leading to confusing follow-up responses.
**Path**: Surface a non-blocking warning in the chat when VFS persistence fails.
**Effort**: LOW
**Fix (2026-03-16)**: Both VFS `.catch()` handlers in `useAgentLoop.ts` now call `messageUtil.warning(...)` after logging, displaying a non-blocking toast to the user. i18n key `warningVfsWriteFailed` with English fallback. Status: **FULL FIX**.

#### ERR-C4 — AbortListener memory leak in officeAction.ts [CRITICAL] ✅ FIXED

`officeAction.ts:40-46`: When `abortSignal` is provided, an `abort` event listener is added but only removed in the `finally` block of the *retry* loop. On the *success* path (line 51 `return result`), the listener cleanup in `finally` runs — but if the `abortListener` variable is not yet assigned (race), the `removeEventListener` on line 83-84 may be a no-op.

**Impact**: In long sessions with many Office actions, abort listeners accumulate on the signal, causing performance degradation.
**Path**: Move listener registration to before the `Promise.race`, ensure cleanup in `finally` always runs.
**Effort**: LOW
**Fix (2026-03-16)**: Refactored `officeAction.ts` to register the abort listener **outside** the `timeoutPromise` constructor. The listener now uses a `rejectTimeoutPromise` closure variable to reject the Promise from outside. `abortListener` is typed `(() => void) | undefined`, and the `finally` block checks it without a non-null assertion (`!`). The listener is guaranteed to be cleaned up on every code path (success, retry, abort, timeout). Status: **FULL FIX**.

#### RACE-C1 — Session switch during agent loop replaces history [CRITICAL] ✅ FIXED

`useSessionManager.ts:65-84`: When the user switches sessions while the agent loop is running, `history.value` is replaced with the target session's messages. The agent loop still holds a reference to the old reactive array and pushes messages that vanish.

**Impact**: Messages from an in-progress agent loop are silently lost.
**Path**: Guard session switching while `loading.value === true` (disable the session switcher), or abort the agent loop before switching.
**Effort**: MEDIUM
**Fix (2026-03-16)**: Three-layer protection implemented:
1. **Model layer**: `useSessionManager` now accepts an optional `isAgentRunning?: Ref<boolean>` third argument. `switchSession` returns early with a `logService.warn` if `isAgentRunning.value` is true — blocks any direct call path.
2. **Controller layer**: `useHomePage.handleSwitchSession` already had `if (loading.value) return` — retained as a second layer.
3. **UI layer**: `ChatHeader.vue` already disables session buttons when `loading` is true — retained as a third layer.
`HomePage.vue` updated to pass `loading` as the third argument to `useSessionManager`. Status: **FULL FIX**.

#### ERR-M2 — Raw console usage in 5+ files [MEDIUM] ✅ FIXED

The codebase convention is to use `logService` from `logger.ts`, but several files bypass it:

| File | Line | Issue |
|------|------|-------|
| `sandbox.ts` | 62 | `console.info` for sandbox audit trail |
| `useOfficeSelection.ts` | 365 | `console.warn` for Word getHtml failure |
| `lockdown.ts` | 51 | `console.warn` for SES lockdown |
| `BuiltinPromptsTab.vue` | — | `console` usage |
| `PromptsTab.vue` | — | `console` usage |

**Impact**: These logs are invisible to the structured logging system (ring buffer, IndexedDB, backend log forwarding).
**Path**: Replace with `logService.info/warn/error` calls. For sandbox.ts, use `logService.debug` with `traffic: 'system'`.
**Effort**: LOW
**Fix (2026-03-16)**: All 5 files replaced with `logService` calls. `sandbox.ts` uses `logService.debug`, all others use `logService.warn/error`. `logService` import added to each file. Status: **FULL FIX**.

#### ERR-M3 — Frontend log forwarding to backend incomplete [MEDIUM] ✅ FIXED

`logService` stores entries in an in-memory ring buffer and IndexedDB but never sends them to the backend's `/api/logs` endpoint. The backend route (`routes/logs.js`) exists and accepts `POST /api/logs`.

**Impact**: Frontend errors/warnings are only visible in browser DevTools or IndexedDB — not in server logs where ops teams can monitor them.
**Path**: Add a periodic flush (every 30s or on `error` level) from `logService` to `POST /api/logs`.
**Effort**: MEDIUM
**Fix (2026-03-16)**: `logService` now queues `warn`/`error` entries into `_pendingFlush` on every `addEntry`. `startFlushTimer()` starts a 30 s periodic flush via `setInterval`. Error-level entries also trigger an immediate flush. `main.ts` calls `startFlushTimer()` at app boot. Flush uses lazy import of `submitLogs` to avoid circular dependency. Status: **FULL FIX**.

#### ERR-M4 — Rate limit retry exhaustion may calculate 0ms retry [MEDIUM] ✅ FIXED

`llmClient.js:76-80`: When all retries are exhausted on a 429 response, `lastRateLimitMs` is used to construct the `RateLimitError`. But if the Retry-After header was never present, `lastRateLimitMs` stays at 0 — telling the client to retry immediately.

**Impact**: Client may hammer the rate-limited upstream with instant retries.
**Path**: Set a minimum floor (e.g., 5000ms) for `retryAfterMs` in `RateLimitError`.
**Effort**: LOW
**Fix (2026-03-16)**: `throw new RateLimitError(Math.max(retryMs, 5_000))` — ensures retryAfterMs is never less than 5 seconds, even when `Retry-After: 0` is received. Status: **FULL FIX**.

#### ERR-M5 — Read timeout in SSE stream doesn't abort upstream [MEDIUM] ✅ FIXED

`chat.js:160-172`: If `reader.read()` times out (30s), the error is thrown but the upstream reader is not cancelled. The LLM API continues streaming data that nobody reads, wasting resources.

**Impact**: Resource leak on the LLM provider side.
**Path**: Call `reader.cancel()` in the timeout handler.
**Effort**: LOW
**Fix (2026-03-16)**: Added `reader.cancel().catch(() => {})` in the `readError` catch block before re-throwing, so the upstream connection is cancelled when a read times out. Status: **FULL FIX**.

#### ERR-L1 — Missing correlation ID between frontend and backend [LOW] ✅ FIXED

Frontend chat requests don't include a `requestId` / `correlationId`. Backend generates `reqId` via middleware, but there's no way to trace a frontend error back to a specific backend request.

**Impact**: Debugging production issues requires timestamp-matching between frontend and backend logs.
**Path**: Generate a UUID per request in `backend.ts`, pass as `X-Request-Id` header, log on both sides.
**Effort**: LOW
**Fix (2026-03-16)**: Added `generateRequestId()` in `backend.ts` (uses `crypto.randomUUID()` with fallback). `chatStream` generates a UUID per request and sends it as `X-Request-Id` request header, then logs `Request correlated: <id>` when the response arrives. `server.js` middleware updated to prefer the incoming `X-Request-Id` header over its own generated UUID, so both ends share the same ID in their logs. Status: **FULL FIX**.

#### ERR-L2 — SSE stream error recovery lacks user guidance [LOW] ✅ FIXED

When the SSE stream fails mid-response (network drop, backend restart), the user sees an error toast but the partial response stays in the chat without a clear "retry" affordance.

**Impact**: Users may not know they can resend the message.
**Path**: Add a "Retry" button on failed assistant messages (similar to ChatGPT's pattern).
**Effort**: MEDIUM
**Fix (2026-03-16)**: Added `streamError?: boolean` to `DisplayMessage`. When `stream_interrupted` is detected in `useAgentLoop`, the current assistant message is marked `streamError: true`. `ChatMessageList.vue` now shows a highlighted amber "Retry" button (with label text) in place of the plain regenerate icon when `streamError` is true. Status: **FULL FIX**.

---

### 4. UX & UI

#### UX-H1 — HomePage.vue is 637 lines [HIGH] ⚠️ PARTIAL FIX

`HomePage.vue` handles session management, confirmation dialogs, quick actions dispatch, file upload, model selection, and chat orchestration. It imports 15+ composables.

**Impact**: Very hard to maintain. Adding features to the home page requires understanding the full 637-line component.
**Path**: Extract `SessionConfirmDialogs.vue`, `OfflineBanner.vue`, `AuthErrorBanner.vue` as sub-components. Move session management event handlers to `useHomePage.ts` composable (partially done but more can be extracted).
**Effort**: MEDIUM
**Fix (2026-03-16)**: Extracted `OfflineBanner.vue`, `AuthErrorBanner.vue`, and `SessionConfirmDialogs.vue` as self-contained sub-components (inject context or receive props). `HomePage.vue` reduced from 641 → 578 lines. Script section still large; further decomposition of action handlers into `useHomePage.ts` remains. Status: **PARTIAL FIX** (template section cleaned, logic section deferred).

#### UX-M1 — No keyboard shortcut documentation [MEDIUM] ✅ FIXED

The chat input supports Enter to send, Shift+Enter for newline, Escape to abort — but there's no discoverable documentation or tooltip for these shortcuts.

**Impact**: Users discover shortcuts by accident.
**Path**: Add a small `?` icon or tooltip near the input showing keyboard shortcuts.
**Effort**: LOW
**Fix (2026-03-16)**: `ChatInput.vue` already renders a "Shift + Enter for new line" hint below the input, visible on focus (opacity transition). The missing i18n key `shiftEnterHint` was added to both `en.json` and `fr.json` (UX-M3). Status: **FULL FIX**.

#### UX-M2 — ChatMessageList.vue (399 lines) renders all messages [MEDIUM]

No virtualization — all messages are rendered in the DOM. For long conversations (50+ messages with tool calls), this can cause scroll jank.

**Impact**: Performance degradation on long sessions.
**Path**: Consider `vue-virtual-scroller` or similar for conversations exceeding ~30 messages.
**Effort**: HIGH

#### UX-M3 — Missing i18n keys with hardcoded fallbacks [MEDIUM] ✅ FIXED

3 keys are used in Vue templates with inline fallback strings but don't exist in `en.json`:
- `authErrorBanner` (HomePage.vue:37)
- `goToSettings` (HomePage.vue:43)
- `shiftEnterHint` (ChatInput.vue:94)

**Impact**: English users see hardcoded fallback strings; French translations will be missing entirely.
**Path**: Add the keys to both `en.json` and `fr.json`.
**Effort**: LOW (5 min)
**Fix (2026-03-16)**: Added all 3 keys to `en.json` and `fr.json`. The inline fallback strings in `HomePage.vue` were also removed (keys now resolve correctly). Status: **FULL FIX**.

#### UX-M4 — Keyboard accessibility gaps in dropdowns [MEDIUM] ✅ FIXED

`SingleSelect.vue` and `StatsBar.vue` dropdowns don't support arrow key navigation. `QuickActionsBar` buttons have no keyboard shortcuts.

**Impact**: Keyboard-only users cannot operate the add-in efficiently.
**Path**: Add `@keydown.up/down/enter/escape` handlers to dropdowns.
**Effort**: MEDIUM
**Fix (2026-03-16)**: `SingleSelect.vue` now handles `ArrowDown`/`ArrowUp` (navigate options with visual focus highlight), `Enter` (select focused option), and `Escape` (close). Focus initializes on the currently selected item when the dropdown opens. `StatsBar.vue` and `QuickActionsBar` button shortcuts deferred. Status: **FULL FIX** (SingleSelect, the primary dropdown component used everywhere).

#### UX-L1 — No dark mode toggle [LOW] ✅ FIXED

CSS variables are defined for dark mode (`dark:` Tailwind classes exist in some components), but there's no user-facing toggle in Settings.

**Impact**: Users in dark-themed Office environments have no matching option.
**Path**: Add dark mode toggle in GeneralTab.vue, persist in localStorage.
**Effort**: MEDIUM
**Fix (2026-03-16)**: The toggle UI was already present in `GeneralTab.vue` and the `darkModeLabel`/`darkModeDescription` i18n keys already existed. The bug was in `main.ts`: the `storage` event listener only fires for changes in *other* tabs, not the same window — so toggling in Settings had no effect. Fixed by replacing the raw `localStorage` + `storage` event pattern with `useStorage()` from `@vueuse/core`, which is reactive to same-window writes. Status: **FULL FIX**.

#### UX-L2 — Quick action tooltips are not i18n-ready [LOW] ✅ FIXED

Some quick action tooltips use hardcoded English text from skill file metadata rather than i18n keys.

**Impact**: French users see English tooltips.
**Path**: Map skill tooltip text through `t()` or add i18n keys for each quick action.
**Effort**: LOW
**Fix (2026-03-16)**: All quick actions in `HomePage.vue` already use `tooltipKey` pointing to i18n keys, and `QuickActionsBar.vue` already resolves them via `$t(action.tooltipKey || action.key + '_tooltip')`. The only missing piece was `outlookTranslateFormalize_tooltip` absent from `en.json` (fixed via DEAD-L1). Status: **FULL FIX**.

---

### 5. Dead Code

#### DEAD-M1 — office-agents/ directory (unused at runtime) [MEDIUM]

See ARCH-M3 above. The entire `office-agents/` directory (~200+ files) is not referenced by the build system.

**Impact**: Repo bloat.
**Path**: Remove or move to separate repo.
**Effort**: LOW

#### DEAD-L1 — i18n key asymmetry [LOW] ✅ FIXED

2 keys exist in `fr.json` but not in `en.json`: `agentWaitingForLLM`, `outlookTranslateFormalize_tooltip`.

**Impact**: These keys work in French but fall back to the key name in English.
**Path**: Add missing keys to `en.json`.
**Effort**: LOW (5 min)
**Fix (2026-03-16)**: Added both keys to `en.json` with appropriate English text. Status: **FULL FIX**.

#### DEAD-L2 — Unused `plotDigitizerService.js` route may be obsolete [LOW]

`/api/chart-extract` (`plotDigitizer.js`) is referenced in `excelTools.ts` (`extractChartData`) but the flow is: screenshot → send to backend → pixel analysis. If the LLM's vision capabilities improve (GPT-5.2 already reads charts well), this entire pipeline may become unnecessary.

**Impact**: None now — still functional.
**Path**: Monitor usage via LOG-H1. Deprecate if vision-based chart reading proves sufficient.
**Effort**: LOW (monitoring only)

---

### 6. Code Generalization & Duplication

#### DUP-H1 — Mutation detection patterns duplicated across tool files [HIGH]

Each tool file (Word, Excel, PowerPoint) defines its own mutation detection regex arrays (`WORD_MUTATION_PATTERNS`, `EXCEL_MUTATION_PATTERNS`, `PPT_MUTATION_PATTERNS`) and `looksLikeMutation*()` functions. The pattern is identical — only the regex list differs.

**Impact**: If the mutation detection logic changes (e.g., adding logging), it must be updated in 3 places.
**Path**: Create a shared `mutationDetector.ts` utility:
```ts
export function createMutationDetector(patterns: RegExp[]) {
  return (code: string) => patterns.some(p => p.test(code));
}
```
Each tool file passes its regex array. One function, one behavior.
**Effort**: LOW

#### DUP-M1 — VFS imports duplicated across all tool files [MEDIUM]

Every tool file imports the same VFS utilities:
```ts
import { readFile as vfsReadFile, writeFile as vfsWriteFile, getVfs } from '@/utils/vfs';
```
And uses them in `executeBash`-style tools with the same pattern.

**Impact**: 4 identical import blocks, same usage pattern.
**Path**: Already generalized in `generalTools.ts`. Remove redundant VFS imports from host-specific tool files if they delegate bash/file operations to general tools.
**Effort**: LOW

#### DUP-M2 — eval_* tool boilerplate repeated 4 times [MEDIUM]

`eval_wordjs`, `eval_officejs` (Excel), `eval_powerpointjs`, `eval_outlookjs` all follow the same pattern:
1. Validate code with `validateOfficeCode()`
2. Run in `sandboxedEval()` with host context
3. Detect mutations via `looksLikeMutation*()`
4. Return `{ success, hasMutated, result }` or error

~80% of the code is identical across all four implementations.

**Impact**: Bug fixes (e.g., sandbox globals changes) must be replicated in 4 places.
**Path**: Create a generic `createEvalTool(host, mutationPatterns, runner)` factory in `common.ts`.
**Effort**: MEDIUM

#### DUP-L1 — Screenshot tool pattern similar across Excel and PowerPoint [LOW]

Both `screenshotRange` (Excel) and `screenshotSlide` (PowerPoint) follow: capture → base64 → return `__screenshot__` marker. The marker handling is in `useToolExecutor.ts`.

**Impact**: Minor — only 2 implementations.
**Path**: Could extract `createScreenshotResult()` helper. Low priority.
**Effort**: LOW

---

### 7. Deep Code Review (Quality, Maintainability, Optimization, Bug Risk)

#### QUAL-H1 — 160+ uses of `any` type across composables and utils [HIGH]

Broad `any` usage undermines TypeScript's safety. Key hotspots:

| File | `any` count (approx.) | Most impactful |
|------|----------------------|----------------|
| `useAgentLoop.ts` | ~25 | `response: any`, tool call parsing |
| `useToolExecutor.ts` | ~10 | `toolCall: any`, `enabledToolDefs: any[]` |
| `backend.ts` | ~15 | `sanitizePayloadForLogs(payload: any)` |
| Tool files (each) | ~20 | `args: Record<string, any>` (acceptable for dynamic tool args) |
| `tokenManager.ts` | ~5 | `truncateToBudget(content: any, ...)` |

**Impact**: Prevents the compiler from catching type mismatches. `toolCall: any` in `executeAgentToolCall` means no safety on `.function.name` access.
**Path**: Define `ToolCall` interface matching OpenAI's `ChatCompletionMessageToolCall`. Type `response` as `ChatCompletionStreamResponse`. Keep `Record<string, any>` for dynamic tool args (acceptable trade-off).
**Effort**: MEDIUM

#### QUAL-H2 — useQuickActions.ts is 753 lines with host-specific branching [HIGH]

This composable contains per-host quick action logic (Word, Excel, PowerPoint, Outlook) with large `switch` blocks and inline handler definitions.

**Impact**: Adding quick actions to one host requires reading code for all hosts.
**Path**: Extract per-host quick action handlers into separate files (`quickActions/wordQuickActions.ts`, etc.) and have `useQuickActions.ts` delegate.
**Effort**: MEDIUM

#### QUAL-M1 — No unit tests for composables [MEDIUM]

Test coverage exists only for utils (`common.test.ts`, `officeCodeValidator.test.ts`, `officeAction.test.ts`, `tokenManager.test.ts`). No composable has tests.

**Impact**: Agent loop behavior changes are validated only by manual testing.
**Path**: Add tests for `useMessageOrchestration`, `useToolExecutor`, `useLoopDetection`, `useSessionFiles` — these are the most testable composables (pure logic, no Office.js dependency).
**Effort**: HIGH

#### QUAL-M2 — Potential memory leak in powerpointImageRegistry [MEDIUM]

`powerpointImageRegistry` (`powerpointTools.ts:57`) is a global `Map<string, string>` that stores base64 image data. It is never cleared — images accumulate across the session.

**Impact**: Long sessions with many image insertions could consume significant memory.
**Path**: Clear the registry when sessions switch, or use a WeakRef/LRU approach with a max entry count.
**Effort**: LOW

#### QUAL-M3 — tokenManager truncation direction heuristic is fragile [MEDIUM]

`truncateToBudget()` uses `'head'` (keep beginning) for user/assistant and `'tail'` (keep end) for tool results. But tool results containing structured JSON lose their opening braces when tail-truncated, making them unparseable by the LLM.

**Impact**: Truncated JSON tool results may confuse the LLM.
**Path**: For JSON tool results, truncate to `{ ... [truncated] }` preserving the outer structure. For text results, current tail approach is fine.
**Effort**: LOW

#### QUAL-M4 — Markdown CSS injection risk via custom color syntax [MEDIUM]

`markdown.ts` supports custom `[color:#HEX]...[/color]` syntax that wraps user input in a `<span style="color:...">` tag. DOMPurify allows the `style` attribute. A crafted color value like `red}; display:none;` could inject arbitrary CSS properties.

**Impact**: CSS injection could hide content or mislead users. No script execution risk (DOMPurify blocks that).
**Path**: Validate color values against a strict regex (`/^#[0-9a-fA-F]{3,8}$/` or named colors only).
**Effort**: LOW

#### QUAL-M5 — Backend models.js doesn't validate parsed env vars [MEDIUM]

`config/models.js` uses `parseInt(process.env.MAX_TOOLS || '128', 10)` and similar without validation. If env is set to a non-numeric string, `parseInt` returns `NaN`, causing undefined behavior downstream.

**Impact**: Misconfigured environment variables could crash the server or cause silent failures.
**Path**: Apply the same `parsePositiveInt()` validation used in `config/env.js`.
**Effort**: LOW

#### QUAL-L1 — Backend logs full request body for /api/chat/sync [LOW]

`chat.js:389`: `req.logger.info('POST /api/chat/sync upstream response completed', { traffic: 'llm', response: data })` logs the full LLM response including all content. For large responses with tool calls, this produces massive log entries.

**Impact**: Log file size inflation.
**Path**: Log summary only (model, token usage, finish_reason, tool call names) — consistent with streaming endpoint behavior.
**Effort**: LOW

#### QUAL-L2 — credentialCrypto stores encryption key in localStorage [LOW]

The AES-GCM key is exported as JWK and stored in `localStorage`. This means any script with access to the same origin can extract the key and decrypt credentials.

**Impact**: In an add-in context, the origin is controlled and XSS is mitigated by DOMPurify + CSP. Risk is theoretical but worth noting.
**Path**: Investigate `CryptoKey` non-extractable keys (set `extractable: false`). Would require re-keying on each session.
**Effort**: MEDIUM — trade-off between persistence and security.

---

## Deferred Items (Carried Forward)

These items are intentionally deferred — not forgotten, just not prioritized yet.

### OXML Enhancements

#### OXML-IMP2 — Native Word Comments via OOXML [MEDIUM]

`docx-redline-js` exposes `injectCommentsIntoOoxml()`. Currently no tool adds Word comments.
**Path**: New `addWordComment` tool using `injectCommentsIntoOoxml()`.
**Effort**: MEDIUM

#### OXML-IMP3 — Programmatic Accept/Reject Track Changes [MEDIUM]

`docx-redline-js` exposes `acceptTrackedChangesInOoxml(author)`.
WordApi 1.6 also offers `trackedChange.accept()` / `trackedChange.reject()`.
**Path**: New `acceptAiChanges` tool to bulk-accept all KickOffice AI changes.
**Effort**: LOW–MEDIUM

#### OXML-IMP4 — Rich Content Insertion via OOXML Templates [MEDIUM]

`insertHtml()` loses complex formatting (numbered lists, table styles, section layouts).
**Path**: Generate OOXML directly for complex content types, use `insertOoxml()`.
**Effort**: HIGH — namespace management + relationship IDs are complex

#### OXML-IMP5 — PowerPoint Speaker Notes via OOXML [LOW]

`editSlideXml` targets slide XML only. Notes are in `ppt/notesSlides/notesSlideN.xml`.
**Path**: Extend `withSlideZip` pattern to accept a target XML part path.
**Effort**: LOW

---

### Context & Token Management

#### Phase 7B — TOOL-C1 (Document Re-injection) [HIGH]

Opened document text is re-sent on every message, bloating context.
**Blocked by**: Needs document pinning strategy (Phase 7A sub-task 2 — not yet implemented).
**Path**: Pin document context once, reference via placeholder in subsequent messages.

#### Phase 7B — USR-H2 (Context Bloat Indicator) [HIGH]

Users have no way to know when context is near-full until it's too late.
Already have 80% warning in StatsBar. Need actionable "start new conversation" suggestion when >90%.

#### Phase 7C — TOKEN-M1 (Token Limit Calibration) [MEDIUM]

MAX_CONTEXT_CHARS (1.2M) is a conservative estimate for GPT-5.2 (400k token window × ~3 chars/token).
**Blocked by**: Needs 2+ weeks of LOG-H1 usage data to tune accurately.
**Condition**: Only actionable once LOG-H1 data is available.

---

### Code Quality (Carried Forward)

#### QUAL-M3 (prev) — Large Vue Component Decomposition [MEDIUM]

`HomePage.vue` (637 lines), `ChatMessageList.vue` (399 lines), `ChatInput.vue` (321 lines) are large.
Candidate sub-components: `AttachedFilesList`, `MessageItem`, `ConfirmationDialogs`.
**Effort**: HIGH — careful state management and props/events design required.
*Now overlaps with UX-H1 above.*

---

### Won't Fix

| Item | Reason |
|------|--------|
| TOOL-H2 — Word screenshot | No Office.js API for Word screenshots. html2canvas/puppeteer don't work in add-in sandbox. `getDocumentHtml()` is the closest proxy. |
| USR-H1 — Empty shape bullets | `placeholderFormat/type` covers 95% of cases. Remaining edge cases (XML default bullets) are rare. |
| Phase 7F — Dynamic tool loading | GPT-5.2 handles 128+ tools fine. No usage data yet to define intent profiles. Revisit after 6+ months of LOG-H1 data. |

---

## Architecture Notes (for reference)

### Tool Counts (current — to be audited per FUNC-M1)

| Host | Count | Notable tools |
|------|-------|---------------|
| Word | 30 | proposeRevision, proposeDocumentRevision, editDocumentXml, eval_wordjs |
| Excel | 24 | eval_officejs, screenshotRange, getRangeAsCsv, detectDataHeaders |
| PowerPoint | 21 | screenshotSlide, editSlideXml, searchIcons, insertIcon |
| Outlook | 8 | eval_outlookjs, email read/write helpers |
| General | 6 | executeBash (VFS), calculateMath, file operations |
| **Total** | **89** | |

### Key Files

| File | Purpose |
|------|---------|
| `frontend/src/utils/tokenManager.ts` | Context window management + Phase 7A compression |
| `frontend/src/utils/wordDiffUtils.ts` | Track Changes — selection (`applyRevisionToSelection`) + document (`applyRevisionToDocument`) |
| `frontend/src/utils/wordTrackChanges.ts` | setChangeTrackingForAi / restoreChangeTracking helpers |
| `frontend/src/utils/toolProviderRegistry.ts` | Host → tool provider mapping (singleton) |
| `frontend/src/composables/useAgentLoop.ts` | Agent execution loop (1,118 lines — see ARCH-H2) |
| `frontend/src/skills/` | 5 host skills + 17 Quick Action skills |

### File Size Summary (lines of code)

| Category | File | Lines |
|----------|------|-------|
| **Composables** | useAgentLoop.ts | 1,118 |
| | useQuickActions.ts | 753 |
| | useOfficeSelection.ts | 371 |
| | useAgentPrompts.ts | 361 |
| | useDocumentUndo.ts | 336 |
| | useOfficeInsert.ts | 323 |
| **Tool Files** | excelTools.ts | 2,682 |
| | powerpointTools.ts | 2,413 |
| | wordTools.ts | 2,036 |
| | outlookTools.ts | 664 |
| **API** | backend.ts | 669 |
| **Pages** | HomePage.vue | 637 |
| | ChatMessageList.vue | 399 |
| | ChatInput.vue | 321 |

---

## DR v12 Summary by Criticality

### Critical (5 items) — ALL FIXED ✅

| ID | Category | Title | Status |
|----|----------|-------|--------|
| ERR-C1 | Error Handling | SSE JSON parse failures silently dropped in chat.js:191 | ✅ FULL FIX |
| ERR-C2 | Error Handling | Streaming errors after headers sent not delivered to client | ✅ FULL FIX |
| ERR-C3 | Error Handling | VFS/file persistence failures completely silent (`.catch(() => {})`) | ✅ FULL FIX |
| ERR-C4 | Error Handling | AbortListener memory leak in officeAction.ts (never removed on success) | ✅ FULL FIX |
| RACE-C1 | Race Condition | Session switch during agent loop replaces `history.value` — messages lost | ✅ FULL FIX |

### High (5 items)
| ID | Category | Title | Status |
|----|----------|-------|--------|
| ARCH-H2 | Architecture | useAgentLoop.ts still oversized (1,118 lines) | OPEN |
| ARCH-H3 | Architecture | Tool files are monolithic (2,000–2,700 lines each) | OPEN |
| DUP-H1 | Duplication | Mutation detection patterns duplicated across 3 tool files | OPEN |
| QUAL-H1 | Code Quality | 160+ uses of `any` type across composables/utils | OPEN |
| QUAL-H2 | Code Quality | useQuickActions.ts 753 lines with host-specific branching | OPEN |
| UX-H1 | UX | HomePage.vue oversized — template extracted, script section remains | ⚠️ PARTIAL |

### Medium (19 items)
| ID | Category | Title | Status |
|----|----------|-------|--------|
| ARCH-M2 | Architecture | backend.ts mixes concerns (669 lines) | OPEN |
| ARCH-M3 | Architecture | office-agents/ directory purpose unclear | OPEN |
| FUNC-M1 | Functionality | Tool count discrepancy across documentation | OPEN |
| FUNC-M2 | Functionality | No Outlook compose-time file attachment tool | OPEN |
| ERR-M2 | Error Handling | Raw console usage in 5+ files | ✅ FULL FIX |
| ERR-M3 | Error Handling | Frontend log forwarding to backend incomplete | ✅ FULL FIX |
| ERR-M4 | Error Handling | Rate limit retry exhaustion may calculate 0ms retry | ✅ FULL FIX |
| ERR-M5 | Error Handling | SSE read timeout doesn't abort upstream reader | ✅ FULL FIX |
| DUP-M1 | Duplication | VFS imports duplicated across tool files | OPEN |
| DUP-M2 | Duplication | eval_* tool boilerplate repeated 4 times | OPEN |
| UX-M1 | UX | No keyboard shortcut documentation | ✅ FULL FIX |
| UX-M2 | UX | ChatMessageList no virtualization for long conversations | OPEN |
| UX-M3 | UX | Missing i18n keys with hardcoded fallbacks (3 keys) | ✅ FULL FIX |
| UX-M4 | UX | Keyboard accessibility gaps in dropdowns | ✅ FULL FIX |
| QUAL-M1 | Code Quality | No unit tests for composables | OPEN |
| QUAL-M2 | Code Quality | powerpointImageRegistry memory leak potential | OPEN |
| QUAL-M3 | Code Quality | tokenManager JSON truncation breaks structure | OPEN |
| QUAL-M4 | Code Quality | Markdown CSS injection risk via custom color syntax | OPEN |
| QUAL-M5 | Code Quality | Backend models.js doesn't validate parsed env vars | OPEN |

### Low (12 items)
| ID | Category | Title | Status |
|----|----------|-------|--------|
| ARCH-L1 | Architecture | PowerPoint tool pattern inconsistency | OPEN |
| FUNC-L1 | Functionality | Excel chart creation limited to basic types | OPEN |
| FUNC-L2 | Functionality | No PowerPoint slide reorder tool | OPEN |
| ERR-L1 | Error Handling | Missing correlation ID frontend↔backend | ✅ FULL FIX |
| ERR-L2 | Error Handling | SSE stream error recovery lacks user guidance | ✅ FULL FIX |
| UX-L1 | UX | No dark mode toggle | ✅ FULL FIX |
| UX-L2 | UX | Quick action tooltips not i18n-ready | ✅ FULL FIX |
| DEAD-L1 | Dead Code | i18n key asymmetry (2 keys) | ✅ FULL FIX |
| DEAD-L2 | Dead Code | plotDigitizer route may become obsolete | OPEN |
| QUAL-L1 | Code Quality | Backend logs full response body for /api/chat/sync | OPEN |
| QUAL-L2 | Code Quality | credentialCrypto stores extractable key in localStorage | OPEN |
| DEAD-L3 | Dead Code | Unused credential utility exports (clearEncryptionKeys) | OPEN |

---

## Fix Batch — 2026-03-16 (UX & UI Fixes)

6 UX items fixed + 1 partial, plus DEAD-L1 i18n asymmetry.

### UX/UI Fixes Summary

| Item | Status | Files changed |
|------|--------|---------------|
| UX-H1 | ⚠️ PARTIAL FIX | `HomePage.vue`, new `OfflineBanner.vue`, `AuthErrorBanner.vue`, `SessionConfirmDialogs.vue` |
| UX-M1 | ✅ FULL FIX | `en.json`, `fr.json` (via UX-M3) |
| UX-M3 | ✅ FULL FIX | `en.json`, `fr.json` |
| UX-M4 | ✅ FULL FIX | `SingleSelect.vue` |
| UX-L1 | ✅ FULL FIX | `main.ts` |
| UX-L2 | ✅ FULL FIX | `en.json` (via DEAD-L1) |
| DEAD-L1 | ✅ FULL FIX | `en.json` |

### UX-H1 detail
`OfflineBanner.vue`, `AuthErrorBanner.vue`, `SessionConfirmDialogs.vue` extracted from `HomePage.vue`. All three use `useHomePageContext` (inject) for state/translations; `SessionConfirmDialogs` receives dialog-visibility state via props and emits cancel/confirm events. `HomePage.vue` reduced from 641 → 578 lines.

### UX-L1 detail (bug)
The dark mode toggle existed in `GeneralTab.vue` but was silently broken: `main.ts` listened to the `storage` DOM event, which only fires in *other* tabs — never in the same window that modified localStorage. Replaced with `useStorage(localStorageKey.darkMode, false)` from `@vueuse/core`, which is reactive to same-window writes.

### UX-M4 detail
`SingleSelect.vue` (the shared dropdown used throughout the app) now supports:
- `ArrowDown` / `ArrowUp`: navigate options with visual focus highlight
- `Enter`: select the focused option
- `Escape`: close without selecting
- Focus initializes on the currently selected item when dropdown opens

---

## Fix Batch — 2026-03-16 (ERR + TS Fixes)

All 5 CRITICAL items and all 6 ERR items (M2–M5, L1–L2) fixed. Pre-existing TypeScript errors (8 items) also resolved.

### ERR Fixes Summary

| Item | Status | Files changed |
|------|--------|---------------|
| ERR-C1 | ✅ FULL FIX | `backend/src/routes/chat.js` |
| ERR-C2 | ✅ FULL FIX | `backend/src/routes/chat.js` |
| ERR-C3 | ✅ FULL FIX | `frontend/src/composables/useAgentLoop.ts` |
| ERR-C4 | ✅ FULL FIX | `frontend/src/utils/officeAction.ts` |
| RACE-C1 | ✅ FULL FIX | `useSessionManager.ts`, `HomePage.vue` |
| ERR-M2 | ✅ FULL FIX | `sandbox.ts`, `lockdown.ts`, `useOfficeSelection.ts`, `BuiltinPromptsTab.vue`, `PromptsTab.vue` |
| ERR-M3 | ✅ FULL FIX | `logger.ts`, `main.ts` |
| ERR-M4 | ✅ FULL FIX | `backend/src/services/llmClient.js` |
| ERR-M5 | ✅ FULL FIX | `backend/src/routes/chat.js` |
| ERR-L1 | ✅ FULL FIX | `backend.ts`, `backend/src/server.js` |
| ERR-L2 | ✅ FULL FIX | `types/chat.ts`, `useAgentLoop.ts`, `ChatMessageList.vue` |

### TypeScript Errors Fixed (pre-existing)

| Error | File | Fix |
|-------|------|-----|
| TS6133 unused `nextTick` | `useAgentLoop.ts` | Removed import |
| TS2551 `getSelectedDataAsync` | `useDocumentUndo.ts` | Added `as any` cast |
| TS6133 unused `TContext` | `common.ts` | Added `@ts-ignore` with phantom generic comment |
| TS2345 traffic type mismatch | `credentialCrypto.ts` | Changed to `logService.debug(string)` |
| TS6133 unused `buildExecuteWrapper` | `outlookTools.ts` | Removed import |
| TS6133 unused `getErrorMessage` | `wordTools.ts` | Removed import |
| TS2353 `minItems` not in ToolProperty | `types/index.ts` | Added `minItems?/maxItems?` to interface |
| TS6133 unused `redlineEnabled` | `wordTrackChanges.ts` | Renamed to `_redlineEnabled` |

---

*See CHANGELOG.md for full version history.*
