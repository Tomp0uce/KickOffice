# Design Review & Code Audit

**Date**: 2026-02-21
**Scope**: KickOffice architecture, security, code quality, and technical debt.

---

## 1. Executive Summary

The KickOffice add-in architecture (Vue 3 + Vite frontend, Express backend proxy) has undergone a successful refactoring cycle. Core UX bottlenecks have been addressed: streaming agent responses, persistent chat history, token pruning, and i18n externalization.

This audit identified **38 issues** across frontend and backend, organized by severity:
- **CRITICAL**: 3 issues (security & correctness) ✅ ALL FIXED
- **HIGH**: 6 issues (stability & data integrity) ✅ ALL FIXED
- **MEDIUM**: 16 issues (code quality & architecture) ✅ ALL FIXED
- **LOW**: 10 issues (polish & optimization) ✅ ALL FIXED
- **BUILD**: 3 warnings (performance, config, testing) ✅ ALL FIXED

---

## 2. Recently Resolved Issues (Fixed)

The following major issues have been successfully addressed:

- ✅ **Tool state desynchronization (Feature Toggle)**: Settings UI now correctly filters tools used dynamically by the agent.
- ✅ **Missing streaming in agent loop**: Sync calls replaced by `chatStream`, providing real-time feedback including during tool execution.
- ✅ **Conversation history persistence**: History is now saved persistently via `localStorage` (isolated per Office Host).
- ✅ **Context pruning (Token management)**: Implemented intelligent context window ensuring token limits are respected while preserving tool call integrity.
- ✅ **Hardcoded translations**: "Thought process" label externalized to i18n. Missing tooltips for Excel/PPT/Outlook quick actions added.
- ✅ **Developer syntax exposure**: Replaced obscure `${text}` syntax with intuitive `[TEXT]` placeholders in settings UI.
- ✅ **Auto-scroll UX**: Added automatic scrolling that keeps the start of AI-generated messages visible during long responses.

---

## 3. Open Issues by Severity

### CRITICAL (C1-C3) — Requires immediate action

#### C1. LiteLLM credentials stored in plain localStorage ✅ FIXED
- **File**: `frontend/src/api/backend.ts:79-80`, `frontend/src/pages/SettingsPage.vue:662-663`
- **Issue**: User API keys (`litellmUserKey`) and emails are stored unencrypted in localStorage.
- **Impact**: If browser is compromised (XSS, malicious extension), credentials can be extracted.
- **Fix applied**: Migrated to `sessionStorage` — credentials now cleared automatically when the browser session ends. Both the read side (`backend.ts`) and write side (`SettingsPage.vue` via `useStorage`) updated.

#### C2. reasoning_effort default value 'none' is invalid ✅ FIXED
- **File**: `backend/src/config/models.js:11,52`
- **Issue**: `reasoningEffort` defaults to `'none'` which is NOT a valid OpenAI API value. Valid values are `'low'`, `'medium'`, `'high'`.
- **Impact**: When tools are used with GPT-5 models and `reasoningEffort='none'`, the API returns empty responses. Line 83 guards against sending `'none'`, but the default assignment creates confusion.
- **Fix applied**: Replaced `|| 'none'` with `|| undefined` on lines 11 and 52. The `canUseSamplingParams` check updated from `=== 'none'` to `!reasoningEffort`. The redundant `!== 'none'` guard on line 83 also removed.

#### C3. Silent JSON parse failure in agent tool arguments ✅ FIXED
- **File**: `frontend/src/composables/useAgentLoop.ts:231`
- **Issue**: `try { toolArgs = JSON.parse(toolCall.function.arguments) } catch {}` — if parsing fails, `toolArgs` is empty object `{}`.
- **Impact**: Malformed tool call arguments from LLM are silently swallowed, causing tools to execute with wrong/missing parameters.
- **Fix applied**: Parse failure now logs the error (tool name + raw arguments) to console and pushes a `tool` error message back into `currentMessages` with `continue` to skip execution — preventing the tool from running with empty/incorrect parameters.

---

### HIGH (H1-H6) — Should fix soon

#### H1. XSS risk via v-html directive ✅ FIXED
- **File**: `frontend/src/components/chat/MarkdownRenderer.vue:2`
- **Issue**: Uses `v-html` with DOMPurify sanitization, but relies on correct configuration.
- **Impact**: If sanitization is bypassed (misconfiguration, DOMPurify vulnerability), arbitrary HTML/JS can execute.
- **Fix applied**: Added strict `ALLOWED_TAGS` and `ALLOWED_ATTR` allowlists to DOMPurify configuration in `markdown.ts`. Disabled `ALLOW_DATA_ATTR` and `ALLOW_ARIA_ATTR` to minimize attack surface.

#### H2. User credentials exposure in error logs ✅ FIXED
- **Files**: `backend/src/routes/chat.js:101,215`, `backend/src/routes/image.js:41`
- **Issue**: `errorText` from upstream API may contain request headers (including `X-User-Key`, `X-User-Email`). These are logged to console.
- **Impact**: Credentials could leak to log aggregation systems.
- **Fix applied**: Added `sanitizeErrorText()` utility in `utils/http.js` that redacts known sensitive headers (`X-User-Key`, `X-User-Email`, `Authorization`, etc.) before logging. Applied to all error logging in chat and image routes.

#### H3. Chat route validation code duplication ✅ FIXED
- **File**: `backend/src/routes/chat.js:19-66, 135-179`
- **Issue**: Both `/api/chat` and `/api/chat/sync` contain identical validation logic (messages, temperature, maxTokens, tools, ChatGPT checks, reasoning checks).
- **Impact**: Maintenance burden; risk of divergent behavior when one route is updated but not the other.
- **Fix applied**: Extracted validation into `validateChatRequest()` function in `middleware/validate.js`. Both routes now use this shared validator, ensuring consistent behavior. Also added empty messages array check (M4 partial fix).

#### H4. Tool storage signature reset causes data loss ✅ FIXED
- **File**: `frontend/src/utils/toolStorage.ts:51-64`
- **Issue**: When tool definitions change (add/remove tool), the signature hash changes, causing all user tool preferences to reset silently.
- **Impact**: User loses custom tool enable/disable state without warning.
- **Fix applied**: Implemented `migrateToolPreferences()` function that preserves enabled state for existing tools, enables new tools by default, and silently drops removed tools. Logs migration info to console for debugging.

#### H5. Race condition in agent loop during abort ✅ FIXED
- **File**: `frontend/src/composables/useAgentLoop.ts:156-159, 246-249`
- **Issue**: If user aborts during tool execution, tool result may be partially added to history, corrupting conversation state.
- **Impact**: Inconsistent message history; potential errors on next agent iteration.
- **Fix applied**: Refactored tool execution loop to collect results in a temporary array before committing. Abort checks added before AND after each tool execution. On abort mid-loop, the incomplete assistant message is rolled back from `currentMessages` to prevent state corruption.

#### H6. Empty API key defaults to empty string ✅ FIXED
- **File**: `backend/src/config/models.js:3`
- **Issue**: `LLM_API_KEY = process.env.LLM_API_KEY || ''` — server starts even if critical API key is missing.
- **Impact**: Requests fail at runtime instead of startup; harder to diagnose misconfiguration.
- **Fix applied**: Added startup validation in `config/models.js`. In production (`NODE_ENV=production`), throws fatal error if `LLM_API_KEY` is not set. In development, logs a warning.

---

### MEDIUM (M1-M16) — Should address

#### M1. Verbose debug logging with REMOVE_ME tags ✅ FIXED
- **Files**: `backend/src/routes/chat.js:8`, `frontend/src/composables/useOfficeInsert.ts:9`
- **Issue**: Production code contains debug tags `[KO-VERBOSE-CHAT][REMOVE_ME]` intended to be temporary.
- **Fix applied**: Replaced with environment-gated `verboseLog()` function controlled by `VERBOSE_LOGGING=true` env var.

#### M2. Stream write error handling missing ✅ FIXED
- **File**: `backend/src/routes/chat.js:114-125`
- **Issue**: `res.write(chunk)` doesn't check for backpressure or write failures. If client disconnects mid-stream, errors may occur.
- **Fix applied**: Added backpressure handling with `res.write()` return value check, `drain` event wait, and client disconnection tracking.

#### M3. No validation of upstream response structure ✅ FIXED
- **File**: `backend/src/routes/chat.js:221-231`
- **Issue**: Sync endpoint returns `res.json(data)` without validating that `data` has expected shape (`choices`, `message`, etc.).
- **Fix applied**: Added validation for `data` object type, `choices` array presence, and `message` object structure with 502 responses on failure.

#### M4. Empty messages array allowed ✅ FIXED
- **File**: `backend/src/routes/chat.js:28-29`
- **Issue**: Validation checks `Array.isArray(messages)` but not `messages.length > 0`.
- **Fix applied**: Added `messages.length === 0` check in `validateChatRequest()`.

#### M5. Message field structure not validated ✅ FIXED
- **File**: `backend/src/routes/chat.js:20,136`
- **Issue**: No validation that each message has required fields (`role`, `content`).
- **Fix applied**: Added `validateMessage()` function that validates role (system/user/assistant/tool), content, tool_calls, and tool_call_id fields based on role.

#### M6. No messages array size limit ✅ FIXED
- **File**: `backend/src/routes/chat.js:28`
- **Issue**: While body parser limits to 4MB, there's no limit on `messages.length`.
- **Fix applied**: Added `MAX_MESSAGES = 200` constant and validation check.

#### M7. No request timeout for Express handlers ✅ FIXED
- **File**: `backend/src/server.js`
- **Issue**: Individual API calls have timeouts, but no overall handler timeout exists.
- **Fix applied**: Added request timeout middleware using `req.setTimeout()` with configurable `REQUEST_TIMEOUT_MS` env var (default 10 minutes).

#### M8. Rate limit env var validation missing ✅ FIXED
- **File**: `backend/src/config/env.js:5-8`
- **Issue**: `parseInt()` on env vars doesn't validate result. If set to `"invalid"`, returns `NaN`.
- **Fix applied**: Added `parsePositiveInt()` helper that throws on NaN or negative values.

#### M9. Excessive prop drilling in useAgentLoop ✅ FIXED
- **File**: `frontend/src/composables/useAgentLoop.ts:45-91`
- **Issue**: `UseAgentLoopOptions` interface has 34 parameters, tightly coupling composable to component structure.
- **Fix applied**: Refactored options into logical sub-interfaces: `AgentLoopRefs`, `AgentLoopModels`, `AgentLoopHost`, `AgentLoopSettings`, `AgentLoopActions`, `AgentLoopHelpers`. Updated `HomePage.vue` caller to use grouped structure.

#### M10. Type safety issues with `any` types ✅ FIXED
- **Files**: `frontend/src/pages/HomePage.vue:388`, `frontend/src/composables/useAgentLoop.ts:127,165`
- **Issue**: Multiple `any` type assertions bypass TypeScript type checking.
- **Fix applied**: Added `StreamResponse`, `StreamResponseChoice`, `AssistantMessage`, `ToolCall`, `ToolCallFunction` interfaces for proper typing of agent loop response handling.

#### M11. Silent clipboard failures ✅ FIXED
- **File**: `frontend/src/composables/useImageActions.ts:56,62,93`
- **Issue**: Empty catch blocks for clipboard operations silently swallow errors.
- **Fix applied**: Added `console.warn` logging for clipboard API failures (already present in useImageActions.ts, added to useOfficeInsert.ts).

#### M12. No HTTPS enforcement ✅ FIXED
- **File**: `backend/src/server.js`
- **Issue**: No HTTPS enforcement or HSTS headers in production.
- **Fix applied**: Added HSTS headers via Helmet config with `maxAge: 1 year`, `includeSubDomains`, and `preload` when `NODE_ENV=production`.

#### M13. TextDecoder stream not flushed ✅ FIXED
- **File**: `backend/src/routes/chat.js:118`
- **Issue**: `decoder.decode(value, { stream: true })` used but final `decoder.decode(undefined)` not called to flush remaining bytes.
- **Fix applied**: Added `decoder.decode()` call after stream loop to flush remaining bytes.

#### M14. Unsafe JSON handling in settings ✅ FIXED
- **File**: `frontend/src/pages/SettingsPage.vue:962`, `frontend/src/utils/savedPrompts.ts:13-15`
- **Issue**: `JSON.parse` with inadequate structure validation. Array existence checked but element types not validated.
- **Fix applied**: Added `isValidSavedPrompt()` type guard function that validates all required fields. Invalid items are logged and skipped.

#### M15. No LLM API service abstraction ✅ FIXED
- **Files**: `backend/src/routes/chat.js:88-97,202-211`, `backend/src/routes/image.js:25-37`
- **Issue**: Each route directly constructs HTTP requests to LLM API. No centralized abstraction.
- **Fix applied**: Created `services/llmClient.js` module with `chatCompletion()`, `imageGeneration()`, and `handleErrorResponse()` functions. Routes now use this centralized service.

#### M16. Timeout values hardcoded in routes ✅ FIXED
- **Files**: `backend/src/routes/chat.js:14-17`, `backend/src/routes/image.js:9-11`
- **Issue**: Timeout values (120s, 300s) defined in route handlers, not centralized config.
- **Fix applied**: Moved timeouts to `services/llmClient.js` with exported `TIMEOUTS` object and `getChatTimeoutMs()`, `getImageTimeoutMs()` functions.

---

### LOW (L1-L10) — Nice to have

#### L1. Hardcoded French strings in quick actions ✅ FIXED
- **File**: `frontend/src/pages/HomePage.vue:251,259,267`
- **Issue**: Draft mode quick action prefixes contain hardcoded French text.
- **Fix applied**: Moved French strings to i18n keys (`excelFormulaPrefix`, `excelTransformPrefix`, `excelHighlightPrefix`) in `fr.json` and `en.json`.

#### L2. Missing i18n fallback pattern ✅ FIXED
- **File**: `frontend/src/pages/SettingsPage.vue:23`
- **Issue**: `{{ $t("settings") || "Settings" }}` uses JS fallback instead of i18n fallback mechanism.
- **Fix applied**: Removed JS fallback, i18n fallback locale handles missing keys.

#### L3. Inconsistent error message text ✅ FIXED
- **File**: `backend/src/routes/chat.js:34,151`
- **Issue**: Different error messages for same condition (`"Unknown model tier"` vs `"Invalid model tier"`).
- **Fix applied**: Standardized to `"Invalid model tier"` throughout `validateChatRequest()`.

#### L4. User credential format not validated ✅ FIXED
- **File**: `backend/src/middleware/auth.js:17-27`
- **Issue**: Email format and key length not validated.
- **Fix applied**: Added `EMAIL_REGEX` validation and `MIN_KEY_LENGTH = 8` check in `ensureUserCredentials()`.

#### L5. Stream cleanup race condition ✅ FIXED
- **File**: `backend/src/routes/chat.js:114-125`
- **Issue**: If client disconnects while streaming, both stream error handler and outer catch block might execute.
- **Fix applied**: Added `clientDisconnected` flag tracking for proper stream cleanup (already done as part of M2 backpressure handling).

#### L6. Inefficient array operations ✅ FIXED
- **Files**: `frontend/src/composables/useAgentLoop.ts:198`, `frontend/src/pages/SettingsPage.vue:934-936`
- **Issue**: Unnecessary `filter(Boolean)` calls; double array search with `findIndex` then index access.
- **Fix applied**: Added index guards to prevent undefined insertion in tool_calls; used `find()` directly in SettingsPage.

#### L7. Modal state not reset on cancel ✅ FIXED
- **File**: `frontend/src/pages/SettingsPage.vue:944-946`
- **Issue**: `cancelEdit()` doesn't reset `editingPrompt` ref; old data persists.
- **Fix applied**: Added `editingPrompt.value = null` reset in `cancelEdit()`.

#### L8. Unnecessary regex compilation ✅ FIXED
- **File**: `frontend/src/composables/useImageActions.ts:10`
- **Issue**: Regex for think tag removal created on every call.
- **Fix applied**: Cached `THINK_TAG_REGEX` as module-level constant.

#### L9. Image base64 splitting unsafe ✅ FIXED
- **File**: `frontend/src/composables/useImageActions.ts:102,112`
- **Issue**: `imageSrc.split(',')[1]` assumes single comma in data URL.
- **Fix applied**: Used regex replacement `imageSrc.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '')` for safe extraction.

#### L10. Models/health endpoints unrated ✅ FIXED
- **File**: `backend/src/server.js:53-59,92-93`
- **Issue**: `/api/models` and `/health` endpoints have no rate limiting.
- **Fix applied**: Added `infoLimiter` with 120 requests per minute limit for `/health` and `/api/models` endpoints.

---

## 4. Build & Environment Warnings

### B1. JavaScript chunk size warnings ✅ FIXED
- **Status**: Resolved
- **Issue**: Vite build reports chunks exceeding 500kB warning threshold.
- **Fix applied**: Added `manualChunks` configuration in `vite.config.js` to split vendor dependencies into separate chunks (`vendor-vue`, `vendor-ui`, `vendor-utils`, `vendor-math`). Increased warning threshold to 600kB.

### B2. Node.js version constraint ✅ FIXED
- **Status**: Resolved
- **Previous**: `package.json` specified `>=20.0.0`
- **Fix applied**: Updated both frontend and backend `package.json` to `>=20.19.0 || >=22.0.0` to align with Node.js LTS versions.

### B3. E2E tests missing ✅ FIXED
- **Status**: Infrastructure in place
- **Issue**: No automated E2E tests for Office Host interactions.
- **Fix applied**: Added Playwright test infrastructure with `playwright.config.ts`, basic navigation tests in `e2e/navigation.spec.ts`, and test scripts (`test:e2e`, `test:e2e:ui`, `test:e2e:report`). Note: Office.js runtime interactions still require manual testing; Playwright tests cover web-only flows.

---

## 5. Tracking Matrix

| ID  | Severity | Category    | Status   | File(s)                          |
|-----|----------|-------------|----------|----------------------------------|
| C1  | CRITICAL | Security    | FIXED    | backend.ts, SettingsPage.vue     |
| C2  | CRITICAL | Correctness | FIXED    | models.js                        |
| C3  | CRITICAL | Correctness | FIXED    | useAgentLoop.ts                  |
| H1  | HIGH     | Security    | FIXED    | MarkdownRenderer.vue, markdown.ts |
| H2  | HIGH     | Security    | FIXED    | chat.js, image.js, http.js       |
| H3  | HIGH     | Quality     | FIXED    | chat.js, validate.js             |
| H4  | HIGH     | UX          | FIXED    | toolStorage.ts                   |
| H5  | HIGH     | Stability   | FIXED    | useAgentLoop.ts                  |
| H6  | HIGH     | Config      | FIXED    | models.js                        |
| M1  | MEDIUM   | Quality     | FIXED    | chat.js, useOfficeInsert.ts      |
| M2  | MEDIUM   | Stability   | FIXED    | chat.js                          |
| M3  | MEDIUM   | Validation  | FIXED    | chat.js                          |
| M4  | MEDIUM   | Validation  | FIXED    | validate.js                      |
| M5  | MEDIUM   | Validation  | FIXED    | validate.js                      |
| M6  | MEDIUM   | Validation  | FIXED    | validate.js                      |
| M7  | MEDIUM   | Stability   | FIXED    | server.js                        |
| M8  | MEDIUM   | Config      | FIXED    | env.js                           |
| M9  | MEDIUM   | Architecture| FIXED    | useAgentLoop.ts, HomePage.vue    |
| M10 | MEDIUM   | Quality     | FIXED    | useAgentLoop.ts                  |
| M11 | MEDIUM   | UX          | FIXED    | useOfficeInsert.ts               |
| M12 | MEDIUM   | Security    | FIXED    | server.js                        |
| M13 | MEDIUM   | Correctness | FIXED    | chat.js                          |
| M14 | MEDIUM   | Validation  | FIXED    | savedPrompts.ts                  |
| M15 | MEDIUM   | Architecture| FIXED    | services/llmClient.js            |
| M16 | MEDIUM   | Architecture| FIXED    | services/llmClient.js            |
| L1  | LOW      | i18n        | FIXED    | HomePage.vue, fr.json, en.json   |
| L2  | LOW      | i18n        | FIXED    | SettingsPage.vue                 |
| L3  | LOW      | Quality     | FIXED    | validate.js                      |
| L4  | LOW      | Validation  | FIXED    | auth.js                          |
| L5  | LOW      | Stability   | FIXED    | chat.js                          |
| L6  | LOW      | Performance | FIXED    | useAgentLoop.ts, SettingsPage.vue|
| L7  | LOW      | UX          | FIXED    | SettingsPage.vue                 |
| L8  | LOW      | Performance | FIXED    | useImageActions.ts               |
| L9  | LOW      | Correctness | FIXED    | useImageActions.ts               |
| L10 | LOW      | Security    | FIXED    | server.js                        |
| B1  | BUILD    | Performance | FIXED    | vite.config.js                   |
| B2  | BUILD    | Config      | FIXED    | package.json (frontend, backend) |
| B3  | BUILD    | Testing     | FIXED    | playwright.config.ts, e2e/       |

---

## 6. Recommended Priority Order

1. **Immediate** (this sprint): C1, C2, C3 ✅ DONE
2. **Next sprint**: H1-H6 ✅ DONE
3. **Backlog**: M1-M16 ✅ DONE, L1-L10 ✅ DONE
4. **Build & Environment**: B1-B3 ✅ DONE

**All 38 issues have been resolved (35 code issues + 3 build/environment warnings).**

---

*Last updated: 2026-02-21*
