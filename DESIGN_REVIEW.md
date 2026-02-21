# Design Review & Code Audit

**Date**: 2026-02-21
**Scope**: KickOffice architecture, security, code quality, and technical debt.

---

## 1. Executive Summary

The KickOffice add-in architecture (Vue 3 + Vite frontend, Express backend proxy) has undergone a successful refactoring cycle. Core UX bottlenecks have been addressed: streaming agent responses, persistent chat history, token pruning, and i18n externalization.

This audit identifies **35 issues** across frontend and backend, organized by severity:
- **CRITICAL**: 3 issues (security & correctness)
- **HIGH**: 6 issues (stability & data integrity)
- **MEDIUM**: 16 issues (code quality & architecture)
- **LOW**: 10 issues (polish & optimization)

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

#### C1. LiteLLM credentials stored in plain localStorage
- **File**: `frontend/src/api/backend.ts:79-80`
- **Issue**: User API keys (`litellmUserKey`) and emails are stored unencrypted in localStorage.
- **Impact**: If browser is compromised (XSS, malicious extension), credentials can be extracted.
- **Fix**: Use sessionStorage with short TTL, or encrypt credentials at rest with a derived key.

#### C2. reasoning_effort default value 'none' is invalid
- **File**: `backend/src/config/models.js:11,52`
- **Issue**: `reasoningEffort` defaults to `'none'` which is NOT a valid OpenAI API value. Valid values are `'low'`, `'medium'`, `'high'`.
- **Impact**: When tools are used with GPT-5 models and `reasoningEffort='none'`, the API returns empty responses. Line 83 guards against sending `'none'`, but the default assignment creates confusion.
- **Fix**: Remove `'none'` as default. Use `undefined` or omit the parameter entirely when reasoning is not needed.
- **Evidence**:
  ```javascript
  // Line 11 - env default is 'none'
  reasoningEffort: process.env.MODEL_STANDARD_REASONING_EFFORT || 'none',
  // Line 52 - fallback is 'none'
  const reasoningEffort = isGpt5Model(modelId) ? (modelConfig.reasoningEffort || 'none') : undefined
  ```

#### C3. Silent JSON parse failure in agent tool arguments
- **File**: `frontend/src/composables/useAgentLoop.ts:231`
- **Issue**: `try { toolArgs = JSON.parse(toolCall.function.arguments) } catch {}` — if parsing fails, `toolArgs` is empty object `{}`.
- **Impact**: Malformed tool call arguments from LLM are silently swallowed, causing tools to execute with wrong/missing parameters.
- **Fix**: Log parse failures and surface error to user; consider aborting tool execution on malformed input.

---

### HIGH (H1-H6) — Should fix soon

#### H1. XSS risk via v-html directive
- **File**: `frontend/src/components/chat/MarkdownRenderer.vue:2`
- **Issue**: Uses `v-html` with DOMPurify sanitization, but relies on correct configuration.
- **Impact**: If sanitization is bypassed (misconfiguration, DOMPurify vulnerability), arbitrary HTML/JS can execute.
- **Fix**: Audit DOMPurify config; consider CSP headers; add tests for XSS payloads.

#### H2. User credentials exposure in error logs
- **Files**: `backend/src/routes/chat.js:101,215`, `backend/src/routes/image.js:41`
- **Issue**: `errorText` from upstream API may contain request headers (including `X-User-Key`, `X-User-Email`). These are logged to console.
- **Impact**: Credentials could leak to log aggregation systems.
- **Fix**: Sanitize `errorText` before logging; redact known sensitive header names.

#### H3. Chat route validation code duplication
- **File**: `backend/src/routes/chat.js:19-66, 135-179`
- **Issue**: Both `/api/chat` and `/api/chat/sync` contain identical validation logic (messages, temperature, maxTokens, tools, ChatGPT checks, reasoning checks).
- **Impact**: Maintenance burden; risk of divergent behavior when one route is updated but not the other.
- **Fix**: Extract validation into shared middleware or helper function.

#### H4. Tool storage signature reset causes data loss
- **File**: `frontend/src/utils/toolStorage.ts:51-64`
- **Issue**: When tool definitions change (add/remove tool), the signature hash changes, causing all user tool preferences to reset silently.
- **Impact**: User loses custom tool enable/disable state without warning.
- **Fix**: Migrate existing preferences instead of resetting; notify user of changes.

#### H5. Race condition in agent loop during abort
- **File**: `frontend/src/composables/useAgentLoop.ts:156-159, 246-249`
- **Issue**: If user aborts during tool execution, tool result may be partially added to history, corrupting conversation state.
- **Impact**: Inconsistent message history; potential errors on next agent iteration.
- **Fix**: Ensure atomic state updates; rollback partial changes on abort.

#### H6. Empty API key defaults to empty string
- **File**: `backend/src/config/models.js:3`
- **Issue**: `LLM_API_KEY = process.env.LLM_API_KEY || ''` — server starts even if critical API key is missing.
- **Impact**: Requests fail at runtime instead of startup; harder to diagnose misconfiguration.
- **Fix**: Throw error at startup if `LLM_API_KEY` is not set in production.

---

### MEDIUM (M1-M16) — Should address

#### M1. Verbose debug logging with REMOVE_ME tags
- **Files**: `backend/src/routes/chat.js:8`, `frontend/src/composables/useOfficeInsert.ts:9`
- **Issue**: Production code contains debug tags `[KO-VERBOSE-CHAT][REMOVE_ME]` intended to be temporary.
- **Fix**: Remove or gate behind environment flag.

#### M2. Stream write error handling missing
- **File**: `backend/src/routes/chat.js:114-125`
- **Issue**: `res.write(chunk)` doesn't check for backpressure or write failures. If client disconnects mid-stream, errors may occur.
- **Fix**: Check `res.write()` return value; handle `drain` event.

#### M3. No validation of upstream response structure
- **File**: `backend/src/routes/chat.js:221-231`
- **Issue**: Sync endpoint returns `res.json(data)` without validating that `data` has expected shape (`choices`, `message`, etc.).
- **Impact**: Malformed upstream response is passed directly to client.
- **Fix**: Validate response structure; return 502 on unexpected format.

#### M4. Empty messages array allowed
- **File**: `backend/src/routes/chat.js:28-29`
- **Issue**: Validation checks `Array.isArray(messages)` but not `messages.length > 0`.
- **Impact**: Empty array sent to LLM API causes error or unexpected behavior.
- **Fix**: Add `messages.length === 0` check.

#### M5. Message field structure not validated
- **File**: `backend/src/routes/chat.js:20,136`
- **Issue**: No validation that each message has required fields (`role`, `content`).
- **Impact**: Invalid message structure forwarded to upstream API.
- **Fix**: Add schema validation for message objects.

#### M6. No messages array size limit
- **File**: `backend/src/routes/chat.js:28`
- **Issue**: While body parser limits to 4MB, there's no limit on `messages.length`.
- **Impact**: Request with thousands of messages could consume excessive memory/processing.
- **Fix**: Add `messages.length` limit (e.g., 200 messages max).

#### M7. No request timeout for Express handlers
- **File**: `backend/src/server.js`
- **Issue**: Individual API calls have timeouts, but no overall handler timeout exists.
- **Impact**: Slow operations could hang Express request indefinitely.
- **Fix**: Add request timeout middleware.

#### M8. Rate limit env var validation missing
- **File**: `backend/src/config/env.js:5-8`
- **Issue**: `parseInt()` on env vars doesn't validate result. If set to `"invalid"`, returns `NaN`.
- **Impact**: Rate limiter breaks silently.
- **Fix**: Validate `parseInt` result; throw on `NaN`.

#### M9. Excessive prop drilling in useAgentLoop
- **File**: `frontend/src/composables/useAgentLoop.ts:25-58`
- **Issue**: `UseAgentLoopOptions` interface has 34 parameters, tightly coupling composable to component structure.
- **Impact**: Hard to test, maintain, and refactor.
- **Fix**: Group related options into sub-objects; consider state management pattern.

#### M10. Type safety issues with `any` types
- **Files**: `frontend/src/pages/HomePage.vue:388`, `frontend/src/composables/useAgentLoop.ts:127,165`
- **Issue**: Multiple `any` type assertions bypass TypeScript type checking.
- **Impact**: Runtime errors not caught at compile time.
- **Fix**: Define proper interfaces; avoid `as any`.

#### M11. Silent clipboard failures
- **File**: `frontend/src/composables/useImageActions.ts:56,62,93`
- **Issue**: Empty catch blocks for clipboard operations silently swallow errors.
- **Impact**: User unaware when clipboard copy fails.
- **Fix**: Log errors; show user notification on failure.

#### M12. No HTTPS enforcement
- **File**: `backend/src/server.js`
- **Issue**: No HTTPS enforcement or HSTS headers in production.
- **Impact**: Credentials transmitted over insecure channels.
- **Fix**: Add HSTS headers; enforce HTTPS in production.

#### M13. TextDecoder stream not flushed
- **File**: `backend/src/routes/chat.js:118`
- **Issue**: `decoder.decode(value, { stream: true })` used but final `decoder.decode(undefined)` not called to flush remaining bytes.
- **Impact**: Potential truncation of final chunk in multi-byte character scenarios.
- **Fix**: Call `decoder.decode()` after loop exits.

#### M14. Unsafe JSON handling in settings
- **File**: `frontend/src/pages/SettingsPage.vue:962`, `frontend/src/utils/savedPrompts.ts:13-15`
- **Issue**: `JSON.parse` with inadequate structure validation. Array existence checked but element types not validated.
- **Impact**: Malformed data could cause runtime errors or data corruption.
- **Fix**: Add schema validation for parsed objects.

#### M15. No LLM API service abstraction
- **Files**: `backend/src/routes/chat.js:88-97,202-211`, `backend/src/routes/image.js:25-37`
- **Issue**: Each route directly constructs HTTP requests to LLM API. No centralized abstraction.
- **Impact**: Error handling, retry logic, and API contract management duplicated.
- **Fix**: Create `llmClient` service module.

#### M16. Timeout values hardcoded in routes
- **Files**: `backend/src/routes/chat.js:14-17`, `backend/src/routes/image.js:9-11`
- **Issue**: Timeout values (120s, 300s) defined in route handlers, not centralized config.
- **Fix**: Move to `config/models.js` alongside model definitions.

---

### LOW (L1-L10) — Nice to have

#### L1. Hardcoded French strings in quick actions
- **File**: `frontend/src/pages/HomePage.vue:251,259,267`
- **Issue**: Draft mode quick action prefixes contain hardcoded French text.
- **Fix**: Move to i18n keys.

#### L2. Missing i18n fallback pattern
- **File**: `frontend/src/pages/SettingsPage.vue:23`
- **Issue**: `{{ $t("settings") || "Settings" }}` uses JS fallback instead of i18n fallback mechanism.
- **Fix**: Configure i18n fallback locale properly.

#### L3. Inconsistent error message text
- **File**: `backend/src/routes/chat.js:34,151`
- **Issue**: Different error messages for same condition (`"Unknown model tier"` vs `"Invalid model tier"`).
- **Fix**: Standardize error messages.

#### L4. User credential format not validated
- **File**: `backend/src/middleware/auth.js:17-27`
- **Issue**: Email format and key length not validated.
- **Fix**: Add basic format validation for email and minimum key length.

#### L5. Stream cleanup race condition
- **File**: `backend/src/routes/chat.js:114-125`
- **Issue**: If client disconnects while streaming, both stream error handler and outer catch block might execute.
- **Fix**: Add connection state tracking.

#### L6. Inefficient array operations
- **Files**: `frontend/src/composables/useAgentLoop.ts:198`, `frontend/src/pages/SettingsPage.vue:934-936`
- **Issue**: Unnecessary `filter(Boolean)` calls; double array search with `findIndex` then index access.
- **Fix**: Prevent undefined insertion; use `find()` instead.

#### L7. Modal state not reset on cancel
- **File**: `frontend/src/pages/SettingsPage.vue:944-946`
- **Issue**: `cancelEdit()` doesn't reset `editingPrompt` ref; old data persists.
- **Fix**: Clear `editingPrompt` on cancel.

#### L8. Unnecessary regex compilation
- **File**: `frontend/src/utils/markdown.ts:39`
- **Issue**: `new RegExp()` created on every render in `cleanContent()`.
- **Fix**: Define as module constant.

#### L9. Image base64 splitting unsafe
- **File**: `frontend/src/composables/useImageActions.ts:100`
- **Issue**: `imageSrc.split(',')[1]` assumes single comma in data URL.
- **Fix**: Use regex match for data URL parsing.

#### L10. Models/health endpoints unrated
- **File**: `backend/src/server.js:62-63`
- **Issue**: `/api/models` and `/health` endpoints have no rate limiting.
- **Fix**: Add basic rate limiting for reconnaissance protection.

---

## 4. Build & Environment Warnings

### B1. JavaScript chunk size warnings (Deferred)
- **Status**: Known issue, non-blocking
- **Issue**: Vite build reports chunks exceeding 500kB warning threshold.
- **Mitigation**: Consider `manualChunks` configuration and `defineAsyncComponent` for code-splitting when bundle size impacts load performance.

### B2. Node.js version constraint
- **Status**: Documented
- **Current**: `package.json` specifies `>=20.0.0`
- **Recommendation**: Consider aligning with LTS versions (20.19+ or 22.x).

### B3. E2E tests missing
- **Status**: Technical debt
- **Issue**: No automated E2E tests for Office Host interactions.
- **Impact**: Increased manual validation time per release.
- **Note**: Office.js runtime complexity makes this challenging; consider Playwright for web-only flows.

---

## 5. Tracking Matrix

| ID  | Severity | Category    | Status   | File(s)                          |
|-----|----------|-------------|----------|----------------------------------|
| C1  | CRITICAL | Security    | OPEN     | backend.ts                       |
| C2  | CRITICAL | Correctness | OPEN     | models.js                        |
| C3  | CRITICAL | Correctness | OPEN     | useAgentLoop.ts                  |
| H1  | HIGH     | Security    | OPEN     | MarkdownRenderer.vue             |
| H2  | HIGH     | Security    | OPEN     | chat.js, image.js                |
| H3  | HIGH     | Quality     | OPEN     | chat.js                          |
| H4  | HIGH     | UX          | OPEN     | toolStorage.ts                   |
| H5  | HIGH     | Stability   | OPEN     | useAgentLoop.ts                  |
| H6  | HIGH     | Config      | OPEN     | models.js                        |
| M1  | MEDIUM   | Quality     | OPEN     | chat.js, useOfficeInsert.ts      |
| M2  | MEDIUM   | Stability   | OPEN     | chat.js                          |
| M3  | MEDIUM   | Validation  | OPEN     | chat.js                          |
| M4  | MEDIUM   | Validation  | OPEN     | chat.js                          |
| M5  | MEDIUM   | Validation  | OPEN     | chat.js                          |
| M6  | MEDIUM   | Validation  | OPEN     | chat.js                          |
| M7  | MEDIUM   | Stability   | OPEN     | server.js                        |
| M8  | MEDIUM   | Config      | OPEN     | env.js                           |
| M9  | MEDIUM   | Architecture| OPEN     | useAgentLoop.ts                  |
| M10 | MEDIUM   | Quality     | OPEN     | HomePage.vue, useAgentLoop.ts    |
| M11 | MEDIUM   | UX          | OPEN     | useImageActions.ts               |
| M12 | MEDIUM   | Security    | OPEN     | server.js                        |
| M13 | MEDIUM   | Correctness | OPEN     | chat.js                          |
| M14 | MEDIUM   | Validation  | OPEN     | SettingsPage.vue, savedPrompts.ts|
| M15 | MEDIUM   | Architecture| OPEN     | chat.js, image.js                |
| M16 | MEDIUM   | Architecture| OPEN     | chat.js, image.js                |
| L1  | LOW      | i18n        | OPEN     | HomePage.vue                     |
| L2  | LOW      | i18n        | OPEN     | SettingsPage.vue                 |
| L3  | LOW      | Quality     | OPEN     | chat.js                          |
| L4  | LOW      | Validation  | OPEN     | auth.js                          |
| L5  | LOW      | Stability   | OPEN     | chat.js                          |
| L6  | LOW      | Performance | OPEN     | useAgentLoop.ts, SettingsPage.vue|
| L7  | LOW      | UX          | OPEN     | SettingsPage.vue                 |
| L8  | LOW      | Performance | OPEN     | markdown.ts                      |
| L9  | LOW      | Correctness | OPEN     | useImageActions.ts               |
| L10 | LOW      | Security    | OPEN     | server.js                        |

---

## 6. Recommended Priority Order

1. **Immediate** (this sprint): C1, C2, C3
2. **Next sprint**: H1-H6
3. **Backlog**: M1-M16, then L1-L10 as capacity allows

---

*Last updated: 2026-02-21*
