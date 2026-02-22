# Design Review & Code Audit — v2

**Date**: 2026-02-22
**Scope**: Full architecture, security, code quality, UX, and technical debt analysis of the KickOffice codebase.

---

## 1. Executive Summary

KickOffice is a production-grade Microsoft Office add-in (Word, Excel, PowerPoint, Outlook) powered by an AI backend proxy. The architecture — Vue 3 + Vite frontend, Express.js backend, OpenAI-compatible LLM API — is sound and well-structured after the v1 audit cycle.

The previous audit (v1, 2026-02-21) identified and resolved **38 issues**. This v2 review is a fresh analysis of the current codebase state, identifying **28 new issues** that were not covered or have emerged since:

- **CRITICAL**: 3 issues (correctness, configuration)
- **HIGH**: 5 issues (architecture, data integrity, documentation accuracy)
- **MEDIUM**: 10 issues (code quality, UX, maintainability)
- **LOW**: 7 issues (polish, performance, resilience)
- **BUILD**: 3 warnings (testing, tooling, CI)

### Previous Audit Status

All 38 issues from the v1 audit (3 CRITICAL, 6 HIGH, 16 MEDIUM, 10 LOW, 3 BUILD) have been resolved and verified. Key wins from v1:
- Session storage for credentials (C1)
- Reasoning effort parameter fix (C2)
- JSON parse failure handling (C3)
- XSS protection, credential sanitization, rate limiting, HSTS
- LLM API service abstraction, validation middleware, E2E test infrastructure

---

## 2. New Issues by Severity

### CRITICAL (C1–C3) — Requires immediate action

#### C1. Agent max iterations setting is silently ignored
- **Files**: `frontend/src/composables/useAgentLoop.ts:214`, `frontend/src/pages/SettingsPage.vue:664-676`
- **Issue**: The Settings UI allows `agentMaxIterations` between 1 and 100, but `useAgentLoop.ts` hardcodes `Math.min(Number(agentMaxIterations.value) || 10, 10)` — capping the effective maximum at **10 iterations** regardless of user setting.
- **Impact**: The settings UI misleads users. Setting iterations to 25 (the default) or 100 has no effect; the agent always stops at 10. For complex multi-step tasks this is a significant functional limitation.
- **Fix**: Replace the hardcoded `10` with the setting value: `Math.min(Number(agentMaxIterations.value) || 10, agentMaxIterations.value)`, or more simply use the setting value directly with a reasonable cap (e.g. 50).

#### C2. `.env.example` still contains invalid `reasoning_effort=none`
- **File**: `backend/.env.example:27`
- **Issue**: `MODEL_STANDARD_REASONING_EFFORT=none` is present in the example env file. The v1 audit (C2) fixed the code default from `'none'` to `undefined`, but the `.env.example` was not updated. Since `models.js:21` reads `process.env.MODEL_STANDARD_REASONING_EFFORT || undefined`, the string `'none'` is truthy and passes through, ultimately sending `reasoning_effort: 'none'` to the OpenAI API.
- **Impact**: Any deployment following the `.env.example` template will send an invalid API parameter, causing empty responses when tools are used with GPT-5 models.
- **Fix**: Change line 27 to `MODEL_STANDARD_REASONING_EFFORT=` (empty, meaning "omit the parameter") and add a comment explaining valid values.

#### C3. Quick actions bypass loading/abort state management
- **Files**: `frontend/src/composables/useAgentLoop.ts:512-617`
- **Issue**: `applyQuickAction()` calls `chatStream` and `runAgentLoop` without setting `loading.value = true` or creating an `AbortController`. The stop button doesn't work during quick actions.
- **Impact**: Users cannot abort a running quick action. If a quick action is triggered while another is running, concurrent stream writes corrupt the chat history. The UI gives no visual feedback that a quick action is in progress.
- **Fix**: Wrap the quick action execution in the same loading/abort pattern as `sendMessage()`.

---

### HIGH (H1–H5) — Should fix soon

#### H1. Chat history grows unbounded in localStorage
- **File**: `frontend/src/pages/HomePage.vue:174`
- **Issue**: `useStorage<DisplayMessage[]>('chatHistory_word', [])` has no size limit. Each message (especially tool results with full document content) can be several KB. Long agent sessions can produce hundreds of messages.
- **Impact**: `localStorage` has a ~5-10 MB limit per origin. When exceeded, `QuotaExceededError` is thrown silently (Vue `useStorage` swallows it), causing data loss or broken persistence.
- **Fix**: Implement a max history size (e.g. 100 messages or 1 MB). Prune oldest messages when the limit is reached. Consider a warning toast when approaching the limit.

#### H2. Token manager uses character count instead of token count
- **File**: `frontend/src/utils/tokenManager.ts:3`
- **Issue**: `MAX_CONTEXT_CHARS = 100_000` uses character count as a proxy for token count. The function name `prepareMessagesForContext` and the variable name suggest token management, but no tokenization occurs.
- **Impact**: For CJK text (1-2 chars per token), the budget is ~2x too generous, risking API token limit errors. For English (~4 chars per token), 100k chars ≈ 25k tokens, which is acceptable but imprecise. Additionally, `getMessageContentLength()` does not account for `tool_calls` JSON payload in assistant messages, leading to budget underestimation in tool-heavy conversations.
- **Fix**: Either rename to `MAX_CONTEXT_CHARS` (already done) and document the approximation, or integrate a lightweight tokenizer (e.g. `tiktoken` WASM build) for accurate counting. At minimum, include `tool_calls` serialized length in the budget calculation.

#### H3. Validation error message references invalid `none` value
- **File**: `backend/src/middleware/validate.js:189`
- **Issue**: Error message says `"temperature is only supported for GPT-5 models when reasoning effort is none"`. The value `'none'` is no longer valid (removed in v1 C2 fix). The correct phrasing is "when reasoning effort is not set".
- **Impact**: Confusing error message for API consumers. Could lead developers to set `reasoning_effort=none` based on this message.
- **Fix**: Change to `"temperature is not supported for GPT-5 models when reasoning effort is enabled"`.

#### H4. `chatSync` function is exported but unused (dead code)
- **File**: `frontend/src/api/backend.ts:228-249`
- **Issue**: The `chatSync()` function is still fully implemented and exported, but the agent loop was migrated to use `chatStream()` with `onToolCallDelta` instead. No code path calls `chatSync`.
- **Impact**: Dead code increases maintenance burden and confusion. The function also contains a useless `try { ... } catch (error) { throw error }` wrapper.
- **Fix**: Remove the `chatSync` function, its interface `ChatSyncOptions`, and the unused import in any consuming file. If kept as future API surface, mark it with a `@deprecated` comment.

#### H5. README model defaults do not match actual code
- **Files**: `README.md:47-49`, `backend/src/config/models.js:17,26,33`, `backend/.env.example:23,30,37`
- **Issue**: README documents default models as `gpt-5.2` and `gpt-image-1.5`, but actual code defaults are `gpt-5.1` and `gpt-image-1`. The `.env.example` correctly shows `gpt-5.1` and `gpt-image-1`.
- **Impact**: Documentation mismatch causes confusion during deployment.
- **Fix**: Update README to match actual defaults (`gpt-5.1`, `gpt-image-1`).

---

### MEDIUM (M1–M10) — Should address

#### M1. No confirmation dialog for "New Chat"
- **File**: `frontend/src/pages/HomePage.vue:490-497`
- **Issue**: `startNewChat()` immediately clears all history without confirmation. One accidental click destroys the entire conversation.
- **Impact**: Users lose conversation data with no recovery path.
- **Fix**: Add a confirmation prompt before clearing, or implement undo functionality.

#### M2. `processChat` has useless try/catch wrapper
- **File**: `frontend/src/composables/useAgentLoop.ts:406-410`
- **Issue**: `try { await runAgentLoop(messages, modelTier) } catch (error) { throw error }` catches and immediately re-throws without transformation.
- **Impact**: Dead code that adds noise. The catch block does nothing useful.
- **Fix**: Remove the try/catch, call `runAgentLoop` directly.

#### M3. Inconsistent quick action type definitions
- **Files**: `frontend/src/pages/HomePage.vue:271,305`, `frontend/src/composables/useAgentLoop.ts:83-86`
- **Issue**: `outlookQuickActions` is a plain array, `excelQuickActions` is a `computed`, `powerPointQuickActions` is a plain array, and `wordQuickActions` is a plain array. The `AgentLoopActions` interface expects `outlookQuickActions` as `Ref<OutlookQuickAction[]>` (optional) and `powerPointQuickActions` as a plain array. Then `HomePage.vue:469` wraps `outlookQuickActions` in `computed(() => outlookQuickActions)` to satisfy the ref requirement.
- **Impact**: Unnecessary complexity and inconsistent patterns make the code harder to understand.
- **Fix**: Standardize all quick action arrays to the same type (either all `computed` or all plain arrays) and update the interface accordingly.

#### M4. Missing accessibility (ARIA) attributes
- **Files**: `frontend/src/pages/SettingsPage.vue`, `frontend/src/components/chat/ChatInput.vue`
- **Issue**: Settings tabs lack `role="tablist"` / `role="tab"` / `role="tabpanel"` attributes. Tool checkboxes have no `aria-label`. The chat input textarea has no `aria-label`.
- **Impact**: Screen reader users cannot properly navigate the UI. May not meet corporate accessibility requirements (WCAG 2.1 AA).
- **Fix**: Add ARIA roles and labels to interactive elements.

#### M5. Built-in prompts serialization uses fragile interpolation
- **Files**: `frontend/src/pages/SettingsPage.vue:968-971,986-987,997-1001`
- **Issue**: Custom built-in prompts are serialized with `${language}` and `${text}` placeholders but the editor UI uses `[LANGUAGE]` and `[TEXT]`. The bidirectional translation works but is fragile. If a user types a literal `${text}` in the editor, it would be interpreted as a replacement variable during deserialization.
- **Impact**: Edge case data corruption. The dual-format approach is a maintenance risk.
- **Fix**: Standardize on a single placeholder format (prefer `[LANGUAGE]`/`[TEXT]` since that's what the user sees) and use it for both serialization and deserialization.

#### M6. `useOfficeInsert.ts` uses `any` for message parameters
- **File**: `frontend/src/composables/useOfficeInsert.ts:20-21,145,153`
- **Issue**: `shouldTreatMessageAsImage: (message: any) => boolean`, `copyMessageToClipboard(message: any)`, `insertMessageToDocument(message: any)` all use `any` type.
- **Impact**: No TypeScript type safety. Callers could pass incorrect objects without compile-time errors.
- **Fix**: Use `DisplayMessage` type from `@/types/chat`.

#### M7. No TypeScript on backend
- **Files**: All `backend/src/*.js` files
- **Issue**: The entire backend is plain JavaScript with no JSDoc type annotations or TypeScript. Request bodies, model configs, and API responses have no type definitions.
- **Impact**: As the project grows, refactoring becomes risky. Runtime type errors are caught late. IDE support (autocomplete, refactoring) is limited.
- **Fix**: Consider migrating to TypeScript incrementally, starting with `config/models.js` and `middleware/validate.js` which handle the most complex data shapes. At minimum, add JSDoc types.

#### M8. Missing `Content-Type` enforcement on backend
- **File**: `backend/src/server.js:87`
- **Issue**: `express.json({ limit: '4mb' })` silently ignores requests without `Content-Type: application/json`. A POST with `text/plain` body will pass through with an empty `req.body`.
- **Impact**: Potentially confusing error messages when content type is wrong (validation says "messages array is required" instead of "invalid content type").
- **Fix**: Add middleware that rejects POST requests without `application/json` content type.

#### M9. Docker Compose `version` field is deprecated
- **File**: `docker-compose.yml:1`
- **Issue**: `version: "3.8"` is deprecated in Docker Compose v2. It's silently ignored but generates warnings.
- **Impact**: Build-time warning noise.
- **Fix**: Remove the `version` line.

#### M10. `officeRichText.ts` markdown parser has `html: true`
- **File**: `frontend/src/utils/officeRichText.ts:9`
- **Issue**: The Office markdown parser has `html: true`, allowing raw HTML pass-through before DOMPurify sanitization. While the sanitization step (line 326-333) properly filters tags, the `ALLOWED_ATTR` list includes `style` attribute, which can carry CSS injection payloads (e.g., `expression()` in older IE, `url()` for data exfiltration).
- **Impact**: Low risk in modern browsers/Office webviews, but the `style` attribute in the allow list is broader than needed. Since the input comes from AI responses (not direct user input), exploitation requires prompt injection.
- **Fix**: Consider using DOMPurify's `FORBID_ATTR: ['style']` or a CSS sanitizer for the `style` attribute values if stricter security is needed.

---

### LOW (L1–L7) — Nice to have

#### L1. `scrollToBottom` couples to Tailwind CSS class
- **File**: `frontend/src/pages/HomePage.vue:393`
- **Issue**: `container.querySelectorAll('.group')` uses the Tailwind `.group` utility class as a DOM selector. A Tailwind version upgrade or class rename would silently break auto-scroll.
- **Fix**: Use a `data-*` attribute or dedicated CSS class for message identification.

#### L2. Custom `debounce` reimplemented despite `@vueuse/core` dependency
- **File**: `frontend/src/main.ts:15-25`
- **Issue**: A custom debounce function is written inline, but `@vueuse/core` (already imported in the same file) provides `useDebounceFn`.
- **Fix**: Replace with `useDebounceFn` from VueUse.

#### L3. Backend polling interval is fixed at 30 seconds
- **File**: `frontend/src/pages/HomePage.vue:539`
- **Issue**: `window.setInterval(checkBackend, 30000)` polls regardless of user activity. No exponential backoff when backend is down, no pause when the add-in is backgrounded.
- **Fix**: Use `document.visibilityState` check to pause polling when the tab is hidden. Consider backoff when offline.

#### L4. `backend.ts:148` non-null assertion on response body
- **File**: `frontend/src/api/backend.ts:148`
- **Issue**: `res.body!.getReader()` uses a non-null assertion. While a 200 response always has a body, a defensive check would prevent crashes on unexpected responses.
- **Fix**: Add a null check: `if (!res.body) throw new Error('Empty response body')`.

#### L5. Undocumented frontend environment variables
- **Files**: `frontend/src/composables/useOfficeInsert.ts:9`, `frontend/src/api/backend.ts:7`
- **Issue**: `VITE_VERBOSE_LOGGING` and `VITE_REQUEST_TIMEOUT_MS` are used in the frontend but not documented in the README's environment variables section or in any `.env.example`.
- **Fix**: Add these to the frontend environment variables table in README.

#### L6. SettingsPage status label has hardcoded French fallback
- **File**: `frontend/src/pages/SettingsPage.vue:75`
- **Issue**: `$t('litellmCredentialsMissing') || 'Statut'` contains a hardcoded French fallback string. This pattern was identified and fixed in the v1 audit (L2) for other instances, but this one was missed.
- **Fix**: Remove the `|| 'Statut'` fallback; the i18n system handles missing keys.

#### L7. Inconsistent error message format in agent tool execution
- **File**: `frontend/src/composables/useAgentLoop.ts:311,323`
- **Issue**: Parse errors return `"Error: malformed tool arguments — JSON parse failed"` while execution errors return `"Error: ${err.message}"`. The first format includes context, the second doesn't include the tool name.
- **Fix**: Standardize error messages to include tool name and error type: `"Error in ${toolName}: ${description}"`.

---

## 3. Build & Environment Warnings

### B1. No unit test infrastructure
- **Status**: Not addressed
- **Issue**: No unit test framework (vitest, jest) is configured. Critical business logic in `tokenManager.ts`, `toolStorage.ts`, `validate.js`, `models.js`, and `buildChatBody()` has zero unit test coverage. Only E2E tests via Playwright exist.
- **Fix**: Add vitest to the frontend. Key candidates for unit testing:
  - `tokenManager.ts:prepareMessagesForContext` — budget calculation and message pruning
  - `toolStorage.ts:migrateToolPreferences` — migration logic
  - `validate.js:validateChatRequest` — input validation
  - `models.js:buildChatBody` — request construction

### B2. No linting or formatting configuration
- **Status**: Not addressed
- **Issue**: No ESLint configuration file (`.eslintrc`, `eslint.config.js`) or Prettier config exists in the repository. No pre-commit hooks (husky, lint-staged) are configured.
- **Fix**: Add ESLint + Prettier with a pre-commit hook. This prevents style drift and catches common errors before commit.

### B3. No CI pipeline for automated testing
- **Status**: Partial
- **Issue**: Only a `bump-version.yml` GitHub Action exists. No automated test execution (lint, build, E2E) on pull requests. Code can be merged without passing any checks.
- **Fix**: Add a CI workflow that runs `npm run build` (both frontend and backend), `npm run lint` (once configured), and `npm run test:e2e` on every PR.

---

## 4. Tracking Matrix

| ID  | Severity | Category       | Status | File(s)                                   |
|-----|----------|----------------|--------|-------------------------------------------|
| C1  | CRITICAL | Correctness    | OPEN   | useAgentLoop.ts, SettingsPage.vue         |
| C2  | CRITICAL | Configuration  | OPEN   | .env.example, models.js                   |
| C3  | CRITICAL | UX / Stability | OPEN   | useAgentLoop.ts                           |
| H1  | HIGH     | Data Integrity | OPEN   | HomePage.vue                              |
| H2  | HIGH     | Architecture   | OPEN   | tokenManager.ts                           |
| H3  | HIGH     | Correctness    | OPEN   | validate.js                               |
| H4  | HIGH     | Quality        | OPEN   | backend.ts                                |
| H5  | HIGH     | Documentation  | OPEN   | README.md, models.js                      |
| M1  | MEDIUM   | UX             | OPEN   | HomePage.vue                              |
| M2  | MEDIUM   | Quality        | OPEN   | useAgentLoop.ts                           |
| M3  | MEDIUM   | Architecture   | OPEN   | HomePage.vue, useAgentLoop.ts             |
| M4  | MEDIUM   | Accessibility  | OPEN   | SettingsPage.vue, ChatInput.vue           |
| M5  | MEDIUM   | Quality        | OPEN   | SettingsPage.vue                          |
| M6  | MEDIUM   | Type Safety    | OPEN   | useOfficeInsert.ts                        |
| M7  | MEDIUM   | Architecture   | OPEN   | backend/src/*.js                          |
| M8  | MEDIUM   | Validation     | OPEN   | server.js                                 |
| M9  | MEDIUM   | Config         | OPEN   | docker-compose.yml                        |
| M10 | MEDIUM   | Security       | OPEN   | officeRichText.ts                         |
| L1  | LOW      | Quality        | OPEN   | HomePage.vue                              |
| L2  | LOW      | Quality        | OPEN   | main.ts                                   |
| L3  | LOW      | Performance    | OPEN   | HomePage.vue                              |
| L4  | LOW      | Resilience     | OPEN   | backend.ts                                |
| L5  | LOW      | Documentation  | OPEN   | README.md                                 |
| L6  | LOW      | i18n           | OPEN   | SettingsPage.vue                          |
| L7  | LOW      | Quality        | OPEN   | useAgentLoop.ts                           |
| B1  | BUILD    | Testing        | OPEN   | (no test framework)                       |
| B2  | BUILD    | Tooling        | OPEN   | (no lint config)                          |
| B3  | BUILD    | CI/CD          | OPEN   | .github/workflows/                        |

---

## 5. Recommended Implementation Priority

### Phase 1 — Immediate (critical correctness)
1. **C1** — Fix agent max iterations cap (5 min fix, high user impact)
2. **C2** — Fix `.env.example` reasoning_effort value (1 min fix, prevents broken deployments)
3. **C3** — Add loading/abort state to quick actions (prevents data corruption)
4. **H5** — Update README model defaults (documentation accuracy)

### Phase 2 — Short-term (stability & data integrity)
5. **H1** — Implement chat history size limits
6. **H3** — Fix validation error message
7. **H4** — Remove dead `chatSync` code
8. **M1** — Add new chat confirmation dialog
9. **M2** — Remove useless try/catch

### Phase 3 — Medium-term (quality & maintainability)
10. **H2** — Improve token budget calculation
11. **M3** — Standardize quick action types
12. **M5** — Unify prompt serialization format
13. **M6** — Replace `any` types with `DisplayMessage`
14. **M8** — Add Content-Type enforcement
15. **M9** — Remove deprecated Docker Compose version
16. **L6** — Remove French fallback string

### Phase 4 — Backlog (polish & infrastructure)
17. **M4** — Accessibility improvements
18. **M7** — Backend TypeScript migration
19. **M10** — Review `style` attribute in DOMPurify config
20. **B1** — Add vitest unit test infrastructure
21. **B2** — Add ESLint + Prettier
22. **B3** — Add CI pipeline
23. **L1–L5, L7** — Minor quality and resilience improvements

---

## 6. Architecture Strengths (preserved from v1)

The following architectural decisions are sound and should be maintained:

- **Proxy pattern**: API keys never reach the client. All LLM communication goes through the Express backend.
- **Modular composables**: `useAgentLoop`, `useAgentPrompts`, `useOfficeInsert`, `useImageActions` provide clear separation of concerns.
- **Host-aware design**: Tool definitions, prompts, quick actions, and insertion logic are all host-specific (Word/Excel/PowerPoint/Outlook).
- **Centralized LLM client**: `services/llmClient.js` is the single point of contact with the upstream API.
- **Validation middleware**: `validate.js` provides thorough input validation with clear error messages.
- **Security posture**: DOMPurify strict allowlists, HSTS in production, credential sanitization, rate limiting, session storage for user credentials.
- **i18n framework**: Full English and French locale support with translation-aware built-in prompts.
- **Tool preference migration**: `toolStorage.ts` gracefully handles tool definition changes without resetting user preferences.

---

*Last updated: 2026-02-22*
