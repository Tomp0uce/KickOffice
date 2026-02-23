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

## 7. New Feature Issues & Implementation Plans (v2.1)

The following three issues were identified during functional testing and require targeted implementation work. Each includes a detailed root-cause analysis and step-by-step implementation plan.

### F1. Quick actions strip images, formatting, and non-text elements from documents

- **Severity**: HIGH
- **Category**: Data Integrity / UX
- **Status**: OPEN
- **Scope**: ALL text-modifying quick actions (translate, polish, academic, summary, proofread, concise) across ALL hosts (Word, Outlook, PowerPoint)
- **Files**: `frontend/src/composables/useOfficeSelection.ts`, `frontend/src/composables/useAgentLoop.ts:615-677`, `frontend/src/api/common.ts`, `frontend/src/utils/wordTools.ts`, `frontend/src/utils/constant.ts`

#### Problem

When ANY text-modifying quick action is triggered (translate, polish, academic, concise, etc.), the current flow is:

1. `getOfficeSelection()` reads the selected content as **plain text only** (`range.load('text')` for Word, `getAsync(CoercionType.Text)` for Outlook, `getPowerPointSelection()` for PowerPoint).
2. The plain-text selection is sent to the LLM for processing.
3. The processed result is displayed in chat. When the user clicks "Replace", it is inserted back via `insertResult()` / `insertFormattedResult()` which overwrites the entire selection.

This flow **destroys all non-text elements** in the original selection:
- **Inline images** (embedded photos, diagrams, logos)
- **Bullet points and list formatting** (indentation levels, custom bullets)
- **Rich formatting** (tables, text boxes, shapes, drawings)
- **Embedded objects** (charts, OLE objects)
- **Text styling** (colors, highlights, custom fonts applied to specific runs)

This affects **all hosts** (Word, Outlook, PowerPoint) and **all text-modifying quick actions**, not just translation.

#### Root Cause

- `useOfficeSelection.ts:56-61`: Word selection reads only `.text`, discarding all rich content.
- `useOfficeSelection.ts:19`: Outlook uses `CoercionType.Text`, losing all HTML formatting and images.
- `common.ts:14-17`: `insertResult()` uses `range.insertText(normalizedResult, 'Replace')` which replaces the entire selection with plain text.
- All quick action prompts (`constant.ts`) produce text-only output with no awareness of non-text elements.

#### Implementation Plan

**Approach**: For all text-modifying quick actions across all hosts, capture the rich content (HTML) of the selection, extract only the text for LLM processing while preserving non-text elements with placeholders, then reassemble the rich content with the LLM's processed text before insertion.

**Step 1 — Create a `richContentPreserver` utility** (new file: `frontend/src/utils/richContentPreserver.ts`)
- Utility to handle HTML parsing and reassembly:
  - `extractTextFromHtml(html: string)`: Parses HTML, replaces `<img>`, `<svg>`, `<table>`, and other non-text elements with unique `{{PRESERVE_N}}` placeholders. Returns `{ cleanText: string, fragments: Map<string, string> }`.
  - `reassembleHtml(processedText: string, fragments: Map<string, string>)`: Takes LLM-processed text and replaces `{{PRESERVE_N}}` placeholders with original HTML fragments.
  - Handles nested structures: images inside table cells, styled lists, etc.

**Step 2 — Add HTML-aware selection readers** (`useOfficeSelection.ts`)
- **Word**: Add `getWordSelectionAsHtml()` using `range.getHtml()` to capture full rich content.
- **Outlook**: Add `getOutlookMailBodyAsHtml()` using `getAsync(CoercionType.Html)` instead of `CoercionType.Text`.
- **PowerPoint**: Verify current `getPowerPointSelection()` behavior — if it strips formatting, add HTML-aware variant.
- Each returns raw HTML that can be fed to the `richContentPreserver`.

**Step 3 — Modify quick action flow for ALL text-modifying actions** (`useAgentLoop.ts:615-677`)
- In `applyQuickAction()`, for all non-agent quick actions on Word/Outlook/PowerPoint:
  1. Get selection as HTML via the new HTML-aware readers.
  2. Call `extractTextFromHtml(html)` to get clean text + preserved fragments.
  3. Send only the clean text to the LLM (with placeholder instruction in prompt).
  4. When LLM responds, call `reassembleHtml(llmResponse, fragments)` to produce final HTML.
  5. Store the reassembled HTML as `richHtml` on the `DisplayMessage` for insertion.

**Step 4 — Update insertion to use rich HTML when available** (`useOfficeInsert.ts`)
- Check if the message has `richHtml` content.
- **Word**: Use `range.insertHtml(richHtml, 'Replace')` instead of `range.insertText()`.
- **Outlook**: Use `setSelectedDataAsync(richHtml, { coercionType: CoercionType.Html })`.
- **PowerPoint**: Use HTML-aware insertion if available, otherwise fall back to text.
- Preserve the existing plain-text path as fallback when no `richHtml` is available.

**Step 5 — Update ALL text-modifying prompts** (`constant.ts`)
- Add to ALL quick action prompts (translate, polish, academic, summary, concise, proofread):
  ```
  If the text contains preservation placeholders like {{PRESERVE_1}}, {{PRESERVE_2}}, etc., keep them exactly in their original position. These represent images and formatting elements that must not be removed or modified.
  ```

**Step 6 — Add `richHtml` field to `DisplayMessage`** (`types/chat.ts`)
- Add optional `richHtml?: string` field to `DisplayMessage` interface.
- This stores the reassembled HTML with preserved non-text elements, ready for insertion.

**Estimated Impact**: Prevents data loss across all text-modifying operations for all Office hosts. Critical for professional workflows where documents contain images, tables, styled lists, and other rich elements.

---

### F2. Outlook "Reply" quick action produces low-quality, wrong-language responses

- **Severity**: HIGH
- **Category**: UX / Quality
- **Status**: OPEN
- **Files**: `frontend/src/utils/constant.ts:323-342`, `frontend/src/pages/HomePage.vue:272-305`, `frontend/src/composables/useAgentLoop.ts:587-599`, `frontend/src/i18n/locales/fr.json:268`, `frontend/src/i18n/locales/en.json:263`

#### Problem

The Outlook "Reply" quick action (`reply` key) currently operates in **draft mode**: it pre-fills the chat input with a short prefix (`"Rédige une réponse à ce mail en disant que : "` / `"Draft a reply to this email saying that: "`) and waits for the user to type a brief intent. The user then presses Send, which triggers the normal `sendMessage()` flow.

The result is a poor-quality reply because:

1. **Insufficient context**: The email thread body is only included as `[Email body: "..."]` context appended by `sendMessage()` via `useSelectedText`. There is no structured analysis of the previous email's tone, formality level, or language.
2. **Wrong language**: The reply prompt (`outlookBuiltInPrompt.reply`) contains language detection instructions, but these are in the **built-in prompt** which is only used when the action is NOT in draft mode. In draft mode, the user's message goes through `processChat()` with the generic `agentPrompt()`, which uses `replyLanguage` setting (defaulting to `Français`) instead of detecting the email language.
3. **No tone analysis**: The generic agent prompt does not analyze the email thread to determine if the tone should be formal (external client) or casual (internal colleague).
4. **No message length calibration**: Short user input like "dis oui" should generate a brief reply, while "explique en détail pourquoi on ne peut pas faire ça" should generate a longer response. There's no instruction for this.

#### Root Cause

- Draft mode (`mode: 'draft'` in `HomePage.vue:301`) bypasses the `outlookBuiltInPrompt.reply` prompt entirely. It only pre-fills the chat input, then `sendMessage()` uses the generic Outlook agent prompt.
- The generic Outlook agent prompt (`useAgentPrompts.ts:120-136`) has basic reply language instructions but no structured email analysis framework.
- The user's brief reply intent (e.g., "dis oui") provides no tone, length, or language guidance.

#### Implementation Plan

**Approach**: Replace the draft mode with a two-phase "smart reply" flow that first analyzes the email thread, then generates a calibrated response.

**Step 1 — Change reply action from draft mode to a new "smart reply" mode** (`HomePage.vue:297-304`)
- Change `mode: 'draft'` to `mode: 'smart-reply'` (new mode).
- Remove the `prefix` property; it's no longer needed as a pre-fill.
- Instead, the smart reply will: (1) open a small inline prompt for the user to describe their reply intent, (2) automatically fetch and analyze the email body.

**Step 2 — Create a dedicated smart reply system prompt** (`constant.ts`)
- Add a new `outlookBuiltInPrompt.smartReply` prompt with a comprehensive analysis framework:

```typescript
smartReply: {
  system: (language: string) =>
    `You are an expert email assistant specialized in drafting context-aware, natural email replies.

BEFORE drafting the reply, you MUST analyze the email thread and determine these parameters:

## Analysis Parameters (internal, do not output these)
1. **Language**: Detect the dominant language of the email thread. Reply in that EXACT language. Ignore interface language "${language}".
2. **Tone**: Determine the formality level from the email context:
   - FORMAL: External clients, senior management, first contact, legal/compliance (use "Monsieur/Madame", "Dear", "Cordialement", "Best regards")
   - SEMI-FORMAL: Known colleagues, recurring contacts (use first name + polite register)
   - CASUAL: Close team members, internal quick exchanges (direct, concise, friendly)
3. **Reply length**: Calibrate based on:
   - User's input length and specificity (short input = short reply, detailed input = detailed reply)
   - Original email complexity (a 3-line email doesn't warrant a 15-line reply)
   - Match the approximate length of the original sender's style
4. **Key points to address**: Identify which points from the original email need to be addressed in the reply.
5. **Sender relationship**: Infer from greeting style, sign-off, and language register.

## Reply Generation Rules
- Address ALL points raised in the original email that relate to the user's intent.
- Match the detected tone and formality level precisely.
- Use appropriate greetings and sign-offs for the detected tone level.
- Keep the reply proportional to the original email's length and the user's intent complexity.
- OUTPUT ONLY the reply text, ready to send. No meta-commentary, no "Here is your reply".
- Do NOT include a subject line ("Objet:", "Subject:"). Start directly with the greeting.
- The user's input is their INTENT for the reply (what they want to say), not the literal text to send.`,
  user: (text: string, language: string) =>
    `## Email thread to reply to:
${text}

## User's reply intent:
[REPLY_INTENT]

Draft the reply now following all analysis rules above.
${GLOBAL_STYLE_INSTRUCTIONS}`,
}
```

**Step 3 — Implement the smart-reply flow in `useAgentLoop.ts`**
- Add handling for `mode === 'smart-reply'` in `applyQuickAction()`:
  1. Pre-fill the chat input with a contextual prompt (keep the existing prefix UX for user intent input).
  2. When the user presses Send, intercept the flow (detect that a smart-reply was pending).
  3. Fetch the full email body via `getOfficeSelection()`.
  4. Build the smart reply prompt by:
     - Using `outlookBuiltInPrompt.smartReply.system(lang)` as the system prompt.
     - Replacing `[REPLY_INTENT]` in the user prompt with the user's typed intent.
     - Replacing `${text}` with the full email body.
  5. Stream the response directly (non-agent mode, no tools needed).

**Step 4 — Add email metadata enrichment** (`useOfficeSelection.ts` or `outlookTools.ts`)
- When building the smart reply context, also fetch:
  - `getEmailSender()` — to determine the sender's name and relationship.
  - `getEmailDate()` — for temporal context.
  - `getEmailSubject()` — for topic context.
- Prepend this metadata to the email thread text:
  ```
  From: {sender}
  Date: {date}
  Subject: {subject}

  {body}
  ```
- This gives the LLM better context for tone and formality analysis.

**Step 5 — Update i18n strings** (`fr.json`, `en.json`)
- Update the reply pre-prompt to be more descriptive:
  - FR: `"Décrivez brièvement ce que vous voulez répondre : "`
  - EN: `"Briefly describe what you want to reply: "`
- Add new i18n keys for smart reply status messages:
  - `smartReplyAnalyzing`: `"Analyzing email tone and language..."` / `"Analyse du ton et de la langue du mail..."`

**Step 6 — Alternative: Keep draft mode but inject smart prompt at send time**
- If the UX of draft mode (user types in chat input) is preferred, then instead of changing the mode:
  - Keep `mode: 'draft'` and `prefix`.
  - In `sendMessage()`, detect that the message starts with the reply prefix.
  - If so, strip the prefix, treat the rest as the reply intent, and route to the smart reply flow from Step 3.
  - This preserves the existing UX while improving the backend prompt quality.

**Estimated Impact**: Dramatically improves reply quality by ensuring correct language detection, tone matching, and proportional response length. The current implementation produces replies that feel generic and robotic; this fix makes them contextually aware and natural.

---

### F3. Excel agent mode processes cells one-by-one with individual LLM calls — extremely inefficient

- **Severity**: HIGH
- **Category**: Performance / Cost
- **Status**: OPEN
- **Files**: `frontend/src/composables/useAgentLoop.ts:224-418`, `frontend/src/utils/excelTools.ts:99-141` (`setCellValue`), `frontend/src/composables/useAgentPrompts.ts:92`

#### Problem

When a user asks the Excel agent to translate, transform, or modify multiple cells (e.g., "translate all cells in column A from French to English"), the agent loop processes each cell individually:

1. The LLM generates one `setCellValue` tool call per cell.
2. Each tool call is executed, and the result is appended to the message history.
3. The updated history is sent back to the LLM for the next iteration.
4. The LLM generates the next `setCellValue` call for the next cell.

For 50 cells, this produces **50+ LLM round-trips** (each iteration often yields only 1-2 tool calls), consuming enormous amounts of tokens and taking a very long time. The `setCellValue` tool already supports multi-cell writes via JSON 2D arrays, but the LLM is not instructed to batch operations.

#### Root Cause

1. **Agent prompt lacks batching instructions** (`useAgentPrompts.ts:92`): The Excel agent prompt says `"Tool First"` and `"use fillFormulaDown when applying same formula across rows"` but has no instruction to batch `setCellValue` calls for multi-cell text transformations.
2. **No batch-oriented tool**: While `setCellValue` accepts a 2D array, there's no dedicated "batch transform" tool that would let the LLM send all transformations in a single call.
3. **LLM behavior**: Without explicit batching instructions, the model defaults to the most "reliable" pattern: one cell at a time, verify, next cell. This is safe but extremely wasteful.
4. **Agent loop design**: The loop processes tool calls sequentially within an iteration, but the LLM typically generates only 1-3 tool calls per response for cell modifications.

#### Implementation Plan

**Approach**: Two-pronged fix: (1) add a dedicated batch cell operation tool, (2) update the agent prompt to strongly prefer batching.

**Step 1 — Add `batchSetCellValues` tool** (`excelTools.ts`)
- Create a new tool designed for bulk cell transformations:

```typescript
batchSetCellValues: {
  name: 'batchSetCellValues',
  category: 'write',
  description:
    'Set values for multiple individual cells in a single operation. This is much more efficient than calling setCellValue repeatedly. Use this tool whenever you need to modify more than 2 cells. Provide an array of {address, value} pairs.',
  inputSchema: {
    type: 'object',
    properties: {
      cells: {
        type: 'array',
        description: 'Array of cell updates. Each item has an "address" (A1 notation) and a "value" (the new cell content).',
        items: {
          type: 'object',
          properties: {
            address: { type: 'string', description: 'Cell address in A1 notation (e.g., "A1")' },
            value: { type: 'string', description: 'New value for the cell' },
          },
          required: ['address', 'value'],
        },
      },
    },
    required: ['cells'],
  },
  executeExcel: async (context, args) => {
    const { cells } = args
    const sheet = context.workbook.worksheets.getActiveWorksheet()
    for (const cell of cells) {
      const range = sheet.getRange(cell.address)
      const num = Number(cell.value)
      range.values = [[isNaN(num) ? cell.value : num]]
    }
    await context.sync()
    return `Successfully updated ${cells.length} cells`
  },
}
```

**Step 2 — Add `batchProcessCells` tool for LLM-powered transformations** (`excelTools.ts`)
- This is an alternative/complementary approach: a tool that reads a range, sends all values for transformation in one shot, and writes back:

```typescript
batchProcessRange: {
  name: 'batchProcessRange',
  category: 'write',
  description:
    'Read all values from a range, apply the same transformation to each cell, and write the results back in one operation. Use this for translations, text cleanup, formatting, or any uniform transformation across multiple cells. You provide the range address and the transformed values as a 2D array matching the range dimensions.',
  inputSchema: {
    type: 'object',
    properties: {
      address: {
        type: 'string',
        description: 'Range address to process (e.g., "A1:A50", "B2:D10")',
      },
      values: {
        type: 'array',
        description: 'A 2D array of new values matching the range dimensions. Example: [["translated1"],["translated2"]] for a single-column range.',
        items: {
          type: 'array',
          items: { type: 'string' },
        },
      },
    },
    required: ['address', 'values'],
  },
  executeExcel: async (context, args) => {
    const { address, values } = args
    const sheet = context.workbook.worksheets.getActiveWorksheet()
    const range = sheet.getRange(address)
    range.values = values
    await context.sync()
    return `Successfully updated range ${address} (${values.length} rows × ${values[0]?.length || 0} columns)`
  },
}
```

**Step 3 — Update the Excel agent prompt** (`useAgentPrompts.ts:92`)
- Add strong batching instructions to the Excel agent prompt:

```typescript
const excelAgentPrompt = (lang: string) => `# Role
You are a highly skilled Microsoft Excel Expert Agent...

# Guidelines
1. **Tool First**
2. **Read First**: Always read the data before modifying it.
3. **BATCH OPERATIONS (CRITICAL)**: When modifying multiple cells:
   - NEVER use setCellValue in a loop to modify cells one by one.
   - For text transformations (translate, clean, format, etc.): use getSelectedCells or getWorksheetData to read ALL values first, then process ALL transformations in your response, and use batchSetCellValues or batchProcessRange to write ALL results in ONE tool call.
   - For formula application: use fillFormulaDown instead of calling insertFormula per row.
   - Example: To translate 50 cells, read all 50 values, translate them all in one response, then write all 50 translated values using batchProcessRange.
4. **Accuracy**
5. **Conciseness**
6. **Language**: You must communicate entirely in ${lang}.
7. **Formula locale**: ${excelFormulaLanguageInstruction()}
8. **Formula duplication**: use fillFormulaDown when applying same formula across rows.`
```

**Step 4 — Add `ExcelToolName` entries** (`excelTools.ts:23-63`)
- Add `'batchSetCellValues'` and `'batchProcessRange'` to the `ExcelToolName` union type.

**Step 5 — Consider token budget for large ranges**
- For very large ranges (100+ cells), the LLM might not be able to process all values in a single response due to output token limits.
- Add a note in the tool description: `"For ranges larger than 100 cells, process in chunks of 50-100 cells at a time."`.
- Alternatively, implement a chunking strategy in the tool itself that splits the range into manageable batches.

**Step 6 — Update tool preferences migration** (`toolStorage.ts`)
- Ensure the new tool names are added to the default enabled set so existing users see them immediately without needing to manually enable them in settings.

**Estimated Impact**: Reduces token consumption by **10-50x** for multi-cell operations. A 50-cell translation that currently takes 50+ LLM round-trips would be reduced to 2-3 (read, process, write). This directly impacts API cost and user wait time.

---

## 8. Updated Tracking Matrix (v2.1 additions)

| ID  | Severity | Category        | Status | File(s)                                                |
|-----|----------|-----------------|--------|--------------------------------------------------------|
| F1  | HIGH     | Data Integrity  | OPEN   | useOfficeSelection.ts, useAgentLoop.ts, useOfficeInsert.ts, richContentPreserver.ts, constant.ts, chat.ts |
| F2  | HIGH     | UX / Quality    | OPEN   | constant.ts, HomePage.vue, useAgentLoop.ts, i18n/*     |
| F3  | HIGH     | Performance     | OPEN   | excelTools.ts, useAgentPrompts.ts, useAgentLoop.ts     |

### Updated Implementation Priority

Insert between Phase 1 and Phase 2 as **Phase 1.5 — Functional quality**:
1. **F2** — Outlook smart reply (highest user-facing quality impact, prompt-only change for quick win)
2. **F3** — Excel batch cell processing (highest cost/performance impact, new tools + prompt change)
3. **F1** — Translation image preservation (requires OOXML/HTML parsing, most complex implementation)

---

*Last updated: 2026-02-23*
