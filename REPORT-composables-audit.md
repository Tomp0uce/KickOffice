# Code Design Review: Frontend Composables

**Date**: 2026-03-01
**Scope**: Deep audit of the 5 core frontend composables — security, logic bugs, code quality, architecture.

---

## 1. Executive Summary

This review covers the 5 composable files that form the core frontend logic of KickOffice:

| File | Lines | Role |
|------|-------|------|
| `useAgentLoop.ts` | ~830 | Agent execution loop, message sending, quick actions |
| `useAgentPrompts.ts` | ~180 | System/user prompt construction per host |
| `useImageActions.ts` | ~160 | Image copy/insert, think-tag parsing, display message helpers |
| `useOfficeInsert.ts` | ~200 | Document insertion (Word/Excel/PowerPoint/Outlook) + clipboard |
| `useOfficeSelection.ts` | ~140 | Office API text/HTML selection retrieval |

**Findings**: 41 issues total across 5 severity levels.

| Severity | Count | Files Affected |
|----------|-------|----------------|
| CRITICAL | 2 | useAgentLoop.ts |
| HIGH | 7 | useAgentLoop.ts, useImageActions.ts |
| MEDIUM | 22 | All files |
| LOW | 10 | All files |

### Top 3 Risk Areas

1. **Prompt injection** — Document/email content is interpolated into LLM prompts without sanitization (C1, C2, M9)
2. **Race conditions & state corruption** — Concurrent `sendMessage` calls and stale `lastIndex` references (H1, H2)
3. **Silent error swallowing** — Multiple `catch` blocks discard errors without logging (M16, L10)

---

## 2. Issues by Severity

### CRITICAL (C1–C2) — Requires immediate action

#### C1. Prompt Injection via Unsanitized User Selection Injected into Prompts

- **File**: `frontend/src/composables/useAgentLoop.ts:613-616`
- **Category**: Security (Prompt Injection)
- **Details**: The Office selection text is directly interpolated into the user message with only a label prefix and double quotes:
  ```typescript
  fullMessage += `\n\n[${selectionLabel}: "${selectedText}"]`
  ```
  A malicious document whose text contains `"]` followed by prompt injection instructions can break out of the bracket framing and inject arbitrary instructions. The same pattern appears at line 611 where `selectedText` is interpolated into the image generation prompt template. There is no escaping or sanitization of `selectedText` before it is placed into the prompt context. This is a classic indirect prompt injection surface: a user opens a document crafted by an attacker and the document content is injected verbatim into the LLM prompt.
- **Recommendation**: Wrap user-controlled content in a robust delimiter that cannot be spoofed (e.g., XML CDATA-style wrapping, base64 encoding, or a unique boundary token). At minimum, escape the closing delimiter characters.

#### C2. Prompt Injection via Quick Action — Selection Text Passed Directly to LLM

- **File**: `frontend/src/composables/useAgentLoop.ts:753, 761`
- **Category**: Security (Prompt Injection)
- **Details**: In `applyQuickAction`, `textForLlm` (derived from `selectedText` or `richContext.cleanText`) is passed directly into `action.user(textForLlm, lang)`. If the built-in prompt templates don't robustly delimit the user-controlled content, the document text can contain instructions that override the system prompt. Additionally at line 525, `emailBody` is injected into a reply prompt with `replyPrompt.user(emailBody, lang)`, where the full email thread body — which may contain attacker-crafted content — goes directly into the prompt.
- **Recommendation**: All prompt template functions should wrap user-provided content in clearly delimited, non-spoofable boundaries. Consider a standard `wrapUserContent(text)` helper used by all prompt builders.

---

### HIGH (H1–H7) — Should be fixed before next release

#### H1. Race Condition: Concurrent `sendMessage` Calls Can Corrupt Shared State

- **File**: `frontend/src/composables/useAgentLoop.ts:466-468, 498-499`
- **Category**: Logic Bug (Race Condition)
- **Details**: The guard `if (loading.value) return` at line 466 is a check-then-act on a reactive ref, which is not atomic. If `sendMessage` is called twice in rapid succession (e.g., double-click, keyboard repeat), both calls could pass the guard before `loading.value = true` is set at line 498. This would cause two concurrent agent loops writing to the same `history` array, corrupting `lastIndex` calculations and creating duplicate `abortController` instances. While Vue's reactivity is single-threaded, `await` points yield to the microtask queue, so a second call can slip in.
- **Recommendation**: Set `loading.value = true` immediately at the top of `sendMessage` before any async work, or use a synchronous mutex flag.

#### H2. `lastIndex` Stale Reference if `history` Is Mutated During Loop

- **File**: `frontend/src/composables/useAgentLoop.ts:242, 267, 317`
- **Category**: Logic Bug (State Management)
- **Details**: `lastIndex` is captured once at line 242 as `history.value.length - 1` and then used throughout the entire `while` loop. However, `history.value.push(...)` is called at line 405 (`agentStoppedByUser` message), and `processChat` is called within `sendMessage` which also pushes messages. If any external code or a concurrent path pushes to `history` during the loop, `lastIndex` becomes stale and `history.value[lastIndex].content = text` would update the wrong message.
- **Recommendation**: Either refresh `lastIndex` after each push, or use a direct reference to the message object rather than an index.

#### H3. Timer Leak — `timeoutId` Reassigned Without Clearing Previous Timer

- **File**: `frontend/src/composables/useAgentLoop.ts:559-568`
- **Category**: Logic Bug (Resource Leak)
- **Details**: At line 559, `timeoutId` is set to a `setTimeout` for `timeoutPromise`. Then at line 568, `timeoutId` is reassigned to a new `setTimeout` for `htmlPromise` without clearing the first timer. The `finally` block at line 588 only clears the last assigned `timeoutId`, leaving the first timer's callback orphaned.
- **Recommendation**: Use separate variables for each timer, or clear the previous timer before reassignment.

#### H4. Error Message Contains Raw `err.message` — Potential Information Leak

- **File**: `frontend/src/composables/useAgentLoop.ts:304, 435, 551, 819`
- **Category**: Security (Information Disclosure)
- **Details**: Multiple error handlers interpolate `err.message` directly into display messages shown to users:
  - Line 304: `Error: The model or API failed to respond. ${err.message || ''}`
  - Line 435: `${t('imageError')}: ${err.message}`
  - Line 551: `Error: ${err.message || t('failedToResponse')}`
  - Line 819: `Error: ${err.message || t('failedToResponse')}`

  Server error messages may contain internal URLs, API keys, stack traces, or other sensitive infrastructure details.
- **Recommendation**: Sanitize or truncate `err.message` before display. Use generic user-facing messages and log full details to console.

#### H5. `any` Type on Error Parameters Disables Type Checking

- **File**: `frontend/src/composables/useAgentLoop.ts:111, 288, 334, 433, 544, 602, 624`
- **Category**: Code Quality
- **Details**: `isCredentialError(error: any)` at line 111, `catch (err: any)` at multiple lines, and `toolArgs: Record<string, any>` at line 334 all use `any` types. The `any` type on function parameters and tool args should be replaced with `unknown` to force explicit narrowing.
- **Recommendation**: Replace `any` with `unknown` and add type guards for property access.

#### H6. XSS via `imageSrc` — Unvalidated URL Used in `fetch` and `img.src`

- **File**: `frontend/src/composables/useImageActions.ts:53-98`
- **Category**: Security (XSS)
- **Details**: `copyImageToClipboard(imageSrc: string)` directly assigns `imageSrc` to both `fetch(imageSrc)` at line 56 and `img.src = imageSrc` at line 70. If `imageSrc` is a `javascript:` URL or crafted data URL, this could lead to XSS. The `fetch` call has no URL validation — it could be used for SSRF if `imageSrc` points to internal resources. The `insertImageToWord` and `insertImageToPowerPoint` functions also use `imageSrc` without validating the URL format.
- **Recommendation**: Validate that `imageSrc` matches expected patterns (e.g., `data:image/...;base64,...` or allowed HTTPS origins) before use.

#### H7. `THINK_TAG_REGEX` Module-Level Regex With `g` Flag — Maintenance Hazard

- **File**: `frontend/src/composables/useImageActions.ts:10`
- **Category**: Performance / Maintenance
- **Details**: The regex uses the `g` flag and is module-scoped. The `g` flag maintains `lastIndex` state between calls. Currently safe because only `replace()` is used (which resets `lastIndex`), but adding a `test()` call later would introduce a subtle bug where alternating matches are skipped.
- **Recommendation**: Remove the `g` flag if only used with `replace()` on the full string, or document the constraint.

---

### MEDIUM (M1–M22) — Should be addressed in upcoming sprints

#### M1. `buildChatMessages` Drops `system` Messages From History

- **File**: `frontend/src/composables/useAgentLoop.ts:209-211`
- **Category**: Logic Bug
- **Details**: The filter `.filter(m => m.role === 'user' || m.role === 'assistant')` strips all `system` messages. System messages pushed at line 405 (`agentStoppedByUser`) are silently dropped from subsequent chat contexts.

#### M2. Hardcoded French String in File Upload Error Path

- **File**: `frontend/src/composables/useAgentLoop.ts:604`
- **Category**: Code Quality (i18n)
- **Details**: `'Erreur lors de l\'extraction du fichier.'` is hardcoded in French rather than using `t()`.
- **Recommendation**: Replace with `t('fileExtractionError')` or equivalent i18n key.

#### M3. Smart Reply Success Path Returns Without User Feedback

- **File**: `frontend/src/composables/useAgentLoop.ts:554`
- **Category**: Logic Bug (Missing Cleanup)
- **Details**: The smart reply success path returns without any status update for the user (no success toast), which is inconsistent with other paths.

#### M4. `applyQuickAction` Loading Guard Returns Silently

- **File**: `frontend/src/composables/useAgentLoop.ts:640-828`
- **Category**: Logic Bug
- **Details**: At line 717, `if (loading.value) return` — if `loading` is already true from a concurrent call, the function returns without any feedback to the user.

#### M5. `selectedText` Fragile Initialization in Normal Path

- **File**: `frontend/src/composables/useAgentLoop.ts:574-575`
- **Category**: Logic Bug
- **Details**: When `htmlContent` is truthy but `richContext.cleanText` is falsy, the fallback `selectedText = richContext.cleanText || selectedText` resolves to `'' || ''`. Correct but fragile if refactored.

#### M6. Overly Large Function: `sendMessage` (~190 lines)

- **File**: `frontend/src/composables/useAgentLoop.ts:449-638`
- **Category**: Architecture
- **Details**: Handles input validation, selection fetching, file upload, prompt construction, model routing, streaming, error handling, and UI state management in one function.
- **Recommendation**: Decompose into focused helpers: `preparePayload()`, `fetchSelection()`, `buildPrompt()`, `executeStream()`.

#### M7. Overly Large Function: `runAgentLoop` (~195 lines)

- **File**: `frontend/src/composables/useAgentLoop.ts:227-421`
- **Category**: Architecture
- **Details**: Handles streaming, tool call parsing, tool execution, abort handling, error handling, and message management. Deeply nested control flow.

#### M8. `response` Mutation-Driven State in Agent Loop

- **File**: `frontend/src/composables/useAgentLoop.ts:254`
- **Category**: Code Quality
- **Details**: `response` is initialized with a complex default object then mutated via callback closures (`onStream`, `onToolCallDelta`). A builder or accumulator pattern would be clearer.

#### M9. Prompt Injection Surface via User Profile Fields

- **File**: `frontend/src/composables/useAgentPrompts.ts:34-41`
- **Category**: Security (Prompt Injection)
- **Details**: `firstName` and `lastName` from user profile are interpolated directly into the system prompt. In shared/managed environments where admins set profiles, this could be exploited.

#### M10. Prompt Bodies Are English-Only Despite `lang` Parameter

- **File**: `frontend/src/composables/useAgentPrompts.ts:60-171`
- **Category**: Code Quality (i18n)
- **Details**: All agent prompt functions contain English-only instructional text. Only `userProfilePromptBlock` uses `t()`. The `lang` parameter controls only the "communicate in ${lang}" instruction. Likely intentional but undocumented.

#### M11. `COMMON_FORMATTING_INSTRUCTIONS` Not Used by `excelAgentPrompt`

- **File**: `frontend/src/composables/useAgentPrompts.ts:97-117`
- **Category**: Code Quality (Inconsistency)
- **Details**: Word, PowerPoint, and Outlook prompts append `COMMON_FORMATTING_INSTRUCTIONS`, but Excel does not. May be intentional (cells don't need Markdown) but inconsistency is undocumented.

#### M12. `insertImageToPowerPoint` Ignores `'NoAction'` Type Semantics

- **File**: `frontend/src/composables/useImageActions.ts:111-157`
- **Category**: Logic Bug
- **Details**: `'NoAction'` semantically means "do nothing" but the function still inserts the image. When `type === 'append'`, it selects the last slide, which may not be the presentation end.

#### M13. Weak Fallback ID Generation in `createDisplayMessage`

- **File**: `frontend/src/composables/useImageActions.ts:36`
- **Category**: Code Quality
- **Details**: The fallback `message-${Date.now()}-${Math.random().toString(36).slice(2, 10)}` used when `crypto.randomUUID()` is unavailable has theoretical collision risk in same-millisecond scenarios. IDs are used as Vue `:key` props.

#### M14. Deprecated `document.execCommand('copy')` in Fallback

- **File**: `frontend/src/composables/useImageActions.ts:84`
- **Category**: Code Quality
- **Details**: `document.execCommand('copy')` is deprecated. Used as fallback for Office Webview environments.

#### M15. HTML Injection via `richHtml` Passed to Office APIs

- **File**: `frontend/src/composables/useOfficeInsert.ts:96, 98, 143-145`
- **Category**: Security (HTML Injection)
- **Details**: `richHtml` from LLM output is passed directly to `item.body.setSelectedDataAsync()` (Outlook) and `range.insertHtml()` (Word). If `richHtml` contains `<script>` tags or event handlers, these could execute. While Office APIs typically sanitize, Outlook webview may not fully strip dangerous constructs.
- **Recommendation**: Sanitize `richHtml` through a whitelist-based HTML sanitizer (e.g., DOMPurify) before passing to Office APIs.

#### M16. `insertToDocument` Silently Swallows All Errors

- **File**: `frontend/src/composables/useOfficeInsert.ts:107, 117-118, 133-134, 156-157`
- **Category**: Code Quality (Error Handling)
- **Details**: Every `catch` block catches all exceptions and falls back to clipboard copy without logging. Makes debugging insertion failures extremely difficult.
- **Recommendation**: Add `console.error` logging in all catch blocks, even when falling back gracefully.

#### M17. Implicit Word Fallback When No Host Flag Matches

- **File**: `frontend/src/composables/useOfficeInsert.ts:33`
- **Category**: Code Quality
- **Details**: `insertToDocument` falls through to the Word path when none of the host flags are true. This would fail confusingly if called in an unexpected context.

#### M18. Hidden Side Effect: `insertType.value` Mutation in `insertToDocument`

- **File**: `frontend/src/composables/useOfficeInsert.ts:140`
- **Category**: Code Quality (Side Effect)
- **Details**: `insertToDocument` mutates the external `insertType` ref as a side effect before passing it to `insertFormattedResult`. Could cause unexpected reactivity triggers.

#### M19. Promise Constructor Anti-Pattern in Outlook Functions

- **File**: `frontend/src/composables/useOfficeSelection.ts:13-67`
- **Category**: Code Quality (Anti-Pattern)
- **Details**: All four Outlook helpers (`getOutlookMailBody`, `getOutlookMailBodyAsHtml`, `getOutlookSelectedText`, `getOutlookSelectedHtml`) use identical `Promise.race` with manual timeout. Should extract a shared `withTimeout(promise, ms)` helper.

#### M20. Timeout Promises Create Orphaned Timer Callbacks

- **File**: `frontend/src/composables/useOfficeSelection.ts:22, 38, 51, 65`
- **Category**: Performance (Memory)
- **Details**: When the main promise wins the `Promise.race`, the timeout's `setTimeout` callback still fires and resolves a promise nobody listens to. Repeated rapid calls accumulate orphaned timers.

#### M21. Excel Selection Returns Unescaped Tab-Separated Values

- **File**: `frontend/src/composables/useOfficeSelection.ts:86-92`
- **Category**: Logic Bug
- **Details**: Cell values containing tabs or newlines make the output ambiguous and unparseable.

#### M22. Inconsistent HTML Fallback Behavior for Outlook Selection

- **File**: `frontend/src/composables/useOfficeSelection.ts:114-115`
- **Category**: Code Quality (Inconsistency)
- **Details**: `getOutlookMailBodyAsHtml()` falls back to `getOutlookMailBody()` (plain text) if HTML is empty. But the selected text path has no fallback — returns empty string if HTML unavailable.

---

### LOW (L1–L10) — Address when touching related code

#### L1. Subtle Destructuring Order Dependency for Scroll Defaults

- **File**: `frontend/src/composables/useAgentLoop.ts:178-179`
- **Details**: `scrollToMessageTop = scrollToBottom` and `scrollToVeryBottom = scrollToBottom` depend on destructuring order.

#### L2. `getActionLabelForCategory` Default Case Merges `'write'` With Unknown

- **File**: `frontend/src/composables/useAgentLoop.ts:186-195`
- **Details**: `case 'write': default:` together means any unknown category silently gets "running" label. Adding a new `ToolCategory` value would fall through without compile-time warning.

#### L3. Inconsistent Quick Action Type Handling With Type Assertions

- **File**: `frontend/src/composables/useAgentLoop.ts:650-652`
- **Details**: Manual type casts (`as ExcelQuickAction | undefined`, etc.) based on host flags. Could use discriminated union types instead.

#### L4. `payload` Parameter Typed as `unknown` — Loose Contract

- **File**: `frontend/src/composables/useAgentLoop.ts:449`
- **Details**: `sendMessage(payload?: unknown)` accepts `unknown` then checks `typeof payload === 'string'`. A `string | undefined` type would be more precise.

#### L5. `hostIsWord` Parameter Accepted But Never Used

- **File**: `frontend/src/composables/useAgentPrompts.ts:13, 26`
- **Category**: Dead Code
- **Details**: `hostIsWord` is declared in `UseAgentPromptsOptions` and destructured but never referenced. Word is the implicit `else` fallback in `agentPrompt`.

#### L6. `cleanContent` and `splitThinkSegments` Implement Different Think-Tag Logic

- **File**: `frontend/src/composables/useImageActions.ts:13-33, 40-42`
- **Details**: `splitThinkSegments` handles unclosed tags gracefully (treats remaining as think segment), while `cleanContent` regex would leave unclosed `<think>` tags in output. Inconsistent behavior for malformed tags.

#### L7. Inconsistent Image Insert Error Reporting Across Hosts

- **File**: `frontend/src/composables/useOfficeInsert.ts:169-198`
- **Details**: Word and PowerPoint fall back to `copyImageToClipboard`, Excel shows info and returns, Outlook falls through and shows misleading "imageInsertWordOnly" message.

#### L8. `VERBOSE_LOGGING_ENABLED` Read Once at Module Load

- **File**: `frontend/src/composables/useOfficeInsert.ts:10-11`
- **Details**: Verbose logging cannot be toggled without full page reload. Standard for Vite but worth noting.

#### L9. `range.values` Typed as `any[]` in Excel Selection

- **File**: `frontend/src/composables/useOfficeSelection.ts:90`
- **Details**: Should use `(string | number | boolean)[]` instead of `any[]` for row type.

#### L10. Word HTML Selection Path Swallows Error Silently

- **File**: `frontend/src/composables/useOfficeSelection.ts:135-136`
- **Details**: `catch` returns `''` without any logging for `Word.run` or `range.getHtml()` failures.

---

## 3. Priority Recommendations

### Immediate (CRITICAL + HIGH Security)

1. **Sanitize all document/email content** before interpolating into LLM prompts. Use clear delimiters with escaping (e.g., XML CDATA-style wrapping or a unique boundary token) to prevent prompt injection from attacker-crafted documents. *Addresses: C1, C2, M9*

2. **Validate `imageSrc` URLs** before using them in `fetch()`, `img.src`, or Office APIs. Enforce that they match expected patterns (`data:image/...;base64,...` or allowed HTTPS origins). *Addresses: H6*

3. **Sanitize `err.message`** before displaying to users. Use generic messages and log full details to console only. *Addresses: H4*

### Short-term (HIGH Logic + Architecture)

4. **Fix the race condition** in `sendMessage` by setting `loading.value = true` immediately before any async work. *Addresses: H1*

5. **Fix the timer leak** where `timeoutId` is reassigned without clearing the first timer. *Addresses: H3*

6. **Use direct message references** instead of `lastIndex` in the agent loop to prevent stale index bugs. *Addresses: H2*

### Medium-term (Architecture + Quality)

7. **Decompose the three large functions** (`sendMessage`, `applyQuickAction`, `runAgentLoop`) into smaller, testable units. *Addresses: M6, M7*

8. **Extract a shared `withTimeout` helper** to deduplicate the four Outlook promise-race patterns. *Addresses: M19, M20*

9. **Add error logging** to all silent `catch` blocks, especially in `useOfficeInsert.ts` and `useOfficeSelection.ts`. *Addresses: M16, L10*

10. **Replace `any` types** with `unknown` and add proper type narrowing. *Addresses: H5, L9*

11. **Sanitize `richHtml`** through DOMPurify or equivalent before passing to Office APIs. *Addresses: M15*

12. **Fix i18n violation** — replace hardcoded French string with `t()` call. *Addresses: M2*
