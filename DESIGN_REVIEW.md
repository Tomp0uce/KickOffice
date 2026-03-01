# Design Review & Code Audit — v3

**Date**: 2026-03-01
**Scope**: Full codebase audit — security, logic bugs, code quality, dead code, architecture across backend, frontend utilities, composables, pages/components, API layer, and infrastructure.

---

## 1. Executive Summary

KickOffice is a Microsoft Office add-in (Word, Excel, PowerPoint, Outlook) powered by a Vue 3 + Vite frontend and Express.js backend proxy for OpenAI-compatible LLM APIs.

**Previous audits**: v1 (2026-02-21, 38 issues — all resolved) and v2 (2026-02-22, 28 issues — all resolved except B1-B3 build warnings and F1-F3 feature issues).

This v3 audit is a **fresh, comprehensive analysis** of the entire codebase. Findings are organized by codebase area, then by severity.

### v2 Open Items Status

| ID | Description | Status |
|----|-------------|--------|
| B1 | No unit test infrastructure | OPEN |
| B2 | No linting or formatting configuration | OPEN |
| B3 | No CI pipeline for automated testing | OPEN |
| F1 | Quick actions strip images/formatting | OPEN |
| F2 | Outlook Reply produces low-quality responses | OPEN |
| F3 | Excel agent processes cells one-by-one | OPEN |

---

## 2. Backend Findings

### CRITICAL

#### BC1. Content-Type enforcement blocks file uploads

- **File**: `backend/src/server.js:91-96`
- **Category**: Broken Functionality
- **Details**: Global middleware rejects all POST/PUT/PATCH requests without `Content-Type: application/json`, but `/api/upload` requires `multipart/form-data` via multer. Every file upload request is blocked with 415 before reaching the route handler.
- **Impact**: The entire file upload feature is non-functional.
- **Fix**: Exempt `/api/upload` from the Content-Type check, e.g. `if (req.path === '/api/upload') return next()`.

#### BC2. Internal LLM API URL exposed in source and .env.example

- **File**: `backend/.env.example:17`, `backend/src/config/models.js:25`
- **Category**: Security / Information Disclosure
- **Details**: The internal company LiteLLM proxy URL `https://litellm.kickmaker.net/v1` is hardcoded as the default value for `LLM_API_BASE_URL`. No validation that the URL is well-formed HTTPS pointing to an expected host.
- **Impact**: Exposes internal infrastructure endpoint. Potential SSRF if the env var is misconfigured.
- **Fix**: Replace with a placeholder like `https://your-llm-api.example.com/v1` and add URL validation.

#### BC3. Sensitive data logged to disk in plaintext

- **File**: `backend/src/utils/logger.js:22-46`, `backend/src/routes/chat.js:49-52, 157-159, 210`
- **Category**: Security / Privacy
- **Details**: `systemLog` writes full request bodies (all user messages) and full LLM responses to `backend/logs/kickoffice.log` as pretty-printed JSON. No redaction of PII, no log rotation, no size limit.
- **Impact**: GDPR risk, credential leakage if users paste secrets, unbounded disk usage.
- **Fix**: Redact message content from logs, add log rotation (e.g. `rotating-file-stream`), set a max file size.

#### BC4. User-supplied credentials forwarded without sanitization

- **File**: `backend/src/services/llmClient.js:34-41`
- **Category**: Security / Header Injection
- **Details**: `X-User-Key` and `X-User-Email` headers are forwarded verbatim to the upstream LLM API. No sanitization against header injection characters (`\r\n`). The user key is accepted with only a minimum length of 8 characters.
- **Impact**: Potential HTTP header injection if the underlying HTTP library has a bypass.
- **Fix**: Strip `\r`, `\n`, and non-printable characters from header values before forwarding.

### HIGH

#### BH1. Drain event listener leak in streaming response

- **File**: `backend/src/routes/chat.js:86-90`
- **Category**: Logic Bug / Resource Leak
- **Details**: When backpressure occurs during SSE streaming, `await new Promise(resolve => res.once('drain', resolve))` will never resolve if the client disconnects between the backpressure check and the drain event.
- **Impact**: Permanently hung request handlers under load with slow/disconnecting clients.
- **Fix**: Race the drain promise with a client-disconnect check or add a timeout.

#### BH2. logAndRespond called after headers already sent (streaming)

- **File**: `backend/src/routes/chat.js:108-116`
- **Category**: Logic Bug
- **Details**: If an error occurs after SSE headers have been sent, the catch block calls `logAndRespond(res, 504/500, ...)` which tries to set headers and write JSON on an active SSE stream, causing `ERR_HTTP_HEADERS_SENT`.
- **Impact**: Unhandled exception, malformed SSE stream, swallowed error.
- **Fix**: Check `res.headersSent` before calling `logAndRespond`; if already streaming, write an SSE error event instead.

#### BH3. Unbounded log file growth

- **File**: `backend/src/utils/logger.js:14, 44`
- **Category**: Security / Availability
- **Details**: Single log file with no rotation, size limits, or cleanup. Full request/response bodies are logged.
- **Impact**: Disk exhaustion on busy servers (potentially gigabytes/day).
- **Fix**: Add log rotation with max file size and retention policy.

#### BH4. Hardcoded version in health endpoint

- **File**: `backend/src/routes/health.js:9`, `backend/package.json:3`
- **Category**: Logic Bug
- **Details**: Health endpoint returns `version: '1.0.0'` but `package.json` declares `"1.0.29"`. These will always drift.
- **Impact**: Impossible to verify deployed version via health check.
- **Fix**: Read version from `package.json` at startup: `const { version } = require('../../package.json')`.

#### BH5. parsePositiveInt allows zero

- **File**: `backend/src/config/env.js:9-20`
- **Category**: Logic Bug
- **Details**: Named `parsePositiveInt` but check is `if (parsed < 0)` which allows zero. `CHAT_RATE_LIMIT_MAX=0` would disable rate limiting.
- **Impact**: Rate limiting bypass via zero config.
- **Fix**: Change to `if (parsed <= 0)`.

#### BH6. Upload route lacks magic-byte file validation

- **File**: `backend/src/routes/upload.js:38-78`
- **Category**: Security
- **Details**: File type detection relies on client-controlled MIME type or extension. No magic-byte validation.
- **Impact**: Attackers can upload crafted files (zip bombs via XLSX, XXE via DOCX) that exploit parsing libraries.
- **Fix**: Add magic-byte validation (e.g. `file-type` package) before processing.

#### BH7. ReDoS potential in sanitizeErrorText

- **File**: `backend/src/utils/http.js:20-21`
- **Category**: Security / Performance
- **Details**: Regex objects are constructed inside a loop on every error response. On pathological input from upstream, this could block the event loop.
- **Fix**: Pre-compile regex patterns at module load time.

### MEDIUM

#### BM1. No graceful shutdown handling

- **File**: `backend/src/server.js`
- **Category**: Architecture
- **Details**: No `SIGTERM`/`SIGINT` handler. In-flight streaming connections are abruptly severed during deployment.
- **Fix**: Add signal handlers that stop accepting new connections and drain existing ones.

#### BM2. Unused `routeName` parameter in `validateChatRequest`

- **File**: `backend/src/middleware/validate.js:159`
- **Category**: Dead Code
- **Details**: Second parameter `routeName` is accepted but never used in the function body.

#### BM3. Exported functions never imported externally

- **File**: `backend/src/middleware/validate.js:219-221`
- **Category**: Dead Code
- **Details**: `validateMaxTokens`, `validateTemperature`, `validateTools` are exported but only used internally.

#### BM4. Exported constants/functions never imported externally

- **File**: `backend/src/services/llmClient.js:10-29`
- **Category**: Dead Code
- **Details**: `TIMEOUTS`, `getChatTimeoutMs`, `getImageTimeoutMs` exported but only used internally.

#### BM5. Validated values discarded in validateChatRequest

- **File**: `backend/src/middleware/validate.js:189-213`
- **Category**: Logic Bug (minor)
- **Details**: `validateTemperature` and `validateMaxTokens` compute `.value` but only `.error` is checked. Callers pass original `req.body` values, bypassing any future normalization.

#### BM6. Inconsistent error logging patterns

- **File**: Multiple backend files
- **Category**: Code Quality
- **Details**: Four different logging patterns (`console.error`, `systemLog`, `logAndRespond`, `process.stderr.write`) used inconsistently. Many errors logged through both `systemLog` and `console.error` on consecutive lines.

#### BM7. `handleErrorResponse` return value discarded

- **File**: `backend/src/routes/chat.js:61, 169`, `backend/src/routes/image.js:31`
- **Category**: Code Quality
- **Details**: `handleErrorResponse` returns sanitized error text but all call sites discard it.

#### BM8. `allCsv` declared with `let` instead of `const`

- **File**: `backend/src/routes/upload.js:58`
- **Category**: Code Quality
- **Details**: Variable is never reassigned, only mutated via `.push()`.

#### BM9. No multer field count limits

- **File**: `backend/src/routes/upload.js:13-18`
- **Category**: Security
- **Details**: No `limits.fields` or `limits.fieldSize` set. Attacker could send thousands of non-file fields.

#### BM10. No request ID / correlation

- **File**: Multiple backend files
- **Category**: Architecture
- **Details**: No request ID generation. Cannot correlate client-side errors with server-side logs.

### LOW

#### BL1. Dead branch: `if (!imageModel)` check

- **File**: `backend/src/routes/image.js:17-19`
- **Category**: Dead Code
- **Details**: `models.image` is always defined in static config. Check can never be true.

#### BL2. French strings hardcoded in backend

- **File**: `backend/src/routes/upload.js:89`, `backend/src/config/models.js:49`
- **Category**: i18n
- **Details**: `'Contenu tronqué en raison de la taille du fichier'`, `'Raisonnement'` hardcoded in French.

#### BL3. Stale comment about character limit

- **File**: `backend/src/routes/upload.js:86-87`
- **Category**: Documentation
- **Details**: Comment says "50k chars" but constant is `MAX_CHARS = 100000`.

#### BL4. `isPlainObject` accepts non-plain objects

- **File**: `backend/src/middleware/validate.js:4-6`
- **Category**: Code Quality
- **Details**: Returns true for Date, RegExp, Map, Set. Safe for JSON-parsed context but imprecise.

---

## 3. Frontend Utilities Findings

### CRITICAL

#### UC1. Prompt injection via custom prompt templates

- **File**: `frontend/src/utils/constant.ts:305-321, 441-466, 469-495, 497-523`
- **Category**: Security
- **Details**: `getBuiltInPrompt()` and per-host variants load custom prompts from localStorage and use `String.replace()` with `${text}` as pattern. The replacement string `text` is interpreted by `String.replace()` — meaning `$&`, `$'`, `` $` ``, and `$<n>` patterns in user text are treated as special replacement patterns, causing silent data corruption.
- **Impact**: User-supplied text containing `$&` or `$'` produces garbled output. Prompt injection via localStorage manipulation.
- **Fix**: Use a function as the replacement argument: `.replace(/\$\{text\}/g, () => text)`.

#### UC2. XOR "obfuscation" provides false security for API keys

- **File**: `frontend/src/utils/credentialStorage.ts:7-36`
- **Category**: Security
- **Details**: API keys XOR'd with hardcoded key `'K1ck0ff1c3'` then base64-encoded in localStorage. Trivially reversible with browser dev tools.
- **Impact**: Any XSS vulnerability allows immediate credential theft. The obfuscation creates a false sense of security.
- **Fix**: Document the limitation clearly. Consider using session-only storage or backend-managed tokens.

#### UC3. Unsanitized HTML injection in Outlook tools

- **File**: `frontend/src/utils/outlookTools.ts:505-534` (insertHtmlAtCursor), `outlookTools.ts:261-291` (setEmailBodyHtml)
- **Category**: Security / XSS
- **Details**: Both tools insert raw HTML directly into Outlook email bodies via `setSelectedDataAsync` with no DOMPurify sanitization. The `html` argument from LLM output is passed verbatim.
- **Impact**: Malicious HTML (scripts, phishing forms, event handlers) injected into outgoing emails.
- **Fix**: Sanitize through DOMPurify before passing to Office APIs.

### HIGH

#### UH1. `eval_officejs` declared in ExcelToolName but never defined

- **File**: `frontend/src/utils/excelTools.ts:66`
- **Category**: Logic Bug / Dead Code
- **Details**: `'eval_officejs'` is in the `ExcelToolName` union but no tool definition exists. The `as unknown as Record<ExcelToolName, ExcelToolDefinition>` cast suppresses the TypeScript error.
- **Impact**: Runtime crash if the agent tries to invoke `eval_officejs` for Excel.
- **Fix**: Either add the tool definition or remove the name from the union type.

#### UH2. Column letter arithmetic overflow

- **File**: `frontend/src/utils/excelTools.ts:1155, 1214`
- **Category**: Logic Bug
- **Details**: `String.fromCharCode(columnLetter.charCodeAt(0) + count - 1)` breaks for multi-character columns (AA, AB) and overflows for single-character columns near Z.
- **Impact**: Data corruption or crash when inserting/deleting columns at/beyond column Z.
- **Fix**: Implement proper column letter arithmetic that handles multi-character references.

#### UH3. Double timeout in Outlook tool execution

- **File**: `frontend/src/utils/outlookTools.ts:67-72`
- **Category**: Logic Bug
- **Details**: Inner 3-second `Promise.race` timeout is extremely aggressive for Outlook API calls (which need server round-trips). It returns an error string rather than rejecting, so the outer 10-second timeout never triggers. Many legitimate operations silently fail.
- **Impact**: Silent failures on slow connections or large emails.
- **Fix**: Remove the inner 3-second timeout or increase it significantly.

#### UH4. `language` parameter ignored in translate prompt

- **File**: `frontend/src/utils/constant.ts:32-48`
- **Category**: Logic Bug
- **Details**: Translate prompt hardcodes "French-English bilingual translation" regardless of the `language` parameter. Explicitly says "Ignore requested output language preferences."
- **Impact**: Users who select a target language other than French/English see their preference silently ignored.
- **Fix**: Make the prompt respect the language parameter.

#### UH5. Host detection caching can return wrong host

- **File**: `frontend/src/utils/hostDetection.ts:3-37`
- **Category**: Logic Bug
- **Details**: `detectOfficeHost()` caches result in a module-level variable. If called before `Office.onReady` fires, may cache an incorrect host permanently.
- **Impact**: Wrong host detection in edge cases, leading to wrong tool set and prompts.
- **Fix**: Only cache after `Office.onReady` has resolved.

#### UH6. Message toast singleton race condition

- **File**: `frontend/src/utils/message.ts:13-43`
- **Category**: Race Condition
- **Details**: `showMessage` uses a module-level `messageInstance` singleton. If called while the 300ms `setTimeout` cleanup is pending, the stale closure may unmount the new instance.
- **Impact**: Toasts prematurely destroyed or DOM container leaks.
- **Fix**: Clear the pending timeout before creating a new instance, or use a unique ID per instance.

#### UH7. `html: true` in MarkdownIt with `style` in DOMPurify allowlist

- **File**: `frontend/src/utils/officeRichText.ts:77, 403`
- **Category**: Security
- **Details**: Raw HTML passes through MarkdownIt, and DOMPurify allows the `style` attribute. Inline styles enable CSS-based attacks (data exfiltration via `background: url(...)`, UI spoofing).
- **Impact**: CSS injection in Office-inserted content.
- **Fix**: Remove `style` from `ALLOWED_ATTR` or use a CSS sanitizer.

### MEDIUM

#### UM1. Massive type unsafety with `as unknown as` casts

- **Files**: `excelTools.ts:21`, `wordTools.ts:194`, `outlookTools.ts:75`, `powerpointTools.ts:39`
- **Category**: Type Safety
- **Details**: All four tool-creation factories use `as unknown as Record<ToolName, ToolDefinition>`, completely disabling TypeScript checking for missing tool definitions.
- **Fix**: Use a type-safe builder that validates all required tool names are present.

#### UM2. Pervasive `any` types in tool definitions

- **Files**: All tool definition files
- **Category**: Type Safety
- **Details**: `args: Record<string, any>`, `mailbox: any`, `context: any` across all tool files. Outlook tools especially: `getMailbox(): any`, `getOfficeAsyncStatus(): any`.

#### UM3. Duplicated `generateVisualDiff` function

- **Files**: `outlookTools.ts:8-20`, `wordTools.ts:8-21`
- **Category**: Code Duplication
- **Details**: Identical function copy-pasted between two files.
- **Fix**: Extract to a shared utility.

#### UM4. Duplicated Office API helpers

- **Files**: `outlookTools.ts:42-66`, `officeOutlook.ts:48-66`
- **Category**: Code Duplication
- **Details**: Both files define their own helpers for `Office.context.mailbox`, `CoercionType`, `AsyncResultStatus`. The typed abstractions in `officeOutlook.ts` go partially unused.

#### UM5. `Ref` without type parameter in WordFormatter

- **File**: `frontend/src/utils/wordFormatter.ts:23, 54`
- **Category**: Type Safety
- **Details**: `insertType: Ref` (bare, unparameterized). Value is typed as `unknown`, requiring implicit `any` comparisons.

#### UM6. `searchAndReplace` tools labeled as category `'read'`

- **Files**: `excelTools.ts:1365`, `wordTools.ts:517`
- **Category**: Inconsistency
- **Details**: Both perform write operations (replacing text/values) but are categorized as `'read'`.

#### UM7. Redundant Set + Array checks in toolStorage

- **File**: `frontend/src/utils/toolStorage.ts:47`
- **Category**: Code Quality
- **Details**: `!storedEnabledSet.has(name) && !storedEnabledNames.includes(name)` — the Set was created from the same array, so these checks are redundant.

#### UM8. No `QuotaExceededError` handling for localStorage

- **Files**: `credentialStorage.ts`, `toolStorage.ts`, `savedPrompts.ts`, `constant.ts`
- **Category**: Error Handling
- **Details**: Multiple files write to localStorage without catching `QuotaExceededError`.

#### UM9. `tokenManager.ts` mutates input messages

- **File**: `frontend/src/utils/tokenManager.ts:105-108`
- **Category**: Code Quality
- **Details**: `delete msg.tool_calls` mutates original message objects from the input array.
- **Fix**: Clone messages before modifying: `const msg = { ...originalMsg }`.

#### UM10. Character-by-character HTML reconstruction in PowerPoint

- **File**: `frontend/src/utils/powerpointTools.ts:121-183`
- **Category**: Performance
- **Details**: Loads font properties for each individual character (up to 100,000) via `range.getSubstring(i, 1)`. Very memory-intensive.

### LOW

#### UL1. Typo in export name `buildInPrompt`

- **File**: `frontend/src/utils/constant.ts:30`
- **Details**: Should be `builtInPrompt`.

#### UL2. `deleteText` reports success when no text selected

- **File**: `frontend/src/utils/wordTools.ts:710-715`
- **Details**: Inserts empty string (no-op) but returns "Successfully deleted text".

#### UL3. Inconsistent error handling strategy across tools

- **Files**: All tool files
- **Details**: Some return error strings, some throw, some return empty strings. Caller must check string prefixes.

#### UL4. `markdown.ts` vs `officeRichText.ts` naming confusion

- **Details**: Both render Markdown but for different targets (chat vs Office). Names don't communicate this.

---

## 4. Frontend Composables Findings

### CRITICAL

#### CC1. Prompt injection via unsanitized document selection

- **File**: `frontend/src/composables/useAgentLoop.ts:613-616`
- **Category**: Security
- **Details**: Office selection text is directly interpolated into the user message: `fullMessage += \`\n\n[${selectionLabel}: "${selectedText}"]\``. A malicious document containing `"]` followed by injection instructions can break out of the framing.
- **Impact**: Indirect prompt injection — attacker crafts document, victim opens it and uses KickOffice.
- **Fix**: Use robust delimiters (XML CDATA-style, base64, unique boundary token).

#### CC2. Prompt injection via quick action selection text

- **File**: `frontend/src/composables/useAgentLoop.ts:753, 761, 525`
- **Category**: Security
- **Details**: `textForLlm` from document selection passed directly to `action.user(textForLlm, lang)`. Email body injected into reply prompt at line 525 with `replyPrompt.user(emailBody, lang)`.
- **Impact**: Attacker-crafted email/document content can override system prompt instructions.

### HIGH

#### CH1. Race condition: concurrent `sendMessage` calls corrupt state

- **File**: `frontend/src/composables/useAgentLoop.ts:466-468, 498-499`
- **Category**: Logic Bug
- **Details**: `if (loading.value) return` is check-then-act on a reactive ref. Two rapid calls can both pass before `loading.value = true` is set, causing concurrent agent loops writing to the same history array.
- **Fix**: Set `loading.value = true` immediately at the top, before any async work.

#### CH2. `lastIndex` stale reference during agent loop

- **File**: `frontend/src/composables/useAgentLoop.ts:242, 267, 317`
- **Category**: Logic Bug
- **Details**: `lastIndex` captured once as `history.value.length - 1` but history is pushed to during the loop. Stale index updates the wrong message.
- **Fix**: Store a direct reference to the message object instead of an index.

#### CH3. Timer leak — `timeoutId` reassigned without clearing

- **File**: `frontend/src/composables/useAgentLoop.ts:559-568`
- **Category**: Resource Leak
- **Details**: First `setTimeout` assigned to `timeoutId`, then reassigned to a new one without clearing the first. `finally` only clears the last.

#### CH4. Raw `err.message` displayed to users

- **File**: `frontend/src/composables/useAgentLoop.ts:304, 435, 551, 819`
- **Category**: Security / Information Disclosure
- **Details**: Server error messages (potentially containing internal URLs, API keys, stack traces) shown directly to users.
- **Fix**: Show generic messages, log details to console only.

#### CH5. `any` types on error parameters and tool args

- **File**: `frontend/src/composables/useAgentLoop.ts:111, 288, 334, 433, 544`
- **Category**: Type Safety
- **Details**: `isCredentialError(error: any)`, multiple `catch (err: any)`, `toolArgs: Record<string, any>`.
- **Fix**: Use `unknown` with type guards.

#### CH6. XSS via unvalidated `imageSrc` URL

- **File**: `frontend/src/composables/useImageActions.ts:53-98`
- **Category**: Security
- **Details**: `imageSrc` directly assigned to `fetch()` and `img.src` with no URL validation. Could be `javascript:` URL or point to internal resources (SSRF).
- **Fix**: Validate URL pattern before use.

#### CH7. `THINK_TAG_REGEX` module-level with `g` flag — maintenance hazard

- **File**: `frontend/src/composables/useImageActions.ts:10`
- **Category**: Maintenance
- **Details**: `g` flag maintains `lastIndex` state. Adding a `test()` call later would introduce subtle bugs.

### MEDIUM

#### CM1. Hardcoded French string in file upload error

- **File**: `frontend/src/composables/useAgentLoop.ts:604`
- **Category**: i18n
- **Details**: `'Erreur lors de l\'extraction du fichier.'` hardcoded instead of using `t()`.

#### CM2. `buildChatMessages` drops system messages

- **File**: `frontend/src/composables/useAgentLoop.ts:209-211`
- **Details**: Filter strips all `system` messages from history.

#### CM3. Overly large functions

- **Files**: `useAgentLoop.ts:449-638` (`sendMessage` ~190 lines), `useAgentLoop.ts:640-828` (`applyQuickAction` ~188 lines), `useAgentLoop.ts:227-421` (`runAgentLoop` ~195 lines)
- **Category**: Architecture
- **Fix**: Decompose into focused helpers.

#### CM4. `insertToDocument` silently swallows all errors

- **File**: `frontend/src/composables/useOfficeInsert.ts:107, 117-118, 133-134, 156-157`
- **Category**: Error Handling
- **Details**: Every catch block falls back to clipboard with no logging.

#### CM5. Promise constructor anti-pattern in Outlook functions

- **File**: `frontend/src/composables/useOfficeSelection.ts:13-67`
- **Details**: Four nearly-identical `Promise.race` + manual timeout patterns.
- **Fix**: Extract a shared `withTimeout(promise, ms)` helper.

#### CM6. Timeout promises create orphaned timers

- **File**: `frontend/src/composables/useOfficeSelection.ts:22, 38, 51, 65`
- **Details**: Losing `Promise.race` timers still fire, resolving promises nobody listens to.

#### CM7. Excel selection returns unescaped tab-separated values

- **File**: `frontend/src/composables/useOfficeSelection.ts:86-92`
- **Details**: Cell values containing tabs/newlines make output ambiguous.

#### CM8. HTML injection via `richHtml` to Office APIs

- **File**: `frontend/src/composables/useOfficeInsert.ts:96, 98, 143-145`
- **Category**: Security
- **Details**: `richHtml` from LLM output passed directly to `insertHtml()` and `setSelectedDataAsync()`.
- **Fix**: Sanitize through DOMPurify before passing to Office APIs.

#### CM9. Prompt injection via user profile fields

- **File**: `frontend/src/composables/useAgentPrompts.ts:34-41`
- **Details**: `firstName`/`lastName` interpolated directly into system prompt.

#### CM10. `insertImageToPowerPoint` ignores `'NoAction'` semantics

- **File**: `frontend/src/composables/useImageActions.ts:111-157`
- **Details**: `'NoAction'` should mean "do nothing" but still inserts the image.

#### CM11. Hidden side effect: `insertType.value` mutation

- **File**: `frontend/src/composables/useOfficeInsert.ts:140`
- **Details**: Mutates external ref as side effect, causing unexpected reactivity triggers.

### LOW

#### CL1. `hostIsWord` parameter accepted but never used

- **File**: `frontend/src/composables/useAgentPrompts.ts:13, 26`
- **Category**: Dead Code

#### CL2. `cleanContent` and `splitThinkSegments` use different think-tag logic

- **File**: `frontend/src/composables/useImageActions.ts:13-33, 40-42`
- **Details**: Inconsistent behavior for malformed tags.

#### CL3. Inconsistent image insert error reporting across hosts

- **File**: `frontend/src/composables/useOfficeInsert.ts:169-198`
- **Details**: Outlook falls through and shows misleading "imageInsertWordOnly" message.

#### CL4. `payload` parameter typed as `unknown` — should be `string | undefined`

- **File**: `frontend/src/composables/useAgentLoop.ts:449`

#### CL5. Word HTML selection swallows errors silently

- **File**: `frontend/src/composables/useOfficeSelection.ts:135-136`

---

## 5. Infrastructure Findings

### CRITICAL

#### IC1. Content-Type middleware blocks uploads (same as BC1)

- **File**: `backend/src/server.js:91-96`
- **Details**: See BC1. The upload route at line 104 is unreachable for multipart requests.

#### IC2. Containers run as root

- **Files**: `backend/Dockerfile:1-15`, `frontend/Dockerfile:1-22`
- **Category**: Security
- **Details**: Neither Dockerfile creates or switches to a non-root user. `PUID`/`PGID` env vars in docker-compose have no effect on standard Node/Nginx images.
- **Fix**: Add `USER node` (backend) and create a non-root user for nginx (frontend).

#### IC3. Internal infrastructure URL as default

- **File**: `backend/.env.example:17`
- **Category**: Security
- **Details**: See BC2. `https://litellm.kickmaker.net/v1` as default `LLM_API_BASE_URL`.

### HIGH

#### IH1. Four different Node.js versions across the project

- **Files**: `docker-compose.yml:3` (Node 18), both `Dockerfile`s (Node 22), CI workflow (Node 20), `engines` (>=20.19.0 || >=22.0.0)
- **Category**: Misconfiguration
- **Details**: The manifest-gen service uses Node 18, violating the project's own engines constraint.
- **Fix**: Standardize on a single Node.js version across all files.

#### IH2. Private IP baked into frontend Docker build

- **File**: `frontend/Dockerfile:9`
- **Category**: Security
- **Details**: Default build arg `VITE_BACKEND_URL=http://192.168.50.10:3003` bakes a private IP into the JS bundle.
- **Fix**: Remove default or use a placeholder that fails visibly.

#### IH3. External DuckDNS domain as default in .env.example

- **File**: `.env.example:10-11`
- **Category**: Misconfiguration
- **Details**: `PUBLIC_FRONTEND_URL` and `PUBLIC_BACKEND_URL` set to `https://kickoffice.duckdns.org` as active values.
- **Fix**: Comment them out or use clearly fake placeholders.

#### IH4. Lock files not copied in Dockerfiles

- **Files**: `backend/Dockerfile:5-6`, `frontend/Dockerfile:4-5`
- **Category**: Misconfiguration
- **Details**: Only `package.json` copied before `npm install`, not `package-lock.json`. Non-deterministic builds.
- **Fix**: `COPY package.json package-lock.json ./` and use `npm ci`.

#### IH5. Nginx missing security headers

- **File**: `frontend/nginx.conf:1-21`
- **Category**: Security
- **Details**: Missing `Content-Security-Policy`, `Referrer-Policy`, `X-Frame-Options`. Uses deprecated `X-XSS-Protection`.
- **Fix**: Add modern security headers, remove `X-XSS-Protection`.

### MEDIUM

#### IM1. Manifest-gen mounts entire project root

- **File**: `docker-compose.yml:5-6`
- **Details**: Grants script access to `.env`, `.git`, all source code when it only needs `manifests-templates/`.

#### IM2. Healthcheck hardcodes port 3003

- **File**: `backend/Dockerfile:12-13`
- **Details**: If `PORT` env var is changed, healthcheck always fails.

#### IM3. `npm install --production` deprecated

- **File**: `backend/Dockerfile:6`
- **Details**: Use `npm ci --omit=dev` with Node 22.

#### IM4. Dev files copied into build context

- **File**: `frontend/Dockerfile:7`
- **Details**: `COPY . .` includes `e2e/`, `playwright.config.ts` unnecessarily.

#### IM5. CORS leaks internal IP

- **File**: `docker-compose.yml:29`
- **Details**: Internal IP always added to CORS origins alongside public URL.

#### IM6. Empty `lang` attribute in index.html

- **File**: `frontend/index.html:2`
- **Details**: `<html lang="">` fails accessibility validation.

#### IM7. Outlook manifest missing AppDomains

- **File**: `manifests-templates/manifest-outlook.template.xml`
- **Details**: Office manifest has `<AppDomains>` but Outlook manifest does not.

#### IM8. CI infinite-loop guard fragile

- **File**: `.github/workflows/bump-version.yml:11, 37`
- **Details**: Relies on commit message prefix + `[skip ci]` suffix — neither fully robust alone.

### LOW

#### IL1. Vite config uses `.js` extension

- **File**: `frontend/vite.config.js`
- **Details**: Should be `.ts` to match the rest of the project.

#### IL2. `@types/diff-match-patch` in dependencies instead of devDependencies

- **File**: `frontend/package.json:15`

#### IL3. `chunkSizeWarningLimit` raised to suppress warnings

- **File**: `frontend/vite.config.js:56-57`
- **Details**: Masks bundle-size regressions.

#### IL4. Obsolete IE meta tag

- **File**: `frontend/index.html:5`
- **Details**: `<meta http-equiv="X-UA-Compatible" content="IE=edge" />` is inert.

#### IL5. Unused PUID/PGID env vars in docker-compose

- **File**: `docker-compose.yml:31-32, 66-67`
- **Category**: Dead Code
- **Details**: Not consumed by standard Docker images.

#### IL6. Dockerfile HEALTHCHECK overridden by compose

- **File**: `backend/Dockerfile:12-13`
- **Category**: Dead Code
- **Details**: Never executed when running via docker-compose.

#### IL7. Legacy entries in .gitignore

- **File**: `.gitignore:31-38`
- **Category**: Dead Code
- **Details**: References to `word-GPT-Plus-master.zip`, `litellm-local-proxy/.auth.env`, `Open_Excel/`.

---

## 6. Dead Code Registry

Consolidated list of all dead code found across the codebase.

### Backend Dead Code

| ID | File | Item | Details |
|----|------|------|---------|
| BD1 | `backend/src/middleware/validate.js:219-221` | `validateMaxTokens`, `validateTemperature`, `validateTools` exports | Only used internally |
| BD2 | `backend/src/services/llmClient.js:10-29` | `TIMEOUTS`, `getChatTimeoutMs`, `getImageTimeoutMs` exports | Only used internally |
| BD3 | `backend/src/middleware/validate.js:159` | `routeName` parameter | Never referenced in function body |
| BD4 | `backend/src/routes/image.js:17-19` | `if (!imageModel)` branch | Can never be true |

### Frontend Utilities Dead Code

| ID | File | Item | Details |
|----|------|------|---------|
| UD1 | `frontend/src/utils/hostDetection.ts:55-57` | `getHostName()` export | Never imported anywhere |
| UD2 | `frontend/src/utils/excelTools.ts:2149-2151` | `getExcelTool()` export | Never imported anywhere |
| UD3 | `frontend/src/utils/wordTools.ts:1889-1891` | `getWordTool()` export | Never imported anywhere |
| UD4 | `frontend/src/utils/outlookTools.ts:571-573` | `getOutlookTool()` export | Never imported anywhere |
| UD5 | `frontend/src/utils/powerpointTools.ts:949-951` | `getPowerPointTool()` export | Never imported anywhere |
| UD6 | `frontend/src/utils/generalTools.ts:98-103` | `getEnabledGeneralTools()` export | Never imported anywhere |
| UD7 | `frontend/src/utils/powerpointTools.ts:63-66` | `normalizePowerPointListText()` export | Never imported; callers use `stripMarkdownListMarkers` directly |
| UD8 | `frontend/src/utils/toolStorage.ts:10` | `buildToolSignature()` export | Only used internally |
| UD9 | `frontend/src/utils/credentialStorage.ts:171-180` | `credentialStorage` object export | All consumers import named exports directly |
| UD10 | `frontend/src/utils/credentialStorage.ts:40-49` | `CredentialStorage` interface export | Never imported |
| UD11 | `frontend/src/utils/excelTools.ts:66` | `'eval_officejs'` in `ExcelToolName` | No tool definition exists |
| UD12 | `frontend/src/utils/excelTools.ts:753` | `dataRange` variable in `sortRange` | Assigned but never read |
| UD13 | `frontend/src/utils/common.ts:3-13` | `getOptionList()` export | Only used internally |

### Frontend Composables Dead Code

| ID | File | Item | Details |
|----|------|------|---------|
| CD1 | `frontend/src/composables/useAgentPrompts.ts:13, 26` | `hostIsWord` option | Destructured but never referenced |

### Infrastructure Dead Code

| ID | File | Item | Details |
|----|------|------|---------|
| ID1 | `docker-compose.yml:31-32, 66-67` | `PUID`/`PGID` env vars | Not consumed by standard images |
| ID2 | `backend/Dockerfile:12-13` | Dockerfile HEALTHCHECK | Overridden by compose |
| ID3 | `backend/Dockerfile:10` | `EXPOSE 3003` | Purely documentary, port set by compose |
| ID4 | `.gitignore:31-38` | Legacy file references | `word-GPT-Plus-master.zip`, `Open_Excel/`, etc. |

---

## 7. v2 Open Feature Issues (carried forward)

### F1. Quick actions strip images/formatting from documents
- **Status**: OPEN
- **Severity**: HIGH
- **Details**: See v2 DESIGN_REVIEW for full implementation plan.

### F2. Outlook Reply produces low-quality responses
- **Status**: OPEN
- **Severity**: HIGH
- **Details**: See v2 DESIGN_REVIEW for full implementation plan.

### F3. Excel agent processes cells one-by-one
- **Status**: OPEN
- **Severity**: HIGH
- **Details**: See v2 DESIGN_REVIEW for full implementation plan.

---

## 8. Build & Environment Warnings (carried forward from v2)

### B1. No unit test infrastructure
- **Status**: OPEN

### B2. No linting or formatting configuration
- **Status**: OPEN

### B3. No CI pipeline for automated testing
- **Status**: OPEN

---

## 9. Pages, Components, API Layer & Types Findings

### CRITICAL

#### PC1. `keep-alive` never caches `HomePage.vue`

- **File**: `frontend/src/App.vue:4`, `frontend/src/pages/HomePage.vue`
- **Category**: Broken Functionality
- **Details**: `<keep-alive include="Home">` filters by component name, but `HomePage.vue` uses `<script setup>` without `defineOptions({ name: 'Home' })`. The auto-inferred name is `"HomePage"`, not `"Home"`. The component is destroyed and recreated on every navigation from Settings back to Home.
- **Impact**: All transient state lost on navigation (scroll position, textarea, `abortController`). Repeated backend health checks, flickering.
- **Fix**: Add `defineOptions({ name: 'Home' })` to `HomePage.vue`.

### HIGH

#### PH1. CSS typo — `itemse-center` instead of `items-center`

- **File**: `frontend/src/pages/HomePage.vue:3`
- **Category**: UI Bug
- **Details**: Tailwind class `"itemse-center"` is silently ignored. Container cross-axis alignment not applied.
- **Impact**: Broken vertical centering on the home page root container.
- **Fix**: Change to `items-center`.

#### PH2. `startNewChat` uses `window.location.reload()` — destructive

- **File**: `frontend/src/pages/HomePage.vue:539`
- **Category**: Architecture / UX
- **Details**: Full page reload instead of reactive state clearing. Discards all in-memory state, is slow, breaks SPA semantics.
- **Impact**: Slow UX, complete app reload, potential loss of Office context.

#### PH3. `agentMaxIterations` not validated on HomePage

- **File**: `frontend/src/pages/HomePage.vue:193`
- **Category**: Logic Bug
- **Details**: `SettingsPage` validates and clamps between 1-100, but `HomePage` reads the raw `useStorage` value. Corrupted localStorage (0, -1, 999999) goes directly to the agent loop.
- **Impact**: Runaway agent loop or zero iterations if localStorage is tampered with.

#### PH4. File upload `.xls` accepted by HTML but rejected by JS

- **File**: `frontend/src/components/chat/ChatInput.vue:89` vs `ChatInput.vue:233`
- **Category**: Logic Bug
- **Details**: HTML `accept` attribute includes `.xls`, but `allowedExtensions` in `processFiles()` does not. Users select `.xls` files, they are silently dropped.
- **Impact**: Silent data loss — user thinks file is attached but it is discarded.

#### PH5. File rejection is completely silent

- **File**: `frontend/src/components/chat/ChatInput.vue:231-256`
- **Category**: UX
- **Details**: When a file exceeds 10MB, has wrong type, or exceeds 3-file limit, no feedback is shown. File is simply not added to `attachedFiles`.
- **Impact**: Users don't know why their file was not attached.
- **Fix**: Show a toast with the rejection reason.

#### AH1. `fetchModels()` missing credential headers

- **File**: `frontend/src/api/backend.ts:90-94`
- **Category**: Security / Logic Bug
- **Details**: `fetchModels()` does not include `getUserCredentialHeaders()`, unlike all other API calls. Will fail if the backend requires authentication.
- **Impact**: Models endpoint may return errors or incomplete data.
- **Fix**: Add `...getUserCredentialHeaders()` to the headers.

#### AH2. `healthCheck()` missing credential headers

- **File**: `frontend/src/api/backend.ts:96-103`
- **Category**: Security / Logic Bug
- **Details**: Same as AH1 — no credential headers on health check.
- **Impact**: Backend appears permanently offline if authentication required.

#### XH1. No CSRF protection on API calls

- **File**: `frontend/src/api/backend.ts` (all POST endpoints)
- **Category**: Security
- **Details**: POST requests include credential headers but no CSRF token. Custom headers provide partial CORS-based protection, but no explicit CSRF defense.
- **Impact**: Potential exploitation if backend uses cookie-based sessions alongside custom headers.

### MEDIUM

#### PM1. Hardcoded French strings in ChatInput

- **File**: `frontend/src/components/chat/ChatInput.vue:47, 79`
- **Category**: i18n
- **Details**: `"Retirer le fichier"` and `"Attacher un document (PDF, DOCX, XLSX)"` hardcoded in French.
- **Fix**: Use `t()` with i18n keys.

#### PM2. Hardcoded English strings with fallback pattern in SettingsPage

- **File**: `frontend/src/pages/SettingsPage.vue:190-193, 200, 470`
- **Category**: i18n
- **Details**: `$t("darkModeLabel") || "Dark mode"` pattern suggests missing i18n keys. Fallbacks mask the issue.

#### PM3. `CustomInput` type flash on mount

- **File**: `frontend/src/components/CustomInput.vue:50-76`
- **Category**: UI Bug
- **Details**: `type` ref initialized to `'text'`, then overridden in `onMounted`. Brief flash where a number input appears as text.
- **Fix**: Initialize from prop: `const type = ref(isPassword ? 'password' : inputType)`.

#### PM4. `CustomInput` model has `any` type

- **File**: `frontend/src/components/CustomInput.vue:36`
- **Category**: Type Safety
- **Details**: `defineModel<any>()` loses all type safety.

#### PM5. `SingleSelect` dropdown positioning without scroll listener

- **File**: `frontend/src/components/SingleSelect.vue:65-96`
- **Category**: UI Bug
- **Details**: Dropdown uses `position: fixed` calculated on toggle, but no scroll/resize recalculation.
- **Impact**: Mispositioned dropdown when settings page is scrolled while open.

#### PM6. Dual emit pattern in `SingleSelect`

- **File**: `frontend/src/components/SingleSelect.vue:42, 48-52`
- **Category**: Code Quality
- **Details**: Both `update:modelValue` and `change` emitted. Redundant and error-prone.

#### PM7. `SettingCard` prop `p1` never used by any consumer

- **File**: `frontend/src/components/SettingCard.vue:2, 9-10`
- **Category**: Dead Code

#### PM8. `Message.vue` setTimeout without cleanup

- **File**: `frontend/src/components/Message.vue:38-45`
- **Category**: Resource Leak
- **Details**: Two `setTimeout` calls in `onMounted` never cleared. Callbacks fire on unmounted components.

#### PM9. `ChatHeader.vue` hardcoded English string

- **File**: `frontend/src/components/chat/ChatHeader.vue:13`
- **Category**: i18n
- **Details**: `"AI Office Assistant"` hardcoded. Not translatable.

#### PM10. Mixed `t()` and `$t()` usage

- **Files**: `HomePage.vue`, `SettingsPage.vue`
- **Category**: Consistency
- **Details**: Inconsistent use of composition API `t()` vs global `$t()` in templates.

#### PM11. `expandedThoughts` grows unbounded

- **File**: `frontend/src/components/chat/ChatMessageList.vue:164`
- **Category**: Memory
- **Details**: Record keyed by message+segment index. Old entries never cleaned up.

#### AM1. Import statement in middle of file

- **File**: `frontend/src/api/backend.ts:79`
- **Category**: Style
- **Details**: `import { getUserKey, getUserEmail }` appears after function definitions.

#### AM2. `chatStream` silently swallows JSON parse errors

- **File**: `frontend/src/api/backend.ts:185-187`
- **Category**: Error Handling
- **Details**: Empty `catch {}` block drops malformed SSE data without logging.

#### AM3. `chatStream` discards remaining buffer after stream ends

- **File**: `frontend/src/api/backend.ts:157-189`
- **Category**: Logic Bug
- **Details**: When `done` is true, final buffer content without trailing newline is lost.
- **Impact**: Potential loss of last streamed token.

#### AM4. Duplicate `ToolDefinition` interface

- **Files**: `frontend/src/api/backend.ts:192-200`, `frontend/src/types/index.d.ts:61-67`
- **Category**: Type Safety
- **Details**: Two different interfaces with the same name — API wire format vs internal tool definition. Name collision causes confusion.

#### TM1. Global ambient types without explicit imports

- **File**: `frontend/src/types/index.d.ts:32-74`
- **Category**: Type Safety
- **Details**: All types declared ambient (no `export`), available everywhere without imports. Bypasses module boundaries.

#### TM2. `OfficeHostType` declared in two files

- **Files**: `frontend/src/types/index.d.ts:74`, `frontend/src/utils/hostDetection.ts:1`
- **Category**: Inconsistency
- **Details**: Two sources of truth for the same type.

#### EM1. `useStorage` called outside Vue component context

- **File**: `frontend/src/main.ts:22`
- **Category**: Code Quality
- **Details**: VueUse composable called in `Office.onReady` callback, outside any component `setup()`. May break with future VueUse versions.

#### EM2. Global `ResizeObserver` monkey-patching

- **File**: `frontend/src/main.ts:15-19`
- **Category**: Code Quality
- **Details**: Global `window.ResizeObserver` replaced with debounced version. Affects all code including third-party libraries.

#### XM1. Deeply nested ternary chains repeated 10+ times

- **Files**: `HomePage.vue:31-38, 67-73, 166-174, 355-361`, `SettingsPage.vue:771-777, 781-787, 789-795, 887-894, 896-903, 916-922`
- **Category**: Code Quality / DRY
- **Details**: `hostIsOutlook ? ... : hostIsPowerPoint ? ... : hostIsExcel ? ... : ...` repeated throughout.
- **Fix**: Extract into a utility function `forHost({ outlook, powerpoint, excel, word })`.

#### XM2. Quick action arrays not reactive to locale changes

- **File**: `frontend/src/pages/HomePage.vue:206-351`
- **Category**: i18n / Reactivity
- **Details**: `wordQuickActions`, `outlookQuickActions`, `powerPointQuickActions` are plain arrays with `t()` at setup time. Only `excelQuickActions` uses `computed()`. Labels won't update on locale change.
- **Fix**: Wrap all quick action arrays in `computed()`.

### LOW

#### PL1. `SettingSection.vue` component never imported or used

- **File**: `frontend/src/components/SettingSection.vue`
- **Category**: Dead Code / Dead File

#### PL2. `CustomButton` `icon` prop typed as `any`

- **File**: `frontend/src/components/CustomButton.vue:43`
- **Details**: Should be `Component | null`.

#### PL3. `SingleSelect` multiple props typed as `any`

- **File**: `frontend/src/components/SingleSelect.vue:44, 107, 117, 119`
- **Details**: `modelValue`, `placeholder`, `icon`, `customFrontIcon` all `any`.

#### PL4. `ChatInput` emits `"input"` event nobody listens to

- **File**: `frontend/src/components/chat/ChatInput.vue:177, 191`
- **Category**: Dead Code

#### PL5. `App.vue` has empty `<script>` block

- **File**: `frontend/src/App.vue:11`
- **Category**: Dead Code

#### AL1. `api/common.ts` is misplaced — contains Word-specific Office logic

- **File**: `frontend/src/api/common.ts`
- **Category**: Architecture
- **Details**: Contains `Word.run`, `insertText`, `insertParagraph` and WordFormatter dependency. Not a generic API utility.

#### TL1. Tool type aliases add no value

- **File**: `frontend/src/types/index.d.ts:69-72`
- **Details**: `WordToolDefinition = ToolDefinition`, `ExcelToolDefinition = ToolDefinition`, etc. — zero differentiation.

#### TL2. `insertTypes` uses lowercase, plural name

- **File**: `frontend/src/types/index.d.ts:34`
- **Details**: Should be `InsertType` (PascalCase, singular) per TypeScript conventions.

### Pages/Components Dead Code

| ID | File | Item | Details |
|----|------|------|---------|
| PD1 | `frontend/src/pages/HomePage.vue:92` | `Briefcase` import | Never used in template or script |
| PD2 | `frontend/src/pages/HomePage.vue:94` | `CheckCircle` import | Never used anywhere |
| PD3 | `frontend/src/components/SettingSection.vue` | Entire component file | Never imported or used |
| PD4 | `frontend/src/components/chat/ChatInput.vue:210` | `handleDragLeave` param `e` | Declared but never read |
| PD5 | `frontend/src/components/chat/ChatInput.vue:177, 191` | `"input"` emit | Emitted but no consumer listens |
| PD6 | `frontend/src/components/SettingCard.vue:2` | `p1` prop | Never passed by any consumer |
| PD7 | `frontend/src/App.vue:11` | Empty `<script>` block | No code inside |

---

## 10. Summary Statistics

| Area | CRITICAL | HIGH | MEDIUM | LOW | Dead Code | Total |
|------|----------|------|--------|-----|-----------|-------|
| Backend | 4 | 7 | 10 | 4 | 4 | 29 |
| Frontend Utils | 3 | 7 | 10 | 4 | 13 | 37 |
| Composables | 2 | 7 | 11 | 5 | 1 | 26 |
| Infrastructure | 3 | 5 | 8 | 7 | 4 | 27 |
| Pages/Components/API | 1 | 8 | 19 | 8 | 7 | 43 |
| **Total** | **13** | **34** | **58** | **28** | **29** | **162** |

---

## 11. Priority Recommendations

### Immediate (CRITICAL — fix now)

1. **BC1/IC1** — Exempt `/api/upload` from Content-Type middleware (upload feature broken)
2. **UC3** — Sanitize HTML before Outlook email injection (XSS in outgoing emails)
3. **BC3** — Add log rotation and redact user content from logs (GDPR/privacy)
4. **CC1/CC2** — Sanitize document content before LLM prompt interpolation (prompt injection)
5. **UC1** — Use function replacement in `String.replace()` (data corruption)
6. **BC2/IC3** — Replace internal URL with placeholder in `.env.example`
7. **IC2** — Add non-root users to Dockerfiles
8. **PC1** — Add `defineOptions({ name: 'Home' })` to fix keep-alive caching

### Short-term (HIGH — fix before next release)

9. **PH1** — Fix CSS typo `itemse-center` → `items-center`
10. **UH1** — Add or remove `eval_officejs` from ExcelToolName
11. **UH2** — Fix column letter arithmetic for multi-char columns
12. **BH1** — Fix drain event listener leak in streaming
13. **BH2** — Check `res.headersSent` before error response
14. **CH1** — Fix `sendMessage` race condition
15. **AH1/AH2** — Add credential headers to `fetchModels()` and `healthCheck()`
16. **PH4** — Synchronize `.xls` between HTML accept and JS validation
17. **PH5** — Add user feedback when files are rejected
18. **IH1** — Standardize Node.js version
19. **IH4** — Copy lock files in Dockerfiles, use `npm ci`

### Medium-term (MEDIUM — address in upcoming sprints)

20. Remove all dead code (29 items across codebase)
21. Fix i18n violations: hardcoded French/English strings (PM1, PM9, CM1, BL2)
22. Fix error handling: add logging to silent catch blocks
23. Replace `any` types with `unknown` + type guards
24. Extract shared utilities (deduplicate `generateVisualDiff`, `withTimeout`, `forHost`)
25. Decompose oversized functions (3 functions >180 lines each)
26. Add security headers to nginx config
27. Wrap all quick action arrays in `computed()` for locale reactivity

---

*Last updated: 2026-03-01*
