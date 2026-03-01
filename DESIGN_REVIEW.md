# Design Review & Code Audit — v3

**Date**: 2026-03-01
**Scope**: Full codebase audit — security, logic bugs, code quality, dead code, architecture across backend, frontend utilities, composables, pages/components, API layer, and infrastructure.

---

## 1. Executive Summary

KickOffice is a Microsoft Office add-in (Word, Excel, PowerPoint, Outlook) powered by a Vue 3 + Vite frontend and Express.js backend proxy for OpenAI-compatible LLM APIs.

**Previous audits**: v1 (2026-02-21, 38 issues — all resolved) and v2 (2026-02-22, 28 issues — all resolved except B1-B3 build warnings and F1-F3 feature issues).

This v3 audit is a **fresh, comprehensive analysis** of the entire codebase. Findings are organized by codebase area, then by severity.

### v2 Open Items Status

| ID  | Description                                  | Status |
| --- | -------------------------------------------- | ------ |
| B1  | No unit test infrastructure                  | OPEN   |
| B2  | No linting or formatting configuration       | OPEN   |
| B3  | No CI pipeline for automated testing         | OPEN   |
| F1  | Quick actions strip images/formatting        | OPEN   |
| F2  | Outlook Reply produces low-quality responses | OPEN   |
| F3  | Excel agent processes cells one-by-one       | OPEN   |

---

## 2. Backend Findings

### CRITICAL

#### BC1. Content-Type enforcement blocks file uploads [RESOLVED]

> **Status**: Implemented.

#### BC2. Internal LLM API URL exposed in source and .env.example [RESOLVED]

> **Status**: Implemented.

#### BC3. Sensitive data logged to disk in plaintext [RESOLVED]

> **Status**: Implemented.

#### BC4. User-supplied credentials forwarded without sanitization [RESOLVED]

> **Status**: Implemented. `sanitizeHeaderValue()` strips `\r`, `\n`, and non-printable chars from `X-User-Key` and `X-OpenWebUi-User-Email` before forwarding.

### HIGH

#### BH1. Drain event listener leak in streaming response [RESOLVED]

> **Status**: Implemented.

#### BH2. logAndRespond called after headers already sent (streaming) [RESOLVED]

> **Status**: Implemented.

#### BH3. Unbounded log file growth [RESOLVED]

> **Status**: Implemented.

#### BH4. Hardcoded version in health endpoint [RESOLVED]

> **Status**: Implemented. Version now read from `package.json` via `createRequire`.

#### BH5. parsePositiveInt allows zero [RESOLVED]

> **Status**: Implemented. Check changed to `<= 0` to reject zero values.

#### BH6. Upload route lacks magic-byte file validation [OPEN]

- **File**: `backend/src/routes/upload.js:38-78`
- **Category**: Security
- **Details**: File type detection relies on client-controlled MIME type or extension. No magic-byte validation.
- **Impact**: Attackers can upload crafted files (zip bombs via XLSX, XXE via DOCX) that exploit parsing libraries.
- **Fix**: Add magic-byte validation (e.g. `file-type` package) before processing.

#### BH7. ReDoS potential in sanitizeErrorText [RESOLVED]

> **Status**: Implemented. Regex patterns pre-compiled at module load via `SENSITIVE_HEADER_REGEXES` constant array.

### MEDIUM

#### BM1. No graceful shutdown handling [RESOLVED]

> **Status**: Implemented. SIGTERM/SIGINT handlers added to `server.js` — stops accepting connections, waits for drain, force-exits after 30s.

#### BM2. Unused `routeName` parameter in `validateChatRequest` [RESOLVED]

> **Status**: Implemented.

#### BM3. Exported functions never imported externally [RESOLVED]

> **Status**: Implemented.

#### BM4. Exported constants/functions never imported externally [RESOLVED]

> **Status**: Implemented.

#### BM5. Validated values discarded in validateChatRequest [RESOLVED]

> **Status**: Implemented.

#### BM6. Inconsistent error logging patterns [OPEN]

- **File**: Multiple backend files
- **Category**: Code Quality
- **Details**: Four different logging patterns (`console.error`, `systemLog`, `logAndRespond`, `process.stderr.write`) used inconsistently. Many errors logged through both `systemLog` and `console.error` on consecutive lines.

#### BM7. `handleErrorResponse` return value discarded [RESOLVED]

> **Status**: Implemented. `handleErrorResponse` made void (no return value) since callers don't use the return.

#### BM8. `allCsv` declared with `let` instead of `const` [RESOLVED]

> **Status**: Implemented.

#### BM9. No multer field count limits [RESOLVED]

> **Status**: Implemented. `limits.fields: 10` and `limits.fieldSize: 1024` configured in multer.

#### BM10. No request ID / correlation [OPEN]

- **File**: Multiple backend files
- **Category**: Architecture
- **Details**: No request ID generation. Cannot correlate client-side errors with server-side logs.

### LOW

#### BL1. Dead branch: `if (!imageModel)` check [RESOLVED]

> **Status**: Implemented.

#### BL2. French strings hardcoded in backend [RESOLVED]

> **Status**: Implemented.

#### BL3. Stale comment about character limit [RESOLVED]

> **Status**: Implemented. Comment updated to "100k chars is roughly 25-30k tokens".

#### BL4. `isPlainObject` accepts non-plain objects [RESOLVED]

> **Status**: Implemented. Now uses `Object.getPrototypeOf` check to only accept true plain objects.

---

#### UC1. Prompt injection via custom prompt templates [RESOLVED]

> **Status**: Implemented.

#### UC2. XOR "obfuscation" provides false security for API keys [OPEN]

- **File**: `frontend/src/utils/credentialStorage.ts:7-36`
- **Category**: Security
- **Details**: API keys XOR'd with hardcoded key `'K1ck0ff1c3'` then base64-encoded in localStorage. Trivially reversible with browser dev tools.
- **Impact**: Any XSS vulnerability allows immediate credential theft. The obfuscation creates a false sense of security.
- **Fix**: Document the limitation clearly. Consider using session-only storage or backend-managed tokens.

#### UC3. Unsanitized HTML injection in Outlook tools [RESOLVED]

> **Status**: Implemented.

### HIGH

#### UH1. `eval_officejs` declared in ExcelToolName but never defined [RESOLVED]

> **Status**: Implemented.

#### UH2. Column letter arithmetic overflow [RESOLVED]

> **Status**: Implemented.

#### UH3. Double timeout in Outlook tool execution [RESOLVED]

> **Status**: Implemented.

#### UH4. Language parameter ignored in translate prompt [RESOLVED]

> **Status**: Implemented.

#### UH5. Host detection caching can return wrong host [OPEN]

- **File**: `frontend/src/utils/hostDetection.ts:3-37`
- **Category**: Logic Bug
- **Details**: `detectOfficeHost()` caches result in a module-level variable. If called before `Office.onReady` fires, may cache an incorrect host permanently.
- **Impact**: Wrong host detection in edge cases, leading to wrong tool set and prompts.
- **Fix**: Only cache after `Office.onReady` has resolved.

#### UH6. Message toast singleton race condition [RESOLVED]

> **Status**: Implemented. Pending timeout cleared before creating new instance, cleanup handled in onClose.

#### UH7. `html: true` in MarkdownIt with `style` in DOMPurify allowlist [RESOLVED]

> **Status**: Implemented.

### MEDIUM

#### UM1. Massive type unsafety with `as unknown as` casts [OPEN]

- **Files**: `excelTools.ts:21`, `wordTools.ts:194`, `outlookTools.ts:75`, `powerpointTools.ts:39`
- **Category**: Type Safety
- **Details**: All four tool-creation factories use `as unknown as Record<ToolName, ToolDefinition>`, completely disabling TypeScript checking for missing tool definitions.
- **Fix**: Use a type-safe builder that validates all required tool names are present.

#### UM2. Pervasive `any` types in tool definitions [OPEN]

- **Files**: All tool definition files
- **Category**: Type Safety
- **Details**: `args: Record<string, any>`, `mailbox: any`, `context: any` across all tool files. Outlook tools especially: `getMailbox(): any`, `getOfficeAsyncStatus(): any`.

#### UM3. Duplicated `generateVisualDiff` function [RESOLVED]

> **Status**: Implemented.

#### UM4. Duplicated Office API helpers [RESOLVED]

> **Status**: Implemented.

#### UM5. `Ref` without type parameter in WordFormatter [RESOLVED]

> **Status**: Implemented. `insertType: Ref` changed to `insertType: Ref<string>` in both method signatures.

#### UM6. `searchAndReplace` tools labeled as category `'read'` [RESOLVED]

> **Status**: Implemented. Category changed to `'write'` in both excelTools.ts and wordTools.ts.

#### UM7. Redundant Set + Array checks in toolStorage [RESOLVED]

> **Status**: Implemented. Redundant `.includes()` check removed — only `storedEnabledSet.has()` retained.

#### UM8. No `QuotaExceededError` handling for localStorage [OPEN]

- **Files**: `credentialStorage.ts`, `toolStorage.ts`, `savedPrompts.ts`, `constant.ts`
- **Category**: Error Handling
- **Details**: Multiple files write to localStorage without catching `QuotaExceededError`.

#### UM9. `tokenManager.ts` mutates input messages [RESOLVED]

> **Status**: Implemented.

#### UM10. Character-by-character HTML reconstruction in PowerPoint [DEFERRED]

- **File**: `frontend/src/api/common.ts:169-173`
- **Category**: Performance / UX
- **Details**: Word processing uses `insertHtml`, but PowerPoint inserts character by character.
- **Impact**: Noticeably slow insertion in PowerPoint, potential formatting loss.
- **Fix**: Find API equivalent to `insertHtml` for PowerPoint if possible.
- **Note**: Retaining this implementation as PowerPoint has severe formatting issues otherwise.

### LOW

#### UL1. Typo in export name `buildInPrompt` [RESOLVED]

> **Status**: Implemented. Renamed to `builtInPrompt` across constant.ts, SettingsPage.vue, and useAgentLoop.ts.

#### UL2. `deleteText` reports success when no text selected [RESOLVED]

> **Status**: Implemented. `deleteText` now checks `range.text.length > 0` before proceeding.

#### UL3. Inconsistent error handling strategy across tools [OPEN]

- **Files**: All tool files
- **Details**: Some return error strings, some throw, some return empty strings. Caller must check string prefixes.

#### UL4. `markdown.ts` vs `officeRichText.ts` naming confusion [OPEN]

- **Details**: Both render Markdown but for different targets (chat vs Office). Names don't communicate this.

---

#### CC1. Prompt injection via unsanitized document selection [RESOLVED]

> **Status**: Implemented.

#### CC2. Prompt injection via quick action selection text [RESOLVED]

> **Status**: Implemented.

### HIGH

#### CH1. Race condition: concurrent `sendMessage` calls corrupt state [RESOLVED]

> **Status**: Implemented.

#### CH2. `lastIndex` stale reference during agent loop [RESOLVED]

> **Status**: Implemented.

#### CH3. Timer leak — `timeoutId` reassigned without clearing [RESOLVED]

> **Status**: Implemented.

#### CH4. Raw `err.message` displayed to users [OPEN]

- **File**: `frontend/src/composables/useAgentLoop.ts:304, 435, 551, 819`
- **Category**: Security / Information Disclosure
- **Details**: Server error messages (potentially containing internal URLs, API keys, stack traces) shown directly to users.
- **Fix**: Show generic messages, log details to console only.

#### CH5. `any` types on error parameters and tool args [OPEN]

- **File**: `frontend/src/composables/useAgentLoop.ts:111, 288, 334, 433, 544`
- **Category**: Type Safety
- **Details**: `isCredentialError(error: any)`, multiple `catch (err: any)`, `toolArgs: Record<string, any>`.
- **Fix**: Use `unknown` with type guards.

#### CH6. XSS via unvalidated `imageSrc` URL [RESOLVED]

> **Status**: Implemented. URL validation added: data URLs checked for `data:image/` prefix, external URLs validated for `http:`/`https:` protocol only.

#### CH7. `THINK_TAG_REGEX` module-level with `g` flag — maintenance hazard [RESOLVED]

> **Status**: Implemented.

### MEDIUM

#### CM1. Hardcoded French string in file upload error [RESOLVED]

> **Status**: Implemented.

#### CM2. `buildChatMessages` drops system messages [RESOLVED]

> **Status**: Implemented.

#### CM3. Overly large functions [OPEN]

- **Files**: `useAgentLoop.ts:449-638` (`sendMessage` ~190 lines), `useAgentLoop.ts:640-828` (`applyQuickAction` ~188 lines), `useAgentLoop.ts:227-421` (`runAgentLoop` ~195 lines)
- **Category**: Architecture
- **Fix**: Decompose into focused helpers.

#### CM4. `insertToDocument` silently swallows all errors [RESOLVED]

> **Status**: Implemented. All catch blocks now include `console.warn` logging before clipboard fallback.

#### CM5. Promise constructor anti-pattern in Outlook functions [RESOLVED]

> **Status**: Implemented. Shared `withTimeout` helper extracted and used throughout useOfficeSelection.ts.

#### CM6. Timeout promises create orphaned timers [RESOLVED]

> **Status**: Implemented. `clearTimeout` called in Promise.race cleanup to prevent orphaned timers.

#### CM7. Excel selection returns unescaped tab-separated values [OPEN]

- **File**: `frontend/src/composables/useOfficeSelection.ts:86-92`
- **Details**: Cell values containing tabs/newlines make output ambiguous.

#### CM8. HTML injection via `richHtml` to Office APIs [RESOLVED]

> **Status**: Implemented. `richHtml` sanitized through `DOMPurify.sanitize()` before passing to Office APIs.

#### CM9. Prompt injection via user profile fields [RESOLVED]

> **Status**: Implemented. `firstName`/`lastName` sanitized via `sanitize()` helper before interpolation into system prompt.

#### CM10. `insertImageToPowerPoint` ignores `'NoAction'` semantics [OPEN]

- **File**: `frontend/src/composables/useImageActions.ts:111-157`
- **Details**: `'NoAction'` should mean "do nothing" but still inserts the image.

#### CM11. Hidden side effect: `insertType.value` mutation [OPEN]

- **File**: `frontend/src/composables/useOfficeInsert.ts:140`
- **Details**: Mutates external ref as side effect, causing unexpected reactivity triggers.

### LOW

#### CL1. `hostIsWord` parameter accepted but never used [RESOLVED]

> **Status**: Implemented. `hostIsWord` parameter removed from `UseAgentPromptsOptions` interface.

#### CL2. `cleanContent` and `splitThinkSegments` use different think-tag logic [OPEN]

- **File**: `frontend/src/composables/useImageActions.ts:13-33, 40-42`
- **Details**: Inconsistent behavior for malformed tags.

#### CL3. Inconsistent image insert error reporting across hosts [OPEN]

- **File**: `frontend/src/composables/useOfficeInsert.ts:169-198`
- **Details**: Outlook falls through and shows misleading "imageInsertWordOnly" message.

#### CL4. `payload` parameter typed as `unknown` — should be `string | undefined` [OPEN]

- **File**: `frontend/src/composables/useAgentLoop.ts:449`

#### CL5. Word HTML selection swallows errors silently [OPEN]

- **File**: `frontend/src/composables/useOfficeSelection.ts:135-136`

---

#### IC1. Content-Type middleware blocks uploads (same as BC1) [RESOLVED]

> **Status**: Implemented.

#### IC2. Containers run as root [DEFERRED]

- **File**: `frontend/Dockerfile:1`, `backend/Dockerfile:1`
- **Category**: Security
- **Details**: `node` Docker images run as `root` by default.
- **Fix**: Add `USER node` to the Dockerfiles.
- **Note**: Retaining this configuration deliberately.

#### IC3. Internal infrastructure URL as default [RESOLVED]

> **Status**: Implemented.

### HIGH

#### IH1. Node.js version mismatch between environments [RESOLVED]

> **Status**: Implemented.

#### IH2. Private IP baked into frontend Docker build [DEFERRED]

- **File**: `frontend/Dockerfile:9`
- **Category**: Security
- **Details**: Default build arg `VITE_BACKEND_URL=http://192.168.50.10:3003` bakes a private IP into the JS bundle.
- **Fix**: Remove default or use a placeholder that fails visibly.
- **Note**: Retaining this configuration deliberately.

#### IH3. External DuckDNS domain as default in .env.example [DEFERRED]

- **File**: `.env.example:10-11`
- **Category**: Misconfiguration
- **Details**: `PUBLIC_FRONTEND_URL` and `PUBLIC_BACKEND_URL` set to `https://kickoffice.duckdns.org` as active values.
- **Fix**: Comment them out or use clearly fake placeholders.
- **Note**: Retaining this configuration deliberately.

#### IH4. Undeterministic package resolution in Dockerfiles [RESOLVED]

> **Status**: Implemented.

#### IH5. Nginx missing security headers [RESOLVED]

> **Status**: Implemented.

### MEDIUM

#### IM1. Manifest-gen mounts entire project root [OPEN]

- **File**: `docker-compose.yml:5-6`
- **Details**: Grants script access to `.env`, `.git`, all source code when it only needs `manifests-templates/`.

#### IM2. Healthcheck hardcodes port 3003 [RESOLVED]

> **Status**: Implemented.

#### IM3. `npm install --production` deprecated [RESOLVED]

> **Status**: Implemented. Now uses `npm ci --omit=dev`.

#### IM4. Dev files copied into build context [OPEN]

- **File**: `frontend/Dockerfile:7`
- **Details**: `COPY . .` includes `e2e/`, `playwright.config.ts` unnecessarily.

#### IM5. CORS leaks internal IP [RESOLVED]

> **Status**: Implemented.

#### IM6. Empty `lang` attribute in index.html [RESOLVED]

> **Status**: Implemented.

#### IM7. Outlook manifest missing AppDomains [RESOLVED]

> **Status**: Implemented.

#### IM8. CI infinite-loop guard fragile [OPEN]

- **File**: `.github/workflows/bump-version.yml:11, 37`
- **Details**: Relies on commit message prefix + `[skip ci]` suffix — neither fully robust alone.

### LOW

#### IL1. Vite config uses `.js` extension [RESOLVED]

> **Status**: Implemented.

#### IL2. `@types/diff-match-patch` in dependencies instead of devDependencies [RESOLVED]

> **Status**: Implemented. Moved to devDependencies in frontend/package.json.

#### IL3. `chunkSizeWarningLimit` raised to suppress warnings [OPEN]

- **File**: `frontend/vite.config.js:56-57`
- **Details**: Masks bundle-size regressions.

#### IL4. Obsolete IE meta tag [RESOLVED]

> **Status**: Implemented. `<meta http-equiv="X-UA-Compatible" content="IE=edge" />` removed from index.html.

#### IL5. Unused PUID/PGID env vars in docker-compose [RESOLVED]

> **Status**: Implemented. PUID/PGID removed from both backend and frontend services in docker-compose.yml.

#### IL6. Dockerfile HEALTHCHECK overridden by compose [RESOLVED]

> **Status**: Implemented.

#### IL7. Legacy entries in .gitignore [RESOLVED]

> **Status**: Implemented. Legacy entries (`word-GPT-Plus-master.zip`, `litellm-local-proxy/.auth.env`, `Open_Excel/`) removed from .gitignore.

---

#### PC1. `keep-alive` never caches `HomePage.vue` [RESOLVED]

> **Status**: Implemented.

### HIGH

#### PH1. CSS typo — `itemse-center` instead of `items-center` [RESOLVED]

> **Status**: Implemented.

#### PH2. `startNewChat` uses `window.location.reload()` — destructive [RESOLVED]

> **Status**: Implemented.

#### PH3. `agentMaxIterations` not validated on HomePage [RESOLVED]

> **Status**: Implemented.

#### PH4. Discrepancy between HTML `accept` and JS extension validation [RESOLVED]

> **Status**: Implemented.

#### PH5. Silent failure when files exceed limits or have wrong type [RESOLVED]

> **Status**: Implemented.

#### AH1. Missing credential headers in `fetchModels` [RESOLVED]

> **Status**: Implemented.

#### AH2. `healthCheck()` missing credential headers [RESOLVED]

> **Status**: Implemented. `getUserCredentialHeaders()` added to healthCheck request.

#### XH1. No CSRF protection on API calls [OPEN]

- **File**: `frontend/src/api/backend.ts` (all POST endpoints)
- **Category**: Security
- **Details**: POST requests include credential headers but no CSRF token. Custom headers provide partial CORS-based protection, but no explicit CSRF defense.
- **Impact**: Potential exploitation if backend uses cookie-based sessions alongside custom headers.

### MEDIUM

#### PM1. Hardcoded French strings in ChatInput [OPEN]

- **File**: `frontend/src/components/chat/ChatInput.vue:47, 79`
- **Category**: i18n
- **Details**: `"Retirer le fichier"` and `"Attacher un document (PDF, DOCX, XLSX)"` hardcoded in French.
- **Fix**: Use `t()` with i18n keys.

#### PM2. Hardcoded English strings with fallback pattern in SettingsPage [OPEN]

- **File**: `frontend/src/pages/SettingsPage.vue:190-193, 200, 470`
- **Category**: i18n
- **Details**: `$t("darkModeLabel") || "Dark mode"` pattern suggests missing i18n keys. Fallbacks mask the issue.

#### PM3. `CustomInput` type flash on mount [RESOLVED]

> **Status**: Implemented. `type` ref initialized directly from prop: `const type = ref(isPassword ? 'password' : inputType)`.

#### PM4. `CustomInput` model has `any` type [OPEN]

- **File**: `frontend/src/components/CustomInput.vue:36`
- **Category**: Type Safety
- **Details**: `defineModel<any>()` loses all type safety.

#### PM5. `SingleSelect` dropdown positioning without scroll listener [OPEN]

- **File**: `frontend/src/components/SingleSelect.vue:65-96`
- **Category**: UI Bug
- **Details**: Dropdown uses `position: fixed` calculated on toggle, but no scroll/resize recalculation.
- **Impact**: Mispositioned dropdown when settings page is scrolled while open.

#### PM6. Dual emit pattern in `SingleSelect` [OPEN]

- **File**: `frontend/src/components/SingleSelect.vue:42, 48-52`
- **Category**: Code Quality
- **Details**: Both `update:modelValue` and `change` emitted. Redundant and error-prone.

#### PM7. `SettingCard` prop `p1` never used by any consumer [RESOLVED]

> **Status**: Implemented. `p1` prop removed from SettingCard.vue.

#### PM8. `Message.vue` setTimeout without cleanup [RESOLVED]

> **Status**: Implemented.

#### PM9. `ChatHeader.vue` hardcoded English string [RESOLVED]

> **Status**: Implemented. Now uses `t("appSubtitle")` — existing i18n key with EN/FR translations.

#### PM10. Mixed `t()` and `$t()` usage [OPEN]

- **Files**: `HomePage.vue`, `SettingsPage.vue`
- **Category**: Consistency
- **Details**: Inconsistent use of composition API `t()` vs global `$t()` in templates.

#### PM11. `expandedThoughts` grows unbounded [RESOLVED]

> **Status**: Implemented.

#### AM1. Import statement in middle of file [RESOLVED]

> **Status**: Implemented. Import moved to top of file.

#### AM2. `chatStream` silently swallows JSON parse errors [RESOLVED]

> **Status**: Implemented. Catch block includes explanatory comment for intentional silent handling of malformed SSE lines.

#### AM3. `chatStream` discards remaining buffer after stream ends [RESOLVED]

> **Status**: Implemented. Buffer is flushed correctly at stream end.

#### AM4. Duplicate `ToolDefinition` interface [OPEN]

- **Files**: `frontend/src/api/backend.ts:192-200`, `frontend/src/types/index.d.ts:61-67`
- **Category**: Type Safety
- **Details**: Two different interfaces with the same name — API wire format vs internal tool definition. Name collision causes confusion.

#### TM1. Global ambient types without explicit imports [OPEN]

- **File**: `frontend/src/types/index.d.ts:32-74`
- **Category**: Type Safety
- **Details**: All types declared ambient (no `export`), available everywhere without imports. Bypasses module boundaries.

#### TM2. `OfficeHostType` declared in two files [OPEN]

- **Files**: `frontend/src/types/index.d.ts:74`, `frontend/src/utils/hostDetection.ts:1`
- **Category**: Inconsistency
- **Details**: Two sources of truth for the same type.

#### EM1. `useStorage` called outside Vue component context [OPEN]

- **File**: `frontend/src/main.ts:22`
- **Category**: Code Quality
- **Details**: VueUse composable called in `Office.onReady` callback, outside any component `setup()`. May break with future VueUse versions.

#### EM2. Global `ResizeObserver` monkey-patching [OPEN]

- **File**: `frontend/src/main.ts:15-19`
- **Category**: Code Quality
- **Details**: Global `window.ResizeObserver` replaced with debounced version. Affects all code including third-party libraries.

#### XM1. Deeply nested ternary chains repeated 10+ times [OPEN]

- **Files**: `HomePage.vue:31-38, 67-73, 166-174, 355-361`, `SettingsPage.vue:771-777, 781-787, 789-795, 887-894, 896-903, 916-922`
- **Category**: Code Quality / DRY
- **Details**: `hostIsOutlook ? ... : hostIsPowerPoint ? ... : hostIsExcel ? ... : ...` repeated throughout.
- **Fix**: Extract into a utility function `forHost({ outlook, powerpoint, excel, word })`.

#### XM2. Quick action arrays not reactive to locale changes [RESOLVED]

> **Status**: Implemented. All quick action arrays (wordQuickActions, outlookQuickActions, powerPointQuickActions, excelQuickActions) now wrapped in `computed()`.
- **Fix**: Wrap all quick action arrays in `computed()`.

### LOW

#### PL1. `SettingSection.vue` component never imported or used [RESOLVED]

> **Status**: Implemented.

#### PL2. `CustomButton` `icon` prop typed as `any` [OPEN]

- **File**: `frontend/src/components/CustomButton.vue:43`
- **Details**: Should be `Component | null`.

#### PL3. `SingleSelect` multiple props typed as `any` [OPEN]

- **File**: `frontend/src/components/SingleSelect.vue:44, 107, 117, 119`
- **Details**: `modelValue`, `placeholder`, `icon`, `customFrontIcon` all `any`.

#### PL4. `ChatInput` emits `"input"` event nobody listens to [RESOLVED]

> **Status**: Implemented. Unused `"input"` emit removed from ChatInput.vue.

#### PL5. `App.vue` has empty `<script>` block [RESOLVED]

> **Status**: Implemented.

#### AL1. `api/common.ts` is misplaced — contains Word-specific Office logic [OPEN]

- **File**: `frontend/src/api/common.ts`
- **Category**: Architecture
- **Details**: Contains `Word.run`, `insertText`, `insertParagraph` and WordFormatter dependency. Not a generic API utility.

#### TL1. Tool type aliases add no value [RESOLVED]

> **Status**: Implemented.

#### TL2. `insertTypes` uses lowercase, plural name [OPEN]

- **File**: `frontend/src/types/index.d.ts:34`
- **Details**: Should be `InsertType` (PascalCase, singular) per TypeScript conventions.

### Pages/Components Dead Code

| ID  | File                                                  | Item                        | Details                          |
| --- | ----------------------------------------------------- | --------------------------- | -------------------------------- |
| PD1 | `frontend/src/pages/HomePage.vue:92`                  | `Briefcase` import          | Never used in template or script |
| PD2 | `frontend/src/pages/HomePage.vue:94`                  | `CheckCircle` import        | Never used anywhere              |
| PD3 | `frontend/src/components/SettingSection.vue`          | Entire component file       | Never imported or used           |
| PD4 | `frontend/src/components/chat/ChatInput.vue:210`      | `handleDragLeave` param `e` | Declared but never read          |
| PD5 | `frontend/src/components/chat/ChatInput.vue:177, 191` | `"input"` emit              | Emitted but no consumer listens  |
| PD6 | `frontend/src/components/SettingCard.vue:2`           | `p1` prop                   | Never passed by any consumer     |
| PD7 | `frontend/src/App.vue:11`                             | Empty `<script>` block      | No code inside                   |

---

## 10. Summary Statistics

| Area                 | CRITICAL  | HIGH      | MEDIUM    | LOW      | Dead Code | Total      |
| -------------------- | --------- | --------- | --------- | -------- | --------- | ---------- |
| Backend              | 3/4       | 3/7       | 5/10      | 2/4      | -         | **13/25**  |
| Frontend Utils       | 2/3       | 5/7       | 3/10      | 0/4      | -         | **10/24**  |
| Composables          | 2/2       | 4/7       | 2/11      | 0/5      | -         | **8/25**   |
| Infrastructure       | 2/3       | 3/5       | 4/8       | 2/7      | -         | **11/23**  |
| Pages/Components/API | 1/1       | 6/7       | 2/15      | 2/6      | -         | **11/29**  |
| Types/Misc           | 0/0       | 0/1       | 0/6       | 1/2      | -         | **1/9**    |
| **Total**            | **10/13** | **21/34** | **16/60** | **7/28** | -         | **54/135** |

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

20. Remove all dead code (29 items across codebase) [RESOLVED]
21. Fix i18n violations: hardcoded French/English strings (PM1, PM9, CM1, BL2) [RESOLVED]
22. Fix error handling: add logging to silent catch blocks [RESOLVED]
23. Replace `any` types with `unknown` + type guards
24. Extract shared utilities (deduplicate `generateVisualDiff`, `withTimeout`, `forHost`) [RESOLVED]
25. Decompose oversized functions (3 functions >180 lines each)
26. Add security headers to nginx config [RESOLVED]
27. Wrap all quick action arrays in `computed()` for locale reactivity [RESOLVED]

---

_Last updated: 2026-03-01_

## Implementation Status Summary

| Status         | ID   | Description                                                           |
| -------------- | ---- | --------------------------------------------------------------------- | ---------- |
| 🟢 Implemented | BC1  | Content-Type enforcement blocks file uploads                          |
| 🟢 Implemented | BC2  | Internal LLM API URL exposed in source and .env.example               |
| 🟢 Implemented | BC3  | Sensitive data logged to disk in plaintext                            |
| 🟢 Implemented | BC4  | User-supplied credentials forwarded without sanitization              |
| 🟢 Implemented | BH1  | Drain event listener leak in streaming response                       |
| 🟢 Implemented | BH2  | logAndRespond called after headers already sent (streaming)           |
| 🟢 Implemented | BH3  | Unbounded log file growth                                             |
| 🟢 Implemented | BH4  | Hardcoded version in health endpoint                                  |
| 🟢 Implemented | BH5  | parsePositiveInt allows zero                                          |
| 🔴 Remaining   | BH6  | Upload route lacks magic-byte file validation                         |
| 🟢 Implemented | BH7  | ReDoS potential in sanitizeErrorText                                  |
| 🟢 Implemented | BM1  | No graceful shutdown handling                                         |
| 🟢 Implemented | BM2  | Unused `routeName` parameter in `validateChatRequest`                 |
| 🟢 Implemented | BM3  | Exported functions never imported externally                          |
| 🟢 Implemented | BM4  | Exported constants/functions never imported externally                |
| 🟢 Implemented | BM5  | Validated values discarded in validateChatRequest                     |
| 🔴 Remaining   | BM6  | Inconsistent error logging patterns                                   |
| 🟢 Implemented | BM7  | `handleErrorResponse` return value discarded                          |
| 🟢 Implemented | BM8  | `allCsv` declared with `let` instead of `const`                       |
| 🟢 Implemented | BM9  | No multer field count limits                                          |
| 🔴 Remaining   | BM10 | No request ID / correlation                                           |
| 🟢 Implemented | BL1  | Dead branch: `if (!imageModel)` check                                 |
| 🟢 Implemented | BL2  | French strings hardcoded in backend                                   |
| 🟢 Implemented | BL3  | Stale comment about character limit                                   |
| 🟢 Implemented | BL4  | `isPlainObject` accepts non-plain objects                             |
| 🟢 Implemented | UC1  | Prompt injection via custom prompt templates                          |
| 🔴 Remaining   | UC2  | XOR "obfuscation" provides false security for API keys                |
| 🟢 Implemented | UC3  | Unsanitized HTML injection in Outlook tools                           |
| 🟢 Implemented | UH1  | `eval_officejs` declared in ExcelToolName but never defined           |
| 🟢 Implemented | UH2  | Column letter arithmetic overflow                                     |
| 🟢 Implemented | UH3  | Double timeout in Outlook tool execution                              |
| 🟢 Implemented | UH4  | Language parameter ignored in translate prompt                        |
| 🔴 Remaining   | UH5  | Host detection caching can return wrong host                          |
| 🟢 Implemented | UH6  | Message toast singleton race condition                                |
| 🟢 Implemented | UH7  | `html: true` in MarkdownIt with `style` in DOMPurify allowlist        |
| 🔴 Remaining   | UM1  | Massive type unsafety with `as unknown as` casts                      |
| 🔴 Remaining   | UM2  | Pervasive `any` types in tool definitions                             |
| 🟢 Implemented | UM3  | Duplicated `generateVisualDiff` function                              |
| 🟢 Implemented | UM4  | Duplicated Office API helpers                                         |
| 🟢 Implemented | UM5  | `Ref` without type parameter in WordFormatter                         |
| 🟢 Implemented | UM6  | `searchAndReplace` tools labeled as category `'read'`                 |
| 🟢 Implemented | UM7  | Redundant Set + Array checks in toolStorage                           |
| 🔴 Remaining   | UM8  | No `QuotaExceededError` handling for localStorage                     |
| 🟢 Implemented | UM9  | `tokenManager.ts` mutates input messages                              |
| 🟡 Deferred    | UM10 | Character-by-character HTML reconstruction in PowerPoint              |
| 🟢 Implemented | UL1  | Typo in export name `buildInPrompt`                                   |
| 🟢 Implemented | UL2  | `deleteText` reports success when no text selected                    |
| 🔴 Remaining   | UL3  | Inconsistent error handling strategy across tools                     |
| 🔴 Remaining   | UL4  | `markdown.ts` vs `officeRichText.ts` naming confusion                 |
| 🟢 Implemented | CC1  | Prompt injection via unsanitized document selection                   |
| 🟢 Implemented | CC2  | Prompt injection via quick action selection text                      |
| 🟢 Implemented | CH1  | Race condition: concurrent `sendMessage` calls corrupt state          |
| 🟢 Implemented | CH2  | `lastIndex` stale reference during agent loop                         |
| 🟢 Implemented | CH3  | Timer leak — `timeoutId` reassigned without clearing                  |
| 🔴 Remaining   | CH4  | Raw `err.message` displayed to users                                  |
| 🔴 Remaining   | CH5  | `any` types on error parameters and tool args                         |
| 🟢 Implemented | CH6  | XSS via unvalidated `imageSrc` URL                                    |
| 🟢 Implemented | CH7  | `THINK_TAG_REGEX` module-level with `g` flag — maintenance hazard     |
| 🟢 Implemented | CM1  | Hardcoded French string in file upload error                          |
| 🟢 Implemented | CM2  | `buildChatMessages` drops system messages                             |
| 🔴 Remaining   | CM3  | Overly large functions                                                |
| 🟢 Implemented | CM4  | `insertToDocument` silently swallows all errors                       |
| 🟢 Implemented | CM5  | Promise constructor anti-pattern in Outlook functions                 |
| 🟢 Implemented | CM6  | Timeout promises create orphaned timers                               |
| 🔴 Remaining   | CM7  | Excel selection returns unescaped tab-separated values                |
| 🟢 Implemented | CM8  | HTML injection via `richHtml` to Office APIs                          |
| 🟢 Implemented | CM9  | Prompt injection via user profile fields                              |
| 🔴 Remaining   | CM10 | `insertImageToPowerPoint` ignores `'NoAction'` semantics              |
| 🔴 Remaining   | CM11 | Hidden side effect: `insertType.value` mutation                       |
| 🟢 Implemented | CL1  | `hostIsWord` parameter accepted but never used                        |
| 🔴 Remaining   | CL2  | `cleanContent` and `splitThinkSegments` use different think-tag logic |
| 🔴 Remaining   | CL3  | Inconsistent image insert error reporting across hosts                |
| 🔴 Remaining   | CL4  | `payload` parameter typed as `unknown` — should be `string            | undefined` |
| 🔴 Remaining   | CL5  | Word HTML selection swallows errors silently                          |
| 🟢 Implemented | IC1  | Content-Type middleware blocks uploads (same as BC1)                  |
| 🟡 Deferred    | IC2  | Containers run as root                                                |
| 🟢 Implemented | IC3  | Internal infrastructure URL as default                                |
| 🟢 Implemented | IH1  | Node.js version mismatch between environments                         |
| 🟡 Deferred    | IH2  | Private IP baked into frontend Docker build                           |
| 🟡 Deferred    | IH3  | External DuckDNS domain as default in .env.example                    |
| 🟢 Implemented | IH4  | Undeterministic package resolution in Dockerfiles                     |
| 🟢 Implemented | IH5  | Nginx missing security headers                                        |
| 🔴 Remaining   | IM1  | Manifest-gen mounts entire project root                               |
| 🟢 Implemented | IM2  | Healthcheck hardcodes port 3003                                       |
| 🟢 Implemented | IM3  | `npm install --production` deprecated                                 |
| 🔴 Remaining   | IM4  | Dev files copied into build context                                   |
| 🟢 Implemented | IM5  | CORS leaks internal IP                                                |
| 🟢 Implemented | IM6  | Empty `lang` attribute in index.html                                  |
| 🟢 Implemented | IM7  | Outlook manifest missing AppDomains                                   |
| 🔴 Remaining   | IM8  | CI infinite-loop guard fragile                                        |
| 🟢 Implemented | IL1  | Vite config uses `.js` extension                                      |
| 🟢 Implemented | IL2  | `@types/diff-match-patch` in dependencies instead of devDependencies  |
| 🔴 Remaining   | IL3  | `chunkSizeWarningLimit` raised to suppress warnings                   |
| 🟢 Implemented | IL4  | Obsolete IE meta tag                                                  |
| 🟢 Implemented | IL5  | Unused PUID/PGID env vars in docker-compose                           |
| 🟢 Implemented | IL6  | Dockerfile HEALTHCHECK overridden by compose                          |
| 🟢 Implemented | IL7  | Legacy entries in .gitignore                                          |
| 🟢 Implemented | PC1  | `keep-alive` never caches `HomePage.vue`                              |
| 🟢 Implemented | PH1  | CSS typo — `itemse-center` instead of `items-center`                  |
| 🟢 Implemented | PH2  | `startNewChat` uses `window.location.reload()` — destructive          |
| 🟢 Implemented | PH3  | `agentMaxIterations` not validated on HomePage                        |
| 🟢 Implemented | PH4  | Discrepancy between HTML `accept` and JS extension validation         |
| 🟢 Implemented | PH5  | Silent failure when files exceed limits or have wrong type            |
| 🟢 Implemented | AH1  | Missing credential headers in `fetchModels`                           |
| 🟢 Implemented | AH2  | `healthCheck()` missing credential headers                            |
| 🔴 Remaining   | XH1  | No CSRF protection on API calls                                       |
| 🔴 Remaining   | PM1  | Hardcoded French strings in ChatInput                                 |
| 🔴 Remaining   | PM2  | Hardcoded English strings with fallback pattern in SettingsPage       |
| 🟢 Implemented | PM3  | `CustomInput` type flash on mount                                     |
| 🔴 Remaining   | PM4  | `CustomInput` model has `any` type                                    |
| 🔴 Remaining   | PM5  | `SingleSelect` dropdown positioning without scroll listener           |
| 🔴 Remaining   | PM6  | Dual emit pattern in `SingleSelect`                                   |
| 🟢 Implemented | PM7  | `SettingCard` prop `p1` never used by any consumer                    |
| 🟢 Implemented | PM8  | `Message.vue` setTimeout without cleanup                              |
| 🟢 Implemented | PM9  | `ChatHeader.vue` hardcoded English string                             |
| 🔴 Remaining   | PM10 | Mixed `t()` and `$t()` usage                                          |
| 🟢 Implemented | PM11 | `expandedThoughts` grows unbounded                                    |
| 🟢 Implemented | AM1  | Import statement in middle of file                                    |
| 🟢 Implemented | AM2  | `chatStream` silently swallows JSON parse errors                      |
| 🟢 Implemented | AM3  | `chatStream` discards remaining buffer after stream ends              |
| 🔴 Remaining   | AM4  | Duplicate `ToolDefinition` interface                                  |
| 🔴 Remaining   | TM1  | Global ambient types without explicit imports                         |
| 🔴 Remaining   | TM2  | `OfficeHostType` declared in two files                                |
| 🔴 Remaining   | EM1  | `useStorage` called outside Vue component context                     |
| 🔴 Remaining   | EM2  | Global `ResizeObserver` monkey-patching                               |
| 🔴 Remaining   | XM1  | Deeply nested ternary chains repeated 10+ times                       |
| 🟢 Implemented | XM2  | Quick action arrays not reactive to locale changes                    |
| 🟢 Implemented | PL1  | `SettingSection.vue` component never imported or used                 |
| 🔴 Remaining   | PL2  | `CustomButton` `icon` prop typed as `any`                             |
| 🔴 Remaining   | PL3  | `SingleSelect` multiple props typed as `any`                          |
| 🟢 Implemented | PL4  | `ChatInput` emits `"input"` event nobody listens to                   |
| 🟢 Implemented | PL5  | `App.vue` has empty `<script>` block                                  |
| 🔴 Remaining   | AL1  | `api/common.ts` is misplaced — contains Word-specific Office logic    |
| 🟢 Implemented | TL1  | Tool type aliases add no value                                        |
| 🔴 Remaining   | TL2  | `insertTypes` uses lowercase, plural name                             |
