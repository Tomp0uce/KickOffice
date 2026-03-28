# KickOffice — Lessons Learned

<!-- Updated after corrections or mistakes. Each entry: what happened, why, prevention rule. -->

---

## 2026-03-28 — Initial seed (migrated from CLAUDE.md knowledge)

### L01 — `reasoning_effort: 'none'` is NOT a valid OpenAI API value
**What happened:** Sending `reasoning_effort: 'none'` to the API caused empty responses from reasoning models.
**Rule:** Valid values are `'low'`, `'medium'`, `'high'`. Omit the parameter entirely when not needed. `buildChatBody` in `config/models.js` is the single source of truth for this logic.

### L02 — ChatGPT-* model IDs do not support temperature / token-limit parameters
**What happened:** Sending `temperature` or `max_tokens` to ChatGPT model IDs caused API errors.
**Rule:** `isChatGptModel` check in `buildChatBody` gates these params. GPT-5 models use `max_completion_tokens`; non-GPT-5 use `max_tokens` (`isGpt5Model` check).

### L03 — Track Changes pattern must disable TC before insertOoxml
**What happened:** Inserting OOXML with `w:ins`/`w:del` while Track Changes was active caused duplicate markings.
**Rule:** Pattern: disable TC → `insertOoxml` with `w:ins`/`w:del` → restore TC. `acceptAiChanges`/`rejectAiChanges` require WordApi 1.6 guard: `Office.context.requirements.isSetSupported('WordApi', '1.6')`. Use `context.document.trackedChanges` (property, not method).

### L04 — Base64 image payloads MUST be sanitized before logging
**What happened:** Logging chat payloads with inline Base64 images caused massive log files and exposed image data.
**Rule:** Always pass payloads through `sanitizePayloadForLogs` before logging. This replaces Base64 content with a placeholder.

### L05 — Session switch must be blocked while agent loop is running
**What happened:** Switching sessions during an active agent loop caused state corruption.
**Rule:** 3-layer protection: `useSessionManager`, `useHomePage`, `ChatHeader.vue` all check `loading.value` before allowing session switch.

### L06 — Never hand-edit generated manifest files
**What happened:** Manual edits to `manifest-office.xml` were overwritten by `generate-manifests.js`.
**Rule:** Update `manifests-templates/` and/or root `.env`, then run `scripts/generate-manifests.js` to regenerate. Docker Compose runs this automatically via the frontend entrypoint.

### L07 — Alpine Docker images are incompatible with Synology DS416play
**What happened:** Node.js Alpine images crashed on the Synology NAS (Intel Celeron N3150) due to musl libc + AVX instruction issues.
**Rule:** Always use Debian-based images: `node:22-slim` for backend, `nginxinc/nginx-unprivileged:stable` for frontend.

---

## 2026-03-28 — Design Review Cycle (feat/user-skills)

### What was found
- [Robustness] MEDIUM: 199 explicit `any` types concentrated in Office tool files (powerpointTools 53, outlookTools 25, excelTools 21)
- [Clean Code] HIGH: 14% frontend test coverage, 0% backend — critical gap for regression safety
- [Architecture] HIGH: Language resolution pattern duplicated 9x across 3 composables
- [Robustness] MEDIUM: `inject*` functions mutated arrays in-place while return type suggested transform
- [Robustness] LOW: Unguarded `JSON.parse` in credentialCrypto caused infinite recursion on corrupted key

### What was fixed
- ARCH-H4: Language resolution — 9x duplication hidden behind different `localStorage` + `navigator.language` fallback chains
- ROB-M1: Office `any` types — many are unavoidable (Office.js types incomplete for PPT/Outlook), typed 135 of 199
- QUAL-H2/H3: Coverage 14%→86% — TDD approach caught a Vue reactivity bug in `streamOneShot` (detached object mutation)
- ROB-L1: credentialCrypto recursive `getEncryptionKey` — replaced with inline key regeneration after corrupted key removal

### Rules to prevent recurrence
- Vue reactivity: Always use `ref.value[idx] = { ...spread }` for array element updates — never mutate a detached object reference
- `vi.clearAllMocks()` vs `vi.restoreAllMocks()`: restoreAllMocks resets `vi.mock()` factory implementations; use clearAllMocks in tests with module-level mocks
- Office.js ambient types: `declare const Word/Excel/PowerPoint: any` in vite-env.d.ts is intentional — @types/office-js doesn't cover all host APIs
- Lazy env validation: Use `{ toString() }` pattern for env vars that may not be set at import time (e.g. VITE_BACKEND_URL)

### Score delta
Strict: 65 → ~76 (+11) | Mechanical: 67 → ~77 (+10) | Subjective: 63 → ~69 (+6)

---

## 2026-03-28 — Design Review v14 Cycle (feat/user-skills)

### What was found
- [CI/CD] HIGH: Zero PR checks, no automated testing, no pre-commit hooks — codebase had excellent local tooling with zero CI enforcement
- [Security] MEDIUM: `x-request-id` header accepted without validation — potential log injection vector via ANSI codes or newlines
- [Robustness] MEDIUM: ExcelJS formula cells with null result produced literal "null" in CSV output
- [Observability] HIGH: logService method signatures inconsistent — traffic param unreachable on warn/error/debug
- [Security] MEDIUM: Backend routes lacked input validation (no category allowlist, no description length limit, no sessionId format check)

### What was fixed
- CI-H1/H2/M1/M2: Complete CI/CD pipeline — single biggest score delta (+65 on CI category)
- OBS-M1: Logger refactor — `toDataRecord()` helper with `Error instanceof` guard prevents empty `{}` serialization
- REV-H1: Formula null — `v.result ?? ''` prevents literal "null" in ExcelJS formula output
- REV-M1: UUID validation — regex check on x-request-id prevents log injection
- REV-M2/M3/M4: Backend input validation — category allowlist, description max length, sessionId format regex

### Rules to prevent recurrence
- Backend input validation: Every POST route must validate all user-provided fields (allowlist, length, format regex) — not just existence checks
- Error serialization: `Error` objects have non-enumerable properties (message, stack, name) — casting to `Record<string, unknown>` produces `{}`; always use `instanceof Error` guard
- Request ID validation: Any header used in logging must be validated (UUID format, length limit) to prevent log injection
- ExcelJS formula cells: `v.result` can be `null` even when `v` is a formula object — always use nullish coalescing

### Score delta
Strict: 69 → ~78 (+9) | Mechanical: 73 → ~84 (+11) | Subjective: 64 (unchanged)
