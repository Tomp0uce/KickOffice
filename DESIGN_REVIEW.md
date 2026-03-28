# DESIGN_REVIEW.md

<!-- dr-state-version: 2 -->
<!-- last-post: 2026-03-28 -->
<!-- branch: feat/user-skills -->
<!-- methodology: desloppify — https://github.com/peteromallet/desloppify -->

---

## Health Score

```
┌────────────���─────────────────────────────────┐
│  STRICT SCORE: ~78/100   Target: 70          │
│  Status: ABOVE TARGET (+8)                   │
├────────────────────────────────────────��─────┤
│  Mechanical (60%):  ~84  [█████████████░░░]  │
│  Subjective (40%):  64   [██████████░░░░░░]  │
│  (subjective unchanged — re-run /dr-audit)   │
├─────────────────────────────────────���────────┤
│  Open issues: 0   CRITICAL: 0   HIGH: 0     │
│  MEDIUM: 0        LOW: 0                     │
│  Fixed: 37   Wontfix: 0   Deferred: 1       │
└──────────���────────────────────���──────────────┘
```

> `~` = estimated. Run `/dr-audit` after merge for validated scores.

---

## Score Evolution

| Date | Branch | Strict | Mechanical | Subjective | Fixed | Notes |
|------|--------|--------|------------|------------|-------|-------|
| 2026-03-28 | feat/user-skills | 65 | 67 | 63 | — | v13 initial audit |
| 2026-03-28 | feat/user-skills | ~76 | ~77 | ~69 | 16 (H:4 M:6 L:6) | v13 post-fix estimate |
| 2026-03-28 | feat/user-skills | 69 | 73 | 64 | — | v14 re-audit (CI/CD scored, new bugs found) |
| 2026-03-28 | feat/user-skills | ~78 | ~84 | 64 (est.) | 37 (H:7 M:14 L:8 + 8 review) | v14 post-fix: 29 plan + 8 review fixes |

---

## Score by Category

| Category | Before | After | Delta | Open | Deferred |
|----------|--------|-------|-------|------|----------|
| Architecture & Data Flow | 62 | ~70 | +8 | 0 | 0 |
| Robustness & Business Logic | 62 | ~76 | +14 | 0 | 0 |
| Observability & Error Mgmt | 72 | ~82 | +10 | 0 | 0 |
| UX/UI & Integration | 88 | ~88 | 0 | 0 | 0 |
| DRY & Modularity | 76 | ~84 | +8 | 0 | 0 |
| Clean Code | 68 | ~76 | +8 | 0 | 1 |
| Documentation | 76 | ~80 | +4 | 0 | 0 |
| Security & Dependencies | 72 | ~88 | +16 | 0 | 0 |
| CI/CD | 15 | ~82 | +67 | 0 | 0 |

---

## Deferred Items

<!-- Items that remain open across cycles. NEVER DELETE this section. -->
<!-- Add rows when items are deferred; remove rows only when items are closed (moved to Resolved History). -->

| Issue ID | Summary | Reason deferred | Deferred on | Target |
|----------|---------|-----------------|-------------|--------|
| ARCH-H2/H3 | Monolithic files (useAgentLoop 1071, excelTools 2829, powerpointTools 2452, wordTools 2175 LOC) | Feature set still evolving post-beta; splitting causes multi-file churn with no functional gain | 2026-03-19 | Revisit when tool additions slow |
| TOKEN-M1 | Token limit calibration — MAX_CONTEXT_CHARS (1.2M) conservative estimate | Requires 2+ weeks usage data | 2026-03-19 | After beta usage data |
| CLN-L1 | 24 raw localStorage calls across 12 files should use localStorageKey enum | T3 scope — 12 files, 24 call sites | 2026-03-28 | Next dr-run cycle |

---

## Backlog (discovered during fix cycles — deferred to next /dr-audit)

<!-- Items found during /dr-run but not CRITICAL/security. Processed by next /dr-audit. -->

| Issue ID | Criticality | File | Problem | Discovered during |
|----------|-------------|------|---------|-------------------|
| _(empty)_ | | | | |

---

## Resolved History

<!-- Compressed 1-line record of every closed issue. NEVER DELETE. Append-only. -->
<!-- ✅ = fixed   ✗ = wontfix (with reason) -->

### v12 (2026-03-16 to 2026-03-19)
- ✅ ARCH-M2 — Split backend.ts into api/types.ts, api/errorCategorization.ts, api/httpClient.ts + facade (2026-03-16)
- ✅ ARCH-L1 — Extracted PowerPoint buildPowerPointExecute from anonymous closure (2026-03-16)
- ✅ ARCH-M3 — Removed legacy office-agents/ directory (2026-03-16)
- ✅ TOOL-C1 — Eliminated file re-injection via contentInjectedAt + VFS fallback (2026-03-16)
- ✅ OXML-IMP3 — Implemented acceptAiChanges/rejectAiChanges with WordApi 1.6 guard (2026-03-16)
- ✅ OXML-IMP4/2/5 — Added insertOoxml, addComment/getComments (Word), native speaker notes (PPT) (2026-03-16)
- ✅ FUNC-M2/L1/L2 — Added addAttachment (Outlook), Waterfall/Treemap/Funnel charts (Excel), reorderSlide (PPT) (2026-03-16)
- ✅ FUNC-M1 — Synchronized tool counts to 100 across all docs (2026-03-16)
- ✅ ERR-C1-C4, RACE-C1 — Hardened SSE error handling, eliminated session-switch race conditions (2026-03-17)
- ✅ ERR-M3/M4/M5 — Frontend log forwarding, rate-limit floor, upstream SSE reader cancellation (2026-03-17)
- ✅ ERR-L1/L2 — Request correlation IDs, stream error Retry button (2026-03-17)
- ✅ UX-H1 — Decomposed HomePage.vue (2026-03-17)
- ✅ UX-M2 — CSS virtualization for ChatMessageList (2026-03-17)
- ✅ UX-M4/L1/L2 — Keyboard nav, dark mode, i18n gaps (2026-03-17)
- ✅ DUP-H1 — Deduplicated mutationDetector.ts (2026-03-17)
- ✅ DUP-M1/M2 — Deduplicated getVfsSandboxContext, created createEvalExecutor factory (2026-03-17)
- ✅ DUP-L1 — Extracted buildScreenshotResult helper (2026-03-17)
- ✅ QUAL-H1 — TypeScript any removal pass (2026-03-17)
- ✅ QUAL-M1 — Added 47 unit tests for useLoopDetection, useSessionFiles, useMessageOrchestration, useToolExecutor (2026-03-17)
- ✅ QUAL-M3/M4/M5 — JSON truncation fix, CSS injection hardening, backend env validation (2026-03-17)
- ✅ QUAL-M2 — Cleared powerpointImageRegistry on session switch (2026-03-17)
- ✅ DEAD-M1/L1 — Removed dead code from legacy i18n and office-agents (2026-03-17)

### v13 (2026-03-28)
- ✅ ARCH-H4 — Extracted getDisplayLanguage() utility, replaced 9x duplication (2026-03-28)
- ✅ ARCH-M6 — Extracted streamOneShot() from handleSmartReply/handleMoM shared tail (2026-03-28)
- ✅ ROB-M3 — inject* functions now return void (explicit mutation contract) (2026-03-28)
- ✅ ARCH-M4 — Single createBuiltInPromptGetter factory replaces 4x copy-paste (2026-03-28)
- ✅ CLN-L1 — Translated French comment to English in models.js (2026-03-28)
- ✅ ROB-L1 — Added try/catch for corrupted crypto key with inline regeneration (2026-03-28)
- ✅ DOC-L1 — Documented VITE_REQUEST_TIMEOUT_MS, VITE_VERBOSE_LOGGING in .env.example (2026-03-28)
- ✅ CLN-L2 — Typed searchIconify return with IconifySearchResult interface (2026-03-28)
- ✅ ROB-H1 — Exported UndoSnapshot, replaced 8x Partial<any> with typed union (2026-03-28)
- ✅ DRY-H1 — Grouped UseQuickActionsOptions into sub-interfaces, removed unused fields (2026-03-28)
- ✅ ROB-M2 — Deferred VITE_BACKEND_URL validation to first API call via lazy toString() (2026-03-28)
- ✅ ROB-M1 — Typed 135 of 199 any types across Office tool files (2026-03-28)
- ✅ QUAL-H2 — Added 50 backend tests (buildChatBody + chat route integration) (2026-03-28)
- ✅ QUAL-H3 — Added 511 frontend tests across 16 files, coverage 14->86% (2026-03-28)
- ✗ TOOL-H2 — WONTFIX: Word screenshot — no Office.js API, html2canvas unsupported in sandbox (2026-03-19)
- ✗ USR-H1 — WONTFIX: Empty shape bullets — placeholderFormat covers 95% of cases (2026-03-19)
- ✗ Phase 7F — WONTFIX: Dynamic tool loading — LLM handles 128+ tools, no usage data for profiles (2026-03-19)
- ✗ DEAD-L2 — WONTFIX: plotDigitizer route — vision insufficient for chart accuracy (2026-03-19)
- ✗ QUAL-L2 — WONTFIX: credentialCrypto in LS — dedicated PCs, XSS mitigated by DOMPurify + CSP (2026-03-19)
- ✗ DEAD-L3 — WONTFIX: clearEncryptionKeys — false positive, still used (2026-03-19)
- ✗ USR-H2 — WONTFIX: Context bloat indicator — already shown live in currentAction + StatsBar (2026-03-19)

### v14 — Plan fixes (2026-03-28)
- ✅ ROB-H1 — Spread-copy messages in prepareMessagesForContext to prevent caller mutation (TDD) (2026-03-28)
- ✅ ROB-H2 — Replaced '\\n' with '\n' in smart-reply XML delimiters (2026-03-28)
- ✅ ARCH-M1 — Split timeoutId into textTimeoutId/htmlTimeoutId, both cleared in finally (2026-03-28)
- ✅ SEC-H1 — Changed trust proxy: true to 1 (single nginx hop) (2026-03-28)
- ✅ SEC-H2 — Migrated xlsx to exceljs (CVE-2023-30533), separated CSV handling (TDD) (2026-03-28)
- ✅ OBS-H1 — Replaced full body logging with metadata-only in sync chat endpoint (2026-03-28)
- ✅ CI-H1 — Created pr-checks.yml: lint, tsc, test, build, Docker build, npm audit (2026-03-28)
- ✅ CI-H2 — Added husky + lint-staged pre-commit hooks (2026-03-28)
- ✅ CI-M1 — Docker build verification in pr-checks.yml (2026-03-28)
- ✅ CI-M2 — npm audit step in pr-checks.yml (2026-03-28)
- ✅ DRY-M3 — Removed dead pendingSmartReply + handleSmartReply (2026-03-28)
- ✅ CLN-M1 — Removed dead injectedContext parameter from full chain (2026-03-28)
- ✅ ROB-M3 — Collapsed identical branch bodies in prepareMessagesForContext (2026-03-28)
- ✅ CLN-L2 — Translated French comments (Tache 4/6) to English (2026-03-28)
- ✅ DOC-L1 — Added warningVfsWriteFailed i18n key (2026-03-28)
- ✅ ARCH-M3 — Extracted focusInputWithGlow() helper, replaces 3 identical 12-line blocks (2026-03-28)
- ✅ DRY-M1 — Extracted BACKEND_URL to httpClient.ts, removed 2 local definitions (2026-03-28)
- ✅ ARCH-M2 — Replaced 50-line inline PPT code with getCurrentSlideNumber+getSlideContentStandalone (2026-03-28)
- ✅ ROB-M1 — generateImage throws Error instead of silent return '' (2026-03-28)
- ✅ OBS-M1 — Unified logService signatures with optional traffic param + toDataRecord helper (2026-03-28)
- ✅ ROB-M4 — Replaced any[] with MessageContentPart[] in truncateToBudget overload (2026-03-28)
- ✅ ROB-M2 — Removed redundant undoSnapshot/canUndo resets from 5 undo sub-functions (2026-03-28)
- ✅ DOC-L2 — Verified README quick action count already correct at 24 (2026-03-28)
- ✅ DOC-L3 — Fixed README tool count 100 to 101 (2026-03-28)
- ✅ SEC-L1 — Added diff-match-patch + @types as direct deps, removed manual shim (2026-03-28)
- ✅ SEC-L2 — Verified focus-trap used by @vueuse/integrations/useFocusTrap, kept (2026-03-28)
- ✅ OBS-L1 — Deferred logCryptoStatus from module load to migrateCredentialsOnStartup (2026-03-28)
- ✅ CLN-L3 — Removed phantom generic TContext, added JSDoc, updated 2 call sites (2026-03-28)
- ✅ DRY-M2 — Documented inject* mutation contract in module JSDoc (2026-03-28)

### v14 — Review fixes (2026-03-28)
- ✅ REV-H1 — Fixed ExcelJS formula cells with null result producing literal "null" in CSV (2026-03-28)
- ✅ REV-H2 — Improved crypto key importKey recovery with try/catch and regeneration (2026-03-28)
- ✅ REV-H3 — Added Error serialization guard in logger toDataRecord (instanceof Error) (2026-03-28)
- ✅ REV-M1 — Added UUID format validation for x-request-id header (log injection prevention) (2026-03-28)
- ✅ REV-M2 — Added description max length (2000) validation in skillCreator route (2026-03-28)
- ✅ REV-M3 — Added category allowlist + sessionId format validation in feedback route (2026-03-28)
- ✅ REV-M4 — Removed continue-on-error from CI audit jobs, changed to audit-level=critical (2026-03-28)
- ✅ REV-M5 — Fixed logService.debug traffic positional arg in backend.ts and sandbox.ts (2026-03-28)

---

## Won't Fix

| Item | Decision |
|------|----------|
| **TOOL-H2** — Word screenshot | No Office.js API. html2canvas/puppeteer unsupported in sandbox. `getDocumentHtml()` is closest proxy. |
| **USR-H1** — Empty shape bullets | `placeholderFormat/type` covers 95% of cases. XML-default-bullet edge cases are rare. |
| **Phase 7F** — Dynamic tool loading | LLM handles 128+ tools fine. No usage data to define intent profiles yet. Revisit after 6+ months. |
| **DEAD-L2** — `plotDigitizer` route | LLM vision tested and found insufficient for chart accuracy. Pixel-analysis pipeline kept as-is. |
| **QUAL-L2** — `credentialCrypto` in LS | Add-in runs on dedicated PCs with per-user Windows login. Re-keying on restart is a UX regression. XSS mitigated by DOMPurify + CSP. |
| **DEAD-L3** — `clearEncryptionKeys` | False positive — still used. |
| **USR-H2** �� Context bloat indicator | Context % shown live in `currentAction`. StatsBar colors change at 70%/90%. A banner is redundant noise. |
