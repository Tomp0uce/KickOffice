# DESIGN_REVIEW.md

<!-- dr-state-version: 2 -->
<!-- last-post: 2026-03-28 -->
<!-- branch: feat/user-skills -->
<!-- methodology: desloppify — https://github.com/peteromallet/desloppify -->

---

## Health Score

```
┌──────────────────────────────────────────────┐
│  STRICT SCORE: ~76/100   Target: 70          │
│  Status: ABOVE TARGET (+6)                   │
├──────────────────────────────────────────────┤
│  Mechanical (60%):  ~77  [████████████░░░░]  │
│  Subjective (40%):  ~69  [███████████░░░░░]  │
│  (subjective unchanged — re-run /dr-audit)   │
├──────────────────────────────────────────────┤
│  Open issues: 0   CRITICAL: 0   HIGH: 0     │
│  MEDIUM: 0        LOW: 0                     │
│  Fixed: 16   Wontfix: 7   Deferred: 2       │
└──────────────────────────────────────────────┘
```

> `~` = estimated. Run `/dr-audit` after merge for validated scores.

---

## Score Evolution

| Date | Branch | Strict | Mechanical | Subjective | Fixed | Notes |
|------|--------|--------|------------|------------|-------|-------|
| 2026-03-28 | feat/user-skills | 65 | 67 | 63 | — | Initial audit (desloppify v13) |
| 2026-03-28 | feat/user-skills | ~76 | ~77 | ~69 (est.) | 16 (H:7 M:5 L:4) | After all 3 phases — coverage 14→86% |

---

## Score by Category

| Category | Before | After | Delta | Open | Deferred |
|----------|--------|-------|-------|------|----------|
| Architecture & Data Flow | 72 | ~76 | +4 | 0 | 1 |
| Robustness & Business Logic | 58 | ~72 | +14 | 0 | 0 |
| Observability & Error Mgmt | 78 | ~80 | +2 | 0 | 0 |
| UX/UI & Integration | 90 | ~90 | 0 | 0 | 0 |
| DRY & Modularity | 70 | ~78 | +8 | 0 | 0 |
| Clean Code | 56 | ~74 | +18 | 0 | 1 |
| Documentation | 78 | ~80 | +2 | 0 | 0 |

---

## Deferred Items

<!-- Items that remain open across cycles. NEVER DELETE this section. -->
<!-- Add rows when items are deferred; remove rows only when items are closed (moved to Resolved History). -->

| Issue ID | Summary | Reason deferred | Deferred on | Target |
|----------|---------|-----------------|-------------|--------|
| ARCH-H2/H3 | Monolithic files (useAgentLoop 1181 LOC, excelTools 2645, powerpointTools 2421, wordTools 2154) | Feature set still evolving post-beta; splitting causes multi-file churn with no functional gain | 2026-03-19 | Revisit when tool additions slow |
| TOKEN-M1 | Token limit calibration — MAX_CONTEXT_CHARS (1.2M) conservative estimate | Requires 2+ weeks usage data | 2026-03-19 | After beta usage data |

---

## Backlog (discovered during fix cycles — deferred to next /dr-audit)

<!-- Items found during /dr-run but not CRITICAL/security. Processed by next /dr-audit. -->

| Issue ID | Criticality | File | Problem | Discovered during |
|----------|-------------|------|---------|-------------------|
| _(empty — no new items discovered during this cycle)_ | | | | |

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
- ✅ QUAL-H3 — Added 511 frontend tests across 16 files, coverage 14→86% (2026-03-28)
- ✗ TOOL-H2 — WONTFIX: Word screenshot — no Office.js API, html2canvas unsupported in sandbox (2026-03-19)
- ✗ USR-H1 — WONTFIX: Empty shape bullets — placeholderFormat covers 95% of cases (2026-03-19)
- ✗ Phase 7F — WONTFIX: Dynamic tool loading — LLM handles 128+ tools, no usage data for profiles (2026-03-19)
- ✗ DEAD-L2 — WONTFIX: plotDigitizer route — vision insufficient for chart accuracy (2026-03-19)
- ✗ QUAL-L2 — WONTFIX: credentialCrypto in LS — dedicated PCs, XSS mitigated by DOMPurify + CSP (2026-03-19)
- ✗ DEAD-L3 — WONTFIX: clearEncryptionKeys — false positive, still used (2026-03-19)
- ✗ USR-H2 — WONTFIX: Context bloat indicator — already shown live in currentAction + StatsBar (2026-03-19)
