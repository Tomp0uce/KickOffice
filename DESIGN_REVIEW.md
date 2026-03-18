# DESIGN_REVIEW.md

**Last updated**: 2026-03-19
**Status**: DR v12 fully triaged. All critical/high/medium/low items resolved or explicitly deferred. Remaining open items are architectural refactors (large-file consolidation — deferred to post-beta) — no functional bugs outstanding.

---

## Completed Work

All items from audit cycles v9–v12 have been addressed. **56 items from v9–v11** are ✅ FIXED (Phases 1A–7A). **All 5 critical items from v12** are ✅ FIXED (2026-03-16). The complete v12 batch (41 items across error handling, UX, dead code, duplication, code quality, and OXML enhancements) was resolved across 2026-03-16 and 2026-03-17.

Key deliverables: file re-injection eliminated — single-pass via `contentInjectedAt` + VFS fallback, images via `/v1/files` fileId to avoid base64 re-send (TOOL-C1, `299e0ca` + `2d91a9d`); SSE error handling hardened (ERR-C1–C4, RACE-C1); session-switch race condition eliminated; frontend log forwarding to backend (ERR-M3); rate-limit floor (ERR-M4); upstream SSE reader cancellation (ERR-M5); request correlation IDs (ERR-L1); stream error Retry button (ERR-L2); HomePage decomposed (UX-H1 + QUAL-H2); dark mode fix (UX-L1); keyboard navigation for dropdowns (UX-M4); i18n gaps closed (UX-M1/M3/L2, DEAD-L1); `mutationDetector.ts` dedup (DUP-H1); `getVfsSandboxContext` dedup (DUP-M1); `createEvalExecutor` factory (DUP-M2); `buildScreenshotResult` helper (DUP-L1); full TypeScript `any` removal (QUAL-H1); JSON truncation fix in tokenManager (QUAL-M3); CSS injection hardening in markdown (QUAL-M4); env var validation in backend (QUAL-M5); backend log summary (QUAL-L1); `addAttachment` Outlook tool (FUNC-M2); Waterfall/Treemap/Funnel chart types (FUNC-L1); `reorderSlide` PPT tool (FUNC-L2); `acceptAiChanges`/`rejectAiChanges` + "Valider" button with proper WordApi 1.6 version guard (OXML-IMP3); `insertOoxml` Word tool (OXML-IMP4); `addComment`/`getComments` Word tools (OXML-IMP2); speaker notes via PPT native API (OXML-IMP5); `powerpointImageRegistry` cleared on session switch (QUAL-M2); `office-agents/` directory removed (ARCH-M3/DEAD-M1); tool counts synchronized to 100 across all docs (FUNC-M1); CSS virtualization for ChatMessageList via `content-visibility: auto` (UX-M2); unit tests for `useLoopDetection`, `useSessionFiles`, `useMessageOrchestration`, `useToolExecutor` — 47 new tests (QUAL-M1); `backend.ts` split into `api/types.ts`, `api/errorCategorization.ts`, `api/httpClient.ts` + facade (ARCH-M2); PowerPoint `buildPowerPointExecute` extracted from anonymous closure (ARCH-L1).

---

## Open Items

These items are acknowledged but not yet prioritized for implementation.

_No open items at this time._

---

## Deferred Items

Intentionally deferred — not forgotten, not yet unblocked.

### Large-file structural refactors (post-beta)

#### ARCH-H2/H3 — Monolithic files consolidation [HIGH]

Deferred until the feature set stabilises post-beta. Splitting now would cause constant multi-file churn with no functional gain.

Files to revisit:

| File | Lines | Suggested split |
|------|-------|-----------------|
| `composables/useAgentLoop.ts` | ~1,100 | Extract `runAgentLoop()` → `useAgentRunner.ts`; image flow → `useImageGeneration.ts`; keep `useAgentLoop` as thin orchestrator |
| `utils/excelTools.ts` | ~2,700 | `tools/excel/` subdirectory + `index.ts` barrel |
| `utils/powerpointTools.ts` | ~2,400 | `tools/powerpoint/` subdirectory + `index.ts` barrel |
| `utils/wordTools.ts` | ~2,100 | `tools/word/` subdirectory + `index.ts` barrel |
| `utils/outlookTools.ts` | ~700 | `tools/outlook/` subdirectory + `index.ts` barrel |

**Trigger**: Revisit when tool additions slow down and the beta feature set is stable.

### Context & Token Management

#### Phase 7C — TOKEN-M1: Token Limit Calibration [MEDIUM]

`MAX_CONTEXT_CHARS` (1.2M) is a conservative estimate. Needs tuning based on real usage data.
**Blocked by**: Requires 2+ weeks of LOG-H1 usage data.

---

## Won't Fix

| Item | Decision |
|------|----------|
| TOOL-H2 — Word screenshot | No Office.js API. html2canvas/puppeteer unsupported in add-in sandbox. `getDocumentHtml()` is the closest proxy. |
| USR-H1 — Empty shape bullets | `placeholderFormat/type` covers 95% of cases. Remaining XML-default-bullet edge cases are rare. |
| Phase 7F — Dynamic tool loading | GPT-5 handles 128+ tools fine. No usage data to define intent profiles yet. Revisit after 6+ months of LOG-H1 data. |
| DEAD-L2 — plotDigitizer route | LLM vision tested and found insufficient for chart data accuracy. Pixel-analysis pipeline kept as-is. |
| QUAL-L2 — credentialCrypto key in localStorage | Add-in runs on dedicated PCs with per-user Windows login. Re-keying on every restart would be a major UX regression. XSS already mitigated by DOMPurify + CSP. |
| DEAD-L3 — clearEncryptionKeys | False positive — still used. |
| USR-H2 — Context bloat indicator | Context % shown live in `currentAction` (e.g. "12s · ctx 73%") since `299e0ca`. StatsBar colors orange at 70%, red at 90% + tooltip at 80%. A separate dismissible banner would be redundant noise. |

---

## Architecture Notes

### Tool Counts (audited 2026-03-16)

| Host | Count | Notable tools |
|------|-------|---------------|
| Word | 34 | `proposeRevision`, `proposeDocumentRevision`, `editDocumentXml`, `insertOoxml`, `acceptAiChanges`, `rejectAiChanges`, `addComment`, `getComments`, `eval_wordjs`, `getDocumentOoxml` |
| Excel | 27 | `eval_officejs`, `screenshotRange`, `getRangeAsCsv`, `detectDataHeaders`, `manageObject` (incl. Waterfall/Treemap/Funnel) |
| PowerPoint | 24 | `screenshotSlide`, `editSlideXml`, `reorderSlide`, `getSpeakerNotes`, `setSpeakerNotes`, `searchIcons`, `insertIcon`, `verifySlides` |
| Outlook | 9 | `eval_outlookjs`, `addAttachment`, email read/write helpers |
| General | 6 | `executeBash` (VFS), `calculateMath`, `getCurrentDate`, file operations |
| **Total** | **100** | |

### Key Files

| File | Purpose |
|------|---------|
| `frontend/src/utils/tokenManager.ts` | Context window management + Phase 7A heuristic compression |
| `frontend/src/utils/wordDiffUtils.ts` | Track Changes — selection (`applyRevisionToSelection`) + document (`applyRevisionToDocument`) |
| `frontend/src/utils/wordTrackChanges.ts` | `setChangeTrackingForAi` / `restoreChangeTracking` helpers |
| `frontend/src/utils/toolProviderRegistry.ts` | Host → tool provider mapping (singleton) |
| `frontend/src/utils/mutationDetector.ts` | Shared `createMutationDetector()` factory (DUP-H1) |
| `frontend/src/composables/useAgentLoop.ts` | Agent execution loop (~1,100 lines — see ARCH-H2) |
| `frontend/src/composables/quickActions/` | Per-host quick action composables (4 files) |
| `frontend/src/skills/` | 5 host skills + 17 Quick Action skills |

### Largest Files (for ARCH-H2/H3 post-beta refactor)

| Category | File | Lines |
|----------|------|-------|
| **Composables** | `useAgentLoop.ts` | ~1,100 |
| **Tool Files** | `excelTools.ts` | ~2,700 |
| | `powerpointTools.ts` | ~2,400 |
| | `wordTools.ts` | ~2,100 |
| | `outlookTools.ts` | ~700 |
| **API** | `backend.ts` | ~670 |
| **Pages** | `HomePage.vue` | ~426 |
| | `ChatMessageList.vue` | ~400 |
| | `ChatInput.vue` | ~320 |

---

*See CHANGELOG.md for full version history.*
