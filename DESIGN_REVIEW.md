# 🏗️ SYSTEM DESIGN REVIEW & AUDIT

**Last updated**: 2026-03-19  
**Audit Cycle**: v12 Complete  
**Overall Health**: 🟢 **STABLE** (No outstanding functional bugs)

## 📊 Executive Status

| Category | Status | 🔴 Critical | 🟠 High | 🟡 Medium | 🔵 Low |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **Open Items** | 🟢 All Clear | 0 | 0 | 0 | 0 |
| **Resolved (v12)** | ✅ Completed | 5 | 4 | 15 | 6 |
| **Deferred** | ⏸️ Post-Beta | 0 | 1 | 1 | 0 |

---

## 📋 Implementation Plan

> **Note:** V12 is complete (resolved across 2026-03-16 and 2026-03-17). The implementation plan is currently empty. The structure below is preserved for the next audit cycle.

### Phase 1: [Placeholder for Next Cycle]
*Goal: [Define the main objective of this phase]*

#### Sub-phase 1A: [Specific Component/Area]
*(Constraint: Max 3 large items or 6 small items to maintain LLM context)*
- [ ] `[ID]` - [Task description]
- [ ] `[ID]` - [Task description]

#### Sub-phase 1B: [Specific Component/Area]
- [ ] `[ID]` - [Task description]

---

## 🔍 Audit Axes (Current State & History)

*This section tracks the state of the codebase across the defined architectural axes. Items marked ✅ were resolved in recent cycles.*

### 1. Architecture & Data Flow
- ✅ **[ARCH-M2]** Split `backend.ts` into `api/types.ts`, `api/errorCategorization.ts`, `api/httpClient.ts` + facade.
- ✅ **[ARCH-L1]** Extracted PowerPoint `buildPowerPointExecute` from anonymous closure.
- ✅ **[ARCH-M3]** Removed legacy `office-agents/` directory.

### 2. Office Add-in Features & Integration
- ✅ **[TOOL-C1]** Eliminated file re-injection. Single-pass via `contentInjectedAt` + VFS fallback, images via `/v1/files` fileId to avoid base64 re-send (`299e0ca` + `2d91a9d`).
- ✅ **[OXML-IMP3]** Implemented `acceptAiChanges`/`rejectAiChanges` + "Valider" button with proper WordApi 1.6 version guard.
- ✅ **[OXML-IMP4/2/5]** Added `insertOoxml`, `addComment`/`getComments` (Word), and native speaker notes API (PowerPoint).
- ✅ **[FUNC-M2/L1/L2]** Added `addAttachment` (Outlook), Waterfall/Treemap/Funnel charts (Excel), `reorderSlide` (PPT).
- ✅ **[FUNC-M1]** Synchronized tool counts to 100 across all docs.

### 3. Observability & Error Handling
- ✅ **[ERR-C1–C4, RACE-C1]** Hardened SSE error handling and eliminated session-switch race conditions.
- ✅ **[ERR-M3/M4/M5]** Implemented frontend log forwarding to backend, rate-limit floor, and upstream SSE reader cancellation.
- ✅ **[ERR-L1/L2]** Added request correlation IDs and stream error Retry button.

### 4. UX & UI
- ✅ **[UX-H1]** Decomposed `HomePage.vue`.
- ✅ **[UX-M2]** CSS virtualization for `ChatMessageList` via `content-visibility: auto`.
- ✅ **[UX-M4/L1/L2]** Fixed keyboard navigation for dropdowns, dark mode, and closed i18n gaps.

### 5. DRY & Modularity (Duplication)
- ✅ **[DUP-H1]** Deduplicated `mutationDetector.ts`.
- ✅ **[DUP-M1/M2]** Deduplicated `getVfsSandboxContext` and created `createEvalExecutor` factory.
- ✅ **[DUP-L1]** Extracted `buildScreenshotResult` helper.

### 6. Clean Code (Dead Code & Quality)
- ✅ **[QUAL-H1]** Full TypeScript `any` removal across the codebase.
- ✅ **[QUAL-M1]** Added 47 new unit tests for `useLoopDetection`, `useSessionFiles`, `useMessageOrchestration`, `useToolExecutor`.
- ✅ **[QUAL-M3/M4/M5]** Fixed JSON truncation in tokenManager, hardened CSS injection in markdown, validated backend env vars.
- ✅ **[QUAL-M2]** Cleared `powerpointImageRegistry` on session switch.
- ✅ **[DEAD-M1/L1]** Removed dead code tied to legacy i18n and office-agents.

### 7. Documentation
- 🟢 *No current documentation drift identified.*

---

## ⏸️ Deferred Items

*Intentionally deferred — not forgotten, not yet unblocked. These are NOT part of the active implementation plan.*

### 🟠 [ARCH-H2/H3] Monolithic Files Consolidation (Post-Beta)
Deferred until the feature set stabilises post-beta. Splitting now would cause constant multi-file churn with no functional gain. **Trigger**: Revisit when tool additions slow down.

| File | Lines | Suggested Refactor |
|------|-------|-----------------|
| `composables/useAgentLoop.ts` | ~1,100 | Extract `runAgentLoop()` → `useAgentRunner.ts`; image flow → `useImageGeneration.ts`; keep `useAgentLoop` as thin orchestrator |
| `utils/excelTools.ts` | ~2,700 | `tools/excel/` subdirectory + `index.ts` barrel |
| `utils/powerpointTools.ts` | ~2,400 | `tools/powerpoint/` subdirectory + `index.ts` barrel |
| `utils/wordTools.ts` | ~2,100 | `tools/word/` subdirectory + `index.ts` barrel |
| `utils/outlookTools.ts` | ~700 | `tools/outlook/` subdirectory + `index.ts` barrel |

### 🟡 [TOKEN-M1] Token Limit Calibration
`MAX_CONTEXT_CHARS` (1.2M) is a conservative estimate. Needs tuning based on real usage data. **Blocked by**: Requires 2+ weeks of `LOG-H1` usage data.

---

## ❌ Won't Fix

| Item | Decision |
|------|----------|
| **TOOL-H2** - Word screenshot | No Office.js API. html2canvas/puppeteer unsupported in sandbox. `getDocumentHtml()` is the closest proxy. |
| **USR-H1** - Empty shape bullets | `placeholderFormat/type` covers 95% of cases. XML-default-bullet edge cases are rare. |
| **Phase 7F** - Dynamic tool loading | LLM handles 128+ tools fine. No usage data to define intent profiles yet. Revisit after 6+ months. |
| **DEAD-L2** - `plotDigitizer` route | LLM vision tested and found insufficient for chart accuracy. Pixel-analysis pipeline kept as-is. |
| **QUAL-L2** - `credentialCrypto` in LS | Add-in runs on dedicated PCs with per-user Windows login. Re-keying on restart is a UX regression. XSS mitigated by DOMPurify + CSP. |
| **DEAD-L3** - `clearEncryptionKeys` | False positive — still used. |
| **USR-H2** - Context bloat indicator | Context % shown live in `currentAction` since `299e0ca`. StatsBar colors change at 70%/90%. A banner is redundant noise. |

---

## 📚 Architecture Notes (Appendices)

### Tool Counts (audited 2026-03-16)

| Host | Count | Notable tools |
|------|-------|---------------|
| Word | 34 | `proposeRevision`, `editDocumentXml`, `insertOoxml`, `acceptAiChanges`, `getDocumentOoxml` |
| Excel | 27 | `eval_officejs`, `screenshotRange`, `getRangeAsCsv`, `detectDataHeaders`, `manageObject` |
| PowerPoint | 24 | `screenshotSlide`, `editSlideXml`, `reorderSlide`, `getSpeakerNotes`, `verifySlides` |
| Outlook | 9 | `eval_outlookjs`, `addAttachment`, email helpers |
| General | 6 | `executeBash` (VFS), `calculateMath`, file operations |
| **Total** | **100** | |

### Core System Files

| File | Purpose |
|------|---------|
| `utils/tokenManager.ts` | Context window management + heuristic compression |
| `utils/wordDiffUtils.ts` | Track Changes — selection & document revision application |
| `utils/wordTrackChanges.ts` | `setChangeTrackingForAi` / `restoreChangeTracking` helpers |
| `utils/toolProviderRegistry.ts` | Host → tool provider mapping (singleton) |
| `utils/mutationDetector.ts` | Shared mutation detection factory |
| `composables/useAgentLoop.ts` | Agent execution loop orchestrator |
| `composables/quickActions/` | Per-host quick action composables (4 files) |
| `skills/` | 5 host skills + 17 Quick Action skills |
