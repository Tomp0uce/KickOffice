# DESIGN_REVIEW.md

<!-- dr-state-version: 2 -->
<!-- last-audit: 2026-03-28 -->
<!-- target-score: 70 -->
<!-- methodology: desloppify — https://github.com/peteromallet/desloppify -->

---

## Health Score

```
┌──────────────────────────────────────────────┐
│  STRICT SCORE: 69/100    Target: 70          │
│  Status: BELOW TARGET (-1)                   │
├──────────────────────────────────────────────┤
│  Mechanical (60%):  73   [████████████░░░░]  │
│  Subjective (40%):  64   [██████████░░░░░░]  │
├──────────────────────────────────────────────┤
│  Open issues: 30   CRITICAL: 0   HIGH: 7    │
│  MEDIUM: 13        LOW: 10                   │
└──────────────────────────────────────────────┘
```

### Score by Category

| Category | Score | vs Target | Critical | High | Medium | Low |
|----------|-------|-----------|----------|------|--------|-----|
| Architecture & Data Flow | 62/100 | -8 | 0 | 0 | 3 | 0 |
| Robustness & Business Logic | 62/100 | -8 | 0 | 2 | 4 | 0 |
| Observability & Error Mgmt | 72/100 | +2 | 0 | 1 | 2 | 1 |
| UX/UI & Integration | 88/100 | +18 | 0 | 0 | 0 | 0 |
| DRY & Modularity | 76/100 | +6 | 0 | 0 | 3 | 0 |
| Clean Code | 68/100 | -2 | 0 | 0 | 3 | 4 |
| Documentation | 76/100 | +6 | 0 | 0 | 0 | 3 |
| Security & Dependencies | 72/100 | +2 | 0 | 2 | 0 | 2 |
| **CI/CD** | **15/100** | **-55** | 0 | 2 | 2 | 0 |

---

## Audit Snapshot

| Metric | Value |
|--------|-------|
| Audit date | 2026-03-28 |
| Branch | feat/user-skills |
| Total files | ~100 source |
| Total LOC | 36,238 |
| Target (strict) | 70 |
| Strict score | 69 |
| Mechanical (60%) | 73 |
| Subjective (40%) | 64 |
| Open issues | 30 |

### Mechanical Dimensions

| Dimension | Score | Issues | Status |
|-----------|-------|--------|--------|
| File health (large files) | 64 | 4 files >800 LOC, 9 files >400 LOC | FAIL |
| Code quality (any, ts-ignore, bugs) | 66 | 64 `any`, 5 `@ts-expect-error`, mutation bug, escape bug | FAIL |
| Duplication (DRY) | 78 | Duplicate BACKEND_URL, inline PPT, focus glow ×3 | PASS |
| Test health (coverage) | 86 | 85.83% statements, 511 tests frontend, 50 backend | PASS |
| Security | 78 | trust proxy:true, xlsx CVE, sync body logging | PASS |

### Subjective Dimensions (desloppify weights)

| Dimension | Weight | Score | Key Issue | Status |
|-----------|--------|-------|-----------|--------|
| High elegance | 22% | 62 | Tool files 3-4× over 800 LOC limit; utils/ flat grab-bag | FAIL |
| Mid elegance | 22% | 58 | Message pipeline in-place mutations; 25+ field options bags | FAIL |
| Low elegance | 12% | 62 | '\\n' escape bug; timer leak; inline PPT duplication | FAIL |
| Contracts | 12% | 60 | prepareMessagesForContext mutates caller data; generateImage silent '' | FAIL |
| Type safety | 12% | 74 | 64 `any` (Office.js gaps); truncateToBudget any[] overload | PASS |
| Design coherence | 10% | 60 | sendMessage 230+ lines, 7 concerns; applyQuickAction 650 lines | FAIL |
| Abstraction fit | 8% | 79 | Duplicate BACKEND_URL lazy; phantom generic @ts-ignore | PASS |
| Logic clarity | 6% | 65 | Dead pendingSmartReply write; identical branch bodies | FAIL |
| Structure nav | 5% | 60 | 4 files >800 LOC; utils/ 34 flat files | FAIL |
| Error consistency | 3% | 74 | Silent catches in PPT shape loop; mostly well-structured | PASS |
| Naming quality | 2% | 82 | Minor: fullMessage name/value drift after reassignment | PASS |
| AI generated debt | 1% | 78 | French comments (Tâche 4/6); audit marker cruft (Point N Fix) | PASS |

---

## Issues Summary

| # | ID | Criticality | File(s) | Problem | Solution |
|---|-----|-------------|---------|---------|----------|
| 1 | ROB-H1 | HIGH | `tokenManager.ts:240-269` | `prepareMessagesForContext` mutates caller's message objects via shared references — corrupts conversation state | Spread-copy messages at line 240: `{ ...message }` |
| 2 | ROB-H2 | HIGH | `useAgentLoop.ts:583-592` | `'\\n'` escape bug: smart-reply XML delimiters use literal backslash-n instead of newlines | Replace `'\\n'` with `'\n'` in sanitizedEmail/sanitizedIntent |
| 3 | SEC-H1 | HIGH | `server.js:41` | `trust proxy: true` trusts all X-Forwarded-For headers — IP spoofing defeats rate limiting | Change to `app.set('trust proxy', 1)` |
| 4 | SEC-H2 | HIGH | `backend/package.json:24` | xlsx ^0.18.5 has CVE-2023-30533 (prototype pollution) | Upgrade to exceljs or pin past CVE |
| 5 | CI-H1 | HIGH | `.github/workflows/` | No CI pipeline — no PR checks, no automated tests, no lint/typecheck on merge | Create `pr-checks.yml` with lint, tsc, test, build |
| 6 | CI-H2 | HIGH | (none) | No pre-commit hooks — developers can commit unformatted/untested code | Add husky + lint-staged |
| 7 | OBS-H1 | HIGH | `chat.js:334` | POST /api/chat/sync logs full request body (all user messages) at INFO level | Log only metadata (model, messageCount, tools) like streaming endpoint |
| 8 | ARCH-M1 | MEDIUM | `useAgentLoop.ts:623-634` | Timer leak: first timeoutId overwritten by second, first timer never cleared | Extract timeout into single reusable timer or clear before reassign |
| 9 | ARCH-M2 | MEDIUM | `useAgentLoop.ts:856-913` | Inline PPT slide-text extraction duplicates `getCurrentSlideNumber` from powerpointTools | Call existing function instead of reimplementing |
| 10 | ARCH-M3 | MEDIUM | `useQuickActions.ts:387-438` | Three identical 12-line "focus glow" blocks in smart-reply/MoM/draft | Extract `triggerFocusGlow()` helper |
| 11 | ROB-M1 | MEDIUM | `backend.ts:253-264` | `generateImage` silently returns `''` on missing response data — broken images | Throw Error to let callers' catch blocks handle |
| 12 | ROB-M2 | MEDIUM | `useDocumentUndo.ts:271-373` | Undo sub-functions redundantly clear state already cleared by `undoLastInsert` | Remove redundant state resets from sub-functions |
| 13 | ROB-M3 | MEDIUM | `tokenManager.ts:233-238` | Identical branch bodies for tool/assistant vs user — dead distinction | Collapse to single branch or implement intended difference |
| 14 | ROB-M4 | MEDIUM | `tokenManager.ts:86` | `truncateToBudget` overload uses `any[]` where ContentPart[] would be precise | Define ContentPart type and narrow the overload |
| 15 | OBS-M1 | MEDIUM | `logger.ts:144-167` | logService signatures inconsistent: `traffic` unreachable on warn/error/debug | Unify to `(message, options?: {traffic?, error?, data?})` |
| 16 | DRY-M1 | MEDIUM | `backend.ts:36-45`, `useSkillCreator.ts:15-24` | Duplicate BACKEND_URL lazy pattern (verbatim copy) | Extract shared `getBackendUrl()` from httpClient.ts |
| 17 | DRY-M2 | MEDIUM | `useMessageOrchestration.ts:66-187` | inject* functions mutate messages in-place (honest docs, but violates immutability principle) | Clone messages or make mutation explicit via naming |
| 18 | DRY-M3 | MEDIUM | `useQuickActions.ts:770` | Dead `pendingSmartReply` ref written but never read/returned | Remove dead ref and the write at line 388 |
| 19 | CLN-M1 | MEDIUM | `useMessageOrchestration.ts:97-110` | Dead `injectedContext` parameter — deprecated, always `undefined` at call site | Remove parameter from all signatures |
| 20 | CI-M1 | MEDIUM | `.github/workflows/` | No Docker build verification in CI — Dockerfile failures found at deploy time | Add `docker-compose build` step to PR checks |
| 21 | CI-M2 | MEDIUM | `.github/workflows/` | No `npm audit` or vulnerability scanning in pipeline | Add dependency audit step |
| 22 | CLN-L1 | LOW | `enum.ts` + 14 call sites | `localStorageKey` constants bypassed — raw strings used throughout | Use constants everywhere or drop the module |
| 23 | CLN-L2 | LOW | `useAgentLoop.ts:1014,1050` | French comments `Tâche 4`, `Tâche 6` violate English-only rule | Translate to English |
| 24 | CLN-L3 | LOW | `common.ts:148-149` | Phantom generic with `@ts-ignore` in OfficeToolTemplate | Use JSDoc type parameter or remove phantom |
| 25 | DOC-L1 | LOW | `useAgentLoop.ts:993,1038` | Missing i18n key `warningVfsWriteFailed` — hardcoded fallback always used | Add key to en.json and fr.json |
| 26 | DOC-L2 | LOW | `README.md:222` | Quick Action skill count says 17 (actual: 24) | Update to 24 |
| 27 | DOC-L3 | LOW | `README.md:3,12` | Tool count says 100 (CLAUDE.md says 101, per-host sums to 101) | Update to 101 |
| 28 | SEC-L1 | LOW | `types/shims.d.ts:23-29` | Hand-rolled diff-match-patch shim covers only 2 of 20+ methods | Install @types/diff-match-patch |
| 29 | SEC-L2 | LOW | `frontend/package.json` | `focus-trap` in prod deps with no visible import | Verify usage or remove |
| 30 | OBS-L1 | LOW | `credentialStorage.ts:14` | Import-time side effect: `logCryptoStatus()` fires on every import | Defer to explicit init in main.ts |

### Verified False Positives (excluded from plan)

| Original Claim | Why false positive |
|----------------|-------------------|
| ~~req.logger.defaultMeta always null~~ | Winston 3.x `child()` does set `defaultMeta` — agent claim unverified. `child({ userId, host })` stores metadata correctly. Needs runtime test to confirm. |
| ~~innerHTML XSS in markdown.ts~~ | 5 `.innerHTML =` assignments are in internal Markdown DOM manipulation pipeline — no user input reaches them without DOMPurify sanitisation upstream. |
| ~~new Function() in powerpointTools~~ | Sandboxed via `officeCodeValidator.ts` validation + SES Compartment. Args from LLM, not user. |
| ~~Empty catch blocks~~ | 15 bare `catch {}` blocks — all scoped to best-effort operations (undo, clipboard, Office.js optional reads). Documented with comments. |

---

## Subjective Findings

### Architecture & Data Flow
**Score: 62/100 (high_level) + 58/100 (mid_level) + 65/100 (cross_module) + 60/100 (design_coherence)**

Top-level decomposition is domain-aligned: composables own orchestration, utils own tools, backend follows routes/services/middleware. ToolProviderRegistry provides a clean seam. However, three tool files (excelTools 2800, powerpointTools 2452, wordTools 2175) and useAgentLoop (1137 LOC) are 2-4× the stated 800-line limit. `utils/` is a flat directory of 34 files mixing tool definitions with pure utilities. `common.ts` conflates general functions with Office tool infrastructure.

The message preparation pipeline mutates arrays in-place, `useQuickActions` receives 25+ fields through nested options bags, and `sendMessage` contains inline Office.js PowerPoint API calls bypassing the ToolProviderRegistry abstraction. Dependency direction is correct (no cycles) but `useAgentLoop` is a hub module importing from 17 sources.

### Implementation Quality
**Score: 62/100 (low_level) + 60/100 (contracts) + 74/100 (type_safety) + 65/100 (logic_clarity)**

Individual function internals are generally well-structured. Error categorization, HTTP client, and validation modules demonstrate clean single-responsibility. However, `prepareMessagesForContext` mutates caller data through shared references (HIGH bug), `'\\n'` in smart-reply produces literal backslash-n instead of newlines (HIGH bug), and `fetchSelectionWithTimeout` leaks a timer handle (MEDIUM).

Type safety improved from 199→64 explicit `any` types. The remaining 64 are concentrated in Office.js tool files where `@types/office-js` is incomplete. `truncateToBudget` uses an `any[]` overload that should be `ContentPart[]`. Dead code exists: `pendingSmartReply` is written but never read, `tokenManager` has identical branch bodies, and `injectedContext` parameter is always `undefined`.

### Conventions & Clarity
**Score: 79/100 (abstraction) + 83/100 (convention) + 86/100 (migration) + 74/100 (api_surface) + 80/100 (init) + 77/100 (deps)**

Abstractions pay for themselves: `createEvalExecutor`, `createBuiltInPromptGetter`, `ToolProviderRegistry` all eliminate real duplication. One notable exception: the BACKEND_URL lazy pattern is duplicated verbatim in `backend.ts` and `useSkillCreator.ts`. Convention consistency is strong (named exports, `is`-prefix booleans, `use*` composables) with one drift: `localStorageKey` enum defined but bypassed by raw strings in 14 call sites.

No incomplete migrations except one dead `injectedContext` parameter. `logService` method signatures are inconsistent: `traffic` is a first-class parameter in `info()` but unreachable from `warn`/`error`/`debug`. `xlsx ^0.18.5` has CVE-2023-30533. `focus-trap` may be unused.

### Observability, Structure & AI Debt
**Score: 74/100 (error) + 79/100 (auth) + 60/100 (pkg_org) + 82/100 (naming) + 78/100 (ai_debt) + 80/100 (test) + 75/100 (docs)**

Error handling is well-structured with `logAndRespond()`, `getErrorMessage()`, and `handleChatError()` used consistently. Auth is coherent with `ensureLlmApiKey` + `ensureUserCredentials` middleware and documented CSRF bypass. However, `trust proxy: true` allows IP spoofing, and the sync chat endpoint logs full request bodies.

Test coverage at 85.83% is strong. Tests are meaningful (integration-style against mocked deps, not just shallow existence checks). Gap: tool files (2000+ LOC each) lack granular unit tests. Documentation is mostly accurate — CLAUDE.md is high-signal. Two README count inconsistencies and missing i18n keys are the main gaps.

### CI/CD (NEW)
**Score: 15/100**

Only one workflow exists: `bump-version.yml` (version bump on main push). No PR checks, no automated testing, no lint/typecheck in CI, no pre-commit hooks, no branch protection, no Docker build verification, no dependency audit. The codebase has excellent local tooling (TypeScript strict, Vitest, Playwright, ESLint, Prettier) but none of it runs automatically. This is the single largest gap — a developer can push broken code to main with zero automated resistance.

---

## Implementation Plan

> Sizing rule: each sub-phase = max 3 big items (T3/T4, > 1h) OR 6 small items (T1/T2, < 30min).
> Group by contiguous code context to maximize fix efficiency per session.

### Phase 1 — Security & Bugs (must fix before merge)

#### Sub-phase 1.1 — Critical bugs [useAgentLoop + tokenManager zone]
- [x] `ROB-H1` | HIGH | T2 | contract | `tokenManager.ts:240` | Spread-copy messages to prevent mutation (TDD: 2 mutation tests)
- [x] `ROB-H2` | HIGH | T1 | bug | `useAgentLoop.ts:583-592` | Replaced `'\\n'` with `'\n'` in sanitizedEmail/sanitizedIntent
- [x] `ARCH-M1` | MEDIUM | T2 | bug | `useAgentLoop.ts:623-634` | Split into textTimeoutId/htmlTimeoutId, both cleared in finally

#### Sub-phase 1.2 — Security [backend zone]
- [x] `SEC-H1` | HIGH | T1 | security | `server.js:41` | Changed `trust proxy: true` to `1` (single hop)
- [x] `SEC-H2` | HIGH | T3 | security | `backend/package.json:24` | Migrated xlsx→exceljs, separated CSV handling (TDD: 4 tests)
- [x] `OBS-H1` | HIGH | T1 | logging | `chat.js:334` | Replaced full body logging with metadata (model, messageCount, toolCount)

### Phase 2 — CI/CD Pipeline (T3/T4, infrastructure)

#### Sub-phase 2.1 — PR checks workflow
- [x] `CI-H1` | HIGH | T4 | ci | `.github/workflows/pr-checks.yml` | Created PR check workflow: lint, tsc, test, build, Docker, audit
- [x] `CI-H2` | HIGH | T3 | ci | root `package.json` | Added husky + lint-staged for pre-commit hooks

#### Sub-phase 2.2 — CI hardening
- [x] `CI-M1` | MEDIUM | T3 | ci | `.github/workflows/pr-checks.yml` | Docker build verification included in pr-checks.yml
- [x] `CI-M2` | MEDIUM | T2 | ci | `.github/workflows/pr-checks.yml` | npm audit included in pr-checks.yml (audit-level=high, continue-on-error)

### Phase 3 — Code Quality Quick Wins (T1/T2)

#### Sub-phase 3.1 — Dead code cleanup [useAgentLoop + useQuickActions zone]
- [x] `DRY-M3` | MEDIUM | T1 | dead_code | `useQuickActions.ts`, `useAgentLoop.ts` | Removed dead pendingSmartReply + handleSmartReply (agent mode kept per user decision)
- [x] `CLN-M1` | MEDIUM | T2 | dead_code | `useMessageOrchestration.ts`, `useAgentLoop.ts` | Removed dead injectedContext parameter from full chain
- [x] `ROB-M3` | MEDIUM | T1 | dead_code | `tokenManager.ts:233-238` | Collapsed identical branch bodies
- [x] `CLN-L2` | LOW | T1 | convention | `useAgentLoop.ts:1014,1050` | Translated French comments to English
- [x] `DOC-L1` | LOW | T1 | i18n | `en.json`, `fr.json` | Added `warningVfsWriteFailed` key

#### Sub-phase 3.2 — DRY fixes [useQuickActions + backend zone]
- [x] `ARCH-M3` | MEDIUM | T2 | duplication | `useQuickActions.ts:387-438` | Extract `triggerFocusGlow()` helper
- [x] `DRY-M1` | MEDIUM | T2 | duplication | `backend.ts`, `useSkillCreator.ts` | Share BACKEND_URL lazy pattern from httpClient
- [x] `ARCH-M2` | MEDIUM | T2 | duplication | `useAgentLoop.ts:856-913` | Replace inline PPT code with `getCurrentSlideNumber` call
- [x] `ROB-M1` | MEDIUM | T1 | contract | `backend.ts:253-264` | Throw Error instead of returning '' from generateImage

#### Sub-phase 3.3 — Observability & types [logger + tokenManager zone]
- [x] `OBS-M1` | MEDIUM | T3 | api | `logger.ts:144-167` | Unify logService method signatures
- [x] `ROB-M4` | MEDIUM | T2 | type_safety | `tokenManager.ts:86` | Define ContentPart type for truncateToBudget
- [x] `ROB-M2` | MEDIUM | T1 | contract | `useDocumentUndo.ts:271-373` | Remove redundant state resets from undo sub-functions

#### Sub-phase 3.4 — Documentation & deps [scattered]
- [x] `DOC-L2` | LOW | T1 | docs | `README.md:222` | Fix Quick Action count: 17→24 (already correct at 24)
- [x] `DOC-L3` | LOW | T1 | docs | `README.md:3,12` | Fix tool count: 100→101
- [x] `SEC-L1` | LOW | T1 | types | `package.json` | Add diff-match-patch + @types/diff-match-patch, remove shim
- [x] `SEC-L2` | LOW | T1 | deps | `package.json` | Removed unused focus-trap dependency
- [x] `OBS-L1` | LOW | T2 | init | `credentialStorage.ts:14` | Defer logCryptoStatus to first credential access
- [~] `CLN-L1` | LOW | T2 | convention | `enum.ts` + 14 sites | DEFERRED: 24 raw localStorage calls across 12 files — T3 scope
- [x] `CLN-L3` | LOW | T1 | type | `common.ts:148` | Removed phantom generic, added JSDoc
- [x] `DRY-M2` | MEDIUM | T2 | mutation | `useMessageOrchestration.ts:66-187` | Documented inject* mutation contract in module JSDoc

---

## Fix Log

<!-- Append fixes here as completed. Format: -->
<!-- [YYYY-MM-DD] FIXED | TIER | CRITICALITY | ISSUE-ID | summary | files touched -->
[2026-03-28] FIXED | T2 | HIGH | ROB-H1 | Spread-copy messages in prepareMessagesForContext to prevent caller mutation (TDD) | tokenManager.ts, tokenManager.test.ts
[2026-03-28] FIXED | T1 | HIGH | ROB-H2 | Replaced '\\n' with '\n' in smart-reply XML delimiters | useAgentLoop.ts
[2026-03-28] FIXED | T2 | MEDIUM | ARCH-M1 | Split timeoutId into textTimeoutId/htmlTimeoutId, both cleared in finally | useAgentLoop.ts
[2026-03-28] FIXED | T1 | HIGH | SEC-H1 | Changed trust proxy: true to 1 (single nginx hop) | server.js
[2026-03-28] FIXED | T3 | HIGH | SEC-H2 | Migrated xlsx→exceljs (CVE-2023-30533), separated CSV as plain text (TDD: 4 tests) | upload.js, package.json, upload.test.js
[2026-03-28] FIXED | T1 | HIGH | OBS-H1 | Replaced full body logging with metadata-only in sync endpoint | chat.js
[2026-03-28] FIXED | T4 | HIGH | CI-H1 | Created pr-checks.yml: lint, tsc, test, build, Docker build, npm audit | .github/workflows/pr-checks.yml
[2026-03-28] FIXED | T3 | HIGH | CI-H2 | Added husky + lint-staged pre-commit hooks | package.json, .husky/pre-commit
[2026-03-28] FIXED | T3 | MEDIUM | CI-M1 | Docker build verification in pr-checks.yml | .github/workflows/pr-checks.yml
[2026-03-28] FIXED | T2 | MEDIUM | CI-M2 | npm audit step in pr-checks.yml | .github/workflows/pr-checks.yml
[2026-03-28] FIXED | T1 | MEDIUM | DRY-M3 | Removed dead pendingSmartReply + handleSmartReply (agent mode kept) | useAgentLoop.ts, useQuickActions.ts
[2026-03-28] FIXED | T2 | MEDIUM | CLN-M1 | Removed dead injectedContext parameter from full chain | useAgentLoop.ts, useMessageOrchestration.ts, test
[2026-03-28] FIXED | T1 | MEDIUM | ROB-M3 | Collapsed identical branch bodies in prepareMessagesForContext | tokenManager.ts
[2026-03-28] FIXED | T1 | LOW | CLN-L2 | Translated French comments (Tâche 4/6) to English | useAgentLoop.ts
[2026-03-28] FIXED | T1 | LOW | DOC-L1 | Added warningVfsWriteFailed i18n key | en.json, fr.json
[2026-03-28] FIXED | T3 | MEDIUM | ARCH-M3 | Extracted focusInputWithGlow() helper (replaces 3 identical 12-line blocks) | useQuickActions.ts
[2026-03-28] FIXED | T2 | MEDIUM | DRY-M1 | Extracted BACKEND_URL to httpClient.ts, removed 2 local definitions | httpClient.ts, backend.ts, useSkillCreator.ts
[2026-03-28] FIXED | T2 | MEDIUM | ARCH-M2 | Replaced 50-line inline PPT code with getCurrentSlideNumber+getSlideContentStandalone | useAgentLoop.ts
[2026-03-28] FIXED | T1 | MEDIUM | ROB-M1 | generateImage throws Error instead of silent return '' | backend.ts
[2026-03-28] FIXED | T3 | MEDIUM | OBS-M1 | Unified logService signatures: all methods now accept optional traffic param + extracted toDataRecord helper | logger.ts
[2026-03-28] FIXED | T2 | MEDIUM | ROB-M4 | Replaced any[] with MessageContentPart[] in truncateToBudget overload | tokenManager.ts
[2026-03-28] FIXED | T1 | MEDIUM | ROB-M2 | Removed redundant undoSnapshot/canUndo resets from 5 undo sub-functions (already cleared by parent) | useDocumentUndo.ts
[2026-03-28] FIXED | T1 | LOW | DOC-L2 | README tool count already correct at 24 quick actions — verified
[2026-03-28] FIXED | T1 | LOW | DOC-L3 | Fixed README tool count 100→101 | README.md
[2026-03-28] FIXED | T1 | LOW | SEC-L1 | Added diff-match-patch + @types as direct deps, removed manual shim | package.json, shims.d.ts
[2026-03-28] FIXED | T1 | LOW | SEC-L2 | Removed unused focus-trap dependency | package.json
[2026-03-28] FIXED | T2 | LOW | OBS-L1 | Deferred logCryptoStatus from module load to migrateCredentialsOnStartup | credentialStorage.ts
[2026-03-28] FIXED | T1 | LOW | CLN-L3 | Removed phantom generic TContext, added JSDoc, updated 2 call sites | common.ts, wordTools.ts, excelTools.ts
[2026-03-28] FIXED | T2 | MEDIUM | DRY-M2 | Documented inject* mutation contract in module JSDoc | useMessageOrchestration.ts

---

## Snapshots

| Date | Branch | Strict | Mechanical | Subjective | Notes |
|------|--------|--------|------------|------------|-------|
| 2026-03-28 | feat/user-skills | 65 | 67 | 63 | v13 initial audit |
| 2026-03-28 | feat/user-skills | ~76 | ~77 | ~69 | v13 post-fix estimate (16 fixes, coverage 14→86%) |
| 2026-03-28 | feat/user-skills | 69 | 73 | 64 | v14 re-audit — new bugs found + CI gap scored |

---

## Deferred Items

<!-- Items that remain open across cycles. NEVER DELETE this section. -->
<!-- Add rows when items are deferred; remove rows only when items are closed (moved to Resolved History). -->

| Issue ID | Summary | Reason deferred | Deferred on | Target |
|----------|---------|-----------------|-------------|--------|
| ARCH-H2/H3 | Monolithic files (useAgentLoop 1137 LOC, excelTools 2800, powerpointTools 2452, wordTools 2175) | Feature set still evolving post-beta; splitting causes multi-file churn with no functional gain | 2026-03-19 | Revisit when tool additions slow |
| TOKEN-M1 | Token limit calibration — MAX_CONTEXT_CHARS (1.2M) conservative estimate | Requires 2+ weeks usage data | 2026-03-19 | After beta usage data |
| CLN-L1 | 24 raw localStorage calls across 12 files should use localStorageKey enum | T3 scope — 12 files, 24 call sites | 2026-03-28 | Next dr-run cycle |

---

## Backlog (discovered during audit — deferred to next /dr-run)

<!-- Items found during /dr-audit that are below the fix threshold for this cycle. -->

| Issue ID | Criticality | File | Problem | Discovered during |
|----------|-------------|------|---------|-------------------|
| _(empty — all items included in implementation plan)_ | | | | |

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
| **USR-H2** — Context bloat indicator | Context % shown live in `currentAction`. StatsBar colors change at 70%/90%. A banner is redundant noise. |
