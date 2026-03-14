# DESIGN_REVIEW.md

**Last updated**: 2026-03-14 — post-PR201 + regression fixes
**Status**: All critical/high/medium items resolved. Remaining items are deferred enhancements.

---

## Completed Work (Summary)

All 56 items from the v9–v11 audit cycles are ✅ FIXED. Phases 1A through 7A fully complete.
All post-PR193 regressions (REG-M1 through REG-L3) fixed on branch `claude/review-recent-changes-1cs54`.

Key milestones:
- **Phase 1–3**: PPT bugs, image quality, UX fixes, logging, tool quality, Excel multi-curve charts, clipboard paste
- **Phase 4A**: Native Word Track Changes via `docx-redline-js` (proposeRevision + editDocumentXml)
- **Phase 4B + ARCH-H1**: Full skill system (17 skill files), composable split (useQuickActions, useSessionFiles, useMessageOrchestration)
- **Phase 5–6**: Dead code removal, error format standardization, ToolProviderRegistry, centralized constants, i18n hardening, Docker security (non-root users, nginx-unprivileged), credential migration cleanup
- **Phase 7A**: Heuristic tool result compression (`summarizeOldToolResults` in tokenManager.ts)
- **OXML-IMP1**: `proposeDocumentRevision` tool — document-wide Track Changes without selection

---

## Deferred Items

These items are intentionally deferred — not forgotten, just not prioritized yet.

### OXML Enhancements

#### OXML-IMP2 — Native Word Comments via OOXML [MEDIUM]

`docx-redline-js` exposes `injectCommentsIntoOoxml()`. Currently no tool adds Word comments.
**Path**: New `addWordComment` tool using `injectCommentsIntoOoxml()`.
**Effort**: MEDIUM

#### OXML-IMP3 — Programmatic Accept/Reject Track Changes [MEDIUM]

`docx-redline-js` exposes `acceptTrackedChangesInOoxml(author)`.
WordApi 1.6 also offers `trackedChange.accept()` / `trackedChange.reject()`.
**Path**: New `acceptAiChanges` tool to bulk-accept all KickOffice AI changes.
**Effort**: LOW–MEDIUM

#### OXML-IMP4 — Rich Content Insertion via OOXML Templates [MEDIUM]

`insertHtml()` loses complex formatting (numbered lists, table styles, section layouts).
**Path**: Generate OOXML directly for complex content types, use `insertOoxml()`.
**Effort**: HIGH — namespace management + relationship IDs are complex

#### OXML-IMP5 — PowerPoint Speaker Notes via OOXML [LOW]

`editSlideXml` targets slide XML only. Notes are in `ppt/notesSlides/notesSlideN.xml`.
**Path**: Extend `withSlideZip` pattern to accept a target XML part path.
**Effort**: LOW

---

### Context & Token Management

#### Phase 7B — TOOL-C1 (Document Re-injection) [HIGH]

Opened document text is re-sent on every message, bloating context.
**Blocked by**: Needs document pinning strategy (Phase 7A sub-task 2 — not yet implemented).
**Path**: Pin document context once, reference via placeholder in subsequent messages.

#### Phase 7B — USR-H2 (Context Bloat Indicator) [HIGH]

Users have no way to know when context is near-full until it's too late.
Already have 80% warning in StatsBar. Need actionable "start new conversation" suggestion when >90%.

#### Phase 7C — TOKEN-M1 (Token Limit Calibration) [MEDIUM]

MAX_CONTEXT_CHARS (1.2M) is a conservative estimate for GPT-5.2 (400k token window × ~3 chars/token).
**Blocked by**: Needs 2+ weeks of LOG-H1 usage data to tune accurately.
**Condition**: Only actionable once LOG-H1 data is available.

---

### Code Quality

#### QUAL-M3 — Large Vue Component Decomposition [MEDIUM]

`HomePage.vue` (592 lines), `ChatMessageList.vue` (336 lines), `ChatInput.vue` (307 lines) are large.
Candidate sub-components: `AttachedFilesList`, `MessageItem`, `ConfirmationDialogs`.
**Effort**: HIGH — careful state management and props/events design required.

---

### Won't Fix

| Item | Reason |
|------|--------|
| TOOL-H2 — Word screenshot | No Office.js API for Word screenshots. html2canvas/puppeteer don't work in add-in sandbox. `getDocumentHtml()` is the closest proxy. |
| USR-H1 — Empty shape bullets | `placeholderFormat/type` covers 95% of cases. Remaining edge cases (XML default bullets) are rare. |
| Phase 7F — Dynamic tool loading | GPT-5.2 handles 128+ tools fine. No usage data yet to define intent profiles. Revisit after 6+ months of LOG-H1 data. |

---

## Architecture Notes (for reference)

### Tool Counts (current)

| Host | Count | Notable tools |
|------|-------|---------------|
| Word | 30 | proposeRevision, proposeDocumentRevision, editDocumentXml, eval_wordjs |
| Excel | 24 | eval_officejs, screenshotRange, getRangeAsCsv, detectDataHeaders |
| PowerPoint | 21 | screenshotSlide, editSlideXml, searchIcons, insertIcon |
| Outlook | 8 | eval_outlookjs, email read/write helpers |
| General | 6 | executeBash (VFS), calculateMath, file operations |
| **Total** | **89** | |

### Key Files

| File | Purpose |
|------|---------|
| `frontend/src/utils/tokenManager.ts` | Context window management + Phase 7A compression |
| `frontend/src/utils/wordDiffUtils.ts` | Track Changes — selection (`applyRevisionToSelection`) + document (`applyRevisionToDocument`) |
| `frontend/src/utils/wordTrackChanges.ts` | setChangeTrackingForAi / restoreChangeTracking helpers |
| `frontend/src/utils/toolProviderRegistry.ts` | Host → tool provider mapping (singleton) |
| `frontend/src/composables/useAgentLoop.ts` | Agent execution loop (881 lines after ARCH-H1 refactor) |
| `frontend/src/skills/` | 5 host skills + 17 Quick Action skills |

---

*See CHANGELOG.md for full version history.*
