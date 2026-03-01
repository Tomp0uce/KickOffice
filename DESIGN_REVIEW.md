# Design Review v3 — Status

**Date**: 2026-03-01 | **Scope**: Full codebase audit (security, logic, quality, infra)
**Progress**: 🟢 105 implemented · 🔴 26 remaining · 🟡 4 deferred · **135 total**

---

## 🔴 Outstanding Items

### Security / Correctness

**BH6** · HIGH · `backend/src/routes/upload.js`
Upload file type validated by extension/MIME only — no magic-byte check. Attacker can rename `.exe` to `.pdf` and bypass.
_Fix_: Use `file-type` npm package to validate actual file bytes.

**UC2** · MEDIUM · `frontend/src/utils/credentialStorage.ts`
XOR "obfuscation" of API keys provides no real security — trivially reversible. Gives false confidence.
_Fix_: Store only in sessionStorage (not localStorage) by default; document the trade-off; remove misleading "secure" comments.

**XH1** · HIGH · `frontend/src/api/backend.ts`
No CSRF token on API calls from the Office add-in. Office add-ins run in a sandboxed iframe, reducing exposure, but remains best practice.
_Fix_: Add CSRF token in request header (backend generates, frontend reads from cookie/header).

**EM1** · MEDIUM · `frontend/src/main.ts`
`useStorage` (VueUse composable) called at module level, outside any Vue component setup context. May cause reactivity warnings or silently fail.
_Fix_: Move `useStorage` calls inside component `setup()` or use raw `localStorage` directly at module level.

**EM2** · LOW · `frontend/src/main.ts`
`ResizeObserver` is monkey-patched globally to suppress loop-limit errors. Can mask legitimate errors from other components.
_Fix_: Remove the global patch; handle `ResizeObserver` errors locally where needed.

---

### Code Quality / Type Safety

**UM1** · MEDIUM · various tool files
Pervasive `as unknown as SomeType` casts throughout tool implementations — bypasses TypeScript safety with no guard.
_Fix_: Replace casts with proper type guards or typed API boundaries.

**UM2** · MEDIUM · `frontend/src/utils/*Tools.ts`
Tool `args` parameters typed as `any` everywhere — no input validation on tool execution.
_Fix_: Use `Record<string, unknown>` + type guards per tool, or generate typed arg interfaces from schemas.

**CH5** · MEDIUM · `frontend/src/composables/useAgentLoop.ts`
`catch (err: any)` and tool args typed `any` throughout the agent loop.
_Fix_: Use `catch (err: unknown)` with `instanceof Error` guards; type tool args at call sites.

**PL2** · LOW · `frontend/src/components/CustomButton.vue`
`icon` prop typed as `any` — accepts anything without type safety.
_Fix_: Type as `Component | null` from vue.

**PL3** · LOW · `frontend/src/components/SingleSelect.vue`
Props `icon`, `customFrontIcon`, `placeholder` all typed as `any`.
_Fix_: Use `Component | null` for icon props; `string` for placeholder.

**CM3** · MEDIUM · `frontend/src/composables/useOfficeInsert.ts`, `useAgentLoop.ts`
Several functions exceed 180 lines — hard to review, test, or maintain.
_Fix_: Extract logical sub-steps into named helpers.

**XM1** · LOW · `frontend/src/composables/useAgentPrompts.ts` and others
Deeply nested ternary chains (`a ? b : c ? d : e ? ...`) repeated 10+ times for host detection.
_Fix_: Replace with a `switch` or lookup map.

---

### Error Handling / Reliability

**BM6** · MEDIUM · `backend/src/routes/`
Some routes use `console.error`, others use `systemLog`, some use neither. No consistent pattern.
_Fix_: Standardize all error paths to `systemLog('ERROR', ...)`.

**BM10** · LOW · `backend/src/server.js`
No request ID generated per request — impossible to correlate logs across middleware.
_Fix_: Add `express-request-id` or a simple UUID middleware; attach to `res.locals`.

**UL3** · LOW · `frontend/src/utils/*Tools.ts`
Some tools return `"Error: ..."` strings, others throw, others return empty string on failure. No consistent contract.
_Fix_: Standardize to always throw on error; let the composable layer handle user-facing messages.

**CL2** · LOW · `frontend/src/composables/useImageActions.ts`
`cleanContent` strips think-tags with regex; `splitThinkSegments` parses them structurally. Malformed tags behave differently in each.
_Fix_: Extract a single `parseThinkTags(text)` utility used by both.

---

### i18n / UI

**PM2** · MEDIUM · `frontend/src/pages/SettingsPage.vue:190-193, 200, 470`
`$t("darkModeLabel") || "Dark mode"` pattern — fallback strings mask missing i18n keys and will always show English.
_Fix_: Add all missing keys to `en.json` and `fr.json`; remove fallback strings.

**PM5** · MEDIUM · `frontend/src/components/SingleSelect.vue`
Dropdown position calculated once on open using `getBoundingClientRect()` — not recalculated on scroll/resize, causing drift.
_Note_: scroll/resize listeners were added in a prior fix but only update on re-open, not continuously.
_Fix_: Recalculate on `scroll` and `resize` events while dropdown is open, or use a `position: sticky` approach.

**PM10** · LOW · `frontend/src/pages/HomePage.vue`, `SettingsPage.vue`
Mix of `t()` (composition API) and `$t()` (global/Options API) in the same template.
_Fix_: Standardize on `t()` from `useI18n()` throughout; avoid `$t()` in `<script setup>` files.

---

### Architecture / Naming

**TM1** · MEDIUM · `frontend/src/types/index.d.ts`
All types declared ambient (no `export` keyword) — available everywhere without imports, bypassing module boundaries. Makes refactoring and dead-code detection harder.
_Fix_: Add `export` to all interfaces; update import sites.

**TM2** · LOW · `frontend/src/types/index.d.ts:74` and `frontend/src/utils/hostDetection.ts:1`
`OfficeHostType` defined in two places — can drift.
_Fix_: Keep one definition (in `hostDetection.ts`), export it, and remove from `index.d.ts`.

**UL4** · LOW · `frontend/src/utils/markdown.ts` vs `officeRichText.ts`
`markdown.ts` processes Office rich text; `officeRichText.ts` does markdown-related work. Names are swapped.
_Fix_: Rename files to match their actual content.

**AL1** · LOW · `frontend/src/api/common.ts`
File is named as an API utility but contains Word-specific `Office.run` logic.
_Fix_: Move content to `frontend/src/utils/wordApi.ts` or similar.

**TL2** · LOW · `frontend/src/utils/insertTypes.ts`
`insertTypes` export uses lowercase plural — inconsistent with other constant naming conventions.
_Fix_: Rename to `INSERT_TYPES` or `InsertType` depending on usage pattern.

---

### Infrastructure

**IM8** · MEDIUM · CI scripts
CI has an infinite-loop guard based on a hardcoded iteration count — fragile if loop runs legitimately longer.
_Fix_: Use a time-based timeout instead of iteration count.

**IL3** · LOW · `frontend/vite.config.ts`
`chunkSizeWarningLimit` raised to 1000 kB to suppress Vite warnings instead of addressing bundle size.
_Fix_: Restore default (500 kB); investigate and split large chunks.

---

## 🟡 Deferred Items

**IC2** · `frontend/Dockerfile`, `backend/Dockerfile`
Containers run as root (Node.js default). Adding `USER node` is the best-practice fix but retained intentionally for this deployment.

**IH2** · `frontend/Dockerfile`
Private IP address baked into the default `VITE_BACKEND_URL` build arg — leaks in the compiled JS.
Retained as-is; users must override at build time.

**IH3** · `.env.example`
External DuckDNS domain used as default example value. Retained; users replace with their own values.

**UM10** · `frontend/src/utils/powerpointTools.ts`
Character-by-character HTML reconstruction for PowerPoint — high complexity, low ROI. Deferred pending PowerPoint feature priority.

---

## 🟢 Implemented (105 items)

### Backend
🟢 BC1–BC4 · BH1–BH5 · BH7 · BM1–BM5 · BM7–BM9 · BL1–BL4

### Frontend Utils
🟢 UC1 · UC3 · UH1–UH7 · UM3–UM9 · UL1 · UL2

### Composables
🟢 CC1 · CC2 · CH1–CH4 · CH6 · CH7 · CM1 · CM2 · CM4–CM11 · CL1 · CL3–CL5

### Infrastructure
🟢 IC1 · IC3 · IH1 · IH4 · IH5 · IM1–IM7 · IL1 · IL2 · IL4–IL7

### Pages / Components / API
🟢 PC1 · PH1–PH5 · AH1 · AH2 · PM1 · PM3 · PM4 · PM6–PM9 · PM11 · AM1–AM4 · PL1 · PL4 · PL5 · TL1 · XM2
