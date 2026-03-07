# DESIGN_REVIEW.md — Code Audit v6

**Date**: 2026-03-07
**Version**: 6.1
**Scope**: Full project design review — Fonctionnalites, Code Quality, Architecture, UX/UI

---

## Etat de sante global

Le projet KickOffice est **solide architecturalement** avec une bonne separation frontend/backend, un sandbox SES pour l'execution de code, et une gestion correcte d'Office.onReady.

| Severite | Total | DONE | OPEN | N/A |
| -------- | ----- | ---- | ---- | --- |
| BLOQUANT/CRITIQUE | 5 | 4 | 0 | 1 |
| MAJEUR/IMPORTANT | 14 | 14 | 0 | 0 |
| MINEUR/AMELIORATION | 16 | 15 | 0 | 1 |
| **Total** | **35** | **33** | **0** | **2** |

---

## AXE 1 — REVUE DES FONCTIONNALITES

### BLOQUANT/CRITIQUE

- **F-C1** — Excel : `safeGetSheet()` avec `getItemOrNullObject` ✅ DONE (PR #167)
- **F-C2** — PowerPoint : description de `proposeShapeTextRevision` corrigee ✅ DONE (PR #167)

### MAJEUR/IMPORTANT

- **F-M1** — Sanitisation prompt injection ✅ DONE (PR #169)
  - `useAgentPrompts.ts` : `sanitize()` echappe newlines, tabs, chars markdown/HTML

- **F-M2** — PowerPoint : outils speaker notes et images ✅ DONE (PR #169)
  - `getSpeakerNotes`, `setSpeakerNotes`, `insertImageOnSlide` ajoutes dans `powerpointTools.ts`

- **F-M3** — Excel : `sortRange` interface consistante ✅ DONE (PR #169)
  - `address` optionnel ; si absent utilise `getSelectedRange()`

- **F-M4** — Excel : `getAllObjects` scan actif par defaut ✅ DONE (PR #169)
  - `allSheets` defaut `false` (etait `true`)

- **F-M5** — Excel : `findData` limite 200 → 2000 avec indication ✅ DONE (PR #169)
  - Quand tronque : retourne `{ matches, totalMatches, truncated: true }`

### MINEUR/AMELIORATION

- **F-L1** — Timeout Outlook 3s → 10s ✅ DONE (PR #169)
- **F-L2** — PowerPoint `insertContent` fallback non silencieux ✅ DONE (PR #169)
- **F-L3** — PowerPoint `hasNativeBullets` verifie tous les paragraphes ✅ DONE (PR #169)
- **F-L4** — PowerPoint `getAllSlidesOverview` truncation 100 → 2000 chars ✅ DONE (PR #169)
- **F-L5** — Excel `findData` validation regex avec try/catch ✅ DONE (PR #169)

---

## AXE 2 — REVUE DE CODE

### BLOQUANT/CRITIQUE

- **C-C1** — Upload route try/catch mammoth + code d'erreur structure ✅ DONE (PR #167)

### MAJEUR/IMPORTANT

- **C-M1** — DRY : factory generique `createOfficeTools<TName, TTemplate, TDef>()` ✅ DONE (PR #170)
  - `common.ts` ; les 4 factories dupliquees supprimees des fichiers host

- **C-M2** — DRY : `LANGUAGE_MATCH_INSTRUCTION` constant ✅ DONE (PR #170)
  - `constant.ts` ; interpole dans toutes les quick actions

- **C-M3** — Double logging dans `chat.js` ✅ DONE (PR #167)

- **C-M4** — DRY : `sanitizeHtml()` centralise ✅ DONE (PR #170)
  - Export dans `markdown.ts` ; `DOMPurify` importe en un seul endroit

- **C-M5** — Backend codes d'erreur structures ✅ DONE (PR #167)

### MINEUR/AMELIORATION

- **C-L1** — `sanitizeExecutionError` non exporte ✅ DONE (PR #170)
- **C-L2** — `OFFICE_ACTION_TIMEOUT_MS` / `OFFICE_BUSY_TIMEOUT_MESSAGE` non exportes ✅ DONE (PR #170)
- **C-L3** — `colToInt()` / `intToCol()` supprimes ✅ DONE (PR #170)
- **C-L4** — `DiffMatchPatch` centralise via `computeTextDiffStats()` dans `common.ts` ✅ DONE (PR #170)
- **C-L5** — Type aliases redondants supprimes (`WordToolDefinition` etc.) ✅ DONE (PR #170)

---

## AXE 3 — REVUE D'ARCHITECTURE

### BLOQUANT/CRITIQUE

- **A-C1** — Registre `ErrorCodes` centralise + `ERROR_CODE_MAP` frontend ✅ DONE (PR #167)

### MAJEUR/IMPORTANT

- **A-M1** — Gestion d'etat sans documentation ✅ DONE (PR #171)
  - `docs/STATE_MANAGEMENT.md` : 5 couches documentees avec diagrammes et regles

- **A-M2** — Credentials storage trop complexe ✅ DONE (PR #171)
  - Chiffrement extrait dans `credentialCrypto.ts` ; `credentialStorage.ts` = routage uniquement

### MINEUR/AMELIORATION

- **A-L1** — Backend sans abstraction retry ✅ DONE (PR #171)
  - `withRetry(fetchFn, maxAttempts=3)` dans `llmClient.js` : backoff 1s/2s/4s sur 429 et 5xx

- **A-L2** — Sandbox sans audit trail ✅ DONE (PR #171)
  - `sandboxedEval()` logge `[sandbox] host=X code=…` avant chaque execution

- **A-L3** — MAX_MESSAGES 200 → 1000 ✅ DONE (PR #171)

---

## AXE 4 — REVUE UX ET UI

### BLOQUANT/CRITIQUE

- **U-C1** — Streaming en boucle agent : **N/A** — `onStream` met bien a jour l'UI progressivement

### MAJEUR/IMPORTANT

- **U-M1** — Layout taskpane < 350px ✅ DONE (PR #172)
  - `ChatInput` : `flex-wrap` sur checkboxes ; `ChatHeader` : dropdown `max-w-[calc(100vw-1rem)]` ; `QuickActionsBar` : `flex-wrap`

- **U-M2** — Labels model tiers ✅ DONE (PR #172)
  - Cles i18n `modelTier.{standard,reasoning,image}` (EN: Quality/Deep Thinking, FR: Qualite/Reflexion profonde)

### MINEUR/AMELIORATION

- **U-L1** — Boutons insertion trop visibles ✅ DONE (PR #172)
  - `opacity-0` par defaut, `group-hover:opacity-100 focus-within:opacity-100` + transition 150ms

- **U-L2** — Pas de bouton Regenerer/Editer ✅ DONE (PR #172)
  - Regenerer (RotateCcw) sur dernier message assistant ; Modifier (Pencil) sur messages utilisateur

- **U-L3** — "Thought process" en anglais ✅ DONE
  - Deja correct : `t('thoughtProcess')` → FR: "Processus de reflexion"

---

## Bugs corriges en cours d'implementation

| Bug | Fichier | Correction |
|-----|---------|-----------|
| C-M2 : valeur constante auto-remplacee par replace_all | `constant.ts` | Restaure la chaine originale |
| Exports `getExcelToolDefinitions` etc. manquants | 4 fichiers tools | Aliases `getXxxToolDefinitions = getToolDefinitions` ajoutes |
| Cles i18n `modelTier.*` non resolues (format plat vs imbrique) | `en.json`, `fr.json` | Format plat → format imbrique (vue-i18n v9 traite les points comme chemin) |

---

## RECAPITULATIF PAR PR

| PR | Contenu |
|----|---------|
| #167 | F-C1, F-C2, C-C1, A-C1, C-M3, C-M5 (bloquants + premiers majeurs) |
| #169 | F-M1..F-M5, F-L1..F-L5 (tous les findings fonctionnels) |
| #170 | C-M1, C-M2, C-M4, C-L1..C-L5 (qualite de code) |
| #171 | A-M1, A-M2, A-L1, A-L2, A-L3 (architecture) |
| #172 | U-M1, U-M2, U-L1, U-L2, U-L3 (UX/UI) |

---

## Deferred Items (carries forward from v1–v5)

- **IC2** — Containers run as root (low priority, deployment simplicity)
- **IH2** — Private IP in build arg (users override at build time)
- **IH3** — DuckDNS domain in example (users replace with their own)
- **UM10** — PowerPoint HTML reconstruction (high complexity, low ROI)

---

## Verification Commands

```bash
cd frontend && npx tsc --noEmit   # 0 erreurs sur chaque branche verifiee
cd frontend && npm run build
cd backend && npm start
```

---

## Changelog

| Version | Date       | Changes                                                                        |
| ------- | ---------- | ------------------------------------------------------------------------------ |
| v6.1    | 2026-03-07 | Mise a jour statut : 33/35 findings resolus (2 N/A). PRs #167-#172             |
| v6.0    | 2026-03-07 | Full 4-axis audit (35 findings). 6 resolus en PR #167                          |
| v5.0    | 2026-03-07 | PR #158 review, audit pipeline Word/PPT, 3 corrections ciblees                |
| v4.0    | 2026-03-03 | Audit complet, 50 problemes tous resolus                                       |
| v3.0    | 2026-02-28 | 162 problemes identifies, 131 resolus                                          |
| v2.0    | 2026-02-22 | 28 nouveaux problemes apres refactoring majeur                                 |
| v1.0    | 2026-02-15 | Audit initial, 38 problemes (tous resolus)                                     |
