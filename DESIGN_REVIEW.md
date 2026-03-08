# DESIGN_REVIEW.md — Code Audit v7

**Date**: 2026-03-08
**Version**: 7.1
**Scope**: Revue complete — Architecture, Fonctionnalites Add-in, Gestion d'erreurs, UX/UI, Code mort, Generalisation, Qualite de code

---

## Etat de sante global

Le projet KickOffice est **mature et bien structure**. Les 5 critiques v7 sont tous resolus. 7 items DONE au total (5 critiques + 2 majeurs bonus).

**Verification :** `npx tsc --noEmit` → 0 erreur | `npx vitest run` → 76/76 tests ✅

| Severite | Total | OPEN | DONE |
| -------- | ----- | ---- | ---- |
| CRITIQUE | 5 | 0 | 5 |
| MAJEUR | 16 | 14 | 2 |
| MINEUR | 13 | 13 | 0 |
| **Total** | **34** | **27** | **7** |

---

## AXE 1 — ARCHITECTURE

### CRITIQUE

- **AR-C1** ✅ DONE — 3 suites Vitest creees (`common`, `tokenManager`, `officeCodeValidator`) — 76 tests passent. Pattern exclude e2e corrige dans `vitest.config.ts`.
- **AR-C2** ✅ DONE — `POST /api/logs` cree (`backend/src/routes/logs.js`) : filtre warn/error, max 200 entrees, ecriture dans `logs/frontend/`. `logsLimiter` (20 req/min) + `submitLogs()` frontend.

### MAJEUR

- **AR-M1** — `useAgentLoop.ts` : composable monolithique (1183 lignes)
  - **Constat** : Boucle agent, streaming, execution des tools, detection de boucle et gestion d'etat concentres dans un seul fichier.
  - **Action** : Extraire en sous-composables : `useAgentStream.ts`, `useToolExecutor.ts`, `useLoopDetection.ts`.

- **AR-M2** ✅ DONE — Header `X-Request-Id` expose par le backend (`server.js`), capture dans `chatStream()`, `generateImage()`, `uploadFile()` via `res.headers.get('x-request-id')`. CORS mis a jour avec `exposedHeaders`.

- **AR-M3** — Pas de health check periodique frontend
  - **Constat** : `backendOnline` detecte ad-hoc (fetch models). Pas de polling.
  - **Action** : Polling `/health` toutes les 30s avec indicateur visuel persistant.

### MINEUR

- **AR-L1** — `noUnusedLocals` / `noUnusedParameters` desactives dans `tsconfig.json`
  - **Action** : Activer ces flags et corriger les warnings resultants.

- **AR-L2** — Pas de limite sur le nombre de messages dans `validateChatRequest()`
  - **Action** : Ajouter une limite de 500 messages max.

---

## AXE 2 — FONCTIONNALITES ADD-IN (Excel / Word / Outlook / PowerPoint)

### MAJEUR

- **FN-M1** — Timeout unique de 10s dans `executeOfficeAction()` pour toutes les operations
  - **Action** : Ajouter un parametre `timeoutMs` optionnel avec des valeurs par categorie (lecture 5s, ecriture 10s, operations lourdes 20s).

- **FN-M2** — Pas de retry sur les operations Office.js
  - **Action** : Retry avec backoff (1s, 2s) sur les erreurs `GeneralException` (max 2 tentatives).

- **FN-M3** — Pas d'AbortSignal pour les operations Office.js longues
  - **Action** : Token d'annulation cooperatif verifie entre les etapes `sync()`.

### MINEUR

- **FN-L1** — Timeout Outlook 10s insuffisant pour pieces jointes volumineuses
  - **Action** : Augmenter a 20s pour les operations `getAttachmentsAsync`.

- **FN-L2** — Pas de tests E2E pour les operations Office.js
  - **Action** : Tests E2E avec mocks Office.js pour les operations critiques.

---

## AXE 3 — GESTION D'ERREURS ET DEBUGGING

### CRITIQUE

- **ER-C1** ✅ DONE — `useSessionDB.ts` : DB_VERSION 1→2, store `logs` ajoute (index sessionId + timestamp). `appendLogEntry()`, `getLogsForSession()`, `pruneOldLogs()` exportes. `logger.ts` : chaque entree persiste dans IndexedDB via `appendLogEntry(entry).catch(() => {})`.
- **ER-C2** ✅ DONE — Voir AR-M2. Backend renvoie `X-Request-Id`, frontend le capture et le logge dans `logService.info()`.

### MAJEUR

- **ER-M1** — JSON malformed ignore silencieusement dans le stream SSE
  - **Constat** : Dans `chatStream()`, les lignes JSON malformees sont skipees sans log.
  - **Action** : Logger au niveau `warn` avec le contenu brut (tronque a 200 chars).

- **ER-M2** ✅ DONE (bonus) — `feedback.js` : `await fs.promises.writeFile(...)` avant `res.json()`. Erreurs d'ecriture desormais capturees par le `try/catch` global.

### MINEUR

- **ER-L1** — Pas de `logLevel` configurable dans `LogService`
  - **Action** : Ajouter un filtre configurable (defaut `warn` en prod, `debug` en dev).

---

## AXE 4 — UX ET UI

### MAJEUR

- **UX-M1** — Accessibilite insuffisante (score estime 66/100)
  - **Constat** : Pas de focus trap dans les dialogues, boutons icones sans `aria-label`, pas de skip-to-content, `outline-hidden` sans alternative visible.
  - **Action** : `aria-label` sur tous les boutons icones, focus trap dans les dialogues, style `focus-visible` coherent, lien skip-to-content.

- **UX-M2** — Pas de skeleton loading ni indicateurs de progression
  - **Action** : Skeleton loaders pour Settings et indicateur de progression pour l'upload.

- **UX-M3** — Erreurs d'authentification affichees inline dans le chat
  - **Action** : Bandeau d'erreur persistant en haut de page avec lien direct vers les settings.

### MINEUR

- **UX-L1** — 5 onglets Settings trop nombreux pour un taskpane 320px
  - **Action** : Labels abreges ou icones avec tooltip sous 400px.

- **UX-L2** — Raccourcis clavier non documentes
  - **Action** : Hint `Shift+Enter for new line` sous le champ de saisie.

---

## AXE 5 — CODE MORT

### MINEUR

- **DC-L1** — 4 scripts Python legacy inutilises (`export_types.py`, `fix_casts.py`, `fix_record.py`, `fix_unknown.py`)
  - **Action** : Supprimer ces 4 fichiers.

- **DC-L2** — Imports inutilises dans `wordTools.ts` (`previewDiffStats`, `hasComplexContent`)
  - **Action** : Supprimer ces imports.

- **DC-L3** — Dependance `type` inutilisee dans `frontend/package.json`
  - **Action** : Supprimer `"type": "^2.7.3"` des devDependencies.

---

## AXE 6 — GENERALISATION DES OUTILS / FONCTIONS

### MAJEUR

- **GN-M1** — Normalisation des fins de ligne dupliquee 5 fois
  - **Constat** : `.replace(/\r\n/g, '\n').replace(/\r/g, '\n')` dans `wordApi.ts`, `wordFormatter.ts`, `useOfficeInsert.ts`, `powerpointTools.ts` (x2).
  - **Action** : Extraire `normalizeLineEndings(text: string): string` dans `common.ts`.

- **GN-M2** — Logique d'insertion Word dupliquee entre `wordApi.ts` et `WordFormatter`
  - **Constat** : Meme switch case (`replace`/`append`/`newLine`) dans deux endroits (~32 lignes).
  - **Action** : Consolider dans `WordFormatter.insertPlainResult()`, `wordApi.insertResult()` devient un wrapper.

### MINEUR

- **GN-L1** — Deux instances `MarkdownIt` avec configs de base dupliquees
  - **Action** : Extraire `createBaseMarkdownParser()` dans `markdown.ts`.

---

## AXE 7 — QUALITE ET MAINTENABILITE DU CODE

### CRITIQUE

- **QC-C1** ✅ DONE — `SettingsPage.vue` decompose : `AccountTab.vue`, `GeneralTab.vue`, `PromptsTab.vue`, `BuiltinPromptsTab.vue`, `ToolsTab.vue` dans `frontend/src/components/settings/`. `SettingsPage.vue` reduit de 1045 a **142 lignes**.

### MAJEUR

- **QC-M1** — Magic numbers dissemines dans le code
  - **Constat** : `120` (textarea), `5` (loop window), `10*1024*1024` (taille max upload), `30000` (timeout stream), `500` (ring buffer)...
  - **Action** : Centraliser dans `frontend/src/constants/limits.ts` et `backend/src/config/limits.js`.

- **QC-M2** — `HomePage.vue` (701 lignes) melange orchestration et logique metier
  - **Action** : Extraire la logique dans un composable `useHomePage.ts`.

- **QC-M3** — Nommage des booleens inconsistant (`loading` vs `showDeleteConfirm` vs `draftFocusGlow`)
  - **Action** : Standardiser avec prefixe `is` ou `has`.

### MINEUR

- **QC-L1** — Tailles d'icones hardcodees (`:size="16"`, `:size="12"`, `:size="20"`)
  - **Action** : Constantes `ICON_SIZE_SM/MD/LG`.

- **QC-L2** — `getGlobalHeaders()` dans `backend.ts` relit les credentials a chaque requete
  - **Action** : Cache avec invalidation sur changement de credentials.

---

## DEFERRED ITEMS (reconduits des versions precedentes)

- **IC2** — Containers run as root (low priority, deployment simplicity)
- **IH2** — Private IP in build arg (users override at build time)
- **IH3** — DuckDNS domain in example (users replace with their own)
- **UM10** — PowerPoint HTML reconstruction (high complexity, low ROI)

---

## MATRICE DE PRIORITE (items OPEN)

| # | ID | Axe | Severite | Effort | Impact |
|---|-----|------|----------|--------|--------|
| 1 | AR-M1 | Architecture | MAJEUR | Eleve | Testabilite agent loop |
| 2 | UX-M1 | UX/UI | MAJEUR | Moyen | Accessibilite |
| 3 | GN-M1 | Generalisation | MAJEUR | Faible | DRY (5 instances) |
| 4 | GN-M2 | Generalisation | MAJEUR | Faible | DRY (logique Word) |
| 5 | QC-M1 | Qualite | MAJEUR | Faible | Lisibilite / constantes |
| 6 | FN-M1 | Fonctionnalites | MAJEUR | Moyen | Robustesse Office.js |
| 7 | FN-M2 | Fonctionnalites | MAJEUR | Moyen | Resilience Office busy |
| 8 | QC-M2 | Qualite | MAJEUR | Moyen | HomePage maintenabilite |
| 9 | ER-M1 | Erreurs | MAJEUR | Faible | Debugging SSE |
| 10 | AR-M3 | Architecture | MAJEUR | Faible | Health check continu |

---

## Verification Commands

```bash
cd frontend && npx tsc --noEmit        # 0 erreurs ✅
cd frontend && npx vitest run          # 76/76 tests ✅
cd frontend && npm run build           # Build OK
cd backend && npm start                # Server demarre
cd frontend && npx playwright test     # Tests E2E navigation
```

---

## Changelog

| Version | Date       | Changes |
| ------- | ---------- | ------- |
| v7.1    | 2026-03-08 | Resolution des 5 critiques (AR-C1, AR-C2, ER-C1, ER-C2, QC-C1) + 2 majeurs bonus (AR-M2, ER-M2). PR #176 |
| v7.0    | 2026-03-08 | Nouvelle revue 7 axes (34 findings). Architecture, erreurs, UX, code mort, generalisation, qualite |
| v6.1    | 2026-03-07 | Mise a jour statut : 33/35 findings resolus (2 N/A). PRs #167-#172 |
| v6.0    | 2026-03-07 | Full 4-axis audit (35 findings). 6 resolus en PR #167 |
| v5.0    | 2026-03-07 | PR #158 review, audit pipeline Word/PPT, 3 corrections ciblees |
| v4.0    | 2026-03-03 | Audit complet, 50 problemes tous resolus |
| v3.0    | 2026-02-28 | 162 problemes identifies, 131 resolus |
| v2.0    | 2026-02-22 | 28 nouveaux problemes apres refactoring majeur |
| v1.0    | 2026-02-15 | Audit initial, 38 problemes (tous resolus) |
