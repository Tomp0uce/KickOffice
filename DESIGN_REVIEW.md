# DESIGN_REVIEW.md — Code Audit v7

**Date**: 2026-03-08

**Version**: 7.0
**Scope**: Revue complete — Architecture, Fonctionnalites Add-in, Gestion d'erreurs, UX/UI, Code mort, Generalisation, Qualite de code

---

## Etat de sante global

Le projet KickOffice est **mature et bien structure** avec une architecture frontend/backend claire, un sandbox SES, une infra de logging Winston + LogService, et 129 outils Office.js couvrant 4 hotes. Les points v6 sont tous resolus (33 DONE, 2 N/A). Cette v7 identifie de **nouveaux axes d'amelioration** sur des sujets non couverts precedemment.

| Severite | Total | OPEN | DONE | Note |
| -------- | ----- | ---- | ---- | ---- |
| CRITIQUE | 5 | 0 | 5 | AR-C1, AR-C2, ER-C1, ER-C2, QC-C1 resolus |
| MAJEUR | 11 | 11 | 0 | ER-M2 resolu (bonus) |
| MINEUR | 12 | 12 | 0 | |
| **Total** | **28** | **23** | **6** | |

---

## AXE 1 — ARCHITECTURE

### CRITIQUE

- **AR-C1** — Aucun test unitaire ni composant ✅ DONE
  - **Constat** : Coverage estimee < 5%. Seul `e2e/navigation.spec.ts` (65 lignes) existe avec des tests superficiels.
  - **Fix** : 3 suites Vitest creees dans `frontend/src/utils/__tests__/` — **76 tests, 76 passent** :
    - `common.test.ts` (24 tests) : generateVisualDiff, computeTextDiffStats, createOfficeTools
    - `tokenManager.test.ts` (13 tests) : prepareMessagesForContext, truncation budget, tool_calls integrity
    - `officeCodeValidator.test.ts` (39 tests) : toutes les regles de validation, formatValidationResult, quickValidate
  - `vitest.config.ts` : pattern exclude e2e corrige

- **AR-C2** — Pas de monitoring / error tracking en production ✅ DONE
  - **Constat** : Ni Sentry, ni LogRocket, ni equivalent. Les erreurs frontend etaient capturees en memoire uniquement.
  - **Fix** : Endpoint `POST /api/logs` cree (`backend/src/routes/logs.js`) : accepte jusqu'a 200 entrees, filtre sur warn/error, ecrit en `logs/frontend/`. Route enregistree dans `server.js` avec `logsLimiter` (20 req/min). `submitLogs()` ajoute dans `frontend/src/api/backend.ts`.

### MAJEUR

- **AR-M1** — `useAgentLoop.ts` : composable monolithique (1183 lignes)
  - **Constat** : Contient la boucle agent, le streaming, l'execution des tools, la detection de boucle, la gestion d'etat — tout dans un seul fichier.
  - **Action** : Extraire en sous-composables :
    - `useAgentStream.ts` — logique de streaming SSE
    - `useToolExecutor.ts` — execution et serialisation des outils
    - `useLoopDetection.ts` — detection des boucles infinies
  - **Priorite** : Haute (maintenabilite, testabilite)

- **AR-M2** — Pas de propagation du `reqId` backend vers le frontend
  - **Constat** : Le backend genere un `reqId` UUID par requete et le logge. Le frontend ne recoit pas ce `reqId` dans les headers de reponse, rendant la correlation front/back impossible.
  - **Action** : Ajouter le header `X-Request-Id` dans la reponse backend. Le stocker cote frontend dans le `LogService` pour chaque requete.

- **AR-M3** — Pas de health check frontend
  - **Constat** : Le backend a un `/health` endpoint. Le frontend detecte `backendOnline` mais de maniere ad-hoc (resultat d'un fetch models). Pas de polling periodique.
  - **Action** : Implementer un polling `/health` toutes les 30s avec indicateur visuel persistant.

### MINEUR

- **AR-L1** — `noUnusedLocals` et `noUnusedParameters` desactives dans tsconfig
  - **Constat** : `tsconfig.json` a `"noUnusedLocals": false, "noUnusedParameters": false`. Des imports et variables inutilisees peuvent s'accumuler silencieusement.
  - **Action** : Activer ces flags et corriger les warnings resultants.

- **AR-L2** — Pas de validation de la longueur du tableau `messages` cote backend
  - **Constat** : `validate.js` verifie la profondeur des tools (max 20) mais pas le nombre de messages. Un client pourrait envoyer des milliers de messages.
  - **Action** : Ajouter une limite (ex: 500 messages max) dans `validateChatRequest()`.

---

## AXE 2 — FONCTIONNALITES ADD-IN (Excel / Word / Outlook / PowerPoint)

### MAJEUR

- **FN-M1** — Pas de timeout differencie par type d'operation Office.js
  - **Constat** : `executeOfficeAction()` dans `officeAction.ts` utilise un timeout unique de 10s pour toutes les operations. Certaines operations (creation de graphique Excel, insertion d'image PPT) peuvent legitimement depasser 10s.
  - **Action** : Ajouter un parametre `timeoutMs` optionnel a `executeOfficeAction()` avec des valeurs par defaut par categorie :
    - Lecture : 5s
    - Ecriture simple : 10s
    - Operations lourdes (chart, image) : 20s

- **FN-M2** — Pas de mecanisme de retry sur les operations Office.js
  - **Constat** : Si Office est occupe (`Office app is busy`), l'operation echoue immediatement sans retry. L'utilisateur doit relancer manuellement.
  - **Action** : Implementer un retry avec backoff (1s, 2s) sur les erreurs `GeneralException` d'Office.js (max 2 tentatives).

- **FN-M3** — Pas de signal d'annulation (AbortSignal) pour les operations Office.js longues
  - **Constat** : Les operations Office.js ne supportent pas l'annulation. Si l'utilisateur change de contexte pendant une operation longue, elle continue en arriere-plan.
  - **Action** : Implementer un token d'annulation cooperatif verifie entre les etapes `sync()`.

### MINEUR

- **FN-L1** — Outlook : timeout 10s peut etre insuffisant pour les pieces jointes volumineuses
  - **Constat** : Le timeout global de 10s s'applique aussi aux operations Outlook sur les pieces jointes. Les emails avec PJ lourdes peuvent timeout.
  - **Action** : Augmenter le timeout pour les operations `getAttachmentsAsync` a 20s.

- **FN-L2** — Pas de tests E2E specifiques aux operations Office.js
  - **Constat** : Les tests E2E existants ne couvrent que la navigation. Aucun test ne valide l'integration avec Office.js (meme mockee).
  - **Action** : Creer des tests E2E avec des mocks Office.js pour valider les operations critiques (insert, select, eval_*).

---

## AXE 3 — GESTION D'ERREURS ET DEBUGGING

### CRITIQUE

- **ER-C1** — Logs frontend perdus au refresh ✅ DONE
  - **Constat** : Le `LogService` utilise un ring buffer en memoire (500 entrees). Au refresh, tous les logs sont perdus.
  - **Fix** : `useSessionDB.ts` : DB_VERSION bumpe a 2, store `logs` ajoute (index sur sessionId + timestamp). `appendLogEntry()`, `getLogsForSession()`, `pruneOldLogs()` exportes. `logger.ts` : chaque entree est persiste dans IndexedDB via `appendLogEntry(entry).catch(() => {})` (fire-and-forget non bloquant).

- **ER-C2** — Pas de correlation front/back pour le debugging ✅ DONE
  - **Constat** : Le frontend logge avec `sessionId` + `userId`. Le backend logge avec `reqId` + `userId`. Pas d'identifiant commun.
  - **Fix** : Backend `server.js` : `res.setHeader('X-Request-Id', res.locals.reqId)` ajoute dans le middleware de contexte. CORS mis a jour avec `exposedHeaders: ['X-Request-Id']`. Frontend `backend.ts` : `chatStream()`, `generateImage()`, `uploadFile()` capturent le `x-request-id` de la reponse et le loggent via `logService.info()`.

### MAJEUR

- **ER-M1** — JSON malformed silencieusement ignore dans le stream SSE
  - **Constat** : Dans `chatStream()` (`backend.ts:320-326`), les lignes JSON malformees sont silencieusement ignorees sauf si le message commence par "Stream error:". Cela peut masquer des erreurs reelles de l'API LLM.
  - **Action** : Logger les lignes JSON malformees au niveau `warn` dans le `logService` avec le contenu brut (tronque a 200 chars) pour faciliter le debugging.

- **ER-M2** — Feedback fire-and-forget : pas de garantie de persistance ✅ DONE (bonus)
  - **Constat** : `feedback.js` envoyait la reponse HTTP avant que `fs.promises.writeFile()` ne termine.
  - **Fix** : `await fs.promises.writeFile(...)` avant `res.json()`. Le `try/catch` du handler capture desormais les erreurs d'ecriture.

### MINEUR

- **ER-L1** — Pas de niveau de log configurable cote frontend
  - **Constat** : Le `LogService` logge tout (error, warn, info, debug). Pas de filtre configurable pour reduire le bruit en production.
  - **Action** : Ajouter un `logLevel` configurable (par defaut `warn` en prod, `debug` en dev).

---

## AXE 4 — UX ET UI

### MAJEUR

- **UX-M1** — Accessibilite insuffisante (score estime 66/100)
  - **Constat** :
    - Pas de focus trap dans les dialogues (FeedbackDialog, modals)
    - Certains boutons icones sans `aria-label` (boutons d'edition, suppression)
    - Pas de lien "skip to content"
    - `outline-hidden` utilise sur certains elements sans alternative visible
  - **Action** :
    - Ajouter des `aria-label` sur tous les boutons icones
    - Implementer un focus trap dans les dialogues
    - Ajouter un style `focus-visible` coherent (ring accent)
    - Ajouter un lien skip-to-content

- **UX-M2** — Pas de skeleton loading ni d'indicateurs de progression
  - **Constat** : Les operations asynchrones (chargement initial, upload fichier, fetch models) n'ont pas d'indicateur visuel de chargement. L'interface semble figee pendant les operations.
  - **Action** : Ajouter des skeleton loaders pour le chargement initial de la page Settings et un indicateur de progression pour l'upload de fichiers.

- **UX-M3** — Erreurs d'authentification affichees inline dans le chat
  - **Constat** : Quand les credentials sont invalides, l'erreur est affichee comme un message assistant normal dans le chat (`⚠️ Credentials required`). L'utilisateur peut ne pas comprendre qu'il doit aller dans les settings.
  - **Action** : Afficher un bandeau d'erreur persistant en haut de page avec un lien direct vers les settings, en plus du message dans le chat.

### MINEUR

- **UX-L1** — Onglets Settings trop nombreux pour un taskpane etroit
  - **Constat** : 5 onglets (Account, General, Prompts, Built-in Prompts, Tools) dans un espace de 320-400px. Les labels sont tronques sur petits ecrans.
  - **Action** : Utiliser des labels abreges ou des icones avec tooltip pour les onglets en dessous de 400px.

- **UX-L2** — Pas de raccourcis clavier documentes
  - **Constat** : `Enter` envoie un message, mais pas de documentation visible des raccourcis disponibles.
  - **Action** : Ajouter un tooltip ou un hint `Shift+Enter for new line` sous le champ de saisie.

---

## AXE 5 — CODE MORT

### MINEUR

- **DC-L1** — Scripts Python legacy inutilises
  - **Constat** : 4 scripts Python avec des chemins Windows hardcodes ne sont references nulle part :
    - `/export_types.py`
    - `/frontend/fix_casts.py`
    - `/frontend/fix_record.py`
    - `/frontend/fix_unknown.py`
  - **Action** : Supprimer ces 4 fichiers.

- **DC-L2** — Imports inutilises dans `wordTools.ts`
  - **Constat** : `previewDiffStats` et `hasComplexContent` sont importes depuis `wordDiffUtils` (ligne 5) mais jamais utilises.
  - **Action** : Supprimer ces imports.

- **DC-L3** — Dependance `type` inutilisee dans `frontend/package.json`
  - **Constat** : `"type": "^2.7.3"` est listee en devDependency mais n'est importee nulle part dans le code source.
  - **Action** : Supprimer cette dependance.

---

## AXE 6 — GENERALISATION DES OUTILS / FONCTIONS

### MAJEUR

- **GN-M1** — Normalisation des fins de ligne dupliquee 5 fois
  - **Constat** : Le pattern `.replace(/\r\n/g, '\n').replace(/\r/g, '\n')` est repete dans :
    - `wordApi.ts` (lignes 8-9)
    - `wordFormatter.ts` (ligne 55)
    - `useOfficeInsert.ts` (lignes 120-121)
    - `powerpointTools.ts` (lignes 159, 672)
  - **Action** : Extraire dans `common.ts` :
    ```typescript
    export function normalizeLineEndings(text: string): string {
      return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    }
    ```

- **GN-M2** — Logique d'insertion Word dupliquee
  - **Constat** : `wordApi.ts:insertResult()` et `WordFormatter.insertPlainResult()` contiennent le meme switch case (`replace`/`append`/`newLine`) avec ~32 lignes identiques.
  - **Action** : Fusionner dans une seule implementation dans `WordFormatter`, et faire de `wordApi.insertResult()` un simple appel vers `WordFormatter.insertPlainResult()`.

### MINEUR

- **GN-L1** — Deux instances MarkdownIt avec configs partiellement similaires
  - **Constat** : `markdown.ts` cree `officeMarkdownParser` (avec plugins) et `officeRichText.ts` cree `markdownParser` (sans plugins). La config de base (breaks, html, linkify, typographer) est dupliquee.
  - **Action** : Extraire une fonction `createBaseMarkdownParser()` dans `markdown.ts` et l'utiliser dans les deux fichiers.

---

## AXE 7 — QUALITE ET MAINTENABILITE DU CODE

### CRITIQUE

- **QC-C1** — `SettingsPage.vue` : 1045 lignes, 5 features dans un composant ✅ DONE
  - **Constat** : Gere les onglets Account, General, Prompts, BuiltinPrompts, Tools dans un seul fichier.
  - **Fix** : Decompose en 5 composants dans `frontend/src/components/settings/` :
    - `AccountTab.vue` — credentials LiteLLM, crypto warning, status badge
    - `GeneralTab.vue` — langue, dark mode, user info, agent max iter, backend status, version, feedback, models
    - `PromptsTab.vue` — CRUD prompts personnalises (etat local)
    - `BuiltinPromptsTab.vue` — edition prompts integres avec reset-to-default (etat local)
    - `ToolsTab.vue` — activation/desactivation des outils par hote (etat local)
  - `SettingsPage.vue` reduit a **143 lignes** (orchestration + navigation des onglets).

### MAJEUR

- **QC-M1** — Magic numbers dissemine dans le code
  - **Constat** : Exemples :
    - `120` (hauteur max textarea, `ChatInput.vue`)
    - `5` (loop detection window, `useAgentLoop.ts`)
    - `10 * 1024 * 1024` (taille max fichier, `ChatInput.vue`)
    - `3` (nombre max fichiers, `ChatInput.vue`)
    - `30000` (timeout stream read, `chat.js`)
    - `500` (taille ring buffer, `logger.ts`)
  - **Action** : Extraire dans un fichier `constants/limits.ts` (frontend) et `config/limits.js` (backend) avec des noms explicites.

- **QC-M2** — `HomePage.vue` (701 lignes) melange orchestration et logique metier
  - **Constat** : Gere la session, l'agent loop, les image actions, les quick actions, et l'UI dans un seul composant.
  - **Action** : Extraire la logique metier dans un composable `useHomePage.ts` (session management, quick actions). Le composant ne garde que le template et les bindings.

- **QC-M3** — Nommage booleens inconsistant
  - **Constat** : Mix de conventions :
    - `loading` (sans prefixe) vs `showDeleteConfirm` (avec prefixe show)
    - `imageLoading` vs `isImageLoading` (prefixe is parfois absent)
    - `draftFocusGlow` (pas clair que c'est un boolean)
  - **Action** : Standardiser avec prefixe `is` ou `has` pour les booleens : `isLoading`, `isImageLoading`, `isDraftFocusGlowActive`.

### MINEUR

- **QC-L1** — Tailles d'icones hardcodees
  - **Constat** : `:size="16"`, `:size="12"`, `:size="20"` etc. dissemines dans les composants sans echelle definie.
  - **Action** : Definir des constantes `ICON_SIZE_SM = 12`, `ICON_SIZE_MD = 16`, `ICON_SIZE_LG = 20`.

- **QC-L2** — Cache manquant pour `getGlobalHeaders()` dans `backend.ts`
  - **Constat** : `getGlobalHeaders()` est async et appele a chaque requete. Il relit les credentials et le contexte du logger a chaque fois.
  - **Action** : Cacher le resultat avec invalidation sur changement de credentials.

---

## DEFERRED ITEMS (reconduits des versions precedentes)

- **IC2** — Containers run as root (low priority, deployment simplicity)
- **IH2** — Private IP in build arg (users override at build time)
- **IH3** — DuckDNS domain in example (users replace with their own)
- **UM10** — PowerPoint HTML reconstruction (high complexity, low ROI)

---

## MATRICE DE PRIORITE

| # | ID | Axe | Severite | Effort | Impact |
|---|-----|------|----------|--------|--------|
| 1 | AR-C1 | Architecture | CRITIQUE | Eleve | Tests = filet de securite |
| 2 | AR-C2 | Architecture | CRITIQUE | Moyen | Visibilite prod |
| 3 | ER-C1 | Erreurs | CRITIQUE | Moyen | Logs persistants |
| 4 | QC-C1 | Qualite | CRITIQUE | Moyen | Maintenabilite |
| 5 | ER-C2 | Erreurs | CRITIQUE | Faible | Correlation debug |
| 6 | AR-M1 | Architecture | MAJEUR | Eleve | Testabilite agent |
| 7 | UX-M1 | UX/UI | MAJEUR | Moyen | Accessibilite |
| 8 | GN-M1 | Generalisation | MAJEUR | Faible | DRY |
| 9 | GN-M2 | Generalisation | MAJEUR | Faible | DRY |
| 10 | QC-M1 | Qualite | MAJEUR | Faible | Lisibilite |

---

## Verification Commands

```bash
cd frontend && npx tsc --noEmit   # 0 erreurs
cd frontend && npm run build      # Build OK
cd backend && npm start           # Server demarre
cd frontend && npx vitest --run   # Tests unitaires
cd frontend && npx playwright test # Tests E2E
```

---

## Changelog

| Version | Date       | Changes |
| ------- | ---------- | ------- |
| v7.0    | 2026-03-08 | Nouvelle revue 7 axes (26 findings). Architecture, erreurs, UX, code mort, generalisation, qualite |
| v6.1    | 2026-03-07 | Mise a jour statut : 33/35 findings resolus (2 N/A). PRs #167-#172 |
| v6.0    | 2026-03-07 | Full 4-axis audit (35 findings). 6 resolus en PR #167 |
| v5.0    | 2026-03-07 | PR #158 review, audit pipeline Word/PPT, 3 corrections ciblees |
| v4.0    | 2026-03-03 | Audit complet, 50 problemes tous resolus |
| v3.0    | 2026-02-28 | 162 problemes identifies, 131 resolus |
| v2.0    | 2026-02-22 | 28 nouveaux problemes apres refactoring majeur |
| v1.0    | 2026-02-15 | Audit initial, 38 problemes (tous resolus) |
