# Plan d'Architecture — Système de Logging & Feedback

> **Statut :** Plan entièrement validé et implémenté avec succès (Toutes les 5 phases terminées).

---

## État des lieux

| Aspect                   | Actuel                                                                                     | Problème                                                                                                      |
| ------------------------ | ------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------- |
| **Backend logging**      | Custom `systemLog()` + `rotating-file-stream` (30 fichiers, 10MB) + `console.*` éparpillés | Pas de niveaux structurés, pas de contexte (user/host/tab), `/health` pollue, prompts/réponses LLM non loggés |
| **Frontend logging**     | `console.*` direct dans 65 fichiers, pas d'abstraction                                     | Aucun buffer récupérable, pas de contexte enrichi, erreurs perdues                                            |
| **Error handling tools** | `officeAction.ts` (timeout 10s) + try/catch ad hoc dans `useAgentLoop.ts`                  | Inconsistant, certains tools sans try/catch, erreurs pas remontées au logger                                  |
| **Sessions/Tabs**        | UUID IndexedDB via `useSessionDB`/`useSessionManager`                                      | Pas de lien logs ↔ session, pas de purge logs à la suppression                                                |
| **Settings UI**          | 5 onglets dans `SettingsPage.vue` (1175 lignes)                                            | Pas de mécanisme de feedback                                                                                  |

---

## Structure de données commune

```typescript
// Entrée de log (front & back)
interface LogEntry {
  timestamp: string // ISO 8601
  level: 'error' | 'warn' | 'info' | 'debug'
  source: 'frontend' | 'backend'
  traffic: 'user' | 'llm' | 'auto' | 'system' // Séparation du trafic
  sessionId?: string // UUID du tab/chat
  userId?: string // Email utilisateur
  host?: string // 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'
  reqId?: string // UUID de requête (backend)
  message: string
  data?: Record<string, unknown>
  error?: {
    name: string
    message: string
    stack?: string
  }
}

// Rapport de feedback (Phase 5)
interface FeedbackReport {
  id: string // UUID
  timestamp: string
  sessionId: string
  userId: string
  host: string
  userMessage: string // Texte libre de l'utilisateur
  chatHistory: DisplayMessage[] // Historique du chat courant
  frontendLogs: LogEntry[] // Dump du ring buffer
  appVersion: string
}
```

---

## Dépendances à ajouter/supprimer

| Côté         | Ajouter                                | Supprimer                        |
| ------------ | -------------------------------------- | -------------------------------- |
| **Backend**  | `winston`, `winston-daily-rotate-file` | `morgan`, `rotating-file-stream` |
| **Frontend** | (rien)                                 | (rien)                           |

---

## Ordre d'implémentation

```
Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5
                                              ↑
                                   Dépend de Phase 3 (ring buffer)
                                   Dépend de Phase 2 (route feedback)
```

---

## Phase 1 — Choix des librairies et structure des données

### Backend : Winston + `winston-daily-rotate-file`

Pourquoi Winston plutôt que Pino :

- **Pretty-print natif** en dev (coloré, lisible) vs JSON strict en prod
- Écosystème de transports riche (`winston-daily-rotate-file` = rotation + suppression automatique)
- API `logger.info()`, `logger.warn()`, `logger.error()` standard et familière
- `child()` loggers avec contexte hérité (user, reqId, host, sessionId)

### Frontend : `LogService` maison (aucune lib externe)

Raison : garder le bundle léger (Office Add-in). Singleton `LogService` qui :

- Wrappe `console.*` (monkey-patch pour capturer les logs existants sans refactor massif)
- Maintient un **ring buffer** de 500 entrées max par session — récupérable pour le feedback
- Enrichit chaque entrée avec `{ host, sessionId, userEmail, timestamp, level }`
- Expose `logService.error()`, `.warn()`, `.info()`, `.debug()`

---

## Phase 2 — Logger Backend

**Fichiers modifiés :** `backend/src/utils/logger.js`, `backend/src/server.js`, `backend/src/routes/chat.js`, `backend/src/routes/image.js`, `backend/src/routes/upload.js`, `backend/src/services/llmClient.js`

### Actions

1. **Remplacer `logger.js`** par une config Winston :
   - Transport console : pretty-print coloré en dev (`NODE_ENV !== 'production'`), JSON en prod
   - Transport fichier : `winston-daily-rotate-file` → `logs/kickoffice-%DATE%.log`, rotation quotidienne, **suppression automatique à 7 jours**, compression gzip
   - Niveaux : `error`, `warn`, `info`, `http`, `debug`

2. **Middleware de contexte** : extrait `reqId`, `userId` (depuis `X-User-Email`), `host` (nouveau header `X-Office-Host` envoyé par le frontend) et les injecte dans `res.locals` + un child logger par requête

3. **Remplacer Morgan** par un middleware Winston `http` level :
   - Log toutes les requêtes (method, url, status, duration, userId)
   - **Filtre `/health`** → log en `debug` uniquement (invisible en prod, ne pollue pas)
   - Tag `traffic: 'auto'` pour `/health`, `traffic: 'llm'` pour `/api/chat` et `/api/image`, `traffic: 'user'` pour le reste

4. **Logger les prompts/réponses LLM** dans `chat.js` :
   - Log le body complet de la requête (`messages`, `tools`, `model`) en `info` avec `traffic: 'llm'`
   - Pour le streaming : log les chunks accumulés à la fin du stream
   - Pour le sync : log la réponse complète

5. **Remplacer tous les `console.*`** du backend par des appels au logger Winston

6. **Supprimer** `rotating-file-stream` et `morgan` de `package.json`

---

## Phase 3 — Logger Frontend

**Nouveau fichier :** `frontend/src/utils/logger.ts`

**Fichiers modifiés :** `frontend/src/main.ts`, `frontend/src/api/backend.ts`

### Actions

1. **Créer `LogService`** (singleton) :
   - Ring buffer de 500 entrées max, indexé par `sessionId`
   - Méthodes : `error()`, `warn()`, `info()`, `debug()` — enrichissement automatique avec contexte
   - `getSessionLogs(sessionId)` → logs du tab courant (pour le feedback)
   - `clearSessionLogs(sessionId)` → purge à la suppression d'un tab
   - `getContext()` → lazy-load du host (via `detectOfficeHost()`), sessionId courant, userEmail

2. **Monkey-patch `console.error` et `console.warn`** dans `main.ts` :
   - Intercepte les appels existants sans casser le comportement natif (les 65 fichiers couverts automatiquement)
   - Route vers `logService.error()` / `logService.warn()`

3. **Améliorer le global error handler** dans `main.ts` :
   - `app.config.errorHandler` → `logService.error('Vue unhandled error', ...)`
   - `window.onerror` → erreurs JS non-Vue
   - `window.onunhandledrejection` → promesses rejetées

4. **Enrichir les appels backend** dans `backend.ts` :
   - Ajouter header `X-Office-Host` à chaque requête (corrélation front ↔ back)
   - Ajouter header `X-Session-Id`
   - Logger les erreurs réseau/timeout/LLM via `logService`

---

## Phase 4 — Audit et sécurisation des Tools, interactions utilisateur et appels LLM

**Fichiers modifiés :** `officeAction.ts`, `useAgentLoop.ts`, `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`, `backend.ts`

### Actions

1. **`officeAction.ts`** — Enrichir le wrapper :
   - Log `logService.warn()` sur timeout (actuellement juste un throw)
   - Log `logService.error()` sur toute erreur Office.js avec le nom de l'action
   - Ajouter un paramètre optionnel `actionName` pour le contexte

2. **`useAgentLoop.ts`** — Standardiser l'exécution des tools :
   - Le try/catch existant → ajouter `logService.error()` avec `{ toolName, toolArgs, sessionId, traffic: 'user' }`
   - Log `logService.info()` sur chaque tool call réussi avec durée d'exécution
   - Log `logService.info()` sur chaque message envoyé/reçu du LLM avec `traffic: 'llm'`

3. **Audit des fichiers tools** (word, excel, ppt, outlook) :
   - Identifier les `executeWord`/`executeExcel`/etc. sans try/catch robuste
   - Wrapper les cas manquants : catch → `logService.error()` + retour d'erreur JSON propre dans le chat

4. **Interactions utilisateur** :
   - Log `logService.info('user_message_sent', ...)` à l'envoi d'un message
   - Log `logService.info('session_created/deleted/switched', ...)` sur les actions de session

5. **Appels LLM côté frontend** dans `chatStream()` :
   - `logService.info('llm_request', { model, messageCount })` au début
   - `logService.info('llm_response_complete', { tokensUsed })` à la fin
   - `logService.error('llm_stream_error', { ... })` sur erreur

---

## Phase 5 — UI de Feedback et cycle de vie des logs par onglet

**Nouveaux fichiers :** `frontend/src/components/settings/FeedbackDialog.vue`, `backend/src/routes/feedback.js`

**Fichiers modifiés :** `SettingsPage.vue`, `useSessionManager.ts`, `server.js`

### Actions

1. **`FeedbackDialog.vue`** — Dialogue dans Settings :
   - Textarea pour le message utilisateur (obligatoire)
   - Affichage du tab courant (nom + nombre de messages)
   - Bouton "Envoyer le rapport"
   - Collecte automatique : `chatHistory`, `frontendLogs` (ring buffer du sessionId), `appVersion`, `host`, `userId`
   - POST vers `/api/feedback`

2. **`SettingsPage.vue`** — Ajouter dans l'onglet "General" :
   - Bouton "Signaler un problème" qui ouvre `FeedbackDialog`

3. **Backend `/api/feedback`** :
   - Route POST dans `feedback.js`
   - Validation du payload (message requis, taille max)
   - Stockage en fichier JSON dans `logs/feedback/` → `{timestamp}_{sessionId}.json`
   - Nettoyage des rapports de plus de 30 jours au démarrage du serveur

4. **Purge à la suppression d'un onglet** dans `useSessionManager.ts` → `deleteCurrentSession()` :
   - POST optionnel vers `/api/feedback/cleanup?sessionId=xxx` — backend supprime les rapports associés

---

## Résumé de l'Implémentation

Toutes les 5 phases du plan ont été implémentées, testées et validées.

**Réalisations principales :**

- **Backend** : Intégration de `winston` et `winston-daily-rotate-file` pour un logging structuré et persistant avec rotation des logs. `morgan` et `rotating-file-stream` ont été complètement supprimés. Un middleware de contexte (`reqId`, `userId`, `host`) a été ajouté.
- **Frontend** : Création d'un `LogService` via un buffer circulaire pour capturer tous les logs, avec monkey-patching transparent de `console.error` et `console.warn`.
- **API & Core** : Interception des erreurs réseau, ajout des headers de traçabilité (`X-Office-Host`, `X-Session-Id`), et sécurisation (try/catch) de tous les factories de tools (Word, Excel, PPT, Outlook) pour renvoyer des erreurs JSON propres au LLM.
- **Feedback UI** : Ajout d'une boîte de dialogue dans les paramètres permettant l'envoi de rapports de bugs/suggestions au backend, incluant automatiquement le dump du _ring buffer_ local.

**Corrections apportées durant l'audit post-implémentation :**

1. **CORS Backend** : Ajout de `X-Office-Host` et `X-Session-Id` dans les `allowedHeaders` pour éviter les échecs de requêtes preflight (OPTIONS).
2. **API Feedback** : Correction du _path_ de la route (suppression du `/feedback` redondant) et réécriture du `FeedbackDialog.vue` pour utiliser une fonction encapsulée `submitFeedback` via l'URL complète du backend.
3. **Typescript** : Correction d'une erreur asynchrone dans `FeedbackDialog.vue` liée à `logService.getContext()`, ainsi que **résolution complète de 9 erreurs TS pré-existantes** dans `HomePage.vue` (typage de `quickActions`, références d'éléments du DOM) et d'un effet de bord dans `useAgentLoop.ts`.

> Le build (`vue-tsc`) est désormais 100% sans erreurs.
