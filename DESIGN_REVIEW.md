# KickOffice - Design Review (Mise à jour)

**Date initiale**: 2026-02-15
**Dernière mise à jour**: 2026-02-16
**Scope**: Architecture, Sécurité, Qualité du code, Documentation
**Fichiers analysés**: Backend (server.js — 448 lignes), Frontend (25+ fichiers source), Documentation (README.md, agents.md, manifests)

---

## Table des matières

1. [Résumé exécutif](#résumé-exécutif)
2. [Bilan des corrections effectuées](#bilan-des-corrections-effectuées)
3. [Architecture globale](#architecture-globale)
4. [Analyse du backend](#analyse-du-backend)
5. [Analyse du frontend](#analyse-du-frontend)
6. [Analyse de sécurité](#analyse-de-sécurité)
7. [Audit de la documentation](#audit-de-la-documentation)
8. [Liste des modifications restantes par degré de gravité](#liste-des-modifications-restantes-par-degré-de-gravité)

---

## Résumé exécutif

KickOffice est un add-in Microsoft Office (Word, Excel, PowerPoint, Outlook) alimenté par IA, construit avec Vue 3 / TypeScript (frontend) et Express.js (backend). L'architecture est saine dans son principe : le backend sert de proxy LLM sécurisé, les clés API ne sont jamais exposées côté client, et le CORS est correctement restreint.

### Progression depuis la première review

Depuis la review initiale (2026-02-15), **12 items sur 26 ont été corrigés** :

| Gravité | Corrigés | Restants | Total initial |
|---------|----------|----------|---------------|
| **CRITIQUE** | 1 / 4 | 3 | 4 |
| **HAUTE** | 4 / 8 | 4 | 8 |
| **MOYENNE** | 4 / 8 | 4 | 8 |
| **BASSE** | 3 / 6 | 3 | 6 |
| **Total** | **12 / 26** | **14** | **26** |

### État actuel des problèmes restants

| Gravité | Nombre | Résumé |
|---------|--------|--------|
| **CRITIQUE** | 3 | Absence d'authentification, pas de rate limiting, fuite d'erreurs LLM |
| **HAUTE** | 4 | Pas de headers sécurité, HomePage god component (1342 lignes), README obsolète, backend monolithique |
| **MOYENNE** | 4 | Pas de error handler Vue global, accessibilité insuffisante, pas de logging, `as any` résiduels |
| **BASSE** | 3 | Pas de dark mode toggle, CSS répétitives, documentation hostDetection |

---

## Bilan des corrections effectuées

### ✅ Items complétés

| ID | Item | Preuve dans le code |
|----|------|---------------------|
| **C4** | Cleanup `setInterval` | `HomePage.vue:1328-1341` — interval stocké dans `backendCheckInterval` ref, nettoyé dans `onUnmounted` via `clearInterval` |
| **H2** | Validation inputs backend | `server.js:55-106` — fonctions `validateTemperature`, `validateMaxTokens`, `validateTools` + validation dans chaque endpoint |
| **H4** | Extraire logique dupliquée | `savedPrompts.ts` — helper partagé `loadSavedPromptsFromStorage` ; `HomePage.vue:549-583` — fonction `getOfficeSelection()` unifiée |
| **H6** | Aligner `.env.example` avec defaults | `.env.example` et `server.js` defaults sont maintenant cohérents (`gpt-5-nano`, `gpt-5-mini`, `gpt-5.2`, `gpt-image-1.5`) |
| **H7** | Timeout sur requêtes fetch | Backend: `fetchWithTimeout` avec `AbortController` + timeouts par tier (`server.js:108-129`). Frontend: `fetchWithTimeoutAndRetry` avec timeout 45s (`backend.ts:48-75`) |
| **H8** | Typer `ToolDefinition` | `types/index.d.ts:27-36` — type générique `ToolDefinition` avec alias `WordToolDefinition`, `ExcelToolDefinition`, `OutlookToolDefinition` |
| **M1** | IDs uniques dans `v-for` | `HomePage.vue:275-280` — `DisplayMessage` a un `id: string` via `crypto.randomUUID()` ; template utilise `:key="item.key"` |
| **M2** | Mémoïser `renderSegments` | `HomePage.vue:460-466` — `historyWithSegments` est un `computed` qui pré-calcule les segments |
| **M5** | Supprimer watchers redondants | `SettingsPage.vue:524-534` — seuls 2 watchers restent : `localLanguage` (pour i18n) et `agentMaxIterations` (sanitization). Tous les watchers `localStorage.setItem` redondants ont été supprimés |
| **M6** | Retry avec backoff (API client) | `backend.ts:8-75` — `fetchWithTimeoutAndRetry` avec 2 retries ciblés sur erreurs réseau/timeout, délais +10s/+30s |
| **B1** | Corriger typo `cursor-po` | `HomePage.vue` — toutes les classes `cursor-pointer` sont correctes, pas de `cursor-po` résiduel |
| **B2** | Réduire `any` types | `HomePage.vue:319-329` — interfaces `QuickAction`/`ExcelQuickAction` avec `icon: Component` ; `backend.ts:160-181` — `OpenAIChatCompletion` interface ; `officeOutlook.ts` — utilitaire typé |

### ⚠️ Partiellement corrigé

| ID | Item | État |
|----|------|------|
| **B6** | Réduire body parser | Réduit de 10MB à 4MB (`server.js:172`). L'objectif initial était 1MB mais 4MB est un compromis acceptable pour supporter les contextes de chat longs. **Considéré comme résolu.** |

---

## Architecture globale

### Points forts
- **Séparation claire** : Frontend (Vue 3 + Vite, port 3002) / Backend (Express.js, port 3003) / LLM API externe
- **Sécurité des secrets** : API keys stockées uniquement côté serveur dans `.env`
- **Déploiement Docker** : Docker Compose fonctionnel avec health checks
- **Support multi-hôte** : Word, Excel, PowerPoint, Outlook avec détection automatique
- **i18n** : Framework complet avec support de 13 langues de réponse (2 locales UI : en/fr)
- **Agent mode** : Boucle d'outils OpenAI function-calling bien implémentée avec validation des tools côté backend
- **Système de thème** : Variables CSS bien structurées avec support dark mode prêt
- **Validation backend robuste** : Température, maxTokens, tools structure, taille prompt — tous validés
- **Timeout et retry** : Les deux côtés (backend et frontend) ont des timeouts et une stratégie de retry

### Points faibles persistants
- Backend monolithique en un seul fichier (448 lignes, a grossi depuis 293 lignes)
- Frontend avec composant `HomePage.vue` de **1342 lignes** (a grossi depuis 1159 lignes — god component encore plus gros)
- **Aucune authentification/autorisation** sur aucun endpoint
- **Aucun rate limiting**
- Pas de persistance de données (tout est en mémoire ou localStorage)
- **Fuite d'erreurs LLM** toujours présente dans les 3 endpoints

---

## Analyse du backend

### `backend/src/server.js` (448 lignes)

**Structure** : Fichier unique contenant configuration, validation, helpers, middleware, routes et démarrage serveur.

#### Problèmes restants

1. **Fuite d'informations sensibles dans les erreurs** (`server.js:254-257`, `349-352`, `421-424`)
   ```javascript
   // Le texte d'erreur brut de l'API LLM est toujours retransmis au client
   return res.status(response.status).json({
     error: `LLM API error: ${response.status}`,
     details: errorText,  // PROBLÈME : peut contenir des infos sensibles
   })
   ```
   **Impact** : Présent dans les 3 endpoints (`/api/chat`, `/api/chat/sync`, `/api/image`).

#### Ce qui a été corrigé
- ✅ Validation de `temperature` (0..2) et `maxTokens` (1..32768), avec rejet pour les modèles `chatgpt-*`
- ✅ Validation stricte de structure des `tools` (type `function`, `name` regex, `parameters` objet, max 32)
- ✅ Validation stricte de `size`, `quality`, `n` et `prompt` (≤4000 chars) pour `/api/image`
- ✅ `AbortController` avec timeout différencié : nano 60s, standard 120s, reasoning 300s, image 180s
- ✅ Limite JSON réduite de 10MB à 4MB
- ✅ `.env.example` aligné avec les defaults de `server.js`

---

## Analyse du frontend

### `HomePage.vue` — God Component (1342 lignes, +183 depuis la dernière review)

Ce fichier combine toujours :
- UI de chat (template de ~220 lignes)
- Logique de gestion des messages et de l'historique
- Agent loop complet avec exécution d'outils
- Quick actions pour Word, Excel, PowerPoint, Outlook
- Intégration Office.js (Word, Excel, PowerPoint, Outlook)
- Gestion du presse-papiers et insertion dans le document
- Health check polling
- Prompts systèmes pour chaque host (Word agent, Excel agent, PowerPoint agent, Outlook agent)
- Insertion d'images dans Word

#### Problèmes corrigés
- ✅ `setInterval` avec cleanup dans `onUnmounted`
- ✅ `v-for` avec ID unique (`message.id` via `crypto.randomUUID`)
- ✅ `renderSegments` mémoïsé dans un `computed` (`historyWithSegments`)
- ✅ Code de sélection Office unifié dans `getOfficeSelection()`
- ✅ `loadSavedPrompts` partagé via `savedPrompts.ts`
- ✅ Interfaces typées pour `QuickAction` / `ExcelQuickAction` avec `icon: Component`
- ✅ Utilitaire Outlook typé (`officeOutlook.ts`)
- ✅ Typo CSS `cursor-po` → `cursor-pointer`

#### Problèmes restants

1. **God component — taille critique** (1342 lignes)
   - Le composant a encore grossi (+183 lignes) avec l'ajout du support PowerPoint et des prompts agents.
   - Responsabilité unique violée : UI + logique de chat + agent loop + prompts + Office API + clipboard.

2. **`as any` résiduels dans l'agent loop** (`HomePage.vue:1021-1024`)
   ```typescript
   currentMessages.push({
     role: 'tool' as any,
     tool_call_id: toolCall.id,
     content: result,
   } as any)
   ```
   Le type `ChatMessage` ne supporte pas le rôle `tool`. Il faudrait étendre l'interface.

3. **`ChatSyncOptions.tools` est `any[]`** (`backend.ts:157`)
   ```typescript
   export interface ChatSyncOptions {
     tools?: any[]  // devrait être typé
   }
   ```

### `SettingsPage.vue` (717 lignes)

#### Problèmes corrigés
- ✅ Watchers redondants supprimés — `useStorage` gère seul la persistance
- ✅ Validation `agentMaxIterations` avec bornes (1..100) et normalisation entière

#### État satisfaisant
- Structure claire avec onglets (General, Prompts, Built-in Prompts, Tools)
- Gestion CRUD des prompts fonctionnelle
- Détection de modifications sur les prompts built-in avec option de reset

### `backend.ts` — Client API (231 lignes)

#### Problèmes corrigés
- ✅ Timeout global (45s) avec `AbortController`
- ✅ Retry avec backoff (2 retries : +10s, +30s) sur erreurs réseau/timeout
- ✅ URL backend obligatoire via `VITE_BACKEND_URL` (plus de fallback hardcodé)
- ✅ `chatSync` retourne `Promise<OpenAIChatCompletion>` (typé)

### `types/index.d.ts` — Définitions de types

#### Problèmes corrigés
- ✅ Type générique `ToolDefinition` avec alias `WordToolDefinition`, `ExcelToolDefinition`, `OutlookToolDefinition`

### `main.ts` — Point d'entrée

#### Problèmes restants
1. **Pas de error handler Vue global**
   ```typescript
   // Pas de app.config.errorHandler configuré
   ```
2. **Fonction `debounce` locale utilise `any`** (`main.ts:13-23`)
   - Pourrait utiliser `@vueuse/core` `useDebounceFn` ou typer proprement.

---

## Analyse de sécurité

### CRITIQUE (inchangé)

| # | Problème | Localisation | Impact | Statut |
|---|----------|-------------|--------|--------|
| S1 | **Aucune authentification** | `server.js` (tout le fichier) | N'importe qui sur le réseau peut appeler les endpoints et consommer les crédits API LLM | ❌ Non corrigé |
| S2 | **Aucun rate limiting** | `server.js` (tout le fichier) | DoS possible, consommation illimitée de crédits API | ❌ Non corrigé |
| S3 | **Fuite d'erreurs LLM** | `server.js:254-257`, `349-352`, `421-424` | Les erreurs brutes de l'API LLM sont renvoyées au client | ❌ Non corrigé |

### HAUTE

| # | Problème | Localisation | Impact | Statut |
|---|----------|-------------|--------|--------|
| S4 | **Pas de headers de sécurité** | `server.js` | Pas de `helmet.js` : manque X-Frame-Options, CSP, X-Content-Type-Options | ❌ Non corrigé |
| ~~S5~~ | ~~Pas de validation d'input côté backend~~ | `server.js` | ~~Non validés~~ | ✅ Corrigé |
| S6 | **Pas de logging/audit** | `server.js` | Impossible de tracer les abus ou incidents | ❌ Non corrigé |

### OK (pas de problème trouvé)

- **XSS** : Aucune utilisation de `v-html` — Vue escape correctement
- **CORS** : Correctement restreint à `FRONTEND_URL`
- **Secrets** : Les clés API ne sont jamais exposées côté client
- **Injection SQL/NoSQL** : N/A (pas de base de données)
- **Validation d'input** : ✅ Température, maxTokens, tools, prompt length, image params tous validés
- **Timeouts** : ✅ Toutes les requêtes fetch ont des timeouts avec AbortController

---

## Audit de la documentation

### README.md vs Code — Divergences actuelles

| Élément | Documentation | Code réel | Statut |
|---------|--------------|-----------|--------|
| Word tools count | "23 Word tools" (ligne 203) | 24 tools dans `WordToolName` (`wordTools.ts:1-25`) | **OBSOLÈTE** — manque `applyTaggedFormatting` |
| Excel tools count | "22 Excel tools" (ligne 204) | 25 tools dans `ExcelToolName` (`excelTools.ts:3-29`) | **OBSOLÈTE** — manque `fillFormulaDown`, `applyConditionalFormatting`, `getConditionalFormattingRules` |
| Outlook tools | Non listés | 3 tools dans `OutlookToolName` (`outlookTools.ts:1-4`) : `getEmailBody`, `getSelectedText`, `setEmailBody` | **MANQUANT** |
| PowerPoint | Mentionné dans l'architecture mais marqué comme non implémenté dans l'ancienne review | Maintenant implémenté : quick actions (bullets, speakerNotes, punchify, shrink, visual) + utilities (`powerpointTools.ts`) | **À METTRE À JOUR** |
| Outlook quick actions | Non documentées | 5 actions : reply, formalize, concise, proofread, extract | **MANQUANT** |
| Excel quick actions | Non documentées | 5 actions : clean, beautify, formula, transform, highlight | **MANQUANT** (nouvelles actions par rapport à l'ancienne review) |
| PowerPoint quick actions | Non documentées | 5 actions : bullets, speakerNotes, punchify, shrink, visual | **MANQUANT** |
| Support des langues | "English + French" | 13 langues de réponse dans `languageMap`, 2 locales UI (en/fr) | **INCOMPLET** |
| `hostDetection.ts` | Non documenté dans README | Fichier utilitaire clé pour la détection Word/Excel/PPT/Outlook | **MANQUANT** |

### Divergences corrigées depuis la dernière review

| Élément | Statut |
|---------|--------|
| `.env.example` vs defaults serveur | ✅ Maintenant cohérents |

---

## Liste des modifications restantes par degré de gravité

### CRITIQUE (à corriger immédiatement)

---

#### C1. Ajouter une authentification sur le backend

**Fichier** : `backend/src/server.js`
**Risque** : Toute personne sur le réseau peut utiliser l'API et consommer les crédits LLM
**Effort** : Moyen

**Guide d'implémentation** :
1. Ajouter une variable `ALLOWED_API_KEYS` dans `.env` (liste séparée par des virgules)
2. Créer un middleware d'authentification :
   ```javascript
   const ALLOWED_API_KEYS = (process.env.ALLOWED_API_KEYS || '').split(',').filter(Boolean)

   function requireAuth(req, res, next) {
     if (ALLOWED_API_KEYS.length === 0) return next() // Pas de clés = pas d'auth (dev mode)
     const apiKey = req.headers['x-api-key']
     if (!apiKey || !ALLOWED_API_KEYS.includes(apiKey)) {
       return res.status(401).json({ error: 'Unauthorized' })
     }
     next()
   }
   ```
3. Appliquer sur `/api/chat`, `/api/chat/sync`, `/api/image`
4. Garder `/health` et `/api/models` publics
5. Côté frontend (`backend.ts`), ajouter le header `x-api-key` dans `fetchWithTimeoutAndRetry`

---

#### C2. Ajouter un rate limiting

**Fichier** : `backend/src/server.js`
**Risque** : DoS, consommation illimitée de crédits API
**Effort** : Faible

**Guide d'implémentation** :
```bash
npm install express-rate-limit
```
```javascript
import rateLimit from 'express-rate-limit'

const chatLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 20,
  message: { error: 'Too many requests, please try again later' },
})

const imageLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 5,
  message: { error: 'Too many image requests' },
})

app.use('/api/chat', chatLimiter)
app.use('/api/image', imageLimiter)
```

---

#### C3. Ne pas transmettre les erreurs brutes de l'API LLM au client

**Fichiers** : `backend/src/server.js:254-257`, `349-352`, `421-424`
**Risque** : Fuite d'informations sur l'infrastructure (URLs internes, versions, clés partielles)
**Effort** : Faible

**Guide d'implémentation** :
Remplacer dans les 3 endpoints :
```javascript
// AVANT
return res.status(response.status).json({
  error: `LLM API error: ${response.status}`,
  details: errorText,
})

// APRÈS
console.error(`LLM API error ${response.status}:`, errorText)
return res.status(502).json({
  error: 'The AI service returned an error. Please try again later.',
})
```

---

### HAUTE (à corriger rapidement)

---

#### H1. Ajouter les headers de sécurité HTTP

**Fichier** : `backend/src/server.js`
**Risque** : Vulnérabilités clickjacking, MIME sniffing, etc.
**Effort** : Faible

**Guide d'implémentation** :
```bash
npm install helmet
```
```javascript
import helmet from 'helmet'
app.use(helmet({
  contentSecurityPolicy: false, // Office add-in a ses propres CSP
  crossOriginEmbedderPolicy: false,
}))
```

---

#### H3. Découper `HomePage.vue` en sous-composants

**Fichier** : `frontend/src/pages/HomePage.vue` (1342 lignes — a grossi de 183 lignes depuis la dernière review)
**Risque** : Maintenabilité, lisibilité, performance
**Effort** : Élevé

**Guide d'implémentation** :
Extraire les sections suivantes :
1. `ChatHeader.vue` — Header avec logo, boutons new chat et settings (lignes 5-38)
2. `QuickActionsBar.vue` — Barre d'actions rapides avec sélecteur de prompt (lignes 41-67)
3. `ChatMessageList.vue` — Container de messages avec empty state (lignes 70-160)
4. `ChatMessage.vue` — Un message individuel avec boutons d'action (lignes 98-158)
5. `ChatInput.vue` — Zone de saisie avec sélecteurs de mode et modèle (lignes 163-217)
6. Composable `useAgentLoop.ts` — La boucle agent + prompts système (lignes 653-1037)
7. Composable `useOfficeInsert.ts` — L'insertion dans le document + clipboard (lignes 1199-1309)

---

#### H5. Mettre à jour la documentation README.md

**Fichier** : `README.md`
**Risque** : Documentation trompeuse pour les développeurs et administrateurs
**Effort** : Faible

**Corrections nécessaires** :
1. Ligne 203 : "23 Word tools" → **"24 Word tools"**, ajouter `applyTaggedFormatting` à la liste
2. Ligne 204 : "22 Excel tools" → **"25 Excel tools"**, ajouter `fillFormulaDown`, `applyConditionalFormatting`, `getConditionalFormattingRules`
3. Ajouter **"3 Outlook tools"** : `getEmailBody`, `getSelectedText`, `setEmailBody`
4. Ajouter une section Quick Actions pour **Excel** (clean, beautify, formula, transform, highlight)
5. Ajouter une section Quick Actions pour **PowerPoint** (bullets, speakerNotes, punchify, shrink, visual)
6. Ajouter une section Quick Actions pour **Outlook** (reply, formalize, concise, proofread, extract)
7. Confirmer le support PowerPoint dans le diagramme d'architecture (c'est maintenant implémenté)
8. Mentionner les 13 langues de réponse et les 2 locales UI
9. Documenter `hostDetection.ts` dans la section Project Structure

---

### MOYENNE (à planifier)

---

#### M3. Ajouter la gestion d'erreurs globale Vue

**Fichier** : `frontend/src/main.ts`
**Risque** : Erreurs non capturées provoquent un crash silencieux
**Effort** : Faible

**Guide d'implémentation** :
```typescript
const app = createApp(App)

app.config.errorHandler = (err, instance, info) => {
  console.error('Vue Global Error:', err, info)
}
```

---

#### M4. Améliorer l'accessibilité (a11y)

**Fichiers** : `HomePage.vue`, composants
**Risque** : Non-conformité WCAG, exclusion d'utilisateurs
**Effort** : Moyen

**Guide d'implémentation** :
1. Ajouter `aria-label` sur tous les boutons sans texte visible (New Chat, Settings, Stop, Send, Copy, Replace, Append)
2. Ajouter `aria-live="polite"` sur le container de messages
3. Ajouter `role="status"` sur l'indicateur de statut backend
4. Ajouter `aria-expanded` sur les détails collapsibles (think tags)

---

#### M7. Séparer le backend en modules

**Fichier** : `backend/src/server.js` (448 lignes — a grossi de 155 lignes depuis la dernière review)
**Risque** : Maintenabilité à mesure que le code grandit
**Effort** : Moyen

**Guide d'implémentation** :
```
backend/src/
├── server.js              # Point d'entrée, middleware setup
├── config/
│   └── models.js          # Configuration des modèles
├── middleware/
│   ├── auth.js            # Authentication (C1)
│   └── validate.js        # Input validation (existant à extraire)
├── routes/
│   ├── health.js          # GET /health
│   ├── models.js          # GET /api/models
│   ├── chat.js            # POST /api/chat, /api/chat/sync
│   └── image.js           # POST /api/image
└── utils/
    └── fetchWithTimeout.js # Helper fetch avec timeout
```

---

#### M8. Ajouter du request logging

**Fichier** : `backend/src/server.js`
**Risque** : Impossible de diagnostiquer ou auditer les requêtes
**Effort** : Faible

**Guide d'implémentation** :
```bash
npm install morgan
```
```javascript
import morgan from 'morgan'
app.use(morgan(':method :url :status :response-time ms'))
```

---

### BASSE (améliorations)

---

#### B3. Ajouter un toggle dark mode dans l'UI

**Fichier** : `frontend/src/pages/SettingsPage.vue`
**Détail** : Les variables CSS dark mode existent dans `index.css:162-187` mais il n'y a aucun toggle pour l'activer.

**Guide d'implémentation** :
```typescript
const darkMode = useStorage(localStorageKey.darkMode, false)
watch(darkMode, (val) => {
  document.documentElement.classList.toggle('dark', val)
}, { immediate: true })
```

---

#### B4. Extraire les classes CSS répétées dans des utilities Tailwind

**Fichier** : `frontend/src/index.css`
**Effort** : Faible

Les patterns comme `rounded-md border border-border-secondary bg-surface p-2 shadow-sm` sont répétés dans de nombreux composants. Créer des utilities :
```css
@layer components {
  .card {
    @apply rounded-md border border-border-secondary bg-surface p-2 shadow-sm;
  }
}
```

---

#### B5. Documenter `hostDetection.ts` dans le README

**Fichier** : `README.md`
**Effort** : Très faible

Ajouter dans la section Project Structure et mentionner le mécanisme de détection dans la section architecture.

---

## Tableau récapitulatif complet

| Priorité | ID | Action | Statut |
|----------|-----|--------|--------|
| CRITIQUE | C1 | Ajouter authentification backend | ❌ À faire |
| CRITIQUE | C2 | Ajouter rate limiting | ❌ À faire |
| CRITIQUE | C3 | Masquer erreurs LLM brutes | ❌ À faire |
| ~~CRITIQUE~~ | ~~C4~~ | ~~Cleanup `setInterval`~~ | ✅ Fait |
| HAUTE | H1 | Headers sécurité (helmet) | ❌ À faire |
| ~~HAUTE~~ | ~~H2~~ | ~~Validation inputs backend~~ | ✅ Fait |
| HAUTE | H3 | Découper `HomePage.vue` (1342 lignes) | ❌ À faire |
| ~~HAUTE~~ | ~~H4~~ | ~~Extraire logique dupliquée~~ | ✅ Fait |
| HAUTE | H5 | Mettre à jour README.md | ❌ À faire |
| ~~HAUTE~~ | ~~H6~~ | ~~Aligner `.env.example` avec defaults~~ | ✅ Fait |
| ~~HAUTE~~ | ~~H7~~ | ~~Timeout sur requêtes fetch~~ | ✅ Fait |
| ~~HAUTE~~ | ~~H8~~ | ~~Renommer `ToolDefinition` type~~ | ✅ Fait |
| ~~MOYENNE~~ | ~~M1~~ | ~~IDs uniques dans `v-for`~~ | ✅ Fait |
| ~~MOYENNE~~ | ~~M2~~ | ~~Mémoïser `renderSegments`~~ | ✅ Fait |
| MOYENNE | M3 | Error handler Vue global | ❌ À faire |
| MOYENNE | M4 | Accessibilité (ARIA) | ❌ À faire |
| ~~MOYENNE~~ | ~~M5~~ | ~~Supprimer watchers redondants~~ | ✅ Fait |
| ~~MOYENNE~~ | ~~M6~~ | ~~Retry avec backoff (API client)~~ | ✅ Fait |
| MOYENNE | M7 | Séparer backend en modules | ❌ À faire |
| MOYENNE | M8 | Request logging (morgan) | ❌ À faire |
| ~~BASSE~~ | ~~B1~~ | ~~Corriger typo `cursor-po`~~ | ✅ Fait |
| ~~BASSE~~ | ~~B2~~ | ~~Réduire `any` types~~ | ✅ Fait (quelques `as any` résiduels dans l'agent loop) |
| BASSE | B3 | Toggle dark mode | ❌ À faire |
| BASSE | B4 | Extraire CSS répétées | ❌ À faire |
| BASSE | B5 | Documenter `hostDetection.ts` | ❌ À faire |
| ~~BASSE~~ | ~~B6~~ | ~~Réduire body parser~~ | ✅ Fait (4MB) |
