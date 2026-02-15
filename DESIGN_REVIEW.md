# KickOffice - Design Review Complète

**Date**: 2026-02-15
**Scope**: Architecture, Sécurité, Qualité du code, Documentation
**Fichiers analysés**: Backend (server.js), Frontend (25 fichiers source), Documentation (README.md, agents.md, manifest.xml)

---

## Table des matières

1. [Résumé exécutif](#résumé-exécutif)
2. [Architecture globale](#architecture-globale)
3. [Analyse du backend](#analyse-du-backend)
4. [Analyse du frontend](#analyse-du-frontend)
5. [Analyse de sécurité](#analyse-de-sécurité)
6. [Audit de la documentation](#audit-de-la-documentation)
7. [Liste des modifications par degré de gravité](#liste-des-modifications-par-degré-de-gravité)

---

## Résumé exécutif

KickOffice est un add-in Microsoft Office (Word, Excel, Outlook) alimenté par IA, construit avec Vue 3 / TypeScript (frontend) et Express.js (backend). L'architecture est saine dans son principe : le backend sert de proxy LLM sécurisé, les clés API ne sont jamais exposées côté client, et le CORS est correctement restreint.

Cependant, l'audit révèle **plusieurs problèmes critiques et importants** qui doivent être traités avant toute mise en production :

| Gravité | Nombre | Résumé |
|---------|--------|--------|
| **CRITIQUE** | 4 | Absence d'authentification, pas de rate limiting, fuite d'erreurs LLM, `setInterval` sans cleanup |
| **HAUTE** | 8 | HomePage trop volumineuse, pas de validation d'input, pas de headers sécurité, duplication de code, documentation obsolète |
| **MOYENNE** | 8 | Pas de retry API, `v-for` avec index comme key, accessibilité insuffisante, pas de gestion d'erreurs globale |
| **BASSE** | 6 | Typage `any` résiduel, CSS verbose, pas de dark mode toggle, `cursor-po` typo |

---

## Architecture globale

### Points forts
- **Séparation claire** : Frontend (Vue 3 + Vite, port 3002) / Backend (Express.js, port 3003) / LLM API externe
- **Sécurité des secrets** : API keys stockées uniquement côté serveur dans `.env`
- **Déploiement Docker** : Docker Compose fonctionnel avec health checks
- **Support multi-hôte** : Word, Excel, Outlook avec détection automatique
- **i18n** : Framework complet avec support de 13 langues (même si 2 sont complètes)
- **Agent mode** : Boucle d'outils OpenAI function-calling bien implémentée
- **Système de thème** : Variables CSS bien structurées avec support dark mode prêt

### Points faibles
- Backend monolithique en un seul fichier (293 lignes, pas de séparation de responsabilités)
- Frontend avec composant HomePage.vue de 1159 lignes (god component)
- Aucune authentification/autorisation sur aucun endpoint
- Aucun rate limiting
- Pas de persistance de données (tout est en mémoire ou localStorage)

---

## Analyse du backend

### `backend/src/server.js` (293 lignes)

**Structure** : Fichier unique contenant configuration, middleware, routes et démarrage serveur.

#### Problèmes identifiés

1. **Fuite d'informations sensibles dans les erreurs** (`server.js:150-152`, `220-222`, `270-272`)
   ```javascript
   // Le texte d'erreur brut de l'API LLM est retransmis au client
   return res.status(response.status).json({
     error: `LLM API error: ${response.status}`,
     details: errorText,  // PROBLÈME : peut contenir des infos sensibles
   })
   ```

2. ~~**Pas de validation des paramètres** (`server.js:109`, `183`, `236`)~~ ✅
   - ✅ Validation de `temperature` (0..2) et `maxTokens` (1..32768), avec rejet explicite pour les modèles `chatgpt-*` qui ne supportent pas ces paramètres.
   - ✅ Limite de longueur sur `prompt` (`<= 4000` caractères).
   - ✅ Validation stricte de `size`, `quality` et `n` pour `/api/image`.

3. ~~**`buildChatBody` accepte n'importe quel tools array** (`server.js:69-72`)~~ ✅
   - ✅ Ajout d'une validation de structure des `tools` (type `function`, `name`, `parameters`, etc.) et sanitization avant envoi au provider.

4. ~~**Pas de timeout sur les requêtes fetch vers l'API LLM** (`server.js:137`, `208`, `252`)~~ ✅
   - ✅ Ajout d'un `AbortController` avec timeout différencié : `nano` 60s, `standard` 120s, `reasoning` 300s, image 120s.

5. ~~**Express 5 avec body parser 10MB** (`server.js:83`)~~ ✅
   - ✅ Réduction de la limite JSON de `10mb` à `4mb`.

---

## Analyse du frontend

### `HomePage.vue` - God Component (1159 lignes)

Ce fichier combine :
- UI de chat (template de 251 lignes)
- Logique de gestion des messages
- Agent loop
- Quick actions
- Intégration Office.js (Word, Excel, Outlook)
- Gestion du presse-papiers
- Health check polling

#### Problèmes spécifiques

1. **`setInterval` sans cleanup** (`HomePage.vue:1157`)
   ```typescript
   setInterval(checkBackend, 30000)  // Jamais nettoyé - fuite mémoire
   ```

2. **`v-for` avec index comme key** (`HomePage.vue:99-100`)
   ```vue
   <div v-for="(msg, index) in history" :key="index">
   ```
   L'index change quand les messages sont ajoutés/supprimés, provoquant des re-renders inutiles.

3. **`renderSegments()` appelé à chaque render** (`HomePage.vue:110`)
   ```vue
   <template v-for="(segment, idx) in renderSegments(msg.content)" :key="idx">
   ```
   Fonction appelée dans le template sans mémoïsation - recalculée à chaque mise à jour du DOM.

4. **Duplication du code de sélection Office** (`HomePage.vue:706-728` et `922-950`)
   Le code pour récupérer le texte sélectionné (Word/Excel/Outlook) est dupliqué dans `sendMessage()` et `applyQuickAction()`.

5. **Duplication de `loadSavedPrompts`** (`HomePage.vue:448-457` vs `SettingsPage.vue:552-566`)
   Logique identique dans deux composants différents.

6. **Watchers manuels au lieu de `useStorage`** (`SettingsPage.vue:522-549`)
   7 watchers qui font `localStorage.setItem(...)` manuellement alors que `useStorage` gère déjà la persistance.

7. **Type `any` pour les icônes** (`HomePage.vue:357`, `369`, `381`)
   ```typescript
   const wordQuickActions: { key: string; label: string; icon: any }[] = [...]
   ```

8. **`chatSync` retourne `Promise<any>`** (`backend.ts:81`)
   ```typescript
   export async function chatSync(options: ChatSyncOptions): Promise<any> {
   ```
   Le type de retour devrait être typé avec l'interface OpenAI ChatCompletion.

9. **Cast dangereux avec `as any`** (`HomePage.vue:642`, `649`, `650`, `675`)
   ```typescript
   const mailbox = (window as any).Office?.context?.mailbox
   ```
   Répété 4 fois - devrait être extrait dans un utilitaire typé.

10. **Typo CSS** (`HomePage.vue:167`, `175`, `183`)
    ```vue
    class="cursor-po flex ..."  <!-- devrait être cursor-pointer -->
    ```

### `SettingsPage.vue` (737 lignes)

1. **ExcelToolDefinitions typées avec `WordToolDefinition`** (`excelTools.ts:35`)
   ```typescript
   const excelToolDefinitions: Record<ExcelToolName, WordToolDefinition> = {
   ```
   Le type `WordToolDefinition` est réutilisé pour Excel - devrait être un type générique `ToolDefinition`.

2. **Pas de validation pour `agentMaxIterations`** (`SettingsPage.vue:124-129`)
   L'utilisateur peut entrer un nombre négatif ou extrêmement grand.

### `backend.ts` - Client API

1. **Pas de timeout sur les requêtes fetch** (`backend.ts:33`, `84`, `105`)
   Si le backend ne répond pas, pas de mécanisme de timeout.

2. **Pas de retry avec backoff** - En cas d'erreur réseau transitoire, la requête échoue immédiatement.

3. **URL par défaut hardcodée** (`backend.ts:1`)
   ```typescript
   const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || 'http://192.168.50.10:3003'
   ```
   L'IP privée est en dur dans le code source.

---

## Analyse de sécurité

### CRITIQUE

| # | Problème | Localisation | Impact |
|---|----------|-------------|--------|
| S1 | **Aucune authentification** | `server.js` (tout le fichier) | N'importe qui sur le réseau peut appeler les endpoints et consommer les crédits API LLM |
| S2 | **Aucun rate limiting** | `server.js` (tout le fichier) | DoS possible, consommation illimitée de crédits API |
| S3 | **Fuite d'erreurs LLM** | `server.js:150-152`, `220-222`, `270-272` | Les erreurs brutes de l'API LLM (potentiellement avec des détails d'infra) sont renvoyées au client |

### HAUTE

| # | Problème | Localisation | Impact |
|---|----------|-------------|--------|
| S4 | **Pas de headers de sécurité** | `server.js` | Pas de `helmet.js` : manque X-Frame-Options, CSP, X-Content-Type-Options, etc. |
| S5 | **Pas de validation d'input côté backend** | `server.js:109`, `183`, `236` | Temperature, maxTokens, prompt length non validés |
| S6 | **Pas de logging/audit** | `server.js` | Impossible de tracer les abus ou incidents |

### OK (pas de problème trouvé)

- **XSS** : Aucune utilisation de `v-html` ou `dangerouslySetInnerHTML` - Vue escape correctement
- **CORS** : Correctement restreint à `FRONTEND_URL`
- **Secrets** : Les clés API ne sont jamais exposées côté client
- **Injection SQL/NoSQL** : N/A (pas de base de données)

---

## Audit de la documentation

### README.md vs Code - Divergences trouvées

| Élément | Documentation | Code réel | Statut |
|---------|--------------|-----------|--------|
| Word tools count | "23 Word tools" (ligne 175) | 24 tools (`WordToolName` dans `wordTools.ts:1-25`) | **OBSOLÈTE** - manque `applyTaggedFormatting` |
| Excel tools count | "22 Excel tools" (ligne 176) | 24 tools (`ExcelToolName` dans `excelTools.ts:3-27`) | **OBSOLÈTE** - manque `applyConditionalFormatting`, `getConditionalFormattingRules` |
| Variables d'env backend | 8 variables documentées (lignes 266-275) | 18+ variables dans le code (`server.js:13-39`) | **INCOMPLET** - manque `MODEL_*_LABEL`, `MODEL_*_MAX_TOKENS`, `MODEL_*_TEMPERATURE` (10 variables) |
| Excel quick actions | Non documentées | 5 actions dans le code (`constant.ts`) : analyze, chart, formula, format, explain | **MANQUANT** |
| Support des langues | "English + French" (lignes 207-208) | 13 langues dans `languageMap` (`constant.ts`) | **INCOMPLET** |
| Locale du manifest | Non mentionnée | `DefaultLocale: fr-FR` (`manifest.xml:12`) | **MANQUANT** |
| `.env.example` modèles | `gpt-5-nano`, `gpt-5-mini`, etc. | Defaults serveur : `gpt-4.1-nano`, `gpt-4.1`, `o3` | **INCOHÉRENT** |
| Architecture diagram | Mentionne "PowerPoint" | PowerPoint non implémenté | **TROMPEUR** |
| `hostDetection.ts` | Non documenté dans README | Fichier utilitaire clé | **MANQUANT** |

### Documentation correcte
- Endpoints API (`/api/chat`, `/api/chat/sync`, `/api/image`, `/api/models`, `/health`)
- Instructions de déploiement Docker
- Quick actions Word et Outlook
- Docker Compose configuration
- Security model (API keys server-side, CORS)
- Feature implementation checklist (sauf les divergences ci-dessus)

---

## Liste des modifications par degré de gravité

### CRITIQUE (à corriger immédiatement)

---

#### C1. Ajouter une authentification sur le backend

**Fichier** : `backend/src/server.js`
**Risque** : Toute personne sur le réseau peut utiliser l'API et consommer les crédits LLM
**Effort** : Moyen

**Guide d'implémentation** :
1. Installer `jsonwebtoken` ou utiliser un système API key simple :
   ```bash
   npm install jsonwebtoken
   ```
2. Ajouter une variable `AUTH_SECRET` ou `ALLOWED_API_KEYS` dans `.env`
3. Créer un middleware d'authentification :
   ```javascript
   // middleware/auth.js
   function requireAuth(req, res, next) {
     const apiKey = req.headers['x-api-key']
     if (!apiKey || !ALLOWED_API_KEYS.includes(apiKey)) {
       return res.status(401).json({ error: 'Unauthorized' })
     }
     next()
   }
   ```
4. Appliquer le middleware sur `/api/chat`, `/api/chat/sync`, `/api/image`
5. Garder `/health` et `/api/models` publics
6. Côté frontend, ajouter le header `x-api-key` dans `backend.ts`

---

#### C2. Ajouter un rate limiting

**Fichier** : `backend/src/server.js`
**Risque** : DoS, consommation illimitée de crédits API
**Effort** : Faible

**Guide d'implémentation** :
1. Installer `express-rate-limit` :
   ```bash
   npm install express-rate-limit
   ```
2. Configurer par endpoint :
   ```javascript
   import rateLimit from 'express-rate-limit'

   const chatLimiter = rateLimit({
     windowMs: 60 * 1000, // 1 minute
     max: 20, // 20 requêtes/minute pour chat
     message: { error: 'Too many requests, please try again later' },
   })

   const imageLimiter = rateLimit({
     windowMs: 60 * 1000,
     max: 5, // 5 images/minute
     message: { error: 'Too many image requests' },
   })

   app.use('/api/chat', chatLimiter)
   app.use('/api/image', imageLimiter)
   ```

---

#### C3. Ne pas transmettre les erreurs brutes de l'API LLM au client

**Fichiers** : `backend/src/server.js:149-152`, `219-222`, `269-272`
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

#### C4. Nettoyer le `setInterval` dans `HomePage.vue`

**Fichier** : `frontend/src/pages/HomePage.vue:1152-1158`
**Risque** : Fuite mémoire - le timer continue même après la destruction du composant
**Effort** : Très faible

**Guide d'implémentation** :
```typescript
// AVANT
onBeforeMount(() => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as insertTypes) || 'replace'
  loadSavedPrompts()
  checkBackend()
  setInterval(checkBackend, 30000)
})

// APRÈS
import { onBeforeMount, onBeforeUnmount, ref } from 'vue'

let healthCheckInterval: ReturnType<typeof setInterval> | null = null

onBeforeMount(() => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as insertTypes) || 'replace'
  loadSavedPrompts()
  checkBackend()
  healthCheckInterval = setInterval(checkBackend, 30000)
})

onBeforeUnmount(() => {
  if (healthCheckInterval) {
    clearInterval(healthCheckInterval)
    healthCheckInterval = null
  }
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

#### H2. Valider les inputs côté backend

**Fichier** : `backend/src/server.js`
**Risque** : Abus de l'API, requêtes malformées
**Effort** : Moyen

**Guide d'implémentation** :
Ajouter avant le proxy dans chaque endpoint :
```javascript
// Pour /api/chat et /api/chat/sync
const temperature = req.body.temperature
if (temperature !== undefined && (typeof temperature !== 'number' || temperature < 0 || temperature > 2)) {
  return res.status(400).json({ error: 'temperature must be a number between 0 and 2' })
}

const maxTokens = req.body.maxTokens
if (maxTokens !== undefined && (!Number.isInteger(maxTokens) || maxTokens < 1 || maxTokens > 128000)) {
  return res.status(400).json({ error: 'maxTokens must be an integer between 1 and 128000' })
}

// Pour /api/image
if (prompt.length > 4000) {
  return res.status(400).json({ error: 'prompt must be 4000 characters or less' })
}

const validSizes = ['256x256', '512x512', '1024x1024', '1024x1792', '1792x1024']
if (!validSizes.includes(size)) {
  return res.status(400).json({ error: `Invalid size. Must be one of: ${validSizes.join(', ')}` })
}

if (n < 1 || n > 4) {
  return res.status(400).json({ error: 'n must be between 1 and 4' })
}
```

---

#### H3. Découper `HomePage.vue` en sous-composants

**Fichier** : `frontend/src/pages/HomePage.vue` (1159 lignes)
**Risque** : Maintenabilité, lisibilité, performance
**Effort** : Élevé

**Guide d'implémentation** :
Extraire les sections suivantes en composants séparés :
1. `ChatHeader.vue` - Header avec logo, boutons new chat et settings (lignes 5-38)
2. `QuickActionsBar.vue` - Barre d'actions rapides avec sélecteur de prompt (lignes 41-67)
3. `ChatMessageList.vue` - Container de messages avec empty state (lignes 70-160)
4. `ChatMessage.vue` - Un message individuel avec boutons d'action (lignes 98-158)
5. `ChatInput.vue` - Zone de saisie avec sélecteurs de mode et modèle (lignes 163-248)
6. Extraire un composable `useChat.ts` pour la logique de chat (sendMessage, processChat, agent loop)
7. Extraire un composable `useOfficeSelection.ts` pour le code de sélection Office dupliqué

---

#### H4. Extraire la logique dupliquée dans des composables

**Fichiers** : `HomePage.vue`, `SettingsPage.vue`, `constant.ts`
**Risque** : Bugs de divergence, maintenabilité
**Effort** : Moyen

**Guide d'implémentation** :
1. **Composable `usePrompts.ts`** :
   ```typescript
   export function usePrompts() {
     const savedPrompts = ref<SavedPrompt[]>([])

     function loadPrompts() { /* logique unifiée */ }
     function savePrompts() { /* ... */ }
     function addPrompt() { /* ... */ }
     function deletePrompt(id: string) { /* ... */ }

     return { savedPrompts, loadPrompts, savePrompts, addPrompt, deletePrompt }
   }
   ```

2. **Composable `useOfficeContext.ts`** :
   ```typescript
   export function useOfficeContext() {
     function getSelectedText(): Promise<string> { /* Word/Excel/Outlook */ }
     function getMailBody(): Promise<string> { /* Outlook */ }
     function getMailSelectedText(): Promise<string> { /* Outlook compose */ }

     return { getSelectedText, getMailBody, getMailSelectedText }
   }
   ```

---

#### H5. Mettre à jour la documentation README.md

**Fichier** : `README.md`
**Risque** : Documentation trompeuse pour les développeurs et administrateurs
**Effort** : Faible

**Guide d'implémentation** :
1. Ligne 175 : Changer "23 Word tools" en "24 Word tools", ajouter `applyTaggedFormatting` à la liste
2. Ligne 176 : Changer "22 Excel tools" en "24 Excel tools", ajouter `applyConditionalFormatting` et `getConditionalFormattingRules`
3. Ajouter une section "Quick Actions (Excel)" entre les sections Word et Outlook :
   ```markdown
   ### Frontend - Quick Actions (Excel)
   - [x] Analyze (data analysis and insights)
   - [x] Chart (chart creation suggestions)
   - [x] Formula (formula assistance)
   - [x] Format (formatting recommendations)
   - [x] Explain (explain formulas/data)
   ```
4. Ajouter les variables d'environnement manquantes dans le tableau :
   - `MODEL_NANO_LABEL`, `MODEL_NANO_MAX_TOKENS`, `MODEL_NANO_TEMPERATURE`
   - Idem pour `STANDARD`, `REASONING`, `IMAGE_LABEL`
5. Ligne 26 : Retirer "PowerPoint" du diagramme d'architecture (non implémenté)
6. Ajouter mention de la locale par défaut du manifest (`fr-FR`)
7. Section Internationalization : Mentionner les 13 langues du `languageMap`

---

#### H6. Corriger l'incohérence `.env.example` vs defaults serveur

**Fichier** : `backend/.env.example`
**Risque** : Confusion pour les administrateurs qui copient le `.env.example`
**Effort** : Très faible

**Guide d'implémentation** :
Aligner les noms de modèles dans `.env.example` avec les defaults de `server.js` :
```env
# Changer :
MODEL_NANO=gpt-5-nano
MODEL_STANDARD=gpt-5-mini
MODEL_REASONING=gpt-5.2
MODEL_IMAGE=gpt-image-1.5

# En :
MODEL_NANO=gpt-4.1-nano
MODEL_STANDARD=gpt-4.1
MODEL_REASONING=o3
MODEL_IMAGE=gpt-image-1
```
Ou inversement mettre à jour les defaults dans `server.js` - l'important est la cohérence.

---

#### H7. Ajouter un timeout sur les requêtes fetch

**Fichiers** : `backend/src/server.js`, `frontend/src/api/backend.ts`
**Risque** : Connexions pendantes indéfiniment, blocage du serveur
**Effort** : Faible

**Guide d'implémentation** :

Backend (`server.js`) :
```javascript
const controller = new AbortController()
const timeout = setTimeout(() => controller.abort(), 120000) // 2 minutes

try {
  const response = await fetch(`${LLM_API_BASE_URL}/chat/completions`, {
    // ... existing options
    signal: controller.signal,
  })
  // ...
} finally {
  clearTimeout(timeout)
}
```

Frontend (`backend.ts`) :
```typescript
export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), 10000) // 10s
  try {
    const res = await fetch(`${BACKEND_URL}/api/models`, { signal: controller.signal })
    if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`)
    return res.json()
  } finally {
    clearTimeout(timeout)
  }
}
```

---

#### H8. Typer correctement `ExcelToolDefinition`

**Fichier** : `frontend/src/utils/excelTools.ts:35`
**Risque** : Confusion sémantique, maintenabilité
**Effort** : Faible

**Guide d'implémentation** :
1. Dans `types/index.d.ts`, renommer `WordToolDefinition` en `ToolDefinition` (ou créer un alias) :
   ```typescript
   type ToolDefinition = {
     name: string
     description: string
     inputSchema: ToolInputSchema
     execute: (args: Record<string, any>) => Promise<string>
   }

   // Backward compatible aliases
   type WordToolDefinition = ToolDefinition
   type ExcelToolDefinition = ToolDefinition
   ```
2. Mettre à jour `excelTools.ts:35` :
   ```typescript
   const excelToolDefinitions: Record<ExcelToolName, ToolDefinition> = {
   ```

---

### MOYENNE (à planifier)

---

#### M1. Utiliser des IDs uniques comme key dans `v-for` au lieu d'index

**Fichier** : `frontend/src/pages/HomePage.vue:99-100`
**Risque** : Re-renders inutiles, état DOM incorrect lors d'ajouts/suppressions
**Effort** : Faible

**Guide d'implémentation** :
1. Ajouter un `id` unique à chaque message :
   ```typescript
   interface DisplayMessage {
     id: string  // Ajouter
     role: 'user' | 'assistant' | 'system'
     content: string
     imageSrc?: string
   }
   ```
2. Générer l'id à la création :
   ```typescript
   history.value.push({
     id: crypto.randomUUID(),
     role: 'user',
     content: fullMessage
   })
   ```
3. Utiliser dans le template :
   ```vue
   <div v-for="msg in history" :key="msg.id">
   ```

---

#### M2. Mémoïser `renderSegments()` dans le template

**Fichier** : `frontend/src/pages/HomePage.vue:110`
**Risque** : Performance - parsing regex à chaque render pour chaque message
**Effort** : Faible

**Guide d'implémentation** :
Utiliser un `computed` ou un `Map` de cache :
```typescript
const segmentsCache = new Map<string, RenderSegment[]>()

function renderSegments(content: string): RenderSegment[] {
  if (segmentsCache.has(content)) return segmentsCache.get(content)!
  const segments = splitThinkSegments(content)
  segmentsCache.set(content, segments)
  return segments
}
```
Ou mieux : déplacer `renderSegments` dans un composant `ChatMessage.vue` et utiliser un `computed` local.

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
  // Optionnel: afficher un toast d'erreur
}

app.config.warnHandler = (msg, instance, trace) => {
  console.warn('Vue Warning:', msg, trace)
}
```

---

#### M4. Améliorer l'accessibilité (a11y)

**Fichiers** : `HomePage.vue`, composants
**Risque** : Non-conformité WCAG, exclusion d'utilisateurs
**Effort** : Moyen

**Guide d'implémentation** :
1. Ajouter `aria-label` sur tous les boutons sans texte visible :
   ```vue
   <CustomButton :icon="Plus" text="" :aria-label="t('newChat')" />
   ```
2. Ajouter `aria-live="polite"` sur le container de messages :
   ```vue
   <div ref="messagesContainer" aria-live="polite" role="log">
   ```
3. Ajouter `role="status"` sur l'indicateur de statut backend
4. Ajouter `aria-expanded` sur les détails collapsibles (think tags)
5. Ajouter `aria-label` sur les modes (ask/agent/image) :
   ```vue
   <button :aria-label="t('askMode')" :aria-pressed="mode === 'ask'">
   ```

---

#### M5. Supprimer les watchers redondants dans `SettingsPage.vue`

**Fichier** : `frontend/src/pages/SettingsPage.vue:522-549`
**Risque** : Duplication de logique, incohérence potentielle avec `useStorage`
**Effort** : Faible

**Guide d'implémentation** :
`useStorage` persiste déjà dans localStorage. Les watchers qui font `localStorage.setItem` sont redondants sauf pour `localLanguage` (qui doit aussi mettre à jour `i18n.global.locale`).

Supprimer les watchers pour :
- `replyLanguage` (ligne 527-529)
- `agentMaxIterations` (ligne 531-533)
- `userGender` (ligne 535-537)
- `userFirstName` (ligne 539-541)
- `userLastName` (ligne 543-545)
- `excelFormulaLanguage` (ligne 547-549)

Garder uniquement :
```typescript
watch(localLanguage, (val) => {
  i18n.global.locale.value = val as 'en' | 'fr'
})
```

---

#### M6. Ajouter du retry avec backoff sur les appels API frontend

**Fichier** : `frontend/src/api/backend.ts`
**Risque** : Échec sur erreurs réseau transitoires
**Effort** : Moyen

**Guide d'implémentation** :
```typescript
async function fetchWithRetry(
  url: string,
  options: RequestInit,
  retries = 3,
  backoff = 1000
): Promise<Response> {
  for (let i = 0; i <= retries; i++) {
    try {
      const res = await fetch(url, options)
      if (res.ok || res.status < 500) return res
    } catch (err) {
      if (i === retries) throw err
    }
    await new Promise(r => setTimeout(r, backoff * Math.pow(2, i)))
  }
  throw new Error('Max retries exceeded')
}
```
Appliquer sur `healthCheck()`, `fetchModels()`, et les appels non-streaming.

---

#### M7. Séparer le backend en modules

**Fichier** : `backend/src/server.js`
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
│   └── validate.js        # Input validation (H2)
├── routes/
│   ├── health.js          # GET /health
│   ├── models.js          # GET /api/models
│   ├── chat.js            # POST /api/chat, /api/chat/sync
│   └── image.js           # POST /api/image
└── services/
    └── llm.js             # Proxy vers l'API LLM
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
app.use(morgan('combined'))
// Ou en format personnalisé:
app.use(morgan(':method :url :status :response-time ms'))
```

---

### BASSE (améliorations)

---

#### B1. Corriger la typo CSS `cursor-po`

**Fichier** : `frontend/src/pages/HomePage.vue:167`, `175`, `183`
**Effort** : Très faible

```vue
<!-- AVANT -->
class="cursor-po flex h-7 w-7 ..."

<!-- APRÈS -->
class="cursor-pointer flex h-7 w-7 ..."
```

---

#### B2. Réduire l'utilisation de `any` dans les types

**Fichiers** : `HomePage.vue`, `backend.ts`, `types/index.d.ts`
**Effort** : Faible

1. Typer les icônes :
   ```typescript
   import type { Component } from 'vue'
   interface QuickAction { key: string; label: string; icon: Component }
   ```
2. Typer `chatSync` return :
   ```typescript
   interface ChatCompletionResponse {
     choices: Array<{
       message: {
         role: string
         content: string | null
         tool_calls?: Array<{
           id: string
           function: { name: string; arguments: string }
         }>
       }
     }>
   }
   export async function chatSync(options: ChatSyncOptions): Promise<ChatCompletionResponse>
   ```

---

#### B3. Ajouter un toggle dark mode dans l'UI

**Fichier** : `frontend/src/pages/SettingsPage.vue`
**Détail** : Les variables CSS dark mode existent déjà dans `index.css:162-187` mais il n'y a aucun toggle pour l'activer

**Guide d'implémentation** :
1. Ajouter dans `SettingsPage.vue` un toggle :
   ```typescript
   const darkMode = useStorage(localStorageKey.darkMode, false)
   watch(darkMode, (val) => {
     document.documentElement.classList.toggle('dark', val)
   })
   ```
2. Ajouter un `SettingCard` avec un checkbox dans la section General

---

#### B4. Extraire les classes CSS répétées dans des composants Tailwind

**Fichier** : `frontend/src/index.css`
**Effort** : Faible

Les classes comme `rounded-md border border-border-secondary bg-surface p-2 shadow-sm` sont répétées 10+ fois. Créer des utilities :
```css
@layer components {
  .card {
    @apply rounded-md border border-border-secondary bg-surface p-2 shadow-sm;
  }
  .card-header {
    @apply rounded-md border border-border-secondary p-1 shadow-sm;
  }
}
```

---

#### B5. Documenter `hostDetection.ts` dans le README

**Fichier** : `README.md`
**Effort** : Très faible

Ajouter dans la section Project Structure :
```
│   ├── utils/
│   │   ├── hostDetection.ts  # Office host detection (Word/Excel/Outlook)
```
Et mentionner le mécanisme de détection dans la section architecture.

---

#### B6. Réduire la limite du body parser à 1MB

**Fichier** : `backend/src/server.js:83`
**Détail** : 10MB est excessif pour un proxy de chat textuel

```javascript
// AVANT
app.use(express.json({ limit: '10mb' }))

// APRÈS
app.use(express.json({ limit: '1mb' }))
```

---

## Résumé des actions par priorité

| Priorité | ID | Action | Fichier(s) | Effort |
|----------|-----|--------|-----------|--------|
| CRITIQUE | C1 | Ajouter authentification backend | `server.js` | Moyen |
| CRITIQUE | C2 | Ajouter rate limiting | `server.js` | Faible |
| CRITIQUE | C3 | Masquer erreurs LLM brutes | `server.js` | Faible |
| CRITIQUE | C4 | Cleanup `setInterval` | `HomePage.vue` | Très faible |
| HAUTE | H1 | Headers sécurité (helmet) | `server.js` | Faible |
| HAUTE | H2 | Validation inputs backend | `server.js` | Moyen |
| HAUTE | H3 | Découper `HomePage.vue` | `HomePage.vue` + nouveaux fichiers | Élevé |
| HAUTE | H4 | Extraire logique dupliquée en composables | `HomePage.vue`, `SettingsPage.vue` | Moyen |
| HAUTE | H5 | Mettre à jour README.md | `README.md` | Faible |
| HAUTE | H6 | Aligner `.env.example` avec defaults | `.env.example` | Très faible |
| HAUTE | H7 | Timeout sur requêtes fetch | `server.js`, `backend.ts` | Faible |
| HAUTE | H8 | Renommer `WordToolDefinition` → `ToolDefinition` | `types/index.d.ts`, `excelTools.ts` | Faible |
| MOYENNE | M1 | IDs uniques dans `v-for` | `HomePage.vue` | Faible |
| MOYENNE | M2 | Mémoïser `renderSegments` | `HomePage.vue` | Faible |
| MOYENNE | M3 | Error handler Vue global | `main.ts` | Faible |
| MOYENNE | M4 | Accessibilité (ARIA) | Multiple fichiers | Moyen |
| MOYENNE | M5 | Supprimer watchers redondants | `SettingsPage.vue` | Faible |
| MOYENNE | M6 | Retry avec backoff (API client) | `backend.ts` | Moyen |
| MOYENNE | M7 | Séparer backend en modules | `server.js` → multiple | Moyen |
| MOYENNE | M8 | Request logging (morgan) | `server.js` | Faible |
| BASSE | B1 | Corriger typo `cursor-po` | `HomePage.vue` | Très faible |
| BASSE | B2 | Réduire `any` types | Multiple | Faible |
| BASSE | B3 | Toggle dark mode | `SettingsPage.vue` | Faible |
| BASSE | B4 | Extraire CSS répétées | `index.css` | Faible |
| BASSE | B5 | Documenter `hostDetection.ts` | `README.md` | Très faible |
| BASSE | B6 | Réduire body parser à 1MB | `server.js` | Très faible |
