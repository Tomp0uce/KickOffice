# KickOffice - Design Review

**Date initiale**: 2026-02-15
**Dernière mise à jour**: 2026-02-16
**Scope**: Architecture, Sécurité, Qualité du code, Bugs fonctionnels, Documentation
**Fichiers analysés**: Backend (`server.js` — 448 lignes), Frontend (25+ fichiers source), Documentation (README.md, agents.md, manifests)

---

## Table des matières

1. [Résumé exécutif](#résumé-exécutif)
2. [Bilan des corrections déjà effectuées](#bilan-des-corrections-déjà-effectuées)
3. [Architecture globale](#architecture-globale)
4. [Liste des issues par criticité](#liste-des-issues-par-criticité)
   - [CRITIQUE — Bloquant / Impact immédiat](#critique--bloquant--impact-immédiat)
   - [HAUTE — À corriger rapidement](#haute--à-corriger-rapidement)
   - [MOYENNE — À planifier](#moyenne--à-planifier)
   - [BASSE — Améliorations](#basse--améliorations)
5. [Tableau récapitulatif](#tableau-récapitulatif)

---

## Résumé exécutif

KickOffice est un add-in Microsoft Office (Word, Excel, PowerPoint, Outlook) alimenté par IA, construit avec Vue 3 / TypeScript (frontend) et Express.js (backend). L'architecture est saine dans son principe : le backend sert de proxy LLM sécurisé, les clés API ne sont jamais exposées côté client, et le CORS est correctement restreint.

### État actuel

| Gravité | Nombre | Résumé |
|---------|--------|--------|
| **CRITIQUE** | 6 | Chat cassé Word/Excel, images insérées en base64, erreurs masquées, sécurité (auth, rate limit, fuite erreurs) |
| **HAUTE** | 5 | Mauvais tiers de modèles, headers sécurité, god component, backend monolithique, agent loop sans abort |
| **MOYENNE** | 5 | SSE parser fragile, error handler Vue, accessibilité, request logging, `as any` résiduels |
| **BASSE** | 3 | Dark mode toggle, CSS répétitives, README obsolète |

### Corrections déjà effectuées (review précédente)

14 items sur 26 initiaux ont été corrigés lors de la première itération. Ce document ne revient pas dessus et se concentre sur les issues restantes + nouvelles issues découvertes.

---

## Bilan des corrections déjà effectuées

| ID ancien | Item | Preuve |
|-----------|------|--------|
| C4 | Cleanup `setInterval` | `HomePage.vue:1338-1342` — `onUnmounted` + `clearInterval` |
| H2 | Validation inputs backend | `server.js:55-106` — `validateTemperature`, `validateMaxTokens`, `validateTools` |
| H4 | Extraire logique dupliquée | `savedPrompts.ts`, `getOfficeSelection()` unifiée |
| H6 | Aligner `.env.example` avec defaults | Cohérent entre `.env.example` et `server.js` |
| H7 | Timeout sur requêtes fetch | Backend: `fetchWithTimeout` + AbortController. Frontend: `fetchWithTimeoutAndRetry` |
| H8 | Typer `ToolDefinition` | `types/index.d.ts:27-36` — type générique avec alias |
| M1 | IDs uniques dans `v-for` | `crypto.randomUUID()` dans `createDisplayMessage` |
| M2 | Mémoïser `renderSegments` | `historyWithSegments` computed |
| M5 | Supprimer watchers redondants | `useStorage` gère la persistance |
| M6 | Retry avec backoff | `backend.ts:8-75` — 2 retries +10s/+30s |
| B1 | Corriger typo `cursor-po` | Toutes classes correctes |
| B2 | Réduire `any` types | Interfaces `QuickAction`, `OpenAIChatCompletion` |
| B5 | Documenter `hostDetection.ts` | Ajouté dans README |
| B6 | Réduire body parser | 4MB (`server.js:172`) |

---

## Architecture globale

### Points forts

- **Séparation claire** : Frontend (Vue 3 + Vite, port 3002) / Backend (Express.js, port 3003) / LLM API externe
- **Sécurité des secrets** : API keys uniquement côté serveur dans `.env`
- **Déploiement Docker** : Docker Compose fonctionnel avec health checks
- **Support multi-hôte** : Word (37 tools), Excel (39 tools), PowerPoint (8 tools), Outlook (13 tools)
- **i18n** : 13 langues de réponse, 2 locales UI (en/fr)
- **Agent mode** : Boucle d'outils OpenAI function-calling avec validation côté backend
- **Validation backend robuste** : Température, maxTokens, tools, prompt length
- **Timeout et retry** : Les deux côtés ont des timeouts et une stratégie de retry

### Points faibles

- **Bug bloquant** : Chat Word/Excel cassé par la limite de 32 tools dans le backend
- **Bug fonctionnel** : Insertion d'images insère du texte base64 au lieu d'images
- **Erreurs silencieuses** : Les erreurs 400 du backend ne sont pas loguées → diagnostic impossible
- **Sécurité** : Aucune authentification, aucun rate limiting
- **Maintenabilité** : `HomePage.vue` = 1344 lignes (god component), `server.js` = 448 lignes (monolithique)
- **Modèles mal tiérés** : gpt-5.2 (raisonnement) est plus rapide que nano/standard → configuration contre-intuitive

---

## Liste des issues par criticité

---

### CRITIQUE — Bloquant / Impact immédiat

---

#### C1. Chat cassé sous Word et Excel — limite de 32 tools dépassée

**Symptôme** : Le chat envoie une "erreur de réponse" dans l'interface sous Word et Excel, mais aucune erreur n'apparaît dans les logs du backend. Le chat fonctionne sous PowerPoint. Les quick actions (boutons) fonctionnent.

**Cause racine** : La validation backend `validateTools()` (`server.js:73`) rejette les requêtes avec plus de 32 tools :
```javascript
if (tools.length > 32) return { error: 'tools supports at most 32 entries' }
```

Mais les tools envoyés par le frontend sont :
- **Word** : 37 tools + 2 general = **39 tools** → ❌ rejeté (> 32)
- **Excel** : 39 tools + 2 general = **41 tools** → ❌ rejeté (> 32)
- **PowerPoint** : 8 tools + 2 general = **10 tools** → ✅ accepté
- **Outlook** : 13 tools + 2 general = **15 tools** → ✅ accepté

Le backend renvoie une erreur 400, qui est bien capturée par le frontend (`backend.ts:192-194`) et affichée comme "erreur de réponse". Mais côté backend, cette erreur 400 n'est pas loguée (pas de `console.error`), d'où l'absence d'erreur dans les logs.

**Pourquoi les quick actions marchent** : Elles utilisent `chatStream()` qui appelle `/api/chat` (streaming) **sans tools**. Seul le chat normal passe par `chatSync()` → `/api/chat/sync` avec tools.

**Fichiers** : `server.js:73`, `HomePage.vue:940-948` (construction des tools)
**Impact** : Chat complètement cassé pour Word et Excel (les 2 hôtes principaux)

**Implémentation proposée** :
1. Augmenter la limite dans `validateTools()` de 32 à 128 (ou la rendre configurable via env) :
   ```javascript
   const MAX_TOOLS = parseInt(process.env.MAX_TOOLS || '128', 10)
   if (tools.length > MAX_TOOLS) return { error: `tools supports at most ${MAX_TOOLS} entries` }
   ```
2. Alternativement, considérer un envoi dynamique des tools pertinents plutôt que l'envoi systématique de tous les tools du host. Cela réduirait aussi la taille du payload et le coût en tokens.

---

#### C2. Boutons image (copier/remplacer/ajouter) insèrent du texte base64 au lieu d'images

**Symptôme** : Après génération d'une image, les boutons "Remplacer", "Ajouter", "Copier" insèrent la chaîne de données base64 brute au lieu de l'image elle-même, ce qui plante Word/PowerPoint.

**Cause racine** : Chaîne de fallback défaillante dans `insertMessageToDocument()` (`HomePage.vue:528-547`).

Pour **Word** (`insertImageToWord`, ligne 513-526) :
- La fonction utilise `insertInlinePictureFromBase64()` qui devrait fonctionner.
- **Mais** si elle échoue (ex : contexte Word pas prêt, range invalide), le fallback appelle `copyImageToClipboard()`.
- `copyImageToClipboard()` tente `ClipboardItem` (souvent bloqué dans l'iframe Office WebView), puis tombe en fallback sur `copyToClipboard(imageSrc)` qui copie la **chaîne data URL complète** (texte de plusieurs Mo) dans le presse-papiers.

Pour **PowerPoint** et **Excel** :
- Pas de chemin d'insertion d'image direct — le code tombe directement sur `copyImageToClipboard()` → même problème de fallback texte.
- Pourtant, PowerPoint dispose d'une API `shapes.addImage(base64)` (déjà implémentée dans `powerpointTools.ts:409-464`), mais elle n'est pas utilisée par les boutons de l'UI.

**Fichiers** : `HomePage.vue:487-547`
**Impact** : Les boutons d'action sur les images sont cassés pour tous les hôtes

**Implémentation proposée** :
1. **Word** : Ajouter un try-catch plus fin dans `insertImageToWord()` avec un log d'erreur explicite. Vérifier que le `base64Payload` extrait est valide (longueur > 0).
2. **PowerPoint** : Ajouter une fonction `insertImageToPowerPoint()` qui utilise `PowerPoint.run()` + `slide.shapes.addImage(base64)` (l'API existe déjà dans les tools).
3. **Excel** : L'insertion d'image n'est pas supportée par l'API Excel JavaScript. Documenter cette limitation et afficher un message clair.
4. **Fallback clipboard** : Ne JAMAIS copier la data URL brute comme texte. Si le `ClipboardItem` échoue, afficher un message d'erreur explicite au lieu de copier la chaîne base64 :
   ```typescript
   async function copyImageToClipboard(imageSrc: string, fallback = false) {
     try {
       const response = await fetch(imageSrc)
       const blob = await response.blob()
       if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
         await navigator.clipboard.write([new ClipboardItem({ [blob.type || 'image/png']: blob })])
         messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))
         return
       }
     } catch (err) {
       console.warn('Image clipboard write failed:', err)
     }
     // Ne PAS tomber sur copyToClipboard(imageSrc) qui copie du texte base64
     messageUtil.error(t('imageClipboardNotSupported'))
   }
   ```

---

#### C3. Erreurs backend non loguées (erreurs 400 silencieuses)

**Symptôme** : Quand le backend rejette une requête (validation tools, température, etc.), l'erreur 400 est renvoyée au client mais **jamais loguée** côté serveur. Le diagnostic des bugs est impossible sans logs.

**Cause racine** : Les réponses de validation retournent directement `res.status(400).json({ error })` sans `console.error` ni `console.warn`. Seules les erreurs 500 et les erreurs LLM sont loguées.

**Fichiers** : `server.js` — toutes les lignes `return res.status(400).json(...)` (environ 15 occurrences)
**Impact** : Impossible de diagnostiquer les problèmes sans accès au frontend (le bug C1 en est la preuve directe)

**Implémentation proposée** :
Ajouter un logging systématique avant chaque réponse d'erreur :
```javascript
// Créer un helper de logging
function logAndRespond(res, status, errorObj) {
  if (status >= 400) {
    console.warn(`[${status}] ${errorObj.error}`)
  }
  return res.status(status).json(errorObj)
}
```
Ou mieux : installer `morgan` pour logger toutes les requêtes (voir M5) et ajouter un middleware de logging des erreurs.

---

#### C4. Fuite d'informations sensibles dans les erreurs LLM

**Symptôme** : Les erreurs brutes de l'API LLM sont retransmises au client avec le champ `details`.

**Fichiers** : `server.js:254-257`, `349-352`, `421-424`
```javascript
return res.status(response.status).json({
  error: `LLM API error: ${response.status}`,
  details: errorText,  // Peut contenir des URLs internes, versions, clés partielles
})
```
**Impact** : Fuite d'informations sur l'infrastructure. Présent dans les 3 endpoints.

**Implémentation proposée** :
```javascript
// Remplacer dans les 3 endpoints :
console.error(`LLM API error ${response.status}:`, errorText)
return res.status(502).json({
  error: 'The AI service returned an error. Please try again later.',
})
```

---

#### C5. Aucune authentification sur le backend

**Fichier** : `server.js` (tout le fichier)
**Impact** : N'importe qui sur le réseau peut appeler les endpoints et consommer les crédits API LLM

**Implémentation proposée** :
1. Variable `ALLOWED_API_KEYS` dans `.env` (liste séparée par virgules)
2. Middleware `requireAuth` vérifiant le header `x-api-key`
3. Appliquer sur `/api/chat`, `/api/chat/sync`, `/api/image`
4. Garder `/health` et `/api/models` publics
5. Frontend : ajouter le header `x-api-key` dans `fetchWithTimeoutAndRetry` via une variable `VITE_API_KEY`

---

#### C6. Aucun rate limiting

**Fichier** : `server.js`
**Impact** : DoS possible, consommation illimitée de crédits API

**Implémentation proposée** :
```bash
npm install express-rate-limit
```
```javascript
import rateLimit from 'express-rate-limit'
const chatLimiter = rateLimit({ windowMs: 60_000, max: 20 })
const imageLimiter = rateLimit({ windowMs: 60_000, max: 5 })
app.use('/api/chat', chatLimiter)
app.use('/api/image', imageLimiter)
```

---

### HAUTE — À corriger rapidement

---

#### H1. Configuration des tiers de modèles inadaptée

**Symptôme** : GPT-5.2 (tier `reasoning`) est beaucoup plus rapide que les modèles `nano` (gpt-5-nano) et `standard` (gpt-5-mini). L'utilisateur doit manuellement sélectionner "Raisonnement" pour obtenir la meilleure performance, ce qui est contre-intuitif.

**Fichiers** : `server.js:13-40`, `backend/.env.example`
**Impact** : UX dégradée, l'utilisateur doit connaître les internals pour choisir le bon modèle

**Configuration actuelle** :
| Tier | Modèle | Usage attendu |
|------|--------|--------------|
| nano | gpt-5-nano | Rapide, basique |
| standard | gpt-5-mini | Chat normal |
| reasoning | gpt-5.2 | Complexe → mais en fait le plus rapide |
| image | gpt-image-1.5 | Génération d'images |

**Implémentation proposée** — Reconfigurer en 3 tiers (supprimer nano) :
| Tier | Modèle | Label | Usage |
|------|--------|-------|-------|
| standard | gpt-5.2 | Standard | Chat normal + agent (rapide et performant) |
| reasoning | gpt-5.2 (reasoning mode) | Raisonnement | Tâches complexes nécessitant un raisonnement approfondi |
| image | gpt-image-1.5 | Image | Génération d'images |

Modifications :
1. `server.js` : Supprimer le tier `nano`, déplacer `gpt-5.2` en `standard`
2. `.env.example` : Mettre à jour les modèles par défaut
3. Frontend `SettingsPage.vue` / `HomePage.vue` : Le sélecteur de modèle n'affichera que 3 options au lieu de 4
4. `getChatTimeoutMs()` : Ajuster les timeouts en conséquence

**Note** : Vérifier si gpt-5.2 supporte le mode reasoning (paramètre `reasoning_effort` ou similaire) et adapter `buildChatBody()` en conséquence.

---

#### H2. Headers de sécurité HTTP manquants

**Fichier** : `server.js`
**Impact** : Vulnérabilités clickjacking, MIME sniffing, etc.

**Implémentation proposée** :
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

#### H3. `HomePage.vue` — god component (1344 lignes)

**Fichier** : `frontend/src/pages/HomePage.vue`
**Impact** : Maintenabilité, lisibilité, testabilité, performance

Le composant combine : UI chat, agent loop, quick actions, Office API (4 hôtes), clipboard, health check polling, prompts systèmes pour chaque hôte, insertion d'images.

**Implémentation proposée** — Extraire en 7 morceaux :
1. `ChatHeader.vue` — Header avec logo, boutons new chat et settings (lignes 5-38)
2. `QuickActionsBar.vue` — Barre d'actions rapides avec sélecteur de prompt (lignes 41-67)
3. `ChatMessageList.vue` — Container de messages avec empty state (lignes 70-160)
4. `ChatInput.vue` — Zone de saisie avec sélecteurs de mode et modèle (lignes 163-217)
5. Composable `useAgentLoop.ts` — La boucle agent + prompts système (lignes 653-1039)
6. Composable `useOfficeInsert.ts` — L'insertion dans le document + clipboard (lignes 1199-1311)
7. Composable `useImageActions.ts` — Génération et insertion d'images (lignes 487-547)

---

#### H4. Backend monolithique (448 lignes dans 1 fichier)

**Fichier** : `backend/src/server.js`
**Impact** : Maintenabilité à mesure que le code grandit

**Implémentation proposée** :
```
backend/src/
├── server.js              # Point d'entrée, middleware setup
├── config/
│   └── models.js          # Configuration des modèles
├── middleware/
│   ├── auth.js            # Authentication (C5)
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

#### H5. Agent loop sans support d'annulation (abort)

**Symptôme** : Quand l'utilisateur clique "Stop" pendant le chat en mode agent, le `abortController` est déclenché mais la requête `chatSync()` en cours n'est pas interrompue car `chatSync` ne reçoit pas le signal d'abort.

**Fichiers** :
- `backend.ts:183-198` : `chatSync()` ne passe pas de `signal` à `fetchWithTimeoutAndRetry()`
- `HomePage.vue:958-965` : La boucle `while` ne vérifie pas `abortController.value?.signal.aborted` entre les itérations

**Impact** : Le bouton "Stop" ne fonctionne pas pendant le mode agent. La requête continue en arrière-plan et les résultats sont ignorés quand ils arrivent, gaspillant des tokens LLM.

**Implémentation proposée** :
1. Ajouter un champ `abortSignal` optionnel à `ChatSyncOptions` et le passer à `fetchWithTimeoutAndRetry()`
2. Dans `runAgentLoop`, passer `abortController.value?.signal` à `chatSync()`
3. Ajouter une vérification `if (abortController.value?.signal.aborted) break` au début de chaque itération de la boucle

---

### MOYENNE — À planifier

---

#### M1. Parser SSE fragile (split de chunks)

**Symptôme potentiel** : Réponses tronquées ou erreurs JSON aléatoires pendant le streaming.

**Fichier** : `backend.ts:124-151`

Le parser SSE split par `\n` mais ne gère pas le cas où une ligne `data: {...}` est coupée entre deux chunks TCP. Si un chunk se termine au milieu d'une ligne JSON, le `JSON.parse()` échoue silencieusement (le catch est vide).

**Implémentation proposée** :
Maintenir un buffer résiduel entre les chunks :
```typescript
let buffer = ''
while (true) {
  const { done, value } = await reader.read()
  if (done) break
  buffer += decoder.decode(value, { stream: true })
  const lines = buffer.split('\n')
  buffer = lines.pop() || '' // Garder la dernière ligne incomplète
  for (const line of lines) {
    if (!line.startsWith('data: ')) continue
    // ... parse comme avant
  }
}
```

---

#### M2. Error handler Vue global manquant

**Fichier** : `frontend/src/main.ts`
**Impact** : Erreurs non capturées provoquent un crash silencieux

**Implémentation proposée** :
```typescript
app.config.errorHandler = (err, instance, info) => {
  console.error('Vue Global Error:', err, info)
  // Optionnel : afficher un toast
}
```

---

#### M3. Accessibilité insuffisante (a11y)

**Fichiers** : `HomePage.vue`, composants
**Impact** : Non-conformité WCAG, exclusion d'utilisateurs

**Implémentation proposée** :
1. `aria-label` sur tous les boutons sans texte (New Chat, Settings, Stop, Send, Copy, Replace, Append)
2. `aria-live="polite"` sur le container de messages
3. `role="status"` sur l'indicateur backend online/offline
4. `aria-expanded` sur les `<details>` (think tags)

---

#### M4. `as any` résiduels dans l'agent loop

**Fichier** : `HomePage.vue:1022-1026`
```typescript
currentMessages.push({
  role: 'tool' as any,
  tool_call_id: toolCall.id,
  content: result,
} as any)
```

**Fichier** : `backend.ts:157`
```typescript
tools?: any[]
```

**Implémentation proposée** :
1. Étendre `ChatMessage` pour supporter le rôle `tool` :
   ```typescript
   export type ChatMessage =
     | { role: 'system' | 'user' | 'assistant'; content: string }
     | { role: 'tool'; tool_call_id: string; content: string }
   ```
2. Typer `tools` dans `ChatSyncOptions` avec le type `ToolDefinition[]` existant.

---

#### M5. Request logging manquant

**Fichier** : `server.js`
**Impact** : Impossible de diagnostiquer ou auditer les requêtes

**Implémentation proposée** :
```bash
npm install morgan
```
```javascript
import morgan from 'morgan'
app.use(morgan(':method :url :status :response-time ms'))
```

---

### BASSE — Améliorations

---

#### B1. Pas de toggle dark mode dans l'UI

**Fichier** : `frontend/src/pages/SettingsPage.vue`
**Détail** : Les variables CSS dark mode existent dans `index.css:162-187` mais il n'y a aucun toggle pour l'activer.

**Implémentation proposée** :
```typescript
const darkMode = useStorage(localStorageKey.darkMode, false)
watch(darkMode, (val) => {
  document.documentElement.classList.toggle('dark', val)
}, { immediate: true })
```

---

#### B2. Classes CSS répétitives

**Fichier** : `frontend/src/index.css`
**Détail** : Les patterns comme `rounded-md border border-border-secondary bg-surface p-2 shadow-sm` sont répétés.

**Implémentation proposée** :
```css
@layer components {
  .card { @apply rounded-md border border-border-secondary bg-surface p-2 shadow-sm; }
}
```

---

#### B3. README.md obsolète

**Fichier** : `README.md`
**Détail** : Plusieurs informations ne correspondent plus au code.

**Corrections nécessaires** :
1. "23 Word tools" → **37 Word tools**
2. "22 Excel tools" → **39 Excel tools**
3. Ajouter **8 PowerPoint tools** et **13 Outlook tools**
4. Ajouter les Quick Actions pour Excel, PowerPoint, Outlook
5. Confirmer le support PowerPoint (marqué comme non implémenté mais l'est maintenant)
6. Mentionner les 13 langues de réponse
7. Documenter la configuration des 4 tiers de modèles (bientôt 3)

---

## Tableau récapitulatif

| Priorité | ID | Action | Statut |
|----------|-----|--------|--------|
| **CRITIQUE** | **C1** | **Chat cassé Word/Excel — limite 32 tools** | ❌ À faire |
| **CRITIQUE** | **C2** | **Boutons image insèrent base64 texte** | ❌ À faire |
| **CRITIQUE** | **C3** | **Erreurs 400 non loguées dans backend** | ❌ À faire |
| **CRITIQUE** | **C4** | **Fuite d'erreurs LLM au client** | ❌ À faire |
| **CRITIQUE** | **C5** | **Aucune authentification backend** | ❌ À faire |
| **CRITIQUE** | **C6** | **Aucun rate limiting** | ❌ À faire |
| HAUTE | H1 | Configuration des tiers de modèles inadaptée | ❌ À faire |
| HAUTE | H2 | Headers sécurité HTTP (helmet) | ❌ À faire |
| HAUTE | H3 | `HomePage.vue` god component (1344 lignes) | ❌ À faire |
| HAUTE | H4 | Backend monolithique (448 lignes) | ❌ À faire |
| HAUTE | H5 | Agent loop sans support abort | ❌ À faire |
| MOYENNE | M1 | Parser SSE fragile (split chunks) | ❌ À faire |
| MOYENNE | M2 | Error handler Vue global | ❌ À faire |
| MOYENNE | M3 | Accessibilité (ARIA) | ❌ À faire |
| MOYENNE | M4 | `as any` résiduels dans agent loop | ❌ À faire |
| MOYENNE | M5 | Request logging (morgan) | ❌ À faire |
| BASSE | B1 | Toggle dark mode | ❌ À faire |
| BASSE | B2 | Extraire CSS répétées | ❌ À faire |
| BASSE | B3 | README.md obsolète | ❌ À faire |

---

## Sécurité — Points OK (pas de problème trouvé)

- **XSS** : Aucune utilisation de `v-html` — Vue escape correctement
- **CORS** : Correctement restreint à `FRONTEND_URL`
- **Secrets** : Les clés API ne sont jamais exposées côté client
- **Injection SQL/NoSQL** : N/A (pas de base de données)
- **Validation d'input** : Température, maxTokens, tools structure, prompt length, image params tous validés
- **Timeouts** : Toutes les requêtes fetch ont des timeouts avec AbortController
