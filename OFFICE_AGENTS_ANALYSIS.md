# Rapport d'Analyse Comparative : KickOffice vs Office-Agents
## Guide d'Implémentation des Fonctionnalités à Reprendre

**Date :** 2026-03-08
**Branche :** `docs/design-review-v6-status`

---

## Table des Matières

1. [Vue d'ensemble](#1-vue-densemble)
2. [Fonctionnalités a implementer (priorite haute)](#2-fonctionnalités-à-implémenter-priorité-haute)
   - 2.1 [Screenshot / Vision (Excel + PowerPoint)](#21-screenshot--vision-excel--powerpoint)
   - 2.2 [Recherche Web (web-search)](#22-recherche-web-web-search)
   - 2.3 [Fetch Web (web-fetch)](#23-fetch-web-web-fetch)
3. [Fonctionnalités a implementer (priorite moyenne)](#3-fonctionnalités-à-implémenter-priorité-moyenne)
   - 3.1 [Enrichissement Excel : get-range-as-csv](#31-enrichissement-excel--get-range-as-csv)
   - 3.2 [Enrichissement Excel : search-data avec pagination](#32-enrichissement-excel--search-data-avec-pagination)
   - 3.3 [Enrichissement Excel : modify-workbook-structure](#33-enrichissement-excel--modify-workbook-structure)
   - 3.4 [Enrichissement Excel : modify-sheet-structure](#34-enrichissement-excel--modify-sheet-structure)
   - 3.5 [Enrichissement PowerPoint : edit-slide-xml (OOXML)](#35-enrichissement-powerpoint--edit-slide-xml-ooxml)
   - 3.6 [Enrichissement PowerPoint : verify-slides](#36-enrichissement-powerpoint--verify-slides)
   - 3.7 [Enrichissement PowerPoint : insert-icon (Iconify)](#37-enrichissement-powerpoint--insert-icon-iconify)
4. [Fonctionnalités a ignorer](#4-fonctionnalités-à-ignorer)
5. [Inventaire comparatif detaille](#5-inventaire-comparatif-détaillé)
6. [Plan d'implémentation recommandé](#6-plan-dimplémentation-recommandé)

---

## 1. Vue d'ensemble

### KickOffice (notre projet)
- **Architecture** : Frontend Vue 3 + Backend Node.js
- **Couverture** : Word, Excel, PowerPoint, Outlook (4 hosts)
- **Forces** : Agent loop mature, loop detection, rich content preservation, Word diff tracking, chart extraction via plotDigitizer, skills system, i18n FR/EN
- **Taille tools** : excelTools.ts (72 KB), powerpointTools.ts (66 KB), wordTools.ts (66 KB)

### Office-Agents (projet analysé)
- **Architecture** : Monorepo React (pnpm workspaces), SDK séparé
- **Couverture** : Excel, PowerPoint uniquement (2 hosts)
- **Forces** : Screenshot/vision, web search/fetch, OOXML XML editing, dirty tracking, pagination search, icon insertion, slide verification
- **Design** : Tools très granulaires avec TypeBox schemas, VFS avec bash intégré

### Verdict global
KickOffice a une **couverture applicative supérieure** (4 hosts vs 2) et un **agent loop plus mature**. Mais office-agents a des **capacités d'interaction externe** (web, vision) et des **outils de manipulation fine** (OOXML, pagination) qui manquent à KickOffice.

---

## 2. Fonctionnalités à implémenter (PRIORITE HAUTE)

### 2.1 Screenshot / Vision (Excel + PowerPoint)

**Pourquoi c'est critique :** Les LLMs multimodaux (Claude 3.5+, GPT-4o) peuvent analyser des images. Capturer visuellement un graphique Excel ou une diapo PowerPoint permet à l'agent de :
- Vérifier le rendu visuel de ses modifications
- Analyser un design existant avant modification
- Détecter des problèmes de mise en page impossible à voir via les données

#### 2.1.1 Screenshot Excel (screenshot-range)

**Ce que fait office-agents :**
- Utilise `range.getImage()` (Office.js API) pour capturer une plage en PNG
- Dessine des en-têtes de colonnes/lignes (A, B, C... / 1, 2, 3...) sur un canvas
- Redimensionne l'image pour optimiser les tokens LLM
- Retourne base64 PNG

**Comment l'implémenter dans KickOffice :**

```
Fichier à modifier : frontend/src/utils/excelTools.ts
Nouveau tool : screenshotRange
```

**Implémentation guidée :**

```typescript
// Dans excelTools.ts - Ajouter ce nouveau tool

async function screenshotRange(params: {
  sheetName: string;   // Nom de la feuille (on utilise les noms, pas les IDs dans KickOffice)
  range: string;       // Notation A1, ex: "A1:F20"
}): Promise<{ base64: string; mimeType: string }> {
  return officeAction(async (context) => {
    const sheet = context.workbook.worksheets.getItem(params.sheetName);
    const range = sheet.getRange(params.range);

    // Office.js API : getImage() retourne base64 PNG
    const imageResult = range.getImage();
    await context.sync();

    // imageResult.value est un base64 PNG sans préfixe
    return {
      base64: imageResult.value,
      mimeType: 'image/png'
    };
  });
}
```

**Points d'attention :**
- `range.getImage()` n'est disponible qu'à partir d'**ExcelApi 1.7** (Excel 2019+, Excel Online)
- Sur Excel Online, l'image est en résolution écran ; sur Desktop, elle est en résolution d'impression
- Il faudra compresser/redimensionner si l'image dépasse ~4.5 MB (utiliser un canvas côté front)
- L'ajout d'en-têtes (colonnes A,B,C et lignes 1,2,3) comme office-agents est un bonus pour aider le LLM, mais pas indispensable dans un premier temps

**Intégration dans l'agent loop :**
```typescript
// Dans useAgentLoop.ts - Ajouter le tool dans les tools Excel
{
  type: "function",
  function: {
    name: "screenshotRange",
    description: "Capture a visual screenshot of an Excel range as PNG image. Use this to verify visual formatting, chart rendering, or analyze existing content visually.",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", description: "Worksheet name" },
        range: { type: "string", description: "A1 notation range, e.g. 'A1:F20'" }
      },
      required: ["sheetName", "range"]
    }
  }
}
```

**Le retour doit être un content block image :**
```typescript
// Dans le handler de résultat tool, si le tool retourne une image :
if (toolResult.base64) {
  // Ajouter comme content image dans le message
  return {
    type: "image",
    source: {
      type: "base64",
      media_type: "image/png",
      data: toolResult.base64
    }
  };
}
```

#### 2.1.2 Screenshot PowerPoint (screenshot-slide)

**Ce que fait office-agents :**
- Utilise `slide.getImageAsBase64({ width: 960 })` pour capturer une diapo
- Gère les différences Desktop vs Office Online via `safeRun()`
- Retourne base64 PNG

**Comment l'implémenter dans KickOffice :**

```
Fichier à modifier : frontend/src/utils/powerpointTools.ts
Nouveau tool : screenshotSlide
```

```typescript
async function screenshotSlide(params: {
  slideIndex: number;  // 0-based index
}): Promise<{ base64: string; mimeType: string }> {
  return officeAction(async (context) => {
    const slide = context.presentation.slides.getItemAt(params.slideIndex);

    // PowerPoint API : getImageAsBase64() avec largeur cible
    const imageResult = slide.getImageAsBase64({ width: 960 });
    await context.sync();

    return {
      base64: imageResult.value,
      mimeType: 'image/png'
    };
  });
}
```

**Points d'attention :**
- `slide.getImageAsBase64()` requiert **PowerPointApi 1.5** (assez récent)
- Sur Office Online, les appels concurrents peuvent échouer → sérialiser les captures
- Largeur 960px est un bon compromis qualité/taille pour le LLM

**Mise à jour du skill PowerPoint :**
```markdown
# Dans powerpoint.skill.md - Ajouter :

## Screenshot Verification Workflow
After making visual changes to a slide:
1. Call `screenshotSlide` to capture the result
2. Analyze the image to verify the changes
3. If issues detected, iterate
```

---

### 2.2 Recherche Web (web-search)

**Pourquoi c'est critique :** L'agent est actuellement limité à ses connaissances internes. Pour remplir un document Excel avec des données récentes ou rédiger un email dans Outlook avec des infos à jour, la recherche web est essentielle.

**Ce que fait office-agents :**
- 4 providers : DuckDuckGo (gratuit), Brave, Serper, Exa
- DuckDuckGo par défaut (scraping HTML, pas d'API key)
- Support de proxy CORS pour les appels depuis le navigateur
- Pagination, filtres temporels, filtres géographiques

**Comment l'implémenter dans KickOffice :**

La meilleure approche est de **passer par le backend Node.js** (pas de problème CORS, pas besoin de proxy).

```
Fichiers à créer/modifier :
- backend/src/routes/webSearch.js       (NOUVEAU)
- backend/src/services/webSearchService.js (NOUVEAU)
- frontend/src/utils/generalTools.ts    (MODIFIER)
- frontend/src/api/backend.ts           (MODIFIER)
```

#### Backend : Service de recherche web

```javascript
// backend/src/services/webSearchService.js

const fetch = require('node-fetch');

/**
 * Recherche DuckDuckGo (gratuit, pas d'API key)
 * Méthode : scraping de la page HTML de résultats
 */
async function searchDuckDuckGo(query, options = {}) {
  const { maxResults = 5, region = 'fr-fr', timelimit } = options;

  const params = new URLSearchParams({
    q: query,
    kl: region,      // Locale (fr-fr, us-en, etc.)
  });
  if (timelimit) params.set('df', timelimit); // d=day, w=week, m=month, y=year

  const response = await fetch('https://html.duckduckgo.com/html/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString()
  });

  const html = await response.text();

  // Parser les résultats depuis le HTML
  // Chaque résultat est dans un bloc <div class="result">
  const results = parseDDGResults(html, maxResults);
  return results;
}

/**
 * Recherche via Serper (plus fiable, nécessite SERPER_API_KEY)
 */
async function searchSerper(query, options = {}) {
  const { maxResults = 5, region, timelimit } = options;

  const body = {
    q: query,
    num: maxResults,
  };
  if (region) body.gl = region.split('-')[1]?.toUpperCase() || 'FR';
  if (timelimit) {
    const tbsMap = { d: 'qdr:d', w: 'qdr:w', m: 'qdr:m', y: 'qdr:y' };
    body.tbs = tbsMap[timelimit];
  }

  const response = await fetch('https://google.serper.dev/search', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-API-KEY': process.env.SERPER_API_KEY
    },
    body: JSON.stringify(body)
  });

  const data = await response.json();
  return (data.organic || []).map(r => ({
    title: r.title,
    href: r.link,
    body: r.snippet
  }));
}

module.exports = { searchDuckDuckGo, searchSerper };
```

#### Backend : Route

```javascript
// backend/src/routes/webSearch.js

const express = require('express');
const router = express.Router();
const { searchDuckDuckGo, searchSerper } = require('../services/webSearchService');

router.post('/search', async (req, res) => {
  try {
    const { query, maxResults = 5, region = 'fr-fr', timelimit, provider = 'duckduckgo' } = req.body;

    if (!query) return res.status(400).json({ error: 'query is required' });

    let results;
    if (provider === 'serper' && process.env.SERPER_API_KEY) {
      results = await searchSerper(query, { maxResults, region, timelimit });
    } else {
      results = await searchDuckDuckGo(query, { maxResults, region, timelimit });
    }

    res.json({ success: true, results, query, provider });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
```

#### Frontend : Tool pour l'agent

```typescript
// Dans generalTools.ts - Ajouter :

async function webSearch(params: {
  query: string;
  maxResults?: number;  // Default 5
  region?: string;      // Default 'fr-fr'
}): Promise<string> {
  const response = await backendApi.post('/api/web/search', {
    query: params.query,
    maxResults: params.maxResults || 5,
    region: params.region || 'fr-fr'
  });

  if (!response.data.success) throw new Error(response.data.error);

  // Formater pour le LLM
  const formatted = response.data.results.map((r, i) =>
    `${i+1}. **${r.title}**\n   URL: ${r.href}\n   ${r.body}`
  ).join('\n\n');

  return formatted || 'Aucun résultat trouvé.';
}
```

**Tool definition pour l'agent :**
```typescript
{
  type: "function",
  function: {
    name: "webSearch",
    description: "Search the web for current information. Use when you need recent data, facts, or references that may not be in your training data. Returns titles, URLs, and snippets.",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query" },
        maxResults: { type: "number", description: "Max results (default 5)" },
        region: { type: "string", description: "Region code, e.g. 'fr-fr', 'us-en'" }
      },
      required: ["query"]
    }
  }
}
```

---

### 2.3 Fetch Web (web-fetch)

**Pourquoi c'est important :** Après une recherche web, l'agent doit pouvoir lire le contenu d'une page trouvée pour extraire des données précises.

**Ce que fait office-agents :**
- Télécharge l'URL
- HTML → extraction via Readability → conversion Markdown via Turndown
- Gestion des fichiers binaires (PDF, images)
- Extraction de métadonnées (titre, auteur, date)

**Comment l'implémenter dans KickOffice :**

```
Fichiers à créer :
- backend/src/routes/webFetch.js          (NOUVEAU)
- backend/src/services/webFetchService.js (NOUVEAU)
```

```javascript
// backend/src/services/webFetchService.js

const fetch = require('node-fetch');
const { Readability } = require('@mozilla/readability');
const { JSDOM } = require('jsdom');
const TurndownService = require('turndown');

async function fetchWebPage(url, options = {}) {
  const { maxLength = 50000 } = options;

  const response = await fetch(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; KickOffice/1.0)',
      'Accept': 'text/html,application/xhtml+xml,*/*'
    },
    timeout: 15000
  });

  const contentType = response.headers.get('content-type') || '';

  // Si pas HTML, retourner le texte brut (JSON, XML, etc.)
  if (!contentType.includes('text/html')) {
    const text = await response.text();
    return {
      title: url,
      content: text.slice(0, maxLength),
      url,
      contentType
    };
  }

  const html = await response.text();

  // Extraction via Readability
  const dom = new JSDOM(html, { url });
  const reader = new Readability(dom.window.document);
  const article = reader.parse();

  // Conversion Markdown
  const turndown = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-'
  });

  const markdown = article
    ? turndown.turndown(article.content)
    : turndown.turndown(html);

  return {
    title: article?.title || dom.window.document.title || url,
    content: markdown.slice(0, maxLength),
    url,
    byline: article?.byline || null,
    contentType: 'text/markdown'
  };
}

module.exports = { fetchWebPage };
```

**Dépendances npm à ajouter au backend :**
```bash
cd backend && npm install @mozilla/readability jsdom turndown
```

**Tool definition :**
```typescript
{
  type: "function",
  function: {
    name: "webFetch",
    description: "Fetch and read the content of a web page. Extracts main article content and converts to readable text. Use after webSearch to read specific pages.",
    parameters: {
      type: "object",
      properties: {
        url: { type: "string", description: "Full URL to fetch" },
        maxLength: { type: "number", description: "Max chars to return (default 50000)" }
      },
      required: ["url"]
    }
  }
}
```

---

## 3. Fonctionnalités à implémenter (PRIORITE MOYENNE)

### 3.1 Enrichissement Excel : get-range-as-csv

**Problème actuel :** KickOffice retourne les données Excel en JSON (`getWorksheetData`, `getDataFromSheet`). Le JSON est **gourmand en tokens** pour les grands jeux de données.

**Ce que fait office-agents :** Retourne les données en CSV, qui est 2-3x plus compact que JSON pour les données tabulaires.

**Comment l'implémenter :**

```
Fichier à modifier : frontend/src/utils/excelTools.ts
Nouveau tool : getRangeAsCsv
```

```typescript
async function getRangeAsCsv(params: {
  sheetName: string;
  range: string;
  maxRows?: number;  // Default 500
}): Promise<string> {
  return officeAction(async (context) => {
    const sheet = context.workbook.worksheets.getItem(params.sheetName);
    const range = sheet.getRange(params.range);
    range.load('values,rowCount,columnCount');
    await context.sync();

    const maxRows = params.maxRows || 500;
    const values = range.values;
    const rows = values.slice(0, maxRows);

    // Conversion CSV
    const csv = rows.map(row =>
      row.map(cell => {
        const str = String(cell ?? '');
        // Échapper si contient virgule, guillemet ou saut de ligne
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');

    const hasMore = values.length > maxRows;
    return `Rows: ${rows.length}/${values.length}${hasMore ? ' (truncated)' : ''}\n\n${csv}`;
  });
}
```

**Avantage token :** Pour un tableau de 100x10 cellules :
- JSON : ~15,000 tokens
- CSV : ~5,000 tokens (3x moins cher)

**Quand utiliser :** Ajouter dans le skill Excel la règle :
```markdown
## Data Reading Strategy
- For data ANALYSIS (formulas, statistics) → use `getRangeAsCsv` (token-efficient)
- For data FORMATTING (styles, colors) → use `getSelectedCells` (includes formatting)
```

---

### 3.2 Enrichissement Excel : search-data avec pagination

**Problème actuel :** `findData` dans KickOffice retourne tous les résultats d'un coup. Sur un fichier de 100,000 lignes, ça explose la fenêtre de contexte.

**Ce que fait office-agents :** Pagination avec `offset` et `maxResults`. L'agent peut parcourir les résultats page par page.

**Comment l'implémenter :**

Modifier `findData` dans `excelTools.ts` pour supporter :

```typescript
async function findData(params: {
  searchTerm: string;
  sheetName?: string;     // Optionnel : limiter à une feuille
  matchCase?: boolean;    // Default false
  matchEntireCell?: boolean; // Default false
  useRegex?: boolean;     // Default false
  maxResults?: number;    // Default 50 (NOUVEAU)
  offset?: number;        // Default 0 (NOUVEAU)
}): Promise<string> {
  // ... implémentation existante ...

  // AJOUTER : Pagination
  const allMatches = [...]; // Résultats existants
  const offset = params.offset || 0;
  const maxResults = params.maxResults || 50;
  const page = allMatches.slice(offset, offset + maxResults);
  const hasMore = offset + maxResults < allMatches.length;

  return JSON.stringify({
    matches: page,
    totalFound: allMatches.length,
    returned: page.length,
    offset,
    hasMore,
    nextOffset: hasMore ? offset + maxResults : null
  });
}
```

---

### 3.3 Enrichissement Excel : modify-workbook-structure

**Problème actuel :** KickOffice a `addWorksheet` mais pas de `deleteWorksheet`, `renameWorksheet`, ni `duplicateWorksheet` en tant que tools dédiés. Ces opérations nécessitent actuellement `eval_officejs`.

**Ce que fait office-agents :** Un seul tool `modify_workbook_structure` avec 4 opérations : create, delete, rename, duplicate.

**Comment l'implémenter :**

```typescript
// Dans excelTools.ts - Enrichir ou créer :

async function modifyWorkbookStructure(params: {
  operation: 'create' | 'delete' | 'rename' | 'duplicate';
  sheetName?: string;      // Feuille cible (delete/rename/duplicate)
  newName?: string;         // Nouveau nom (create/rename/duplicate)
  tabColor?: string;        // Couleur onglet (hex, optionnel)
}): Promise<string> {
  return officeAction(async (context) => {
    const sheets = context.workbook.worksheets;

    switch (params.operation) {
      case 'create': {
        const newSheet = sheets.add(params.newName);
        if (params.tabColor) newSheet.tabColor = params.tabColor;
        await context.sync();
        return `Worksheet "${params.newName}" created.`;
      }
      case 'delete': {
        const sheet = sheets.getItem(params.sheetName!);
        sheet.delete();
        await context.sync();
        return `Worksheet "${params.sheetName}" deleted.`;
      }
      case 'rename': {
        const sheet = sheets.getItem(params.sheetName!);
        sheet.name = params.newName!;
        await context.sync();
        return `Worksheet renamed to "${params.newName}".`;
      }
      case 'duplicate': {
        const sheet = sheets.getItem(params.sheetName!);
        const copy = sheet.copy();
        if (params.newName) copy.name = params.newName;
        await context.sync();
        return `Worksheet "${params.sheetName}" duplicated.`;
      }
    }
  });
}
```

---

### 3.4 Enrichissement Excel : modify-sheet-structure

**Problème actuel :** KickOffice a `modifyStructure` qui gère insert/delete de lignes/colonnes, mais pas hide/unhide ni freeze panes.

**Ce que fait office-agents :** Ajoute hide, unhide, freeze, unfreeze.

**Comment enrichir :**

```typescript
// Dans excelTools.ts - Ajouter à modifyStructure ou créer un nouveau tool :

// Cas freeze :
case 'freeze': {
  const sheet = context.workbook.worksheets.getItem(params.sheetName!);
  sheet.freezePanes.freezeAt(sheet.getRange(params.reference!)); // ex: "B3"
  await context.sync();
  return `Panes frozen at ${params.reference}.`;
}
case 'unfreeze': {
  const sheet = context.workbook.worksheets.getItem(params.sheetName!);
  sheet.freezePanes.unfreeze();
  await context.sync();
  return `Panes unfrozen.`;
}

// Cas hide/unhide :
case 'hideRows': {
  const sheet = context.workbook.worksheets.getItem(params.sheetName!);
  const range = sheet.getRange(params.reference!); // ex: "5:8"
  range.rowHidden = true;
  await context.sync();
  return `Rows ${params.reference} hidden.`;
}
case 'hideColumns': {
  const sheet = context.workbook.worksheets.getItem(params.sheetName!);
  const range = sheet.getRange(params.reference!); // ex: "C:D"
  range.columnHidden = true;
  await context.sync();
  return `Columns ${params.reference} hidden.`;
}
```

---

### 3.5 Enrichissement PowerPoint : edit-slide-xml (OOXML)

**Pourquoi c'est important :** L'API Office.js pour PowerPoint est **notoirement limitée**. De nombreuses opérations (graphiques, diagrammes, animations, SmartArt) ne sont PAS accessibles via Office.js. L'édition directe du XML OOXML est le seul moyen.

**Ce que fait office-agents :**
1. Exporte la slide en base64 PPTX via `slide.exportAsBase64()`
2. Ouvre le PPTX comme archive ZIP avec JSZip
3. Permet de modifier n'importe quel fichier XML dans l'archive
4. Réinsère la slide modifiée via `insertSlidesFromBase64()`
5. Supprime l'originale

**Comment l'implémenter dans KickOffice :**

C'est le changement le plus complexe mais aussi le plus puissant.

```
Fichiers à modifier/créer :
- frontend/src/utils/powerpointTools.ts   (MODIFIER)
- frontend/src/utils/pptxZipUtils.ts      (NOUVEAU - utilitaires ZIP)
```

**Étape 1 : Ajouter JSZip comme dépendance frontend**
```bash
cd frontend && npm install jszip
```

**Étape 2 : Utilitaires ZIP/XML**

```typescript
// frontend/src/utils/pptxZipUtils.ts

import JSZip from 'jszip';

/**
 * Workflow atomique pour éditer une slide via son XML :
 * 1. Exporter la slide en base64
 * 2. Charger dans JSZip
 * 3. Appeler le callback avec l'archive
 * 4. Si modifié, réinsérer et supprimer l'originale
 */
export async function withSlideZip(
  context: PowerPoint.RequestContext,
  slideIndex: number,
  callback: (zip: JSZip, markDirty: () => void) => Promise<any>
): Promise<any> {
  // 1. Charger les slides
  const slides = context.presentation.slides;
  slides.load('items/id');
  await context.sync();

  const targetSlide = slides.items[slideIndex];
  const slideId = targetSlide.id;

  // 2. Exporter en base64
  const base64Result = targetSlide.exportAsBase64();
  await context.sync();

  // 3. Charger dans JSZip
  const zip = await JSZip.loadAsync(base64Result.value, { base64: true });

  let dirty = false;
  const markDirty = () => { dirty = true; };

  // 4. Appeler le callback
  const result = await callback(zip, markDirty);

  // 5. Si modifié, réinsérer
  if (dirty) {
    const newBase64 = await zip.generateAsync({ type: 'base64' });

    // Trouver la slide précédente pour l'insertion
    const prevSlideId = slideIndex > 0 ? slides.items[slideIndex - 1].id : undefined;

    // Insérer la slide modifiée
    context.presentation.insertSlidesFromBase64(newBase64, {
      targetSlideId: prevSlideId
    });
    await context.sync();

    // Supprimer l'originale (qui est maintenant décalée)
    // Recharger pour avoir les IDs à jour
    slides.load('items/id');
    await context.sync();

    const originalSlide = slides.items.find(s => s.id === slideId);
    if (originalSlide) {
      originalSlide.delete();
      await context.sync();
    }
  }

  return result;
}

// Utilitaires XML
export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function sanitizeXmlAmpersands(xml: string): string {
  return xml.replace(/&(?!amp;|lt;|gt;|apos;|quot;|#\d+;|#x[0-9a-fA-F]+;)/g, '&amp;');
}
```

**Étape 3 : Tool editSlideXml**

```typescript
// Dans powerpointTools.ts :

async function editSlideXml(params: {
  slideIndex: number;
  code: string;         // Code JS à exécuter sur le ZIP
  explanation?: string;
}): Promise<string> {
  return officeAction(async (context) => {
    const result = await withSlideZip(context, params.slideIndex,
      async (zip, markDirty) => {
        // Exécuter le code utilisateur dans un sandbox
        // Le code a accès à : zip, markDirty, DOMParser, XMLSerializer, escapeXml
        const fn = new Function('zip', 'markDirty', 'escapeXml', 'DOMParser', 'XMLSerializer',
          `return (async () => { ${params.code} })()`
        );
        return fn(zip, markDirty, escapeXml, DOMParser, XMLSerializer);
      }
    );
    return JSON.stringify({ success: true, result });
  });
}
```

**Points d'attention critiques :**
- `slide.exportAsBase64()` requiert **PowerPointApi 1.5**
- `insertSlidesFromBase64()` requiert **PowerPointApi 1.4**
- Sur Office Online, les appels concurrents sont problématiques → sérialiser
- Ne JAMAIS permettre d'ajouter des références externes (sécurité)
- Toujours sanitiser les ampersands dans le XML avant réinsertion

---

### 3.6 Enrichissement PowerPoint : verify-slides

**Pourquoi c'est utile :** Après des modifications, l'agent peut vérifier automatiquement les problèmes de layout.

**Ce que fait office-agents :**
- Itère toutes les slides et shapes
- Détecte les chevauchements (AABB collision)
- Détecte les débordements hors slide
- Retourne un rapport détaillé

**Comment l'implémenter :**

```typescript
// Dans powerpointTools.ts :

async function verifySlides(): Promise<string> {
  return officeAction(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    // Dimensions de la slide
    // Note: pageSetup n'est pas toujours accessible, utiliser des valeurs par défaut
    const slideWidth = 12192000;  // 10 inches en EMU (standard 16:9)
    const slideHeight = 6858000;  // 7.5 inches en EMU

    const issues = [];

    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items/id,items/name,items/left,items/top,items/width,items/height');
      await context.sync();

      const shapes = slide.shapes.items.map(s => ({
        id: s.id, name: s.name,
        left: s.left, top: s.top,
        width: s.width, height: s.height
      }));

      // Vérifier débordements
      for (const s of shapes) {
        if (s.left + s.width > slideWidth || s.top + s.height > slideHeight) {
          issues.push(`Slide ${i+1}: Shape "${s.name}" overflows slide boundaries`);
        }
      }

      // Vérifier chevauchements
      for (let a = 0; a < shapes.length; a++) {
        for (let b = a + 1; b < shapes.length; b++) {
          const sa = shapes[a], sb = shapes[b];
          if (sa.left < sb.left + sb.width && sa.left + sa.width > sb.left &&
              sa.top < sb.top + sb.height && sa.top + sa.height > sb.top) {
            issues.push(`Slide ${i+1}: "${sa.name}" overlaps with "${sb.name}"`);
          }
        }
      }
    }

    return issues.length === 0
      ? 'All slides verified: no overlaps or overflows detected.'
      : `Found ${issues.length} issue(s):\n` + issues.join('\n');
  });
}
```

---

### 3.7 Enrichissement PowerPoint : insert-icon (Iconify)

**Pourquoi c'est utile :** L'agent peut insérer des icônes professionnelles (Material Design, Fluent, Feather, etc.) directement dans les slides sans que l'utilisateur fournisse des fichiers.

**Ce que fait office-agents :**
- Recherche d'icônes via l'API Iconify (https://api.iconify.design)
- Récupération du SVG
- Rasterisation en PNG (pour compatibilité)
- Insertion dans la slide via ZIP/XML

**Comment l'implémenter :**

**Backend : proxy Iconify**

```javascript
// backend/src/routes/icons.js

const fetch = require('node-fetch');
const router = require('express').Router();

// Recherche d'icônes
router.get('/search', async (req, res) => {
  const { query, limit = 10, prefix } = req.query;
  const params = new URLSearchParams({ query, limit });
  if (prefix) params.set('prefix', prefix);

  const response = await fetch(`https://api.iconify.design/search?${params}`);
  const data = await response.json();
  res.json(data);
});

// Récupération SVG d'une icône
router.get('/svg/:prefix/:name', async (req, res) => {
  const { prefix, name } = req.params;
  const { color } = req.query;

  let url = `https://api.iconify.design/${prefix}/${name}.svg`;
  if (color) url += `?color=${encodeURIComponent(color)}`;

  const response = await fetch(url);
  const svg = await response.text();
  res.type('image/svg+xml').send(svg);
});

module.exports = router;
```

**Frontend : tools**

```typescript
// Deux tools complémentaires :

// 1. Rechercher des icônes
async function searchIcons(params: { query: string; limit?: number }): Promise<string> {
  const response = await backendApi.get('/api/icons/search', { params });
  return JSON.stringify(response.data);
}

// 2. Insérer une icône dans une slide
async function insertIcon(params: {
  iconId: string;        // ex: "mdi:home"
  slideIndex: number;
  x?: number;            // Position en points
  y?: number;
  width?: number;
  height?: number;
  color?: string;        // Couleur hex, ex: "#FF5733"
}): Promise<string> {
  // Récupérer le SVG via le backend
  const [prefix, name] = params.iconId.split(':');
  const svgResponse = await backendApi.get(`/api/icons/svg/${prefix}/${name}`, {
    params: { color: params.color },
    responseType: 'text'
  });

  // Convertir SVG en base64 pour insertion
  const svgBase64 = btoa(svgResponse.data);

  // Utiliser insertImageOnSlide existant ou Office.js addImage
  // (réutiliser le mécanisme d'insertion d'image déjà en place)
  return insertImageOnSlide({
    slideIndex: params.slideIndex,
    imageBase64: svgBase64,
    mimeType: 'image/svg+xml',
    x: params.x,
    y: params.y,
    width: params.width || 50,
    height: params.height || 50
  });
}
```

---

## 4. Fonctionnalités à IGNORER

### 4.1 Exécution Bash système (bash.ts, read-file.ts local)
**Raison :** office-agents tourne en local/CLI. KickOffice est un add-in distribué (navigateur/Webview). Pas d'accès filesystem réel ni exécution shell système. Notre VFS + sandbox est la bonne approche.

### 4.2 Séparation Monorepo strict
**Raison :** Notre architecture single-app multi-host est plus simple à maintenir. Un monorepo pnpm workspaces ajouterait de la complexité de build sans gain fonctionnel pour une petite équipe.

### 4.3 Migration vers React
**Raison :** Notre architecture Vue 3 (composables, skills system) est propre et fonctionnelle. Aucun avantage à migrer.

### 4.4 OAuth multi-provider (Anthropic/OpenAI direct)
**Raison :** Notre backend proxie déjà les appels LLM. Pas besoin de connexion OAuth directe depuis le front vers les providers.

### 4.5 SES Lockdown (Secure EcmaScript)
**Raison :** Intéressant pour la sécurité mais très complexe à intégrer (modifie les prototypes globaux). Notre sandbox actuel est suffisant car le code exécuté est généré par le LLM, pas par l'utilisateur final.

### 4.6 Dirty Tracking (Proxy-based)
**Raison :** Concept élégant (tracking automatique des mutations) mais complexité énorme pour un gain marginal. Notre approche directe (chaque tool sait ce qu'il modifie) est suffisante.

### 4.7 Sheet Stable IDs (GUID → integer mapping)
**Raison :** Office-agents utilise des IDs stables car les GUIDs changent au reload. Nous utilisons les noms de feuilles directement, ce qui est plus lisible et suffisant pour notre cas d'usage.

---

## 5. Inventaire Comparatif Détaillé

### Excel Tools

| Capacité | KickOffice | Office-Agents | Action |
|----------|-----------|---------------|--------|
| Lire cellules | `getSelectedCells`, `getWorksheetData` | `get_cell_ranges` | OK - déjà couvert |
| Lire en CSV | -- | `get_range_as_csv` | **AJOUTER** (priorité moyenne) |
| Écrire cellules | `setCellRange` | `set_cell_range` | OK - déjà couvert |
| Effacer | `clearRange` | `clear_cell_range` | OK - déjà couvert |
| Rechercher | `findData` (pas de pagination) | `search_data` (paginé) | **ENRICHIR** avec pagination |
| Copier plage | -- | `copy_to` | Faisable via `eval_officejs`, faible priorité |
| Screenshot | -- | `screenshot_range` | **AJOUTER** (priorité haute) |
| Chartes/Pivots | `manageObject` | `modify_object` | OK - déjà couvert |
| Lister objets | `getAllObjects` | `get_all_objects` | OK - déjà couvert |
| Ajouter feuille | `addWorksheet` | `modify_workbook_structure` | **ENRICHIR** (delete/rename/duplicate) |
| Lignes/Colonnes | `modifyStructure` | `modify_sheet_structure` | **ENRICHIR** (hide/freeze) |
| Redimensionner | via `formatRange` | `resize_range` | OK - couvert autrement |
| Eval JS | `eval_officejs` | `eval_officejs` | OK - identique |
| Formatting conditionnel | `applyConditionalFormatting` | -- | **Avantage KickOffice** |
| Named Ranges | `getNamedRanges`, `setNamedRange` | -- | **Avantage KickOffice** |
| Sort | `sortRange` | -- | **Avantage KickOffice** |
| Protection | `protectWorksheet` | -- | **Avantage KickOffice** |
| Chart extraction (image→data) | plotDigitizer | -- | **Avantage KickOffice** |

### PowerPoint Tools

| Capacité | KickOffice | Office-Agents | Action |
|----------|-----------|---------------|--------|
| Lire texte slide | `getSlideContent` | `read_slide_text` | OK |
| Écrire texte | `insertContent`, `proposeShapeTextRevision` | `edit_slide_text` | OK |
| Lister shapes | `getShapes` | `list_slide_shapes` | OK |
| Insérer image | `insertImageOnSlide` | `insert-image` (VFS cmd) | OK |
| Screenshot | -- | `screenshot_slide` | **AJOUTER** (priorité haute) |
| Edit XML/OOXML | -- | `edit_slide_xml` | **AJOUTER** (priorité moyenne) |
| Edit chart XML | -- | `edit_slide_chart` | **AJOUTER** (avec edit_slide_xml) |
| Edit master | -- | `edit_slide_master` | **AJOUTER** (avec edit_slide_xml) |
| Dupliquer slide | -- | `duplicate_slide` | **AJOUTER** (simple) |
| Vérifier layout | -- | `verify_slides` | **AJOUTER** (priorité moyenne) |
| Insérer icône | -- | `insert-icon` (Iconify) | **AJOUTER** (priorité moyenne) |
| Speaker notes | `getSpeakerNotes`, `setSpeakerNotes` | -- | **Avantage KickOffice** |
| Vue d'ensemble | `getAllSlidesOverview` | -- | **Avantage KickOffice** |
| Eval JS | `eval_powerpointjs` | `execute_office_js` | OK |
| Ajouter slide | `addSlide` | -- (via code) | OK |
| Supprimer slide | `deleteSlide` | -- (via code) | OK |

### Capacités Transversales

| Capacité | KickOffice | Office-Agents | Action |
|----------|-----------|---------------|--------|
| Word | Complet (diff, track changes) | -- | **Avantage KickOffice** |
| Outlook | Complet | -- | **Avantage KickOffice** |
| Web Search | -- | DuckDuckGo/Brave/Serper/Exa | **AJOUTER** (priorité haute) |
| Web Fetch | -- | Readability + Turndown | **AJOUTER** (priorité haute) |
| VFS | `vfsWriteFile`, `vfsReadFile` | VFS complet + bash | OK - suffisant |
| File conversion | PDF/DOCX/XLSX upload | pdf-to-text, docx-to-text, xlsx-to-csv | OK - via upload route |
| Image search | -- | Serper image search | Optionnel |
| Icon search | -- | Iconify API | **AJOUTER** (priorité moyenne) |
| Session persistence | Partiel (fichiers, images) | Complet (IndexedDB sessions) | OK pour notre usage |
| Loop detection | Oui (sliding window) | -- | **Avantage KickOffice** |
| i18n | FR/EN | -- | **Avantage KickOffice** |
| Skills system | Oui (skill.md par host) | Oui (installable) | Comparable |

---

## 6. Plan d'Implémentation Recommandé

### Phase 1 : Quick Wins (1-2 jours)
**Impact maximum, effort minimum**

| # | Feature | Effort | Impact | Fichiers |
|---|---------|--------|--------|----------|
| 1 | `screenshotSlide` (PPT) | 2h | Tres haut | powerpointTools.ts |
| 2 | `screenshotRange` (Excel) | 2h | Tres haut | excelTools.ts |
| 3 | `duplicateSlide` (PPT) | 1h | Moyen | powerpointTools.ts |
| 4 | `verifySlides` (PPT) | 2h | Moyen | powerpointTools.ts |
| 5 | `getRangeAsCsv` (Excel) | 1h | Moyen | excelTools.ts |

### Phase 2 : Web Integration ⏸️ DEFERRED (2-3 jours)
**Capacité externe la plus demandée - planifiée pour future release**

| # | Feature | Effort | Impact | Statut |
|---|---------|--------|--------|--------|
| 6 | `webSearch` (backend route) | 3h | Tres haut | 🔲 DEFERRED |
| 7 | `webFetch` (backend route) | 3h | Haut | 🔲 DEFERRED |
| 8 | Integration frontend tools | 2h | -- | 🔲 DEFERRED |
| 9 | Deps npm backend | 30min | -- | 🔲 DEFERRED |
| 10 | Mise à jour skills/prompts | 1h | -- | 🔲 DEFERRED |

### Phase 3 : Excel Enrichment (1-2 jours)
**Améliorer les outils existants**

| # | Feature | Effort | Impact | Fichiers |
|---|---------|--------|--------|----------|
| 11 | Pagination `findData` | 2h | Moyen | excelTools.ts |
| 12 | `modifyWorkbookStructure` enrichi | 2h | Moyen | excelTools.ts |
| 13 | Hide/unhide/freeze dans `modifyStructure` | 2h | Moyen | excelTools.ts |

### Phase 4 : PowerPoint OOXML (3-5 jours)
**Le plus complexe mais le plus puissant**

| # | Feature | Effort | Impact | Fichiers |
|---|---------|--------|--------|----------|
| 14 | Utilitaires ZIP/XML | 4h | -- | pptxZipUtils.ts (nouveau) |
| 15 | `editSlideXml` tool | 4h | Tres haut | powerpointTools.ts |
| 16 | `editSlideChart` (même base) | 2h | Haut | powerpointTools.ts |
| 17 | `editSlideMaster` | 4h | Moyen | powerpointTools.ts |
| 18 | Icon search/insert | 3h | Moyen | icons.js (backend), powerpointTools.ts |

### Phase 5 : Polish & Skills (1 jour)
**Finalisation**

| # | Feature | Effort | Impact | Fichiers |
|---|---------|--------|--------|----------|
| 19 | Mise à jour excel.skill.md | 1h | Haut | excel.skill.md |
| 20 | Mise à jour powerpoint.skill.md | 1h | Haut | powerpoint.skill.md |
| 21 | Mise à jour common.skill.md | 30min | Moyen | common.skill.md |
| 22 | Tests manuels E2E | 4h | -- | -- |

---

### Estimation Totale
- **Phase 1** : 1-2 jours ✅ À implémenter
- **Phase 2** : 2-3 jours ⏸️ DEFERRED (web search/fetch)
- **Phase 3** : 1-2 jours ✅ À implémenter
- **Phase 4** : 3-5 jours ✅ À implémenter
- **Phase 5** : 1 jour ✅ À implémenter

**Total pour release actuelle : 6-10 jours** (sans Phase 2)
**Total avec web integration : 8-13 jours** (Phase 2 déferred pour future release)

### Prérequis Techniques
- **npm packages backend** : `@mozilla/readability`, `jsdom`, `turndown`
- **npm packages frontend** : `jszip` (pour OOXML editing)
- **API keys optionnels** : `SERPER_API_KEY` (pour recherche web premium)
- **Versions Office.js minimales** : ExcelApi 1.7, PowerPointApi 1.4-1.5

---

## Annexe : Fichiers de Référence dans office-agents

Pour implémenter ces features, les fichiers sources les plus utiles à consulter sont :

| Feature | Fichier office-agents | Chemin |
|---------|----------------------|--------|
| Screenshot Excel | screenshot-range.ts | packages/excel/src/lib/tools/ |
| Screenshot PPT | screenshot-slide.ts | packages/powerpoint/src/lib/tools/ |
| Web Search | search.ts | packages/sdk/src/web/ |
| Web Fetch | fetch.ts | packages/sdk/src/web/ |
| CSV Export | get-range-as-csv.ts | packages/excel/src/lib/tools/ |
| Search Pagination | search-data-pagination.ts | packages/excel/src/lib/excel/ |
| Edit Slide XML | edit-slide-xml.ts | packages/powerpoint/src/lib/tools/ |
| Slide ZIP workflow | slide-zip.ts | packages/powerpoint/src/lib/pptx/ |
| XML Utilities | xml-utils.ts | packages/powerpoint/src/lib/pptx/ |
| Verify Slides | verify-slides.ts | packages/powerpoint/src/lib/tools/ |
| Icon Insert | custom-commands.ts | packages/powerpoint/src/lib/vfs/ |
| Master Cleanup | master-cleanup.ts | packages/powerpoint/src/lib/pptx/ |

Le ZIP est disponible dans `office-agents/office-agents.zip` pour consultation directe.
