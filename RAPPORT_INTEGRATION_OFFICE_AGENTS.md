# Rapport d'Intégration : Office Agents → KickOffice

## Table des matières
1. [Analyse critique des 4 opportunités identifiées](#1-analyse-critique-des-4-opportunités-identifiées)
2. [Opportunités supplémentaires découvertes](#2-opportunités-supplémentaires-découvertes)
3. [Plan d'intégration détaillé par fonctionnalité](#3-plan-dintégration-détaillé)
4. [Ordre de priorité recommandé](#4-ordre-de-priorité)

---

## 1. Analyse critique des 4 opportunités identifiées

### 1.1 Screenshot Excel avec En-têtes (Row/Column Headers)

**Verdict : VALIDER - Impact élevé**

**Pourquoi c'est pertinent :**
- Actuellement `screenshotRange` dans KickOffice (`excelTools.ts:1737-1769`) retourne une image brute sans aucun repère de cellule.
- Le modèle de vision reçoit une image PNG mais n'a aucun moyen de savoir que la colonne visible est "B" ou "D". Il fait des approximations qui mènent à des erreurs de ciblage.
- Office Agents l'implémente avec un compositing Canvas qui dessine les lettres de colonnes (A, B, C...) et les numéros de lignes (1, 2, 3...) autour de l'image.
- **Impact mesurable** : réduit drastiquement les erreurs de ciblage de cellules par le modèle vision.

**Challenge / Points d'attention :**
- Le code Canvas nécessite que le navigateur supporte `document.createElement("canvas")`. Dans l'environnement Office Add-in (taskpane WebView), c'est bien supporté.
- Il faut récupérer les `columnWidth` et `rowHeight` pour chaque colonne/ligne du range, ce qui nécessite des `.load()` et `context.sync()` supplémentaires — impact de performance mineur mais acceptable.
- Attention à parser correctement le range A1 pour extraire `startRow` et `startCol` (index 0-based) afin que les en-têtes soient corrects.

---

### 1.2 Feedback d'Erreur Office.js Enrichi (Auto-correction)

**Verdict : VALIDER - Impact élevé**

**Pourquoi c'est pertinent :**
- Actuellement dans KickOffice, les blocs `catch` des outils `eval_officejs` / `eval_wordjs` / `eval_powerpointjs` utilisent `getErrorMessage(err)` qui ne renvoie que `error.message`.
- Les erreurs `OfficeExtension.Error` contiennent des informations de débogage critiques : `debugInfo.errorLocation`, `debugInfo.statement`, `debugInfo.surroundingStatements`.
- Sans ces infos, le LLM ne sait pas **quelle ligne** a planté et fait des corrections "à l'aveugle", gaspillant des tours d'agent.

**Challenge / Points d'attention :**
- `OfficeExtension` n'est pas toujours typé globalement. Il faut utiliser du duck-typing : `err && typeof err === 'object' && 'debugInfo' in err`.
- La modification est simple et localisée : enrichir `getErrorMessage()` dans `common.ts` ou créer une nouvelle fonction `getOfficeErrorMessage()`.
- Ne pas casser le typage existant — `getErrorMessage` est utilisé partout, il vaut mieux créer une fonction spécialisée `getDetailedOfficeError()` et l'appeler uniquement dans les blocs catch de `eval_*`.

---

### 1.3 Tracker de Mutations ("Dirty Tracker" statique)

**Verdict : VALIDER PARTIELLEMENT - Impact moyen**

**Ce qu'il faut intégrer :**
- La version **regex statique** (pattern matching) : scanner le code pour détecter les mutations (`looksLikeMutation()`). C'est simple, sans effet de bord, et immédiatement utile.
- Renvoyer `hasMutated: true/false` dans la réponse JSON des outils `eval_*`.

**Ce qu'il NE faut PAS intégrer :**
- La version **Proxy-based** complète d'Office Agents (`tracked-context.ts`) qui intercepte toutes les mutations via des Proxies ES6. Raisons :
  - KickOffice utilise SES (Secure ECMAScript) qui **bloque explicitement Proxy** dans la sandbox (`Proxy: undefined` dans `sandbox.ts:43`). Intégrer les Proxies nécessiterait de désactiver une protection de sécurité.
  - La complexité du `tracked-context.ts` (~400 lignes, gestion asynchrone des sheet IDs, résolution post-sync) est disproportionnée pour le bénéfice dans le contexte KickOffice.
  - Le dirty tracking par Proxy sert principalement à mettre à jour l'UI (highlights de cellules modifiées), ce qui n'est pas un besoin exprimé dans KickOffice.

**La version regex suffit** pour informer l'agent qu'il a bien modifié le document.

---

### 1.4 Injection VFS dans la Sandbox

**Verdict : VALIDER - Impact moyen-élevé**

**Pourquoi c'est pertinent :**
- Actuellement, le code exécuté dans `eval_officejs` n'a accès qu'à `context`, `Excel`, et `Office`. Il ne peut pas :
  - Encoder/décoder en Base64 (`btoa`/`atob` sont bloqués par SES)
  - Lire un fichier uploadé par l'utilisateur dans le VFS
  - Écrire un résultat intermédiaire dans le VFS
- Office Agents expose `readFile()`, `readFileBuffer()`, `writeFile()`, `btoa`, `atob` dans le contexte d'exécution.
- **Cas d'usage concret** : l'utilisateur uploade une image PNG, le LLM génère un script qui la lit depuis le VFS et l'insère dans Excel/Word/PowerPoint via le code Office.js.

**Challenge / Points d'attention :**
- `btoa` et `atob` sont des APIs navigateur qui fonctionnent dans le taskpane WebView mais sont bloquées par SES. Il faut les passer explicitement dans les `globals` du Compartment.
- Pour les fonctions VFS (`readFile`, `writeFile`), il faut les wrapper et les passer comme globals au `sandboxedEval()` depuis chaque outil `eval_*`.
- Attention à ne PAS exposer `fetch`, `XMLHttpRequest`, ou d'autres APIs réseau — uniquement les helpers VFS locaux.

---

## 2. Opportunités supplémentaires découvertes

Après analyse approfondie du code d'Office Agents, voici les fonctionnalités supplémentaires pertinentes à intégrer dans KickOffice :

### 2.1 Commandes Bash personnalisées : `csv-to-sheet` et `sheet-to-csv`

**Impact : ÉLEVÉ**

**Ce qu'Office Agents implémente :**
- `csv-to-sheet <file> <sheetId> [startCell] [--force]` : importe un CSV depuis le VFS directement dans une feuille Excel avec :
  - Détection de types (booléens, nombres, texte)
  - Padding automatique des lignes/colonnes
  - Protection contre l'écrasement (sauf `--force`)
- `sheet-to-csv <sheetId> [range] [file]` : exporte une feuille/range en CSV dans le VFS

**Pourquoi l'intégrer :**
- KickOffice a déjà `getRangeAsCsv` pour l'export, mais pas d'import CSV.
- L'import CSV est un workflow courant : l'utilisateur uploade un CSV, l'agent doit l'insérer dans Excel. Aujourd'hui, le LLM doit générer un script complexe via `eval_officejs` pour parser le CSV et écrire les cellules — sujet à erreurs.
- Avoir une commande bash native simplifie drastiquement ce workflow.

**Ce qu'il faut adapter :**
- Dans KickOffice, le VFS existe déjà (`vfs.ts`). Il faut enregistrer des commandes bash custom dans le `Bash` de `just-bash`.
- L'accès à l'API Excel (`Excel.run()`) depuis une commande bash est le point technique à résoudre : la commande bash doit pouvoir déclencher une action Office.js.

---

### 2.2 Commande `image-to-sheet` (Pixel Art Excel)

**Impact : MOYEN (mais impressionnant pour les démos et utilisateurs créatifs)**

**Ce qu'Office Agents implémente :**
- Convertit une image (PNG/JPG) en "pixel art" dans Excel : chaque cellule reçoit la couleur d'un pixel.
- Downsampling intelligent, run-length encoding pour les cellules adjacentes de même couleur.
- Batching des opérations (1000 ranges par batch) pour éviter les timeouts.
- Max 200x200 pixels, taille de cellule configurable (1-50 points).

**Pourquoi l'intégrer :**
- Effet "wow" pour les démos et les utilisateurs créatifs.
- Montre la puissance de l'intégration LLM+Office.
- Utilise déjà le VFS et Canvas (mêmes briques que le screenshot avec headers).

**Attention :** C'est une feature "fun" plus qu'utilitaire. À prioriser après les features critiques.

---

### 2.3 Extraction OOXML avec structure (`get_ooxml` pour Word)

**Impact : ÉLEVÉ pour Word**

**Ce qu'Office Agents implémente :**
- Un outil `get_ooxml` qui extrait le XML Office Open XML du document Word avec :
  - Un résumé structuré des enfants du `<w:body>` (type, index, numéro de ligne)
  - Mapping vers les index Office.js (paragraphIndex, tableIndex)
  - Nettoyage des attributs RSID (bruit)
  - Inclusion des styles et définitions de numérotation référencées
  - Écriture dans le VFS pour inspection ultérieure

**Pourquoi c'est intéressant :**
- KickOffice a déjà `editDocumentXml` pour Word mais PAS d'outil pour **lire** le OOXML de façon structurée.
- Pour les documents complexes (avec mise en forme directe, listes multi-niveaux, champs), le OOXML est le seul moyen de comprendre la structure réelle.
- L'agent peut ensuite utiliser `editDocumentXml` pour faire des modifications ciblées basées sur l'analyse OOXML.

**Remarque :** KickOffice a déjà `getDocumentHtml` et `getDocumentContent`. L'outil OOXML serait complémentaire pour les cas avancés (documents juridiques, templates, mise en forme complexe).

---

### 2.4 Recherche de données avec pagination (`search_data` amélioré)

**Impact : MOYEN-ÉLEVÉ pour Excel**

**Ce qu'Office Agents implémente :**
- Recherche dans les données avec support de :
  - Expressions régulières
  - Recherche case-sensitive/insensitive
  - Match cellule entière vs partiel
  - Recherche dans les formules (pas juste les valeurs)
  - **Pagination** avec `offset` pour les grands jeux de données
  - Limite de résultats configurable

**Comparaison avec KickOffice :**
- KickOffice a `findData` qui fait une recherche basique via `worksheet.findAllOrNullObject()`.
- Il manque : la regex, la recherche dans les formules, et surtout la pagination.

**Pourquoi l'intégrer :**
- Pour les grands tableaux (10k+ lignes), l'agent a besoin de paginer les résultats de recherche.
- La recherche dans les formules est critique pour le débogage de spreadsheets.

---

### 2.5 Document Metadata enrichi (Word)

**Impact : MOYEN**

**Ce qu'Office Agents implémente via `getDocumentMetadata()` :**
- Détection de run-level overrides (mise en forme directe vs styles)
- Échantillonnage des 20 premiers paragraphes pour détecter le pattern de formatage
- Flag `hasRunLevelOverrides` qui change la stratégie d'édition de l'agent
- Mode de suivi des modifications (tracking mode)
- Informations de style (font/taille/couleur pour les styles clés)

**Pourquoi c'est intéressant :**
- KickOffice a `officeDocumentContext.ts` qui injecte un contexte document basique.
- Enrichir ce contexte avec la détection de run-level overrides permettrait à l'agent de choisir automatiquement entre édition standard et OOXML.
- **À considérer comme une amélioration future** plutôt qu'une intégration immédiate.

---

### 2.6 Résumé : Fonctionnalités analysées et NON retenues

| Fonctionnalité | Raison du rejet |
|---|---|
| Web Fetch / Web Search | Exclu explicitement par le demandeur |
| Bring Your Own Key | Exclu explicitement par le demandeur |
| Dirty Tracking par Proxy | Incompatible avec SES (Proxy bloqué), complexité disproportionnée |
| Screenshot document Word (PDF→PNG) | Nécessite PDF.js (~300KB), desktop only, KickOffice Word a déjà getDocumentHtml |
| Slide ZIP manipulation (JSZip) | KickOffice a déjà `editSlideXml` — architecture différente mais même résultat |
| Slide Master editing | Cas d'usage très avancé, faible demande |
| Tracked Changes UI (indicator) | Feature UI React vs Vue, non prioritaire |
| Sheet ID mapping stable | Nécessaire seulement pour le dirty tracking Proxy |

---

## 3. Plan d'intégration détaillé

### FEATURE 1 : Screenshot Excel avec En-têtes

**Fichiers à modifier :** `frontend/src/utils/excelTools.ts`

**Étape 1 — Ajouter les constantes et utilitaires Canvas en haut du fichier :**

```typescript
// ============================================================
// Screenshot headers composition — ported from Office Agents
// ============================================================
const HEADER_WIDTH = 40;
const HEADER_HEIGHT = 20;
const HEADER_BG = '#f0f0f0';
const HEADER_BORDER = '#c0c0c0';
const HEADER_FONT = 'bold 11px Calibri, Arial, sans-serif';
const HEADER_TEXT_COLOR = '#333333';

/**
 * Convert 0-based column index to Excel column letter (0→A, 1→B, 26→AA, etc.)
 * Ported from Office Agents screenshot-range.ts
 */
function columnIndexToLetter(index: number): string {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Parse start row and column from an A1-style range string.
 * Examples: "B3:D10" → { startRow: 2, startCol: 1 }, "A1" → { startRow: 0, startCol: 0 }
 */
function parseRangeStart(rangeAddress: string): { startRow: number; startCol: number } {
  // rangeAddress may look like "Sheet1!B3:D10" or "B3:D10" or "B3"
  const addr = rangeAddress.includes('!') ? rangeAddress.split('!')[1] : rangeAddress;
  const startCell = addr.split(':')[0];
  const match = startCell.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { startRow: 0, startCol: 0 };

  const colStr = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1; // 0-based

  const row = parseInt(match[2], 10) - 1; // 0-based
  return { startRow: row, startCol: col };
}

/**
 * Composite an Excel range screenshot with row/column headers using Canvas.
 * Draws column letters (A, B, C...) on top and row numbers (1, 2, 3...) on the left.
 * Ported from Office Agents screenshot-range.ts
 *
 * @param imageBase64 - The original range image (base64 PNG without data URI prefix)
 * @param startRow - 0-based row index of the first row in the range
 * @param startCol - 0-based column index of the first column in the range
 * @param colWidths - Array of column widths (in points) for each column in the range
 * @param rowHeights - Array of row heights (in points) for each row in the range
 * @returns Promise<string> - New base64 PNG string (without data URI prefix)
 */
function compositeWithHeaders(
  imageBase64: string,
  startRow: number,
  startCol: number,
  colWidths: number[],
  rowHeights: number[],
): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const totalColWidth = colWidths.reduce((a, b) => a + b, 0);
      const totalRowHeight = rowHeights.reduce((a, b) => a + b, 0);
      const scaleX = totalColWidth > 0 ? img.width / totalColWidth : 1;
      const scaleY = totalRowHeight > 0 ? img.height / totalRowHeight : 1;

      const canvas = document.createElement('canvas');
      canvas.width = HEADER_WIDTH + img.width;
      canvas.height = HEADER_HEIGHT + img.height;
      const ctx = canvas.getContext('2d');
      if (!ctx) return reject(new Error('Failed to get 2d canvas context'));

      // White background
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Draw the original image offset by headers
      ctx.drawImage(img, HEADER_WIDTH, HEADER_HEIGHT);

      // --- Column headers ---
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(HEADER_WIDTH, 0, img.width, HEADER_HEIGHT);
      ctx.font = HEADER_FONT;
      ctx.fillStyle = HEADER_TEXT_COLOR;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';

      let x = HEADER_WIDTH;
      for (let i = 0; i < colWidths.length; i++) {
        const w = colWidths[i] * scaleX;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.strokeRect(x, 0, w, HEADER_HEIGHT);
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(columnIndexToLetter(startCol + i), x + w / 2, HEADER_HEIGHT / 2);
        x += w;
      }

      // --- Row headers ---
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, HEADER_HEIGHT, HEADER_WIDTH, img.height);
      let y = HEADER_HEIGHT;
      for (let i = 0; i < rowHeights.length; i++) {
        const h = rowHeights[i] * scaleY;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.strokeRect(0, y, HEADER_WIDTH, h);
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(String(startRow + i + 1), HEADER_WIDTH / 2, y + h / 2);
        y += h;
      }

      // Top-left corner (empty cell)
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
      ctx.strokeStyle = HEADER_BORDER;
      ctx.strokeRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);

      resolve(canvas.toDataURL('image/png').split(',')[1]);
    };
    img.onerror = () => reject(new Error('Failed to load range image for header composition'));
    img.src = `data:image/png;base64,${imageBase64}`;
  });
}
```

**Étape 2 — Modifier l'outil `screenshotRange` (remplacer le contenu de `executeExcel`) :**

Remplacer le code actuel de `executeExcel` dans `screenshotRange` (lignes ~1756-1768) par :

```typescript
executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
  const sheet = await safeGetSheet(context, args.sheetName);
  const targetRange = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();

  // Load range address and dimensions to determine row/column count
  targetRange.load('address, rowCount, columnCount');
  await context.sync();

  const numCols = targetRange.columnCount;
  const numRows = targetRange.rowCount;

  // Load column widths and row heights for header composition
  const cols: Excel.Range[] = [];
  for (let i = 0; i < numCols; i++) {
    const col = targetRange.getColumn(i);
    col.format.load('columnWidth');
    cols.push(col);
  }
  const rows: Excel.Range[] = [];
  for (let i = 0; i < numRows; i++) {
    const row = targetRange.getRow(i);
    row.format.load('rowHeight');
    rows.push(row);
  }

  // Get the raw image
  const imageResult = (targetRange as any).getImage();
  await context.sync();

  const base64 = imageResult.value as string;
  const colWidths = cols.map(c => c.format.columnWidth);
  const rowHeights = rows.map(r => r.format.rowHeight);

  // Parse the range address to get start row/col for header labels
  const { startRow, startCol } = parseRangeStart(targetRange.address);

  // Composite the image with headers
  const composited = await compositeWithHeaders(
    base64,
    startRow,
    startCol,
    colWidths,
    rowHeights,
  );

  return JSON.stringify({
    __screenshot__: true,
    base64: composited,
    mimeType: 'image/png',
    description: `Screenshot of range ${args.range || 'used range'} on sheet ${args.sheetName || 'active'} (with row/column headers)`,
  });
},
```

**Points critiques pour l'implémenteur :**
- `targetRange.getColumn(i)` et `targetRange.getRow(i)` sont disponibles dans ExcelApi 1.1+.
- `format.columnWidth` et `format.rowHeight` sont en points (pas en pixels). Le scaling se fait automatiquement via `scaleX`/`scaleY` dans `compositeWithHeaders`.
- `(targetRange as any).getImage()` nécessite ExcelApi 1.7+ (comme c'est déjà le cas).
- Le `parseRangeStart` gère le format `Sheet1!B3:D10` que retourne `.address`.

---

### FEATURE 2 : Feedback d'Erreur Office.js Enrichi

**Fichiers à modifier :** `frontend/src/utils/common.ts`, `frontend/src/utils/excelTools.ts`, `frontend/src/utils/wordTools.ts`, `frontend/src/utils/powerpointTools.ts`

**Étape 1 — Ajouter une nouvelle fonction dans `common.ts` (après `getErrorMessage`) :**

```typescript
/**
 * Extract detailed error information from Office.js OfficeExtension.Error objects.
 * Provides the LLM with error location, failing statement, and surrounding context
 * for better auto-correction.
 * Ported from Office Agents error handling pattern.
 *
 * Falls back to getErrorMessage() for non-Office errors.
 */
export function getDetailedOfficeError(error: unknown): string {
  // Duck-type check for OfficeExtension.Error (may not be typed globally)
  if (
    error &&
    typeof error === 'object' &&
    'debugInfo' in error &&
    'message' in error
  ) {
    const officeError = error as {
      message: string;
      code?: string;
      debugInfo?: {
        errorLocation?: string;
        statement?: string;
        surroundingStatements?: string[];
      };
    };

    const parts = [officeError.message];

    if (officeError.code) {
      parts.push(`Code: ${officeError.code}`);
    }

    if (officeError.debugInfo) {
      const { errorLocation, statement, surroundingStatements } = officeError.debugInfo;
      if (errorLocation) {
        parts.push(`Location: ${errorLocation}`);
      }
      if (statement) {
        parts.push(`Failing statement: ${statement}`);
      }
      if (surroundingStatements?.length) {
        parts.push(`Surrounding context: ${surroundingStatements.join('; ')}`);
      }
    }

    return parts.join('\n');
  }

  // Fallback to standard error extraction
  return getErrorMessage(error);
}
```

**Étape 2 — Modifier les blocs catch des outils `eval_*` dans chaque fichier tools :**

Dans `excelTools.ts`, dans le catch de `eval_officejs` (actuellement lignes ~1722-1734), remplacer :

```typescript
// AVANT :
error: getErrorMessage(err),

// APRÈS :
error: getDetailedOfficeError(err),
```

Faire la même chose dans :
- `wordTools.ts` — bloc catch de `eval_wordjs`
- `powerpointTools.ts` — bloc catch de `eval_powerpointjs`
- `outlookTools.ts` — bloc catch de `eval_outlookjs` (si existant)

**Important :** Il faut aussi ajouter `getDetailedOfficeError` à l'import depuis `./common` dans chaque fichier.

Le diff complet pour `excelTools.ts` (catch block) :

```typescript
// Fichier: excelTools.ts — dans le catch de eval_officejs
} catch (err: unknown) {
  return JSON.stringify(
    {
      success: false,
      error: getDetailedOfficeError(err),  // ← CHANGEMENT ICI
      explanation,
      codeExecuted: code.slice(0, 200) + '...',
      hint: 'Check that all properties are loaded before access, and context.sync() is called.',
    },
    null,
    2,
  );
}
```

---

### FEATURE 3 : Mutation Tracker statique (Regex)

**Fichiers à modifier :** `frontend/src/utils/excelTools.ts`, `frontend/src/utils/wordTools.ts`, `frontend/src/utils/powerpointTools.ts`

**Étape 1 — Ajouter les patterns de détection de mutation dans chaque fichier :**

Les patterns sont différents par host. Voici les patterns recommandés :

**Pour Excel** (dans `excelTools.ts`, en haut du fichier après les imports) :

```typescript
// ============================================================
// Mutation detection patterns — ported from Office Agents dirty-tracker
// Used by eval_officejs to flag write operations for the agent.
// ============================================================
const EXCEL_MUTATION_PATTERNS = [
  /\.(values|formulas|formulasLocal|numberFormat|numberFormatLocal)\s*=/,
  /\.clear\s*\(/,
  /\.delete\s*\(/,
  /\.insert\s*\(/,
  /\.copyFrom\s*\(/,
  /\.add\s*\(/,
  /\.merge\s*\(/,
  /\.unmerge\s*\(/,
  /\.format\.\w+\s*=/,     // format.fill.color = ..., format.font.bold = ...
  /\.set\s*\(/,
];

function looksLikeMutation(code: string): boolean {
  return EXCEL_MUTATION_PATTERNS.some((p) => p.test(code));
}
```

**Pour Word** (dans `wordTools.ts`) :

```typescript
const WORD_MUTATION_PATTERNS = [
  /\.insertText\s*\(/,
  /\.insertHtml\s*\(/,
  /\.insertOoxml\s*\(/,
  /\.insertParagraph\s*\(/,
  /\.insertBreak\s*\(/,
  /\.insertTable\s*\(/,
  /\.insertInlinePictureFromBase64\s*\(/,
  /\.delete\s*\(/,
  /\.clear\s*\(/,
  /\.font\.\w+\s*=/,
  /\.style\s*=/,
  /\.styleBuiltIn\s*=/,
  /\.alignment\s*=/,
  /\.set\s*\(/,
];

function looksLikeMutation(code: string): boolean {
  return WORD_MUTATION_PATTERNS.some((p) => p.test(code));
}
```

**Pour PowerPoint** (dans `powerpointTools.ts`) :

```typescript
const PPT_MUTATION_PATTERNS = [
  /\.insertSlide\s*\(/,
  /\.delete\s*\(/,
  /\.text\s*=/,
  /\.insertText\s*\(/,
  /\.font\.\w+\s*=/,
  /\.fill\.\w+\s*=/,
  /\.set\s*\(/,
  /\.add\s*\(/,
];

function looksLikeMutation(code: string): boolean {
  return PPT_MUTATION_PATTERNS.some((p) => p.test(code));
}
```

**Étape 2 — Modifier le retour JSON des outils `eval_*` pour inclure `hasMutated` :**

Dans chaque fichier, dans le bloc `try` de l'outil `eval_*` (après l'exécution réussie), ajouter le flag :

```typescript
// DANS le try block de eval_officejs (excelTools.ts), AVANT le return :
const hasMutated = looksLikeMutation(code);

return JSON.stringify(
  {
    success: true,
    result: result ?? null,
    explanation,
    hasMutated,  // ← AJOUT
    warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
  },
  null,
  2,
);
```

Appliquer le même pattern pour `eval_wordjs` et `eval_powerpointjs`.

---

### FEATURE 4 : Injection VFS + btoa/atob dans la Sandbox

**Fichiers à modifier :** `frontend/src/utils/sandbox.ts`, `frontend/src/utils/excelTools.ts`, `frontend/src/utils/wordTools.ts`, `frontend/src/utils/powerpointTools.ts`

**Étape 1 — Exposer `btoa` et `atob` dans le Compartment SES (`sandbox.ts`) :**

Modifier le Compartment dans `sandboxedEval()` pour ajouter `btoa` et `atob` :

```typescript
// Dans sandbox.ts, modifier le Compartment (ligne ~26-54) :
// @ts-ignore - Compartment is provided by SES
const compartment = new Compartment({
  globals: {
    ...filteredGlobals,
    // Safe built-ins
    console,
    Math,
    Date,
    JSON,
    Array,
    Object,
    String,
    Number,
    Boolean,
    Promise,
    // Base64 utilities — needed for image insertion and binary data handling
    // Ported from Office Agents VFS integration
    btoa: typeof btoa !== 'undefined' ? btoa : undefined,
    atob: typeof atob !== 'undefined' ? atob : undefined,
    // Blocked APIs
    Function: undefined,
    Reflect: undefined,
    Proxy: undefined,
    Compartment: undefined,
    harden: undefined,
    lockdown: undefined,
    eval: undefined,
    fetch: undefined,
    XMLHttpRequest: undefined,
    WebSocket: undefined,
  },
  __options__: true,
});
```

**Étape 2 — Passer les helpers VFS comme globals depuis les outils `eval_*` :**

Dans `excelTools.ts`, modifier l'appel à `sandboxedEval` dans `eval_officejs` :

```typescript
// EN HAUT DU FICHIER, ajouter l'import VFS :
import { readFile as vfsReadFile, writeFile as vfsWriteFile } from '@/utils/vfs';

// DANS eval_officejs, modifier l'appel sandboxedEval :
const result = await sandboxedEval(
  code,
  {
    context,
    Excel: typeof Excel !== 'undefined' ? Excel : undefined,
    Office: typeof Office !== 'undefined' ? Office : undefined,
    // VFS helpers — ported from Office Agents
    // Allows LLM-generated code to read uploaded files and write results
    readFile: vfsReadFile,
    readFileBuffer: async (path: string) => {
      const vfs = (await import('@/utils/vfs')).getVfs();
      const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;
      return vfs.readFileBuffer(fullPath);
    },
    writeFile: vfsWriteFile,
  },
  'Excel',
);
```

Faire la même chose pour `eval_wordjs` (avec `Word` au lieu de `Excel`) et `eval_powerpointjs` (avec `PowerPoint`).

**Exemple pour `wordTools.ts` :**

```typescript
import { readFile as vfsReadFile, writeFile as vfsWriteFile } from '@/utils/vfs';

// Dans eval_wordjs :
const result = await sandboxedEval(
  code,
  {
    context,
    Word: typeof Word !== 'undefined' ? Word : undefined,
    Office: typeof Office !== 'undefined' ? Office : undefined,
    readFile: vfsReadFile,
    readFileBuffer: async (path: string) => {
      const vfs = (await import('@/utils/vfs')).getVfs();
      const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;
      return vfs.readFileBuffer(fullPath);
    },
    writeFile: vfsWriteFile,
  },
  'Word',
);
```

**Points critiques :**
- NE PAS ajouter `fetch`, `XMLHttpRequest`, ou toute API réseau.
- `readFileBuffer` retourne `Uint8Array` — utile pour lire des images binaires.
- `writeFile` accepte `string | Uint8Array`.
- `readFile` retourne `string` — utile pour lire des CSV, JSON, XML.

---

### FEATURE 5 (bonus) : Commande Bash `csv-to-sheet`

**Fichiers à créer/modifier :** `frontend/src/utils/excelCustomCommands.ts` (nouveau), `frontend/src/utils/excelTools.ts`

**Cette feature est plus complexe car elle nécessite un pont entre le VFS bash et l'API Office.js.**

**Approche recommandée :** Plutôt qu'une commande bash native (complexe à intégrer avec Office.js), créer un **nouvel outil dédié** `importCsvToSheet` dans `excelTools.ts`.

```typescript
// Nouveau tool dans excelToolDefinitions :
importCsvToSheet: {
  name: 'importCsvToSheet',
  category: 'write',
  description:
    'Import a CSV file from the VFS into an Excel worksheet. Reads the CSV from the virtual filesystem and writes the data to the specified sheet and starting cell. Automatically detects data types (numbers, booleans, text).',
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: 'Path to CSV file in VFS (e.g., "/home/user/uploads/data.csv")',
      },
      sheetName: {
        type: 'string',
        description: 'Target worksheet name. Uses active sheet if omitted.',
      },
      startCell: {
        type: 'string',
        description: 'Starting cell in A1 notation (e.g., "A1"). Defaults to "A1".',
      },
      delimiter: {
        type: 'string',
        description: 'CSV delimiter character. Defaults to ",".',
      },
      overwrite: {
        type: 'boolean',
        description: 'If true, overwrites existing cell data. Default is false (fails if cells contain data).',
      },
    },
    required: ['filePath'],
  },
  executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
    const {
      filePath,
      sheetName,
      startCell = 'A1',
      delimiter = ',',
      overwrite = false,
    } = args;

    // Read CSV from VFS
    const { readFile } = await import('@/utils/vfs');
    const csvContent = await readFile(filePath);
    if (!csvContent || !csvContent.trim()) {
      throw new Error(`File "${filePath}" is empty or not found.`);
    }

    // Parse CSV (basic parser handling quoted fields)
    const lines = csvContent.trim().split('\n');
    const data: (string | number | boolean)[][] = [];

    for (const line of lines) {
      const row: (string | number | boolean)[] = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
          inQuotes = !inQuotes;
        } else if (char === delimiter && !inQuotes) {
          row.push(coerceValue(current.trim()));
          current = '';
        } else {
          current += char;
        }
      }
      row.push(coerceValue(current.trim()));
      data.push(row);
    }

    if (data.length === 0) {
      throw new Error('CSV file contains no data rows.');
    }

    // Pad rows to equal length
    const maxCols = Math.max(...data.map(r => r.length));
    for (const row of data) {
      while (row.length < maxCols) row.push('');
    }

    const sheet = await safeGetSheet(context, sheetName);
    const targetRange = sheet.getRange(startCell).getResizedRange(data.length - 1, maxCols - 1);

    if (!overwrite) {
      targetRange.load('values');
      await context.sync();
      const hasData = targetRange.values.some((row: any[]) =>
        row.some((cell: any) => cell !== '' && cell !== null),
      );
      if (hasData) {
        throw new Error(
          'Target range contains existing data. Use overwrite=true to replace, or choose a different startCell.',
        );
      }
    }

    targetRange.values = data as any[][];
    await context.sync();

    return JSON.stringify({
      success: true,
      rowsImported: data.length,
      columnsImported: maxCols,
      range: `${startCell} to ${columnIndexToLetter(maxCols - 1)}${data.length}`,
      hasMutated: true,
    });
  },
},
```

**Fonction helper pour la coercion de types (à ajouter en haut de `excelTools.ts`) :**

```typescript
/**
 * Coerce a CSV string value to its native type.
 * Ported from Office Agents csv-to-sheet custom command.
 */
function coerceValue(value: string): string | number | boolean {
  if (value === '') return '';
  const lower = value.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;
  const num = Number(value);
  if (!isNaN(num) && value.trim() !== '') return num;
  return value;
}
```

**N'oubliez pas** d'ajouter `'importCsvToSheet'` au type `ExcelToolName`.

---

### FEATURE 6 (bonus) : Recherche avec pagination et regex

**Fichier à modifier :** `frontend/src/utils/excelTools.ts` — outil `findData`

**Amélioration de l'outil existant `findData` :**

Ajouter les paramètres suivants au `inputSchema` de `findData` :
- `useRegex` (boolean) : activer la recherche par regex
- `searchInFormulas` (boolean) : chercher dans les formules au lieu des valeurs
- `offset` (number) : pagination — sauter les N premiers résultats
- `limit` (number) : nombre max de résultats (défaut 50)

Le code de `executeExcel` de `findData` doit être modifié pour :

```typescript
executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
  const {
    searchText,
    sheetName,
    range: rangeAddress,
    matchCase = false,
    matchEntireCell = false,
    useRegex = false,
    searchInFormulas = false,
    offset = 0,
    limit = 50,
  } = args;

  const sheet = await safeGetSheet(context, sheetName);
  const searchRange = rangeAddress ? sheet.getRange(rangeAddress) : sheet.getUsedRange();

  if (useRegex || searchInFormulas) {
    // Manual search: iterate over values/formulas and apply regex or formula search
    const propertyToLoad = searchInFormulas ? 'formulas' : 'values';
    searchRange.load(`address, ${propertyToLoad}, rowCount, columnCount`);
    await context.sync();

    const data = searchInFormulas ? searchRange.formulas : searchRange.values;
    const regex = useRegex ? new RegExp(searchText, matchCase ? 'g' : 'gi') : null;
    const results: { cell: string; value: any }[] = [];
    const { startRow, startCol } = parseRangeStart(searchRange.address);

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cellValue = String(data[r][c]);
        let isMatch = false;

        if (regex) {
          regex.lastIndex = 0;
          isMatch = regex.test(cellValue);
        } else {
          const compareValue = matchCase ? cellValue : cellValue.toLowerCase();
          const compareSearch = matchCase ? searchText : searchText.toLowerCase();
          isMatch = matchEntireCell ? compareValue === compareSearch : compareValue.includes(compareSearch);
        }

        if (isMatch) {
          results.push({
            cell: `${columnIndexToLetter(startCol + c)}${startRow + r + 1}`,
            value: data[r][c],
          });
        }
      }
    }

    const paginated = results.slice(offset, offset + limit);
    return JSON.stringify({
      totalMatches: results.length,
      returned: paginated.length,
      offset,
      hasMore: offset + limit < results.length,
      matches: paginated,
    }, null, 2);
  }

  // Default: use built-in findAllOrNullObject (fast but limited)
  // ... keep existing implementation with pagination wrapper ...
}
```

**N'oubliez pas** d'ajouter les nouveaux paramètres au `inputSchema.properties` et de mettre à jour la `description` de l'outil.

---

## 4. Ordre de priorité recommandé

| Priorité | Feature | Effort | Impact | Justification |
|----------|---------|--------|--------|---------------|
| **P0** | Feature 2 : Erreur enrichie | 30 min | Élevé | Modification minimale, gain immédiat sur l'auto-correction |
| **P0** | Feature 3 : Mutation tracker regex | 30 min | Moyen | Simple regex, pas de risque, informe l'agent |
| **P1** | Feature 4 : VFS + btoa/atob sandbox | 1h | Moyen-élevé | Débloque les workflows images/binaires |
| **P1** | Feature 1 : Screenshot avec headers | 1h30 | Élevé | Améliore fortement la vision, plus de code |
| **P2** | Feature 5 : Import CSV | 2h | Moyen-élevé | Nouvel outil, plus de tests nécessaires |
| **P3** | Feature 6 : Recherche paginée/regex | 2h | Moyen | Amélioration d'un outil existant |

---

## Annexe : Récapitulatif des fichiers à modifier

| Fichier | Features concernées |
|---------|-------------------|
| `frontend/src/utils/common.ts` | Feature 2 (ajout `getDetailedOfficeError`) |
| `frontend/src/utils/sandbox.ts` | Feature 4 (ajout `btoa`/`atob`) |
| `frontend/src/utils/excelTools.ts` | Features 1, 2, 3, 4, 5, 6 |
| `frontend/src/utils/wordTools.ts` | Features 2, 3, 4 |
| `frontend/src/utils/powerpointTools.ts` | Features 2, 3, 4 |
| `frontend/src/utils/outlookTools.ts` | Feature 2 (si eval_outlookjs existe) |
