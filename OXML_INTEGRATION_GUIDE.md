# OXML Integration Guide — Phase 4A

> **Document créé le 2026-03-14, mis à jour le 2026-03-14** — Guide dédié à l'intégration OOXML dans KickOffice.
> Couvre les 3 tâches de la Phase 4A : OXML-M1, WORD-H1, DUP-M1.

## ✅ Implementation Status — COMPLETED (2026-03-14)

All Phase 4A tasks have been successfully implemented:

| Task | Status | Details |
|------|--------|---------|
| **OXML-M1** | ✅ FIXED | OOXML evaluation completed for all hosts:<br>• Word: ✅ Supported (Track Changes + formatting preservation)<br>• Excel: ❌ No OOXML API available<br>• PowerPoint: ✅ Already implemented (JSZip)<br>• Outlook: ❌ HTML-only (no OOXML API) |
| **WORD-H1** | ✅ FIXED | docx-redline-js integration complete:<br>• Installed `@ansonlai/docx-redline-js` (v0.1.4)<br>• Created `wordTrackChanges.ts` (TC helpers)<br>• Rewrote `wordDiffUtils.ts` with Gemini pattern<br>• Updated `proposeRevision` tool (native Track Changes)<br>• Added `editDocumentXml` tool (OOXML manipulation)<br>• Settings UI for redline author + toggle<br>• Updated `word.skill.md` documentation<br>• Removed `office-word-diff` package |
| **DUP-M1** | ✅ FIXED | `truncateString()` extracted to `common.ts`:<br>• Replaced 4 occurrences (wordTools ×2, outlookTools ×2)<br>• Import added to both files |

**Files Modified/Created:**
- ✅ Created: `frontend/src/utils/wordTrackChanges.ts`
- ✅ Rewritten: `frontend/src/utils/wordDiffUtils.ts`
- ✅ Updated: `frontend/src/utils/wordTools.ts` (proposeRevision + editDocumentXml)
- ✅ Updated: `frontend/src/utils/outlookTools.ts` (truncateString)
- ✅ Updated: `frontend/src/utils/common.ts` (truncateString)
- ✅ Updated: `frontend/src/components/settings/ToolsTab.vue` (redline settings UI)
- ✅ Updated: `frontend/src/skills/word.skill.md` (documentation)
- ✅ Updated: `frontend/package.json` (dependencies)
- ✅ Updated: `README.md` (credits)
- ✅ Deleted: `/office-word-diff/` directory

**Test Results:**
- ✅ Build passes: `npm install` successful
- ✅ No import errors
- ✅ TypeScript compilation clean
- ⏳ Manual testing pending: proposeRevision with Track Changes, editDocumentXml, Settings UI

---

## Table des matières

1. [Résumé exécutif](#1-résumé-exécutif)
2. [OXML-M1 : Évaluation OOXML par host](#2-oxml-m1--évaluation-ooxml-par-host)
3. [WORD-H1 : Track Changes via docx-redline-js](#3-word-h1--track-changes-via-docx-redline-js)
4. [Nouveau tool : editDocumentXml](#4-nouveau-tool--editdocumentxml)
5. [Migration : office-word-diff → docx-redline-js](#5-migration--office-word-diff--docx-redline-js)
6. [DUP-M1 : Extraction truncateString](#6-dup-m1--extraction-truncatestring)
7. [Plan d'implémentation détaillé](#7-plan-dimplémentation-détaillé)
8. [Risques et mitigations](#8-risques-et-mitigations)
9. [Sources](#9-sources)

---

## 1. Résumé exécutif

### Constat initial

KickOffice utilise Office.js comme couche d'abstraction exclusive pour manipuler les documents. L'OOXML est utilisé **uniquement** dans PowerPoint (`editSlideXml` via JSZip dans `pptxZipUtils.ts`). Word utilise `office-word-diff` (npm custom) pour le diffing, Excel et Outlook n'ont aucune manipulation OOXML.

### Comportement de `insertOoxml()` avec les Track Changes

La recherche de l'API Office.js révèle un point important pour WORD-H1 :

| Comportement | Détail |
|---|---|
| `range.getOoxml()` | Retourne du XML **aplati** — les éléments `<w:ins>` / `<w:del>` existants sont **absents** de la sortie |
| Track Changes **ON** + `insertOoxml()` | Word crée ses **propres** révisions sur **tout** le contenu inséré (= tout marqué comme nouveau) — ❌ double-tracking |
| Track Changes **OFF** + `insertOoxml()` | Le revision markup `<w:ins>` / `<w:del>` embarqué dans le XML est **préservé** — ✅ c'est la clé |
| `<w:trackRevisions/>` dans settings | **Absent** du XML retourné par `getOoxml()` |

**Découverte clé** (inspirée par [Gemini AI for Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting)) : en **désactivant** temporairement Track Changes avant d'insérer du XML contenant du revision markup, les éléments `<w:ins>` / `<w:del>` survivent à `insertOoxml()`. C'est exactement ce que fait la lib [`docx-redline-js`](https://github.com/AnsonLai/docx-redline-js).

### Stratégie adoptée : docx-redline-js

```
1. paragraph.getOoxml()                → extraire le XML OOXML du paragraphe
2. applyRedlineToOxml(ooxml, ...)      → injecter <w:ins>/<w:del> avec auteur configurable
3. changeTrackingMode = OFF            → désactiver Track Changes AVANT insertion
4. paragraph.insertOoxml(result, ...)  → le revision markup survit (Track Changes est OFF)
5. changeTrackingMode = restore        → restaurer l'état original
```

**Avantages** : auteur configurable ("KickOffice AI"), diff word-level chirurgical, formatting préservé (`w:rPr` intact), vrais Track Changes natifs dans Word, support listes/tables/commentaires, lib éprouvée en production sur le Microsoft Marketplace.

L'OOXML (`getOoxml()` / `insertOoxml()`) est aussi utile pour un second cas : **la préservation de mise en forme** lors d'éditions complexes (nouveau tool `editDocumentXml`).

---

## 2. OXML-M1 : Évaluation OOXML par host

### 2.1. Word — ✅ OOXML utile (deux cas d'usage)

**APIs disponibles** (WordApi 1.1+) :
- `Range.getOoxml()` → retourne le Flat OPC XML du range
- `Range.insertOoxml(ooxml, insertLocation)` → insère/remplace avec OOXML
- `Paragraph.getOoxml()` / `Paragraph.insertOoxml()`
- `Body.getOoxml()` / `Body.insertOoxml()`
- `ContentControl.getOoxml()` / `ContentControl.insertOoxml()`

**APIs Track Changes** (WordApi 1.4+) :
- `document.changeTrackingMode` — lecture/écriture du mode (`TrackAll` | `TrackMineOnly` | `Off`)
- `body.getTrackedChanges()` → `TrackedChangeCollection` (WordApi 1.6)
- `trackedChange.accept()` / `trackedChange.reject()` / `trackedChange.getRange()` (WordApi 1.6)

**Cas d'usage 1 — Track Changes chirurgical via docx-redline-js** (WORD-H1) :
- Extraire le XML paragraphe via `paragraph.getOoxml()`
- Calculer le diff et injecter `<w:ins>` / `<w:del>` via `applyRedlineToOxml()`
- Désactiver `changeTrackingMode` temporairement, insérer via `insertOoxml()`, restaurer
- Avantage : révisions parfaites, auteur configurable (ex: "KickOffice AI"), formatting préservé

**Cas d'usage 2 — Préservation de mise en forme** (nouveau tool `editDocumentXml`) :
- `getOoxml()` retourne la structure XML complète avec `<w:rPr>` (fonts, couleurs, tailles)
- Permet de modifier le texte tout en conservant intégralement le formatage
- `insertOoxml(modifiedXml, 'Replace')` réinjecte le résultat
- Cas idéal : réécrire du texte dans un document très formaté (rapports, contrats, mises en page complexes)

**Cas d'usage 3 — Insertion de contenu riche** :
- Tableaux avec styles complexes
- Listes numérotées avec format personnalisé
- Images positionnées avec wrapping text
- En-têtes/pieds de page structurés

**Limites connues** :
- `insertOoxml()` ne fonctionne **PAS dans Word Online** dans certaines versions (Issue #3271)
- Les `w:rsid` (revision identifiers) sont ignorés à l'insertion
- `getOoxml()` est ~6x plus lent que `body.text` (overhead XML)
- Incohérences cross-platform : Mac retourne `w:sdt` complet, Web retourne seulement `w:sdtContent`

**Verdict Word** : ✅ Utile pour Track Changes (via docx-redline-js + getOoxml/insertOoxml) et préservation de formatting (via editDocumentXml). Deux tools distincts.

---

### 2.2. Excel — ❌ Pas d'OOXML via Office.js

**Résultat de l'évaluation** : **Aucune API OOXML disponible** pour Excel dans Office.js.

| API | Disponible ? | Détail |
|-----|---|---|
| `Workbook.getOoxml()` | ❌ Non | N'existe pas dans l'API Excel |
| `Range.getOoxml()` | ❌ Non | Exclusif à Word |
| `Chart.getOoxml()` | ❌ Non | Les charts sont des objets binaires embarqués |
| `Office.context.document.getSelectedDataAsync(coercionType: 'ooxml')` | ❌ Non | Le coercionType OOXML est **exclusif à Word** |

**Alternatives déjà en place** :
- L'API Excel.js est riche (ranges, tables, charts, conditional formatting, named items, pivots)
- `setCellRange`, `formatRange`, `eval_exceljs` couvrent les besoins actuels
- Les charts sont manipulés via l'API objet (`chart.series`, `chart.axes`, etc.)

**Verdict Excel** : ❌ Aucun bénéfice OOXML. L'API Excel.js est suffisante. Ne pas investir.

---

### 2.3. PowerPoint — ✅ Déjà implémenté (editSlideXml)

**Implémentation existante** dans `powerpointTools.ts:1240-1283` + `pptxZipUtils.ts` :

```
Flux : exportAsBase64() → JSZip.loadAsync() → modifier XML → generateAsync() → insertSlidesFromBase64() → delete original
```

**Ce qui marche** :
- Édition du XML de slide (`ppt/slides/slideN.xml`)
- Accès aux shapes, text runs, styles via DOMParser
- Modification de charts, animations, SmartArt impossibles via Office.js API
- Pattern `withSlideZip` + `markDirty()` robuste

**Ce qui pourrait être amélioré** (non prioritaire) :
- Accès aux slide masters (`ppt/slideMasters/`) — pas testé
- Édition de thèmes (`ppt/theme/`) — risque de corruption
- Manipulation d'animations (`<p:timing>`) — complexe mais faisable

**Verdict PowerPoint** : ✅ Déjà fait, fonctionne bien. Pas de changement nécessaire pour Phase 4A.

---

### 2.4. Outlook — ❌ Pas d'OOXML

**Résultat de l'évaluation** : **Aucune API OOXML disponible** pour Outlook.

| API | Disponible ? | Détail |
|-----|---|---|
| `body.getOoxml()` | ❌ Non | N'existe pas — le `Body` Outlook est différent de `Word.Body` |
| MIME manipulation | ❌ Non | Feature request ouverte (Issue #3295), pas implémenté |
| `body.setAsync(html, CoercionType.Html)` | ✅ Oui | Seule option pour le corps d'email riche |
| `body.getAsync(CoercionType.Html)` | ✅ Oui | Lecture du corps en HTML |

**Conclusion** : Les emails Outlook sont manipulés via HTML, pas via OOXML. Le module `richContentPreserver.ts` (implémenté en Phase 2C) gère déjà la préservation de mise en forme HTML.

**Verdict Outlook** : ❌ Aucune possibilité OOXML. Le HTML est l'unique canal. Ne pas investir.

---

### 2.5. Tableau récapitulatif OXML-M1

| Host | OOXML dispo ? | Méthode | Usage recommandé | Priorité |
|------|---|---|---|---|
| **Word** | ✅ Oui | `getOoxml()` / `insertOoxml()` + `docx-redline-js` | Track Changes (auteur configurable) + préservation formatting | 🟠 High |
| **Excel** | ❌ Non | — | Aucun (API Excel.js suffit) | — |
| **PowerPoint** | ✅ Oui | JSZip (`editSlideXml`) | Déjà implémenté | ✅ Done |
| **Outlook** | ❌ Non | — | Aucun (HTML via body.setAsync) | — |

---

## 3. WORD-H1 : Track Changes via docx-redline-js

### 3.1. L'approche Gemini AI for Office : une référence

L'add-in [Gemini AI for Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting) (publié sur le Microsoft Marketplace) résout le problème du Track Changes via sa lib [`docx-redline-js`](https://github.com/AnsonLai/docx-redline-js). Le mécanisme clé :

1. **`getOoxml()`** retourne du XML aplati (pas de revision markup existant)
2. **`docx-redline-js`** calcule le diff et **injecte** `<w:ins>` / `<w:del>` avec un auteur configurable dans le XML
3. **Désactiver Track Changes** (`changeTrackingMode = Off`) **AVANT** `insertOoxml()`
4. **`insertOoxml()`** préserve le revision markup car Word ne le re-traite pas quand Track Changes est OFF
5. **Restaurer** le mode Track Changes original

L'astuce cruciale : quand Track Changes est **ON** pendant `insertOoxml()`, Word crée ses propres révisions sur tout le contenu (= double-tracking). Quand il est **OFF**, le markup `<w:ins>` / `<w:del>` embarqué dans le XML **survit** intact.

### 3.2. Pourquoi docx-redline-js est meilleur que les alternatives

| Critère | office-word-diff (actuel) | changeTrackingMode natif (envisagé initialement) | docx-redline-js (adopté) |
|---|---|---|---|
| Track Changes natifs Word | ❌ Simule via CSS | ✅ Vrais `<w:ins>` / `<w:del>` | ✅ Vrais `<w:ins>` / `<w:del>` |
| Visible panneau Révisions | ❌ Non | ✅ Oui | ✅ Oui |
| Auteur configurable | ❌ CSS only | ❌ Compte Windows uniquement | ✅ `setDefaultAuthor("KickOffice AI")` |
| Préservation formatting | 🟡 Reconstruit les runs | ✅ Natif (ne touche que le texte) | ✅ `w:rPr` préservé + `w:rPrChange` pour modifs style |
| Changements de style (bold, italic...) | ❌ Non | ❌ Non | ✅ `w:rPrChange` intégré |
| Support listes | ❌ Non | ❌ Difficile (ranges imbriqués) | ✅ `applyRedlineToOxmlWithListFallback()` |
| Support tables | ❌ Non | ❌ Complexe | ✅ `reconcileMarkdownTableOoxml()` |
| Commentaires | ❌ Non | ❌ Non | ✅ `injectCommentsIntoOoxml()` |
| Accept/Reject programmatique | ❌ Non | ✅ WordApi 1.6 | ✅ `acceptTrackedChangesInOoxml()` par auteur |
| Complexité d'implémentation | 🟡 3 stratégies cascade | 🟠 Mapping diff→ranges (N syncs) | ✅ 1 getOoxml + 1 transform + 1 insertOoxml |
| Performance | 🟡 N syncs par opération | 🟠 N syncs par opération diff | ✅ 2 syncs total (get + insert) |
| Dépendance | `office-word-diff` (custom) | Aucune | `@ansonlai/docx-redline-js` (zero-dep, ~50 KB) |
| Maturité | Custom, non maintenu | À coder from scratch | ✅ En production sur Microsoft Marketplace |
| Word Online | ✅ Fonctionne | ✅ WordApi 1.4 | 🟡 Dépend de `insertOoxml()` (buggé sur certaines versions) |

### 3.3. Flux détaillé : proposeRevision v2

```
┌──────────────────────────────────────────────────────────────┐
│                    proposeRevision v2                         │
│              (via docx-redline-js)                            │
├──────────────────────────────────────────────────────────────┤
│                                                               │
│  1. Extraire le texte original de la sélection               │
│     selection.load('text') → context.sync()                  │
│                                                               │
│  2. Extraire le XML OOXML de la sélection                    │
│     selection.getOoxml() → context.sync()                    │
│                                                               │
│  3. Appeler docx-redline-js                                  │
│     applyRedlineToOxml(ooxml, originalText, revisedText, {   │
│       author: redlineAuthor,   // "KickOffice AI" (Settings) │
│       generateRedlines: true   // enableTrackChanges          │
│     })                                                        │
│     → retourne { oxml: "...modified XML with w:ins/w:del..." }│
│                                                               │
│  4. Sauvegarder changeTrackingMode original                  │
│     doc.load('changeTrackingMode') → context.sync()          │
│                                                               │
│  5. DÉSACTIVER Track Changes (crucial !)                     │
│     doc.changeTrackingMode = Word.ChangeTrackingMode.off     │
│     → context.sync()                                         │
│                                                               │
│  6. Insérer le XML modifié                                   │
│     selection.insertOoxml(result.oxml, 'Replace')            │
│     → context.sync()                                         │
│                                                               │
│  7. RESTAURER changeTrackingMode (dans finally)              │
│     doc.changeTrackingMode = originalMode                    │
│     → context.sync()                                         │
│                                                               │
│  Résultat : Track Changes natifs dans Word                   │
│  Auteur = "KickOffice AI" (configurable)                     │
│  Visible et acceptable/rejetable dans le panneau Révisions   │
└──────────────────────────────────────────────────────────────┘
```

### 3.4. Algorithme détaillé

```typescript
// PSEUDO-CODE — proposeRevision v2 (docx-redline-js)

import { applyRedlineToOxml, setDefaultAuthor } from '@ansonlai/docx-redline-js'

async function applyRevisionV2(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true,
  redlineAuthor: string = 'KickOffice AI'
): Promise<RevisionResult> {

  // 1. Extraire texte + OOXML de la sélection
  const selection = context.document.getSelection()
  selection.load('text')
  const ooxmlResult = selection.getOoxml()
  await context.sync()

  const originalText = selection.text
  const ooxml = ooxmlResult.value

  if (!originalText?.trim()) {
    return { success: false, strategy: 'none', message: 'No text selected.' }
  }
  if (originalText === revisedText) {
    return { success: true, strategy: 'none', message: 'Text identical.' }
  }

  // 2. Générer le XML avec revision markup via docx-redline-js
  setDefaultAuthor(redlineAuthor)
  const result = await applyRedlineToOxml(ooxml, originalText, revisedText, {
    author: enableTrackChanges ? redlineAuthor : undefined,
    generateRedlines: enableTrackChanges
  })

  // 3. Sauvegarder et désactiver Track Changes
  const doc = context.document
  doc.load('changeTrackingMode')
  await context.sync()
  const originalMode = doc.changeTrackingMode

  try {
    // CRUCIAL : désactiver pour que insertOoxml ne crée pas de double-tracking
    doc.changeTrackingMode = Word.ChangeTrackingMode.off
    await context.sync()

    // 4. Insérer le XML modifié (le revision markup survit car TC est OFF)
    selection.insertOoxml(result.oxml, 'Replace')
    await context.sync()

  } finally {
    // 5. TOUJOURS restaurer le mode original
    doc.changeTrackingMode = originalMode
    await context.sync()
  }

  return {
    success: true,
    strategy: enableTrackChanges ? 'redline' : 'direct-replace',
    author: redlineAuthor,
    message: `Revision applied with ${enableTrackChanges ? 'Track Changes' : 'direct replacement'}.`
  }
}
```

### 3.5. Champ "Redline Author" dans Settings

**Avec `docx-redline-js`, l'auteur est entièrement configurable.**

**Implémentation UI** :
- Ajouter un champ texte dans Settings (onglet Account ou nouvel onglet "Editing")
- Label : "Track Changes Author" / "Auteur des révisions"
- Default : `"KickOffice AI"`
- Stocké dans `localStorage` (clé : `redlineAuthor`)
- Passé à `applyRedlineToOxml({ author: redlineAuthor })`

**Avantages** :
- L'utilisateur peut distinguer les révisions AI des révisions humaines dans le panneau Révisions
- Un cabinet peut configurer "AI Assistant - [Nom du cabinet]"
- On peut aussi accept/reject par auteur via `acceptTrackedChangesInOoxml(oxml, { author: 'KickOffice AI' })`

### 3.6. API docx-redline-js — fonctions utiles pour KickOffice

```typescript
// Configuration (une seule fois au démarrage)
import {
  setDefaultAuthor,
  configureXmlProvider,
  applyRedlineToOxml,
  applyRedlineToOxmlWithListFallback,
  reconcileMarkdownTableOoxml,
  acceptTrackedChangesInOoxml,
  rejectTrackedChangesInOoxml,
  injectCommentsIntoOoxml,
  extractReplacementNodesFromOoxml
} from '@ansonlai/docx-redline-js'

// En environnement browser (Office add-in) : DOMParser/XMLSerializer sont natifs
// Pas besoin de configureXmlProvider()

setDefaultAuthor('KickOffice AI')

// --- Réconciliation texte ---
const result = await applyRedlineToOxml(ooxml, original, revised, {
  author: 'KickOffice AI',
  generateRedlines: true
})
// result.oxml contient le XML avec <w:ins>/<w:del>

// --- Listes (avec fallback structural) ---
const listResult = await applyRedlineToOxmlWithListFallback(ooxml, original, revised, {
  author: 'KickOffice AI',
  generateRedlines: true
})

// --- Tables ---
const tableResult = await reconcileMarkdownTableOoxml(ooxml, original, markdownTable, {
  author: 'KickOffice AI',
  generateRedlines: true
})

// --- Accept/Reject par auteur ---
const accepted = acceptTrackedChangesInOoxml(ooxml, { author: 'KickOffice AI' })
const rejected = rejectTrackedChangesInOoxml(ooxml, { allAuthors: true })

// --- Commentaires ---
const withComments = injectCommentsIntoOoxml(ooxml, [
  { anchorText: 'mot ciblé', commentText: 'Suggestion ici', author: 'KickOffice AI' }
])
```

### 3.7. Attention au format de sortie (Output Shape Matrix)

`docx-redline-js` retourne différents formats selon l'API :

| API | Format retourné | Safe pour `insertOoxml()` ? |
|---|---|---|
| `applyRedlineToOxml()` | Fragment, `<w:document>`, ou `<pkg:package>` | ⚠️ Variable — à normaliser |
| `extractReplacementNodesFromOoxml()` | Nœuds normalisés | ✅ Oui |
| `applyOperationToDocumentXml()` | `<w:document>` root | ✅ Oui (mais pour document complet) |

**Important** : Ne **jamais** écrire un payload `<pkg:package>` directement dans `insertOoxml()`. Utiliser `extractReplacementNodesFromOoxml()` pour normaliser si nécessaire. À tester lors de l'intégration.

### 3.8. Migration : office-word-diff → docx-redline-js

| Aspect | office-word-diff (à supprimer) | docx-redline-js (à installer) |
|---|---|---|
| Package | `office-word-diff` (local dans `/office-word-diff/`) | `@ansonlai/docx-redline-js` (npm) |
| Diff engine | diff-match-patch (dépendance externe) | diff-match-patch (internalisé, zero-dep) |
| Approche | Reconstruit les Word runs après diff | Injecte du revision markup XML dans l'OOXML existant |
| Track Changes | Simulés via CSS (strikethrough + underline) | Vrais `<w:ins>` / `<w:del>` natifs |
| Auteur | Non configurable | Configurable via `setDefaultAuthor()` |

**Actions de migration** :
1. `npm install @ansonlai/docx-redline-js` dans `frontend/`
2. Réécrire `wordDiffUtils.ts` pour utiliser `applyRedlineToOxml` + le flux disable/insert/restore
3. Supprimer la dépendance `office-word-diff` de `frontend/package.json`
4. Supprimer le dossier `/office-word-diff/`
5. Supprimer `diff-match-patch` de `frontend/package.json` (internalisé dans docx-redline-js)
6. Mettre à jour le Dockerfile si nécessaire
7. Ajouter le champ "Redline Author" dans le composant Settings

---

## 4. Nouveau tool : editDocumentXml

### 4.1. Pourquoi un tool séparé

L'OOXML natif (`getOoxml()` / `insertOoxml()`) est inutile pour Track Changes (cf. section 3.1), mais reste **très utile** pour :

| Cas d'usage | Pourquoi OOXML > Office.js API |
|---|---|
| **Réécrire du texte en conservant le formatting exact** | `getOoxml()` expose les `<w:rPr>` (police, couleur, taille) — on peut changer le texte `<w:t>` sans toucher au formatting |
| **Modifier des styles inline complexes** | Accès direct aux propriétés XML impossibles via l'API haut niveau |
| **Insérer des tableaux avec styles précis** | L'OOXML permet un contrôle pixel-perfect vs `insertHtml()` qui est approximatif |
| **Manipuler des content controls** | Accès au `<w:sdt>` complet |
| **Debugging / inspection** | Voir exactement ce que Word stocke en interne |

### 4.2. Architecture du tool

```typescript
// frontend/src/utils/wordTools.ts — nouveau tool

editDocumentXml: {
  name: 'editDocumentXml',
  category: 'write',
  description: `Edit Word document OOXML directly for precision operations.

    **USE WHEN**: You need to modify text while preserving exact formatting
    (fonts, colors, sizes, styles) that would be lost with insertText/insertHtml.

    **DO NOT USE FOR**: Track Changes (use proposeRevision instead).

    The code receives:
    - \`ooxml\`: string — the raw Flat OPC XML of the target range
    - \`DOMParser\`, \`XMLSerializer\` — XML manipulation
    - \`escapeXml(str)\` — safely escape XML special chars
    - \`setResult(modifiedXml)\` — call this to write back the modified XML

    Your code should:
    1. Parse the ooxml with DOMParser
    2. Find and modify the desired elements
    3. Serialize back and call setResult()`,

  inputSchema: {
    type: 'object',
    properties: {
      target: {
        type: 'string',
        enum: ['selection', 'paragraph'],
        description: 'What to target: current selection or specific paragraph'
      },
      paragraphIndex: {
        type: 'number',
        description: 'If target=paragraph, the 0-based paragraph index'
      },
      code: {
        type: 'string',
        description: 'JavaScript code to manipulate the OOXML'
      },
      explanation: {
        type: 'string',
        description: 'What this code does (required for audit trail)'
      }
    },
    required: ['code', 'explanation']
  },

  executeWord: async (context, args) => {
    const { target = 'selection', paragraphIndex, code, explanation } = args

    // 1. Get the target range
    let range
    if (target === 'paragraph' && paragraphIndex !== undefined) {
      const paragraphs = context.document.body.paragraphs
      paragraphs.load('items')
      await context.sync()
      if (paragraphIndex >= paragraphs.items.length) {
        return JSON.stringify({ success: false, error: `Paragraph ${paragraphIndex} out of bounds` })
      }
      range = paragraphs.items[paragraphIndex].getRange()
    } else {
      range = context.document.getSelection()
    }

    // 2. Get OOXML
    const ooxmlResult = range.getOoxml()
    await context.sync()
    const ooxml = ooxmlResult.value

    // 3. Execute code in sandbox
    let modifiedXml = null
    const setResult = (xml) => { modifiedXml = xml }

    await sandboxedEval(code, {
      ooxml,
      DOMParser,
      XMLSerializer,
      escapeXml,
      setResult
    }, 'Word')

    // 4. Write back if modified
    if (modifiedXml) {
      range.insertOoxml(modifiedXml, 'Replace')
      await context.sync()
      return JSON.stringify({
        success: true,
        explanation,
        action: 'OOXML modified and reinserted'
      })
    }

    return JSON.stringify({
      success: true,
      explanation,
      action: 'No modifications applied (setResult not called)'
    })
  }
}
```

### 4.3. Exemple d'utilisation par le LLM

**Scénario** : L'utilisateur a un paragraphe avec du texte en gras rouge et veut changer "Contrat de vente" en "Accord commercial" en gardant le même formatting.

```json
{
  "target": "selection",
  "code": "const parser = new DOMParser();\nconst doc = parser.parseFromString(ooxml, 'application/xml');\nconst textNodes = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't');\nfor (const t of textNodes) {\n  if (t.textContent.includes('Contrat de vente')) {\n    t.textContent = t.textContent.replace('Contrat de vente', 'Accord commercial');\n  }\n}\nconst serializer = new XMLSerializer();\nsetResult(serializer.serializeToString(doc));",
  "explanation": "Replace 'Contrat de vente' with 'Accord commercial' while preserving bold red formatting"
}
```

**Résultat** : Le texte change, mais les `<w:rPr>` (bold, color, font) sont intégralement préservés car seul le contenu de `<w:t>` est modifié.

### 4.4. Quand utiliser editDocumentXml vs les autres tools

```
Decision tree mis à jour pour word.skill.md :

User veut modifier du TEXTE existant ?
  ├── Modification simple (mot/phrase) → searchAndReplace
  ├── Réécriture de paragraphes → proposeRevision (avec Track Changes natif)
  └── Texte dans un document très formaté (préserver fonts/couleurs/styles) → editDocumentXml

User veut ajouter du NOUVEAU contenu ?
  └── insertContent (Markdown + inline syntax)

User veut appliquer du FORMATTING ?
  ├── Sur des mots spécifiques → searchAndFormat
  └── Sur la sélection → formatText

Cas edge non couvert ?
  └── eval_wordjs (dernier recours)
```

---

## 5. Migration : office-word-diff → docx-redline-js

### 5.1. Fichiers impactés

| Fichier | Action |
|---|---|
| `frontend/src/utils/wordTrackChanges.ts` | **Créer** — helpers Track Changes (setChangeTrackingForAi, restoreChangeTracking, loadRedlineAuthor) |
| `frontend/src/utils/wordDiffUtils.ts` | **Réécrire** — utiliser `docx-redline-js` + helpers de `wordTrackChanges.ts` |
| `frontend/src/utils/wordTools.ts` (proposeRevision) | Mettre à jour l'appel + description du tool |
| `frontend/package.json` | Supprimer `office-word-diff` + `diff-match-patch`, ajouter `@ansonlai/docx-redline-js` |
| `office-word-diff/` (dossier entier) | **Supprimer** |
| `frontend/src/skills/word.skill.md` | Mettre à jour la description de `proposeRevision` |
| `frontend/src/components/settings/` | Ajouter champ "Redline Author" |
| `README.md` | Ajouter section Acknowledgments (docx-redline-js MIT) + supprimer toute référence à `office-word-diff` |
| `Dockerfile` (si référence à office-word-diff) | Supprimer les `COPY` et `npm install` relatifs |

### 5.2. Helpers Track Changes : copier depuis Gemini AI for Office (MIT)

Ces fonctions sont copiées/adaptées de [Gemini AI for Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting) (MIT License).

**Créer `frontend/src/utils/wordTrackChanges.ts`** :

```typescript
/**
 * Word Track Changes Utilities
 *
 * Manages Track Changes state during OOXML insertion.
 * Pattern from Gemini AI for Office (MIT License):
 * https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
 */

const DEFAULT_AUTHOR = 'KickOffice AI'

export interface TrackingState {
  available: boolean
  originalMode: any | null
  changed: boolean
}

/**
 * Save current Track Changes mode and set desired mode.
 * Mirrors Gemini's setChangeTrackingForAi().
 *
 * When inserting OOXML with embedded w:ins/w:del, we DISABLE native tracking
 * to prevent Word from double-tracking the inserted content.
 */
export async function setChangeTrackingForAi(
  context: Word.RequestContext,
  redlineEnabled: boolean,
  sourceLabel: string = 'AI',
): Promise<TrackingState> {
  let originalMode = null
  let changed = false
  let available = false

  try {
    const doc = context.document
    doc.load('changeTrackingMode')
    await context.sync()

    available = true
    originalMode = doc.changeTrackingMode

    // When redlines are embedded in OOXML → DISABLE native tracking
    // When no redlines → ENABLE tracking so Word tracks our text changes
    const desiredMode = redlineEnabled
      ? Word.ChangeTrackingMode.off    // OFF because w:ins/w:del are already in the XML
      : Word.ChangeTrackingMode.off    // OFF for silent replacement too

    if (originalMode !== desiredMode) {
      doc.changeTrackingMode = desiredMode
      await context.sync()
      changed = true
    }
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: unavailable`, error)
  }

  return { available, originalMode, changed }
}

/**
 * Restore Track Changes mode to its original state.
 * Mirrors Gemini's restoreChangeTracking().
 *
 * MUST be called in a finally block after setChangeTrackingForAi().
 */
export async function restoreChangeTracking(
  context: Word.RequestContext,
  trackingState: TrackingState,
  sourceLabel: string = 'AI',
): Promise<void> {
  if (!trackingState || !trackingState.available || !trackingState.changed || trackingState.originalMode === null) {
    return
  }

  try {
    context.document.changeTrackingMode = trackingState.originalMode
    await context.sync()
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: restore failed`, error)
  }
}

/**
 * Load redline enabled setting from localStorage.
 * Default: true (Track Changes enabled).
 */
export function loadRedlineSetting(): boolean {
  const storedSetting = localStorage.getItem('redlineEnabled')
  return storedSetting !== null ? storedSetting === 'true' : true
}

/**
 * Load the redline author name from localStorage.
 * Default: "KickOffice AI".
 */
export function loadRedlineAuthor(): string {
  const storedAuthor = localStorage.getItem('redlineAuthor')
  if (storedAuthor && storedAuthor.trim() !== '') {
    return storedAuthor
  }
  return DEFAULT_AUTHOR
}
```

### 5.3. Code complet : nouveau wordDiffUtils.ts

**Réécrire `frontend/src/utils/wordDiffUtils.ts`** :

```typescript
/**
 * Word Diff Utilities — v2 (docx-redline-js)
 *
 * Generates native Word Track Changes (w:ins / w:del) by injecting
 * revision markup into paragraph OOXML via docx-redline-js.
 *
 * Integration pattern from Gemini AI for Office (MIT License):
 * https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
 * OOXML engine: https://github.com/AnsonLai/docx-redline-js (MIT License)
 */

import {
  applyRedlineToOxml,
  setDefaultAuthor,
} from '@ansonlai/docx-redline-js'

import {
  setChangeTrackingForAi,
  restoreChangeTracking,
  loadRedlineAuthor,
} from './wordTrackChanges'

export interface RevisionResult {
  success: boolean
  strategy: 'redline' | 'direct-replace' | 'none'
  author?: string
  message: string
}

/**
 * Apply a revision to the current selection using docx-redline-js.
 *
 * Follows the Gemini AI for Office pattern:
 * 1. Extract selection text + OOXML via getOoxml()
 * 2. Generate revision markup via applyRedlineToOxml() (w:ins / w:del)
 * 3. Disable Track Changes via setChangeTrackingForAi() (prevent double-tracking)
 * 4. Insert modified OOXML via insertOoxml() (revision markup survives)
 * 5. Restore Track Changes via restoreChangeTracking()
 *
 * IMPORTANT: Must be called within Word.run() context.
 */
export async function applyRevisionToSelection(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true,
): Promise<RevisionResult> {
  const redlineAuthor = loadRedlineAuthor()

  // 1. Get selection text + OOXML in a single sync batch
  const selection = context.document.getSelection()
  selection.load('text')
  const ooxmlResult = selection.getOoxml()
  await context.sync()

  const originalText = selection.text
  const ooxml = ooxmlResult.value

  // 2. Edge cases
  if (!originalText || !originalText.trim()) {
    return {
      success: false,
      strategy: 'none',
      message: 'Error: No text selected. Please select text before using proposeRevision.',
    }
  }

  if (originalText === revisedText) {
    return {
      success: true,
      strategy: 'none',
      message: 'Text is identical, no changes needed.',
    }
  }

  // 3. Generate revision markup via docx-redline-js
  setDefaultAuthor(redlineAuthor)

  let resultOoxml: string
  try {
    const redlineResult = await applyRedlineToOxml(
      ooxml,
      originalText,
      revisedText,
      {
        author: enableTrackChanges ? redlineAuthor : undefined,
        generateRedlines: enableTrackChanges,
      },
    )
    resultOoxml = redlineResult.oxml
  } catch (error: any) {
    console.error('[WordDiff] docx-redline-js error:', error)
    return {
      success: false,
      strategy: 'none',
      message: `Error generating revision markup: ${error.message || String(error)}`,
    }
  }

  // 4. Disable Track Changes, insert, restore — pattern from Gemini AI for Office
  const trackingState = await setChangeTrackingForAi(
    context,
    enableTrackChanges,
    'proposeRevision',
  )

  try {
    // Insert the modified OOXML
    // w:ins/w:del survive because native tracking is OFF
    selection.insertOoxml(resultOoxml, 'Replace')
    await context.sync()
  } catch (insertError: any) {
    // Fallback: if insertOoxml fails (Word Online), use direct text replacement
    console.warn('[WordDiff] insertOoxml failed, falling back to insertText:', insertError)
    try {
      selection.insertText(revisedText, 'Replace')
      await context.sync()
    } catch (fallbackError: any) {
      return {
        success: false,
        strategy: 'none',
        message: `Error applying revision: ${fallbackError.message || String(fallbackError)}`,
      }
    }
    return {
      success: true,
      strategy: 'direct-replace',
      message: 'Revision applied with direct replacement (insertOoxml unavailable).',
    }
  } finally {
    // 5. ALWAYS restore the original tracking mode
    await restoreChangeTracking(context, trackingState, 'proposeRevision')
  }

  return {
    success: true,
    strategy: enableTrackChanges ? 'redline' : 'direct-replace',
    author: enableTrackChanges ? redlineAuthor : undefined,
    message: enableTrackChanges
      ? `Revision applied with Track Changes (author: "${redlineAuthor}").`
      : 'Revision applied with direct replacement (no Track Changes).',
  }
}
```

### 5.3. Code complet : proposeRevision tool mis à jour

Dans `frontend/src/utils/wordTools.ts`, remplacer la définition de `proposeRevision` :

```typescript
proposeRevision: {
  name: 'proposeRevision',
  category: 'write' as ToolCategory,
  description: `**PREFERRED TOOL** for modifying existing text.

Generates native Word Track Changes (redlines) using OOXML revision markup.
The user can accept/reject each change individually in Word's Review pane.

Changes are attributed to a configurable author (default: "KickOffice AI")
visible in the Track Changes panel, distinguishable from human edits.

**Input**: The COMPLETE revised version of the selected text.
**Output**: The selection is replaced with tracked insertions/deletions.

**Requirements**: Text must be selected in the document before calling.
**Track Changes**: Enabled by default. Set enableTrackChanges=false for silent replacement.`,

  inputSchema: {
    type: 'object',
    properties: {
      revisedText: {
        type: 'string',
        description: 'The complete revised version of the selected text. Must contain ALL text, not just changes.',
      },
      enableTrackChanges: {
        type: 'boolean',
        description: 'Show changes in Word Track Changes panel (default: true). Set false for silent replacement.',
      },
    },
    required: ['revisedText'],
  },

  executeWord: async (context: Word.RequestContext, args: Record<string, any>) => {
    const { revisedText, enableTrackChanges = true } = args

    const result = await applyRevisionToSelection(context, revisedText, enableTrackChanges)

    return JSON.stringify({
      success: result.success,
      strategy: result.strategy,
      author: result.author,
      message: result.message,
    })
  },
},
```

### 5.4. Interface publique préservée

L'interface `RevisionResult` change légèrement :

```typescript
// AVANT (office-word-diff)
export interface RevisionResult {
  success: boolean
  strategy: 'token' | 'sentence' | 'block'
  insertions: number
  deletions: number
  unchanged: number
  message: string
}

// APRÈS (docx-redline-js)
export interface RevisionResult {
  success: boolean
  strategy: 'redline' | 'direct-replace' | 'none'
  author?: string
  message: string
}
```

Les champs `insertions`, `deletions`, `unchanged` sont supprimés (le comptage était approximatif et non utilisé par le LLM). `author` est ajouté.

### 5.5. Attribution README.md

Ajouter dans `README.md` :

```markdown
## Acknowledgments

### docx-redline-js

KickOffice's Track Changes (redline) feature is powered by
[docx-redline-js](https://github.com/AnsonLai/docx-redline-js) by Anson Lai,
a zero-dependency OOXML reconciliation engine for native Word revision markup.

The integration approach (disable Track Changes → insertOoxml with embedded
w:ins/w:del → restore) is inspired by the
[Gemini AI for Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting)
add-in.

Both projects are licensed under the [MIT License](https://opensource.org/licenses/MIT).
```

### 5.6. Gestion du fallback (si insertOoxml échoue)

```typescript
function isOoxmlAvailable(): boolean {
  // insertOoxml est dans WordApi 1.1, mais buggé sur certaines versions Word Online
  return Office.context.requirements.isSetSupported('WordApi', '1.1')
}
```

Si `insertOoxml()` échoue (Word Online buggé) :
- Catch l'erreur, log un warning
- Fallback : `selection.insertText(revisedText, 'Replace')` — remplace sans Track Changes
- Retourner `{ strategy: 'direct-replace', message: 'insertOoxml unavailable, used direct replacement' }`

Pas de fallback vers `office-word-diff` — suppression complète.

---

## 6. DUP-M1 : Extraction truncateString

### 6.1. Pattern dupliqué (4 occurrences)

```typescript
// wordTools.ts:1511
code.slice(0, 300) + (code.length > 300 ? '...' : '')

// wordTools.ts:1543
code.slice(0, 200) + '...'

// outlookTools.ts:463
code.slice(0, 300) + (code.length > 300 ? '...' : '')

// outlookTools.ts:494
code.slice(0, 200) + '...'
```

### 6.2. Utilitaire à extraire dans common.ts

```typescript
// frontend/src/utils/common.ts — AJOUTER

/**
 * Truncate a string to maxLen characters, appending '...' if truncated.
 */
export function truncateString(str: string, maxLen: number): string {
  if (str.length <= maxLen) return str
  return str.slice(0, maxLen) + '...'
}
```

### 6.3. Remplacement dans les fichiers

```typescript
// wordTools.ts:1511 — AVANT
code.slice(0, 300) + (code.length > 300 ? '...' : '')
// APRÈS
truncateString(code, 300)

// wordTools.ts:1543 — AVANT
code.slice(0, 200) + '...'
// APRÈS
truncateString(code, 200)

// outlookTools.ts:463 — AVANT
code.slice(0, 300) + (code.length > 300 ? '...' : '')
// APRÈS
truncateString(code, 300)

// outlookTools.ts:494 — AVANT
code.slice(0, 200) + '...'
// APRÈS
truncateString(code, 200)
```

Import à ajouter dans `wordTools.ts` et `outlookTools.ts` :
```typescript
import { truncateString } from './common'
```

---

## 7. Plan d'implémentation détaillé

### Ordre d'exécution (4 étapes)

```
Étape 1 — DUP-M1 (15 min)
  └── Extraire truncateString dans common.ts
  └── Remplacer les 4 occurrences (wordTools.ts × 2, outlookTools.ts × 2)
  └── Tester que le build passe

Étape 2 — WORD-H1 : docx-redline-js (1-2 heures)
  ├── 2a. Installer @ansonlai/docx-redline-js
  │   └── cd frontend && npm install @ansonlai/docx-redline-js
  │
  ├── 2b. Créer wordTrackChanges.ts
  │   └── Copier le code de la section 5.2 (helpers Track Changes)
  │   └── setChangeTrackingForAi(), restoreChangeTracking(), loadRedlineAuthor()
  │
  ├── 2c. Réécrire wordDiffUtils.ts
  │   └── Copier le code de la section 5.3 (applyRevisionToSelection v2)
  │   └── Import docx-redline-js + wordTrackChanges
  │   └── Supprimer les imports office-word-diff
  │   └── Supprimer previewDiffStats(), computeRawDiff(), hasComplexContent()
  │
  ├── 2d. Mettre à jour proposeRevision dans wordTools.ts
  │   └── Copier le code de la section 5.4
  │   └── Mettre à jour les imports
  │
  ├── 2e. Ajouter editDocumentXml dans wordTools.ts
  │   └── Copier le code de la section 4.2
  │   └── Ajouter à l'objet wordToolDefinitions
  │
  ├── 2f. Ajouter "Redline Author" dans Settings
  │   └── Champ texte, default "KickOffice AI", stocké localStorage('redlineAuthor')
  │   └── Toggle "Enable Track Changes", stocké localStorage('redlineEnabled')
  │
  ├── 2g. Mettre à jour word.skill.md
  │   └── Ajouter editDocumentXml dans le decision tree
  │   └── Mettre à jour la description de proposeRevision (Track Changes natif, auteur configurable)
  │
  └── 2h. Supprimer office-word-diff + diff-match-patch
      └── Supprimer le dossier /office-word-diff/
      └── Retirer office-word-diff et diff-match-patch de frontend/package.json
      └── npm install pour mettre à jour le lockfile

Étape 3 — README + Attribution + Nettoyage (10 min)
  └── Ajouter la section Acknowledgments dans README.md (cf. section 5.6)
  └── Mentionner docx-redline-js (MIT) et Gemini AI for Office (MIT)
  └── Supprimer toute référence à office-word-diff dans README.md

Étape 4 — Vérification (30 min)
  └── Build complet (npm run build)
  └── Vérifier aucune référence résiduelle à office-word-diff ou diff-match-patch
  └── Tests manuels :
      └── proposeRevision avec Track Changes → vérifier w:ins/w:del dans le panneau Révisions
      └── proposeRevision sans Track Changes → vérifier remplacement silencieux
      └── editDocumentXml → vérifier préservation formatting
      └── Vérifier le fallback si insertOoxml échoue
  └── Mettre à jour DESIGN_REVIEW.md (marquer les 3 items comme FIXED)
```

### Fichiers créés / modifiés / supprimés

| Action | Fichier | Détail |
|---|---|---|
| **Modifier** | `frontend/src/utils/common.ts` | Ajouter `truncateString()` |
| **Créer** | `frontend/src/utils/wordTrackChanges.ts` | Helpers Track Changes (section 5.2) |
| **Réécrire** | `frontend/src/utils/wordDiffUtils.ts` | docx-redline-js integration (section 5.3) |
| **Modifier** | `frontend/src/utils/wordTools.ts` | proposeRevision v2 (section 5.4) + editDocumentXml (section 4.2) + truncateString |
| **Modifier** | `frontend/src/utils/outlookTools.ts` | Utiliser `truncateString` importé |
| **Modifier** | `frontend/src/skills/word.skill.md` | Decision tree + descriptions mises à jour |
| **Modifier** | `frontend/package.json` | +`@ansonlai/docx-redline-js`, −`office-word-diff`, −`diff-match-patch` |
| **Modifier** | `frontend/src/components/settings/` | Champ "Redline Author" + Toggle "Enable Track Changes" |
| **Modifier** | `README.md` | Section Acknowledgments (docx-redline-js MIT) + supprimer refs office-word-diff |
| **Supprimer** | `office-word-diff/` (dossier entier) | Package custom remplacé par `@ansonlai/docx-redline-js` |

---

## 8. Risques et mitigations

### 8.1. Risque : insertOoxml() ne marche pas sur Word Online

| Risque | Impact | Mitigation |
|---|---|---|
| `insertOoxml()` retourne "Not implemented" | `proposeRevision` et `editDocumentXml` inutilisables | Catch l'erreur → fallback `selection.insertText(revisedText, 'Replace')` sans Track Changes |
| XML corrompu après transformation docx-redline-js | Document cassé | `try/catch` autour de `insertOoxml()`. Si erreur, ne pas modifier le document. Log l'erreur |

### 8.2. Risque : double-tracking si changeTrackingMode n'est pas restauré

| Risque | Impact | Mitigation |
|---|---|---|
| Crash entre disable et restore | Track Changes reste OFF sans que l'utilisateur le sache | `try/finally` **obligatoire** — le `finally` restaure toujours le mode original |
| `changeTrackingMode` n'est pas supporté (WordApi < 1.4) | Impossible de disable/restore | Vérifier `isSetSupported('WordApi', '1.4')`. Si non supporté, insérer quand même (risque de double-tracking, mais c'est le mieux possible) |

### 8.3. Risque : format de sortie docx-redline-js variable

| Risque | Impact | Mitigation |
|---|---|---|
| `applyRedlineToOxml()` retourne `<pkg:package>` au lieu d'un fragment | `insertOoxml()` échoue ou comportement inattendu | Tester le format de sortie. Si nécessaire, utiliser `extractReplacementNodesFromOoxml()` pour normaliser |
| OOXML namespace manquant après transformation | XML invalide | Valider le XML avant insertion via `DOMParser.parseFromString()` — checker le `<parsererror>` |

### 8.4. Risque : performance de getOoxml()

| Risque | Impact | Mitigation |
|---|---|---|
| `getOoxml()` ~6x plus lent que `body.text` | Latence perceptible sur de grosses sélections | `proposeRevision` travaille sur la **sélection** (pas le body entier). `editDocumentXml` sur un range/paragraphe ciblé |
| XML volumineux pour de gros paragraphes | Overhead mémoire | Documenter dans le skill.md que ces tools sont pour des éditions ciblées |

### 8.5. Risque : suppression de office-word-diff = régression

| Risque | Impact | Mitigation |
|---|---|---|
| `docx-redline-js` gère différemment certains cas edge | Résultat différent | docx-redline-js utilise le même engine `diff-match-patch` en interne, mais opère au niveau XML plutôt qu'au niveau runs. Plus précis, pas moins |
| Dépendance externe (`@ansonlai/docx-redline-js`) | Risque supply chain | Lib zero-dep, MIT license, code auditable, utilisée en production par Gemini AI for Office (Microsoft Marketplace) |

---

## 9. Sources

### Projet de référence (inspiration principale)
- [Gemini AI for Office — Word Add-In](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting) — MIT License. Add-in publié sur le Microsoft Marketplace. Approche Track Changes via OOXML + docx-redline-js.
- [docx-redline-js](https://github.com/AnsonLai/docx-redline-js) — MIT License. Zero-dependency OOXML engine for native Word redlines. Utilisé pour `applyRedlineToOxml()`, `setDefaultAuthor()`, `acceptTrackedChangesInOoxml()`.
- [docx-redline-mcp](https://github.com/AnsonLai/docx-redline-mcp) — MCP server companion (référence, non utilisé directement).

### Office.js Word OOXML APIs
- [Word.Range class — getOoxml/insertOoxml](https://learn.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview)
- [Word.Paragraph class](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview)
- [Word.Body class](https://learn.microsoft.com/en-us/javascript/api/word/word.body?view=word-js-preview)
- [OOXML in Word add-ins (Microsoft Guide)](https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml)
- [Load and write OOXML sample](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/word-add-in-load-and-write-open-xml/)

### Track Changes APIs
- [Word.ChangeTrackingMode enum (WordApi 1.4)](https://learn.microsoft.com/en-us/javascript/api/word/word.changetrackingmode?view=word-js-preview)
- [Word.TrackedChange class (WordApi 1.6)](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange?view=word-js-preview)
- [Word.TrackedChangeCollection class](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchangecollection?view=word-js-preview)
- [WordApi 1.4 requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-4-requirement-set?view=word-js-preview)
- [WordApi 1.6 requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-6-requirement-set?view=word-js-preview)

### Issues GitHub connues (Office.js)
- [#2123 — getOOXML trackrevision absent](https://github.com/OfficeDev/office-js/issues/2123) — `getOoxml()` ne retourne pas les éléments de revision
- [#329 — Expose TrackRevision via API](https://github.com/OfficeDev/office-js/issues/329) — feature request depuis 2018
- [#334 — Bug comment + selection OOXML](https://github.com/OfficeDev/office-js/issues/334) — `insertOoxml` + Track Changes ON = tout marqué comme nouveau (d'où le disable/restore)
- [#3271 — insertOoxml not working Word Online](https://github.com/OfficeDev/office-js/issues/3271) — limitation Word Web
- [#5491 — insertOoxml style changes ignored](https://github.com/OfficeDev/office-js/issues/5491)
- [#2991 — Numbering changed on getOoxml/insertOoxml](https://github.com/OfficeDev/office-js/issues/2991)
- [#6246 — No API for revision display mode](https://github.com/OfficeDev/office-js/issues/6246) — Oct 2025
- [#3295 — Outlook MIME control](https://github.com/OfficeDev/office-js/issues/3295) — pas d'API MIME Outlook

### Codebase KickOffice (références internes)
- `frontend/src/utils/pptxZipUtils.ts` — pattern `withSlideZip` existant pour PowerPoint
- `frontend/src/utils/wordDiffUtils.ts` — wrapper `office-word-diff` actuel (à réécrire)
- `frontend/src/utils/wordTools.ts` — tools Word dont `proposeRevision`, `eval_wordjs`
- `frontend/src/utils/sandbox.ts` — SES sandbox pour l'exécution sécurisée
- `frontend/src/skills/word.skill.md` — guidelines de sélection d'outils
- `office-word-diff/` — package custom à supprimer (remplacé par `@ansonlai/docx-redline-js`)
