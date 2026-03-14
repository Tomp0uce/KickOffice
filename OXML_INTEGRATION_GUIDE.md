# OXML Integration Guide — Phase 4A

> **Document créé le 2026-03-14** — Guide dédié à l'intégration OOXML dans KickOffice.
> Couvre les 3 tâches de la Phase 4A : OXML-M1, WORD-H1, DUP-M1.

---

## Table des matières

1. [Résumé exécutif](#1-résumé-exécutif)
2. [OXML-M1 : Évaluation OOXML par host](#2-oxml-m1--évaluation-ooxml-par-host)
3. [WORD-H1 : Track Changes natif (stratégie révisée)](#3-word-h1--track-changes-natif-stratégie-révisée)
4. [Nouveau tool : editDocumentXml](#4-nouveau-tool--editdocumentxml)
5. [Migration : suppression de office-word-diff](#5-migration--suppression-de-office-word-diff)
6. [DUP-M1 : Extraction truncateString](#6-dup-m1--extraction-truncatestring)
7. [Plan d'implémentation détaillé](#7-plan-dimplémentation-détaillé)
8. [Risques et mitigations](#8-risques-et-mitigations)
9. [Sources](#9-sources)

---

## 1. Résumé exécutif

### Constat initial

KickOffice utilise Office.js comme couche d'abstraction exclusive pour manipuler les documents. L'OOXML est utilisé **uniquement** dans PowerPoint (`editSlideXml` via JSZip dans `pptxZipUtils.ts`). Word utilise `office-word-diff` (npm) pour le diffing, Excel et Outlook n'ont aucune manipulation OOXML.

### Découverte critique : `insertOoxml()` ne supporte PAS le revision markup

La recherche approfondie de l'API Office.js révèle un point bloquant pour la stratégie initiale de WORD-H1 :

| Comportement | Détail |
|---|---|
| `range.getOoxml()` | Retourne du XML **aplati** — les éléments `<w:ins>` / `<w:del>` sont **absents** |
| `range.insertOoxml(xml, 'Replace')` | **Supprime** tout revision markup injecté (`<w:ins>`, `<w:del>`) |
| Track Changes + insertOoxml | Si Track Changes est ON, Word crée ses **propres** révisions sur tout le contenu inséré (= tout marqué comme nouveau) |
| `<w:trackRevisions/>` dans settings | **Absent** du XML retourné par `getOoxml()` |

**Conséquence** : L'approche décrite dans DESIGN_REVIEW.md (injecter `<w:ins>` / `<w:del>` directement dans le XML paragraphe) **ne fonctionne pas** via Office.js. Il faut une stratégie alternative.

### Stratégie révisée

Au lieu d'injecter du revision markup OOXML, utiliser les **APIs natives Track Changes** de Word (WordApi 1.4+) :

```
1. Activer Track Changes via document.changeTrackingMode = 'TrackAll'
2. Faire les modifications texte via Office.js (insertText, delete, etc.)
3. Word crée automatiquement des révisions natives
4. Restaurer le mode Track Changes original
```

L'OOXML (`getOoxml()` / `insertOoxml()`) reste utile pour un autre cas : **la préservation de mise en forme** lors d'éditions complexes (nouveau tool `editDocumentXml`).

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

**Cas d'usage 1 — Track Changes chirurgical** (WORD-H1) :
- Activer `changeTrackingMode = 'TrackAll'`
- Appliquer les diffs via `range.insertText()` ou `range.delete()`
- Word crée des `<w:ins>` / `<w:del>` natifs automatiquement
- Avantage : révisions parfaites, auteur = compte Windows actif

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

**Verdict Word** : ✅ Utile pour Track Changes (via changeTrackingMode) et préservation de formatting (via getOoxml/insertOoxml). Deux tools distincts.

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
| **Word** | ✅ Oui | `getOoxml()` / `insertOoxml()` + `changeTrackingMode` | Track Changes natif + préservation formatting | 🟠 High |
| **Excel** | ❌ Non | — | Aucun (API Excel.js suffit) | — |
| **PowerPoint** | ✅ Oui | JSZip (`editSlideXml`) | Déjà implémenté | ✅ Done |
| **Outlook** | ❌ Non | — | Aucun (HTML via body.setAsync) | — |

---

## 3. WORD-H1 : Track Changes natif (stratégie révisée)

### 3.1. Pourquoi la stratégie initiale ne marche pas

Le DESIGN_REVIEW.md proposait d'injecter `<w:ins>` / `<w:del>` directement dans le XML paragraphe via `insertOoxml()`. Cette approche est **impossible** car :

1. **`insertOoxml()` strip le revision markup** — les éléments `<w:ins>` et `<w:del>` sont supprimés par Word lors de l'insertion
2. **`getOoxml()` retourne du XML aplati** — pas de revision markup dans la sortie
3. **Pas de contrôle sur l'auteur** — même si le markup était préservé, l'auteur serait celui du compte Windows, pas une valeur configurable

### 3.2. Nouvelle stratégie : ChangeTrackingMode + éditions atomiques

```
┌─────────────────────────────────────────────────────┐
│                 proposeRevision v2                    │
├─────────────────────────────────────────────────────┤
│                                                      │
│  1. Sauvegarder changeTrackingMode original          │
│  2. Activer changeTrackingMode = 'TrackAll'          │
│  3. Calculer le diff (original → revised)            │
│  4. Appliquer chaque opération atomiquement :        │
│     • Deletion → range.delete()                      │
│     • Insertion → range.insertText()                 │
│     • Unchanged → skip (préservé intact)             │
│  5. Restaurer changeTrackingMode original            │
│                                                      │
│  Résultat : Word crée des révisions NATIVES          │
│  avec auteur = compte Windows actif                  │
│  Visible dans le panneau "Révisions" de Word         │
└─────────────────────────────────────────────────────┘
```

### 3.3. Avantages vs approche actuelle (office-word-diff)

| Critère | office-word-diff (actuel) | ChangeTrackingMode (nouveau) |
|---|---|---|
| Track Changes natifs | ❌ Simule via styles CSS (rouge barré / vert souligné) | ✅ Vrais `<w:ins>` / `<w:del>` dans le document |
| Visible dans panneau Révisions | ❌ Non | ✅ Oui — accepter/rejeter individuellement |
| Préservation formatting | 🟡 Partiel (reconstruit les runs) | ✅ Total (ne touche que le texte modifié) |
| Auteur de révision | ❌ "KickOffice AI" en CSS | ✅ Compte Windows actif (natif) |
| Compatibilité Word Online | ✅ Fonctionne | 🟡 `changeTrackingMode` supporté (WordApi 1.4) mais vérifier |
| Complexité | 🟡 3 stratégies en cascade (token/sentence/block) | ✅ 1 seule stratégie (diff + apply atomique) |
| Dépendance npm | ❌ `office-word-diff` (package custom) | ✅ Aucune — tout natif Office.js |
| Performance | 🟡 Calcul diff + reconstruction runs | ✅ Diff léger + API native directe |

### 3.4. Algorithme détaillé de proposeRevision v2

```typescript
// PSEUDO-CODE — proposeRevision v2

async function applyRevisionV2(
  context: Word.RequestContext,
  revisedText: string,
  enableTrackChanges: boolean = true
): Promise<RevisionResult> {

  // 1. Obtenir le texte original de la sélection
  const selection = context.document.getSelection()
  selection.load('text')
  await context.sync()
  const originalText = selection.text

  // 2. Calculer le diff mot-par-mot
  //    Utiliser diff-match-patch (déjà en dépendance via office-word-diff)
  const dmp = new diff_match_patch()
  const diffs = dmp.diff_main(originalText, revisedText)
  dmp.diff_cleanupSemantic(diffs)

  // 3. Sauvegarder et activer Track Changes
  const doc = context.document
  doc.load('changeTrackingMode')
  await context.sync()
  const originalMode = doc.changeTrackingMode

  if (enableTrackChanges) {
    doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll
    await context.sync()
  }

  // 4. Appliquer les diffs de droite à gauche (pour préserver les offsets)
  //    Chaque opération diff est appliquée sur le range correspondant
  try {
    let cursor = selection.getRange('Start')

    for (const [op, text] of diffs) {
      if (op === 0) {
        // UNCHANGED — avancer le curseur de text.length caractères
        cursor = cursor.expandTo(/* advance by text.length */)
        cursor = cursor.getRange('End')
      } else if (op === -1) {
        // DELETION — sélectionner text.length chars et supprimer
        const deleteRange = cursor.expandToOrNullObject(/* text.length chars */)
        deleteRange.delete()
        await context.sync()
      } else if (op === 1) {
        // INSERTION — insérer le texte au curseur
        cursor = cursor.insertText(text, 'Before')
        await context.sync()
      }
    }
  } finally {
    // 5. TOUJOURS restaurer le mode original
    doc.changeTrackingMode = originalMode
    await context.sync()
  }

  return { success: true, strategy: 'native-track-changes', ... }
}
```

### 3.5. Gestion du champ "Redline Author"

**Changement par rapport à DESIGN_REVIEW.md** : le champ "Redline Author" configurable dans Settings n'est **plus nécessaire**.

Quand on utilise `changeTrackingMode = 'TrackAll'`, l'auteur des révisions est automatiquement le **compte Windows/Microsoft 365** de l'utilisateur connecté. C'est le comportement natif de Word — identique à ce qu'il se passe quand l'utilisateur active Track Changes manuellement.

**Avantage** : pas de configuration supplémentaire, l'auteur est authentique (pas un faux "KickOffice AI").

**Si le besoin d'un auteur custom revient** : ce n'est pas possible via `changeTrackingMode`. Il faudrait soit :
- Accepter l'auteur Windows natif (recommandé)
- Revenir à un mécanisme visuel (CSS-like) pour un auteur custom (pas recommandé — régression)

### 3.6. Diff engine : garder diff-match-patch, supprimer office-word-diff

`office-word-diff` est un wrapper autour de `diff-match-patch` avec 3 stratégies (token/sentence/block). Avec la nouvelle approche :

- **diff-match-patch** (déjà en dépendance : `"diff-match-patch": "^1.0.5"` dans `frontend/package.json`) suffit pour calculer les diffs
- **office-word-diff** n'est plus nécessaire — ses 3 stratégies de cascade n'ont plus de raison d'être car on ne reconstruit plus les runs manuellement

**Migration** :
1. Réécrire `wordDiffUtils.ts` pour utiliser `diff-match-patch` directement + `changeTrackingMode`
2. Supprimer la dépendance `office-word-diff` de `frontend/package.json`
3. Supprimer le dossier `/office-word-diff/`
4. Mettre à jour le Dockerfile si nécessaire

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

## 5. Migration : suppression de office-word-diff

### 5.1. Fichiers impactés

| Fichier | Action |
|---|---|
| `frontend/src/utils/wordDiffUtils.ts` | Réécrire — utiliser `diff-match-patch` + `changeTrackingMode` |
| `frontend/src/utils/wordTools.ts` (proposeRevision) | Mettre à jour l'appel vers le nouveau `applyRevisionV2` |
| `frontend/package.json` | Supprimer `"office-word-diff"` de dependencies |
| `office-word-diff/` (dossier entier) | Supprimer |
| `frontend/src/skills/word.skill.md` | Mettre à jour la description de `proposeRevision` |
| `Dockerfile` (si référence) | Vérifier et nettoyer |

### 5.2. Interface publique préservée

L'interface `RevisionResult` de `wordDiffUtils.ts` reste identique :

```typescript
export interface RevisionResult {
  success: boolean
  strategy: 'native-track-changes' | 'direct-replace'  // CHANGÉ: plus de token/sentence/block
  insertions: number
  deletions: number
  unchanged: number
  message: string
}
```

- `strategy` passe de `'token' | 'sentence' | 'block'` à `'native-track-changes' | 'direct-replace'`
- `'native-track-changes'` = quand enableTrackChanges=true et WordApi 1.4+ disponible
- `'direct-replace'` = quand enableTrackChanges=false ou API indisponible (fallback simple insertText)

### 5.3. Gestion du fallback (si WordApi 1.4 indisponible)

```typescript
function isChangeTrackingAvailable(context: Word.RequestContext): boolean {
  return Office.context.requirements.isSetSupported('WordApi', '1.4')
}
```

Si WordApi 1.4 n'est pas disponible (vieux Word, certaines versions Web) :
- **enableTrackChanges=true** → log un warning, faire un replace simple (pas de Track Changes)
- **enableTrackChanges=false** → comportement identique (replace simple)

Pas de fallback vers `office-word-diff` — on simplifie.

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

### Ordre d'exécution (3 étapes)

```
Étape 1 — DUP-M1 (15 min)
  └── Extraire truncateString dans common.ts
  └── Remplacer les 4 occurrences
  └── Tester que le build passe

Étape 2 — WORD-H1 (2-4 heures)
  ├── 2a. Réécrire wordDiffUtils.ts
  │   └── Nouveau applyRevisionV2 avec diff-match-patch + changeTrackingMode
  │   └── Fallback direct-replace si WordApi < 1.4
  │   └── Garder previewDiffStats() et computeRawDiff() (utilisent diff-match-patch directement)
  │
  ├── 2b. Mettre à jour proposeRevision dans wordTools.ts
  │   └── Appeler applyRevisionV2 au lieu de applyRevisionToSelection
  │   └── Mettre à jour la description du tool (mentionner Track Changes natif)
  │
  ├── 2c. Ajouter editDocumentXml dans wordTools.ts
  │   └── Nouveau tool pour manipulation OOXML directe
  │   └── Sandboxed eval avec DOMParser/XMLSerializer
  │
  ├── 2d. Mettre à jour word.skill.md
  │   └── Ajouter editDocumentXml dans le decision tree
  │   └── Mettre à jour la description de proposeRevision
  │
  ├── 2e. Supprimer office-word-diff
  │   └── Supprimer le dossier /office-word-diff/
  │   └── Retirer de frontend/package.json
  │   └── npm install pour mettre à jour le lockfile
  │
  └── 2f. Tests manuels
      └── Tester proposeRevision avec Track Changes
      └── Tester editDocumentXml avec un paragraphe formaté
      └── Vérifier le fallback sans WordApi 1.4

Étape 3 — Vérification (30 min)
  └── Build complet (npm run build)
  └── Vérifier aucune référence résiduelle à office-word-diff
  └── Mettre à jour DESIGN_REVIEW.md (marquer les 3 items comme FIXED)
```

### Fichiers créés / modifiés / supprimés

| Action | Fichier |
|---|---|
| **Modifier** | `frontend/src/utils/common.ts` — ajouter `truncateString()` |
| **Modifier** | `frontend/src/utils/wordTools.ts` — proposeRevision v2 + editDocumentXml + truncateString |
| **Modifier** | `frontend/src/utils/outlookTools.ts` — truncateString |
| **Réécrire** | `frontend/src/utils/wordDiffUtils.ts` — diff-match-patch + changeTrackingMode |
| **Modifier** | `frontend/src/skills/word.skill.md` — decision tree + descriptions |
| **Modifier** | `frontend/package.json` — retirer office-word-diff |
| **Supprimer** | `office-word-diff/` (dossier entier) |

---

## 8. Risques et mitigations

### 8.1. Risque : changeTrackingMode non disponible (vieux Word / Word Online)

| Risque | Impact | Mitigation |
|---|---|---|
| WordApi 1.4 indisponible | Pas de Track Changes natif | Fallback `direct-replace` : le texte est remplacé sans tracking. Message d'info à l'utilisateur |
| WordApi 1.4 dispo mais buggé | Comportement inattendu | `try/finally` pour toujours restaurer le mode original. Log exhaustif |

### 8.2. Risque : insertOoxml() ne marche pas sur Word Online

| Risque | Impact | Mitigation |
|---|---|---|
| `insertOoxml()` retourne "Not implemented" | `editDocumentXml` inutilisable | Détecter au runtime via `try/catch`. Retourner un message clair au LLM pour qu'il utilise un autre tool |
| XML corrompu après modification | Document cassé | Validation XML avant insertion. `DOMParser.parseFromString()` lève une erreur si le XML est invalide |

### 8.3. Risque : suppression de office-word-diff = régression

| Risque | Impact | Mitigation |
|---|---|---|
| Un cas edge que office-word-diff gérait | Édition incorrecte | `diff-match-patch` est le même engine sous-jacent. Le diff est identique, seule l'application change |
| Les 3 stratégies (token/sentence/block) étaient utiles | Perte de robustesse | La nouvelle stratégie est plus simple ET plus robuste car elle laisse Word gérer les runs en natif |

### 8.4. Risque : performance de getOoxml()

| Risque | Impact | Mitigation |
|---|---|---|
| `getOoxml()` ~6x plus lent que `body.text` | Latence perceptible | `editDocumentXml` travaille sur des **ranges** (pas le body entier), donc le volume XML est limité |
| XML volumineux pour de gros paragraphes | Overhead mémoire | Documenter dans le skill.md que `editDocumentXml` est pour des éditions ciblées, pas des documents entiers |

---

## 9. Sources

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

### Issues GitHub connues
- [#2123 — getOOXML trackrevision absent](https://github.com/OfficeDev/office-js/issues/2123) — `getOoxml()` ne retourne pas les éléments de revision
- [#329 — Expose TrackRevision via API](https://github.com/OfficeDev/office-js/issues/329) — feature request depuis 2018
- [#334 — Bug comment + selection OOXML](https://github.com/OfficeDev/office-js/issues/334) — `insertOoxml` + Track Changes = tout marqué comme nouveau
- [#3271 — insertOoxml not working Word Online](https://github.com/OfficeDev/office-js/issues/3271) — limitation Word Web
- [#5491 — insertOoxml style changes ignored](https://github.com/OfficeDev/office-js/issues/5491)
- [#2991 — Numbering changed on getOoxml/insertOoxml](https://github.com/OfficeDev/office-js/issues/2991)
- [#6246 — No API for revision display mode](https://github.com/OfficeDev/office-js/issues/6246) — Oct 2025
- [#3295 — Outlook MIME control](https://github.com/OfficeDev/office-js/issues/3295) — pas d'API MIME Outlook

### Codebase KickOffice (références internes)
- `frontend/src/utils/pptxZipUtils.ts` — pattern `withSlideZip` existant pour PowerPoint
- `frontend/src/utils/wordDiffUtils.ts` — wrapper `office-word-diff` actuel
- `frontend/src/utils/wordTools.ts` — tools Word dont `proposeRevision`, `eval_wordjs`
- `frontend/src/utils/sandbox.ts` — SES sandbox pour l'exécution sécurisée
- `frontend/src/skills/word.skill.md` — guidelines de sélection d'outils
- `office-word-diff/` — package custom à supprimer
