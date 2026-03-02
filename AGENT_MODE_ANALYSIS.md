# Analyse Architecturale : Pourquoi Open_Excel gère mieux les requêtes complexes

> Date : 2 Mars 2026  
> Périmètre : Architecture agent, prompts, outils, boucle d'agent — Applicable à Excel, Word, PowerPoint, Outlook

---

## Résumé Exécutif

L'analyse approfondie des deux codebases révèle que la supériorité d'Open_Excel pour les requêtes complexes (ex : « trace-moi des graphiques intéressants ») ne provient **pas** de la quantité d'outils (KickOffice en a 8x plus), mais de **9 différences architecturales fondamentales** dans la façon dont l'agent comprend le contexte, planifie ses actions et exécute de façon autonome.

> [!CAUTION]
> Ces problèmes sont **structurels** et affectent **toutes les applications Office** (Word, PowerPoint, Outlook), pas uniquement Excel.

---

## 1. 🔴 CRITIQUE — Pas d'injection automatique de contexte

### Le problème

**Open_Excel** appelle `getWorkbookMetadata()` **automatiquement** à chaque message utilisateur. Le modèle reçoit systématiquement :

```json
{
  "fileName": "Budget_2026.xlsx",
  "totalSheets": 4,
  "sheetsMetadata": [
    { "id": 1, "name": "Ventes", "maxRows": 150, "maxColumns": 8 },
    { "id": 2, "name": "Produits", "maxRows": 45, "maxColumns": 5 },
    ...
  ],
  "activeSheetId": 1,
  "activeSheetName": "Ventes",
  "selectedRange": "A1"
}
```

**KickOffice** n'injecte **aucun contexte automatique**. L'agent ne sait rien du document actuel — il ne connaît ni les feuilles, ni les données, ni les dimensions. Il doit deviner ce qu'il doit lire ou demander à l'utilisateur de sélectionner quelque chose.

### Impact

C'est la **cause principale** de la différence dans l'exemple « graphiques ». Quand l'utilisateur dit « trace-moi des graphiques » :

- **Open_Excel** : Le modèle voit 4 feuilles avec leurs dimensions → il appelle `get_cell_ranges` sur chaque feuille → il comprend les données → il crée 9 graphiques ciblés
- **KickOffice** : Le modèle ne sait rien → il demande de sélectionner → ou il appelle `getWorksheetData` sur la feuille active seulement

### Applicabilité multi-app

| App            | Contexte à injecter automatiquement                                                   |
| -------------- | ------------------------------------------------------------------------------------- |
| **Excel**      | Feuilles, dimensions, usedRange, sélection, noms définis                              |
| **Word**       | Nombre de pages, sections, style actif, position curseur, présence de tableaux/images |
| **PowerPoint** | Nombre de slides, slide active, titres/layout de chaque slide                         |
| **Outlook**    | Type d'item (lecture/compose), sujet, destinataires, corps actuel, pièces jointes     |

### Implémentation recommandée

Créer des fonctions `getDocumentMetadata()` par host, appelées automatiquement dans `sendMessage()`, injectées en `<doc_context>` dans le message utilisateur.

---

## 2. 🔴 CRITIQUE — Outils dépendants de la sélection vs basés sur des adresses

### Le problème

**Open_Excel** : Les outils utilisent des `sheetId` + `range` explicites. Le modèle spécifie exactement _où_ il opère :

```typescript
// Open_Excel: modify_object — source range explicite
modify_object({
  operation: "create",
  sheetId: 1,
  objectType: "chart",
  properties: {
    source: "A1:D50", // ← plage source explicite
    chartType: "line",
    anchor: "F1", // ← position du graphique
    title: "Tendances ventes",
  },
});
```

**KickOffice** : `createChart` opère sur `context.workbook.getSelectedRange()` — il utilise **la sélection Excel actuelle** :

```typescript
// KickOffice: createChart — dépend de la sélection
createChart({ chartType: "Line", title: "Trends" });
// ↪ Utilise getSelectedRange() en interne
```

### Impact

- **Problème séquentiel** : Quand l'agent veut créer 9 graphiques, chaque `createChart` opère sur la sélection. Après le 1er graphique, la sélection change (Excel sélectionne le graphique créé). Les graphiques suivants échouent ou sont vides.
- **Problème de ciblage** : Sans pouvoir spécifier `sheetId` + `source`, l'agent ne peut pas créer un graphique à partir de données d'une autre feuille ou plage non contiguë.

> [!IMPORTANT]
> KickOffice a un `range.select()` après `createChart` pour re-sélectionner la plage d'origine, mais cela ne résout pas le problème de fond : l'agent ne peut pas cibler des plages différentes pour chaque graphique.

### Outils concernés (Excel)

| Outil KickOffice   | Dépendance sélection           | Équivalent Open_Excel                     |
| ------------------ | ------------------------------ | ----------------------------------------- |
| `createChart`      | `getSelectedRange()`           | `modify_object` avec `source` explicite   |
| `getSelectedCells` | Sélection seule                | `get_cell_ranges` avec `sheetId`+`ranges` |
| `sortRange`        | `getSelectedRange()`           | N/A (via `eval_officejs`)                 |
| `formatRange`      | Optionnel (fallback sélection) | `set_cell_range` avec `sheetId`           |

### Applicabilité multi-app

Ce pattern se retrouve dans **tous les outils KickOffice** :

- **Word** : Outils insèrent au curseur actuel ; pas de ciblage par paragraphe/section
- **PowerPoint** : Outils opèrent sur la slide active ; pas de `slideIndex` paramètre
- **Outlook** : Moins impactant (item compose unique)

### Implémentation recommandée

Refactorer tous les outils avec un paramètre `address`/`range` explicite + `sheetId`/`slideIndex` quand applicable. L'agent doit pouvoir dire « écrire dans `Sheet2!B5:C20` » sans toucher à la sélection.

---

## 3. 🔴 CRITIQUE — `modifyObject` limité à la suppression

### Le problème

**Open_Excel** `modify_object` supporte 3 opérations : **create**, **update**, **delete** — pour les charts ET les pivot tables. C'est un outil unique et puissant pour toute manipulation d'objets.

**KickOffice** `modifyObject` ne supporte que la **suppression**. Pour créer un graphique, il y a `createChart` séparé (selection-dependent). **Pas d'update**, pas de pivot table en création directe.

### Impact direct

L'agent ne peut pas :

1. Modifier un graphique existant (changer le type, la source, le titre)
2. Créer un pivot table via un outil dédié
3. Créer un graphique avec une source précise (sans sélection)

### Implémentation recommandée

Remplacer `createChart` + `modifyObject` par un seul outil `manageObject` avec `operation: create|update|delete`, `objectType: chart|pivotTable`, et les `properties` appropriées (source, chartType, anchor, title, rows/columns/values pour pivots).

---

## 4. 🟠 HAUTE — Prompt système trop générique / pas de workflow agent

### Le problème

**Open_Excel** inclut dans son prompt système un **inventaire concis et structuré** de tous ses outils avec des exemples :

```
EXCEL READ:
- get_cell_ranges: Read cell values, formulas, and formatting
- get_range_as_csv: Get data as CSV (great for analysis)
- search_data: Find text across the spreadsheet
- get_all_objects: List charts, pivot tables, etc.

EXCEL WRITE:
- set_cell_range: Write values, formulas, and formatting
...
```

Plus des **instructions d'usage claires** :

- « When the user asks about their data, **read it first** »
- « Use csv-to-sheet over reading file content to avoid wasting tokens »

**KickOffice** prompt :

- Pas d'inventaire des outils (le LLM ne les voit que via les `tool definitions` JSON)
- Instructions vagues : « Tool First: Always use the available tools »
- Pas de workflow de découverte (« lis d'abord, agis ensuite »)
- Pas de stratégie de batch (mentionne `batchSetCellValues` mais pas de workflow complet)

### Impact

Le modèle ne sait pas quels outils combiner ni dans quel ordre. Pour une requête complexe comme « analyse mes données et crée des graphiques », il ne sait pas qu'il doit :

1. D'abord appeler `getWorksheetInfo` pour découvrir les feuilles
2. Puis `getWorksheetData` / `getDataFromSheet` pour lire les données
3. Puis créer des graphiques ciblés

### Applicabilité multi-app

| App            | Workflow agent manquant                                             |
| -------------- | ------------------------------------------------------------------- |
| **Excel**      | Lecture → Analyse → Action (graphiques, formules, formatage)        |
| **Word**       | Lecture sections → Compréhension structure → Modification ciblée    |
| **PowerPoint** | Inventaire slides → Compréhension contenu → Génération/modification |
| **Outlook**    | Lecture email → Analyse contexte → Draft/Reply pertinent            |

### Implémentation recommandée

1. **Enrichir les prompts système** par host avec un inventaire des outils et des workflows typiques
2. Ajouter des « agent directives » : « Toujours lire le contexte avant d'agir », « Pour les graphiques, d'abord lire les données avec getWorksheetData puis utiliser createChart avec la source appropriée »
3. Ajouter des exemples de workflows multi-étapes dans le prompt

---

## 5. 🟠 HAUTE — Pas de `get_all_objects` de découverte globale

### Le problème

**Open_Excel** : `get_all_objects` liste **tous les charts et pivot tables** du workbook (ou d'une feuille spécifique), retournant `id`, `type`, `name`, `sheetId`, `sheetName`.

**KickOffice** : `getAllObjects` ne fonctionne que sur la **feuille active** et retourne seulement `name` et `id` — pas de `sheetId`, pas de scope workbook.

### Impact

L'agent ne peut pas faire un inventaire complet de ce qui existe déjà dans le classeur, ce qui empêche des actions comme « mets à jour tous mes graphiques », « supprime les graphiques obsolètes », etc.

### Applicabilité multi-app

- **PowerPoint** : Besoin d'un `getAllSlides` retournant layout/titre par slide
- **Word** : Besoin d'un `getDocumentStructure` retournant headings/tables/images

---

## 6. 🟡 MOYENNE — Boucle agent hand-rolled vs libraire dédiée

### Le problème

**Open_Excel** utilise `@mariozechner/pi-agent-core`, une bibliothèque d'agent dédiée qui gère bien la boucle outil et le streaming.

**KickOffice** a une boucle écrite à la main. Bien que fonctionnelle, elle doit s'assurer de transporter bien la structure complète `{role, content, tool_calls}` pour les messages assistant, et `{role: 'tool', tool_call_id, content}` pour les résultats.

### Implémentation recommandée

S'assurer que la sérialisation vers le backend préserve l'historique complet des appels d'outils et de leurs résultats.

---

## 7. 🟡 MOYENNE — Pas de « read first » automatique pour le mode agent

### Le problème

Même avec un outil `getWorksheetData` disponible, l'agent KickOffice ne l'appelle pas systématiquement avant d'agir. Open_Excel force cela via l'injection de `wb_context` et des directives claires.

### Implémentation recommandée

Ajouter dans le prompt système une directive claire : « **TOUJOURS** commencer par lire le contexte du document courant avec les outils de lecture avant d'effectuer des modifications. »

---

## 8. 🟠 HAUTE — Trop d'outils dédiés, pas assez de scripting agent

### Le constat

Open_Excel n'a que **~15 outils dédiés** + `eval_officejs` comme « escape hatch ». KickOffice a **116+ outils**. Cette prolifération a 2 effets négatifs :

1. **Token cost** : Chaque outil envoyé coûte des tokens (~15K tokens perdus à chaque requête avec 116 outils).
2. **Decision fatigue** : Le modèle a trop de choix.

### Stratégie : « Keep or Script »

L'idée est d'identifier quels outils **garder comme outils dédiés** et lesquels l'agent devrait plutôt réaliser via les outils `eval_*`.

| App            | Stratégie                                                                                                                      |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| **Excel**      | Garder ~18 outils (lecture, écriture valeurs, structure), scripter ~24 outils (formatage spécifique, lignes/colonnes simples)  |
| **Word**       | Garder ~15 outils (lecture, insertion texte/table, styles), scripter ~22 outils (formatage paragraphe, sauts de page, signets) |
| **PowerPoint** | Garder ~7 outils (lecture, gestion slides, texte), scripter ~7 outils (formes, images, notes)                                  |
| **Outlook**    | Garder ~8 outils (lecture email, sujet, corps, destinataires), scripter ~5 outils (pièces jointes, headers)                    |

---

## 9. 🔴 CRITIQUE — Problèmes de mise en forme de texte (surtout PowerPoint)

### PowerPoint — La plus fragile

Le pipeline de mise en forme PowerPoint a **4 problèmes majeurs** :

1. **`addTextBox`** : Insère du texte brut puis _essaie_ un upgrade HTML (échoue silencieusement si API < 1.5).
2. **`replaceSelectedText`** : Utilise `Office.context.document.setSelectedDataAsync` (Common API) dont le support HTML est limité et écrase souvent les styles existants.
3. **Lecture char-par-char** : `getPowerPointSelectionAsHtml` lit 1000 caractères via 1000 appels API, provoquant des lenteurs extrêmes.
4. **Pas d'outil `setSlideText`** : Tout dépend de la sélection utilisateur.

### Word — Plus robuste mais des fragilités

Bien que meilleur, le formatage Word dépend encore trop de la sélection, empêchant l'agent de formater un paragraphe spécifique sans intervention utilisateur.

### Implémentation recommandée

| App            | Fix                                                                                                                   | Effort |
| -------------- | --------------------------------------------------------------------------------------------------------------------- | ------ |
| **PowerPoint** | Utiliser `textRange.insertHtml()` (Modern API) en principal, ajouter `setShapeText(slideNumber, shapeIdOrName, text)` | Moyen  |
| **PowerPoint** | Optimiser la lecture HTML (par paragraphe au lieu de char-par-char)                                                   | Moyen  |
| **Word**       | Ajouter `paragraphIndex` aux outils de formatage                                                                      | Faible |
| **Outlook**    | Ajouter `appendToEmailBody` (insert sans écraser)                                                                     | Faible |

---

## Tableau Récapitulatif des Priorités

| #   | Problème                            | Criticité   | Apps concernées    | Effort |
| --- | ----------------------------------- | ----------- | ------------------ | ------ |
| 1   | Injection auto de contexte document | 🔴 CRITIQUE | Toutes             | Moyen  |
| 2   | Outils dépendants de la sélection   | 🔴 CRITIQUE | Excel, Word, PPT   | Élevé  |
| 3   | `modifyObject` create/update/delete | 🔴 CRITIQUE | Excel              | Moyen  |
| 4   | Prompt système enrichi + workflows  | 🟠 HAUTE    | Toutes             | Faible |
| 5   | `getAllObjects` global discovery    | 🟠 HAUTE    | Excel, PPT, Word   | Faible |
| 6   | Boucle agent robuste                | 🟡 MOYENNE  | Toutes             | Moyen  |
| 7   | Directive « read first »            | 🟡 MOYENNE  | Toutes             | Faible |
| 8   | Trop d'outils dédiés                | 🟠 HAUTE    | Toutes             | Moyen  |
| 9   | Mise en forme texte cassée          | 🔴 CRITIQUE | PPT, Word, Outlook | Moyen  |

---

## Plan d'implémentation recommandé

### Phase 1 — Quick Wins (1-2 jours)

1. Enrichir les prompts systèmes.
2. Implémenter l'injection automatique de contexte.

### Phase 2 — Refactoring outils (3-5 jours)

3. Fusioner `createChart` et `modifyObject` dans `manageObject`.
4. Passer à des outils adressables (`address`, `range`, `sheetId`).
5. Enrichir `getAllObjects` pour un scope workbook.

### Phase 3 — Découverte multi-app (3-5 jours)

6. Créer des outils de découverte de structure pour Word, PowerPoint et Outlook.

### Phase 4 — Fiabilité du formatage (2-3 jours)

7. Corriger le pipeline de mise en forme PowerPoint et rendre les outils Word non dépendants de la sélection.

---

## Conclusion

La force d'Open_Excel réside dans **l'intelligence de son architecture** plutôt que dans le nombre de ses outils. En adoptant l'auto-injection de contexte, des outils adressables et des prompts guidant le workflow, KickOffice deviendra un agent proactif capable de transformer radicalement l'expérience utilisateur sur toute la suite Office.

---

## Suivi d'implémentation

> Dernière mise à jour : 2 Mars 2026 — Branch `claude/implement-agent-mode-analysis-kesh2`

### Tableau de statut

| #   | Problème                            | Criticité   | Statut          | Commit(s) |
| --- | ----------------------------------- | ----------- | --------------- | --------- |
| 1   | Injection auto de contexte document | 🔴 CRITIQUE | ✅ Implémenté   | `1af3379`, `da09ee5` |
| 2   | Outils dépendants de la sélection   | 🔴 CRITIQUE | ✅ Implémenté   | `901f47f`, `2450a84`, `da09ee5` |
| 3   | `modifyObject` create/update/delete | 🔴 CRITIQUE | ✅ Implémenté   | `901f47f` |
| 4   | Prompt système enrichi + workflows  | 🟠 HAUTE    | ✅ Implémenté   | `ce070e0`, `7a088d3` |
| 5   | `getAllObjects` global discovery    | 🟠 HAUTE    | ✅ Implémenté   | `901f47f` |
| 6   | Boucle agent robuste                | 🟡 MOYENNE  | ✅ Vérifié      | (existant) |
| 7   | Directive « read first »            | 🟡 MOYENNE  | ✅ Implémenté   | `ce070e0`, `7a088d3` |
| 8   | Trop d'outils dédiés                | 🟠 HAUTE    | ✅ Implémenté   | `7a088d3` |
| 9   | Mise en forme texte cassée          | 🔴 CRITIQUE | ✅ Implémenté   | `2450a84`, `0b8977f`, `da09ee5` |

### Détail par point

---

#### ✅ Point 1 — Injection automatique de contexte

**Fichier :** `frontend/src/utils/officeDocumentContext.ts`
**Injection :** `frontend/src/composables/useAgentLoop.ts` (lignes ~601-614)

Quatre fonctions créées, appelées automatiquement avant chaque requête LLM :

| App            | Données injectées dans `<doc_context>`                                                          |
| -------------- | ----------------------------------------------------------------------------------------------- |
| **Excel**      | `activeSheet`, `selectedRange`, `totalSheets`, `sheets[]` (name, rows, columns)                |
| **PowerPoint** | `totalSlides`, `slides[]` (slideNumber, title)                                                  |
| **Outlook**    | `itemType` (compose/read), `subject`, `sender`, `recipients`, `bodySnippet`                     |
| **Word**       | `pageCount`, `wordCount`, `paragraphCount`, `tableCount`, `hasImages`                           |

---

#### ✅ Point 2 — Outils adressables (non dépendants de la sélection)

- **Excel** : `createChart` + `modifyObject` → remplacés par `manageObject(sheetName, source)` — cible n'importe quelle feuille/plage sans toucher à la sélection.
- **PowerPoint** : `setShapeText(slideNumber, shapeIdOrName)` — cible une forme par ID/nom sans sélection utilisateur.
- **Word** : `applyStyle(paragraphIndex)` — cible un paragraphe par index sans dépendre du curseur.

---

#### ✅ Point 3 — `manageObject` create / update / delete

**Fichier :** `frontend/src/utils/excelTools.ts`

Outil unique `manageObject` avec :
- `operation : 'create' | 'update' | 'delete'`
- `objectType : 'chart' | 'pivotTable'`
- Paramètres explicites : `sheetName`, `source`, `chartType`, `title`, `anchor`, `name`

Remplace complètement `createChart` (dépendant sélection) et `modifyObject` (suppression seule).

---

#### ✅ Point 4 — Prompts système enrichis + workflows agent

**Fichier :** `frontend/src/composables/useAgentPrompts.ts`

Les 4 prompts hôtes ont été réécrits avec :
- Section **"Agent Workflow — ALWAYS Follow This Order"** : séquence read → explore → act explicite
- **Tool Inventory** complet par catégorie (READ / WRITE / FORMAT / ADVANCED)
- Mention des opérations couvertes par `eval_*` pour guider le modèle sur l'escape hatch
- Directives de batch, de formatage Markdown et de langue

---

#### ✅ Point 5 — `getAllObjects` scope workbook

**Fichier :** `frontend/src/utils/excelTools.ts`

Paramètre `allSheets` (défaut `true`) : scan workbook-wide en 2 syncs groupés.
Retourne pour chaque objet : `name`, `id`, `sheetName` — permettant d'agir sur des graphiques de n'importe quelle feuille.

---

#### ✅ Point 6 — Boucle agent robuste

**Fichier :** `frontend/src/composables/useAgentLoop.ts`

Sérialisation conforme au format OpenAI :
- Messages assistant : `{ role: 'assistant', content, tool_calls[] }`
- Résultats d'outils : `{ role: 'tool', tool_call_id, content }`

Protections supplémentaires : détection de boucle (même signature d'outil → message d'erreur), rollback atomique si abandon en cours d'exécution.

---

#### ✅ Point 7 — Directive « read first »

Présente dans les 4 prompts hôtes :
- **Excel** : « Read doc_context before calling any tool »
- **Word** : « ALWAYS start by reading the document »
- **PowerPoint** : « ALWAYS call getAllSlidesOverview first »
- **Outlook** : « Read doc_context block before calling any tool »

---

#### ✅ Point 8 — Réduction du nombre d'outils (116 → 57)

**Commit :** `7a088d3`

| App            | Avant | Après | Réduction |
| -------------- | ----- | ----- | --------- |
| **Excel**      | 43    | 21    | −51%      |
| **Word**       | 42    | 17    | −60%      |
| **PowerPoint** | 16    | 9     | −44%      |
| **Outlook**    | 15    | 10    | −33%      |
| **Total**      | **116** | **57** | **−51%** |

Ajout de **`eval_officejs`** pour Excel (manquant) — escape hatch couvrant : tri, filtres, gel, validation, formats numériques, liens, commentaires, lignes/colonnes, protection, etc.

Outils supprimés routés vers `eval_*` (exemples) :
- **Excel** : `insertRow/Column`, `sortRange`, `freezePanes`, `addDataValidation`, `setCellNumberFormat`, `addHyperlink`, `autoFitColumns`, `duplicateWorksheet`…
- **Word** : `insertParagraph`, `setFontName`, `insertPageBreak`, `insertBookmark`, `insertHeaderFooter`, `setParagraphFormat`…
- **PPT** : `setSlideNotes`, `insertTextBox`, `insertImage`, `deleteShape`, `setShapeFill`, `moveResizeShape`
- **Outlook** : `getEmailDate`, `getAttachments`, `insertHtmlAtCursor`, `setEmailBodyHtml`

---

#### ✅ Point 9 — Mise en forme texte corrigée

**PowerPoint — lecture HTML** (`powerpointTools.ts`) :
- Avant : 1 000 appels API pour lire 1 000 caractères (char-by-char)
- Après : chargement par paragraphe en batch (2 syncs pour toute la sélection)

**PowerPoint — `setShapeText`** (`powerpointTools.ts`) :
- Nouvel outil ciblant une forme par `slideNumber` + `shapeIdOrName`
- Utilise `textRange.insertHtml()` (API 1.5+) avec fallback texte brut

**Outlook — `appendToEmailBody`** (`outlookTools.ts`) :
- Nouveau : lit le corps existant, ajoute le nouveau contenu à la fin (non destructif)
- Supporte le Markdown + `DOMPurify.sanitize` pour la sécurité

**Word — `applyStyle`** (bonus, `wordTools.ts`) :
- Ajout de `paragraphIndex` : formate un paragraphe ciblé sans dépendre de la sélection
