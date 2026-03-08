# DESIGN_REVIEW.md — Code Audit v8

**Date**: 2026-03-08
**Version**: 8.0
**Scope**: Revue utilisateur — Bugs fonctionnels, UX/UI, Actions rapides, Coherence inter-applications

---

## Etat de sante global

Revue basee sur les retours utilisateur couvrant Word, Excel, PowerPoint, Outlook et les elements transverses.

**Verification :** `npx tsc --noEmit` → 0 erreur ✅ | **Tests :** `npm run test:unit` → all pass ✅

| Severite    | Total    | OPEN     | DONE   |
| ----------- | -------- | -------- | ------ |
| CRITIQUE    | 3        | 0        | 3      |
| MAJEUR      | 6        | 0        | 6      |
| MINEUR      | 3        | 0        | 3      |
| **Total**   | **12**   | **0**    | **12** |

---

## Résumé des Racines Traitées

### Loop Detection — Correction de la Compréhension

**Problème initial identifié** : J'avais ajouté `trackToolCallAndCheckStuck` basé sur le nom du tool seul (sans arguments), ce qui cassait les opérations légitimes multi-cellules/multi-shapes (Excel: `setCellRange(A1), setCellRange(B1), setCellRange(C1)` → bloquerait à tort).

**Solution correcte** : Le mécanisme `addSignatureAndCheckLoop` (basé sur `toolName + JSON.stringify(args)`) est déjà optimal et suffisant :
- Détecte les vraies boucles (même outil, mêmes args répétés)
- Ne bloque PAS les itérations légitimes (mêmes outil, args différents)
- `trackToolCallAndCheckStuck` supprimé (était redondant + nocif)

**Vérification par scénario** :
| Cas | Signature | Comportement |
|---|---|---|
| `getAllSlidesOverview({})` × 2 | Identique | Trigger après 2 calls ✅ |
| `setCellRange(addr="A1")` puis `setCellRange(addr="B1")` | Différentes | Pas de trigger ✅ |
| `insertContent(...)` × 3 slides | Différentes | Pas de trigger ✅ |

---

## ITEMS CRITIQUES

### PPT-C1 ✅ DONE — Boucle infinie lors de la generation de slide a partir d'une image

**Fichiers concernes:**
- `frontend/src/utils/powerpointTools.ts` (lignes 872-914) — `getAllSlidesOverview`
- `frontend/src/composables/useAgentLoop.ts` (lignes 650-703) — boucle agent
- `frontend/src/composables/useLoopDetection.ts` (lignes 1-22)

**Probleme:** Quand l'utilisateur demande de creer une slide a partir d'une image, l'agent appelle `getAllSlidesOverview` en boucle. Le premier appel reussit, les suivants retournent des donnees vides. La detection de boucle ne capte pas la repetition (signatures legerement differentes a chaque fois) et le nombre de requetes finit par faire couper la connexion.

**Cause racine:** `getAllSlidesOverview` fait 2 `context.sync()` par slide dans une boucle (lignes 893, 899), ce qui est couteux. L'agent ne sait pas quoi faire avec l'image et re-tente le meme workflow de decouverte en boucle.

**Strategie validee — Option C (les deux) :**

1. **Detection de boucle (filet de securite)** : Renforcer `useLoopDetection.ts` pour detecter les appels repetitifs au meme outil (meme outil appele 3+ fois consecutivement = arreter la boucle et afficher un message d'erreur explicite).
2. **Corriger le workflow image PowerPoint** : Modifier `powerpoint.skill.md` pour donner a l'agent des instructions claires sur comment gerer une image (creer une slide, inserer l'image via un outil dedie, ne PAS boucler sur `getAllSlidesOverview`).

---

### XL-C1 ✅ DONE — Generation de graphiques : etiquettes d'abscisses traitees comme serie de donnees

**Fichiers concernes:**
- `frontend/src/utils/excelTools.ts` (ligne 429) — `manageObject`

**Probleme:** Lors de la creation d'un graphique, le code utilise :
```javascript
const chart = sheet.charts.add(excelChartType, dataRange, Excel.ChartSeriesBy.auto)
```
Le mode `auto` laisse Excel deviner comment interpreter les donnees. La premiere colonne/ligne (souvent les etiquettes d'abscisses) est interpretee comme une serie de donnees supplementaire au lieu d'etre utilisee comme etiquettes d'axe.

**Strategie validee — Option A (parametre explicite) :**

1. Ajouter les parametres `seriesBy` (`'columns'` | `'rows'`) et `hasHeaders` (`boolean`) a l'outil `manageObject` pour la creation de graphiques.
2. L'agent choisit les valeurs en fonction du contexte des donnees selectionnees.
3. Apres creation, si `hasHeaders=true`, extraire la premiere ligne/colonne pour `chart.axes.categoryAxis.setCategoryNames()` et la retirer des series de donnees.
4. Mettre a jour `excel.skill.md` pour documenter ces parametres et guider l'agent.

---

### WD-C1 ✅ DONE — Crash formatText sur texte insere depuis un PDF (pas de selection)

**Fichiers concernes:**
- `frontend/src/utils/wordTools.ts` (lignes 248-304) — `formatText`
- `frontend/src/utils/wordTools.ts` (lignes 178-244) — `insertContent`
- `frontend/src/skills/word.skill.md` (lignes 30-34) — regle critique
- `frontend/src/utils/wordTools.ts` (lignes 355-461) — `searchAndFormat`

**Probleme:** `formatText` requiert une selection active (`context.document.getSelection()`). Quand du texte est insere via `insertContent`, il n'est PAS selectionne. L'agent tente quand meme `formatText` → crash. Le skill documente 3 workflows alternatifs (A: syntax inline, B: applyTaggedFormatting, C: searchAndFormat) mais l'agent ne les suit pas systematiquement.

**Cause racine:** `insertContent` a acces a `insertedRange` (ligne 230) mais n'appelle jamais `insertedRange.select()` pour le selectionner. Il n'existe aucun outil `selectText`/`selectRange`.

**Strategie validee — Option A + C (auto-select + prompt renforcé) :**

1. **Auto-select apres insertion** : Dans `insertContent` (ligne 230 de `wordTools.ts`), ajouter `insertedRange.select()` apres `context.sync()` pour que le texte insere soit automatiquement selectionne. `formatText` fonctionnera directement apres.
2. **Renforcement du prompt agent** : Dans `word.skill.md` et `useAgentPrompts.ts`, ajouter une regle explicite : "Apres `insertContent`, le texte est auto-selectionne. Tu peux utiliser `formatText` directement. Pour du formatage cible sur du texte existant, utilise `searchAndFormat`."

---

## ITEMS MAJEURS

### GEN-M1 ✅ DONE — Taille du taskpane bloquee a 300px (manifest desynchronise)

**Fichiers concernes:**
- `manifests-templates/manifest-office.template.xml` (ligne 31) — `<RequestedWidth>450</RequestedWidth>` ✅
- `manifest-office.xml` (lignes 29-31) — MANQUE `<RequestedWidth>` ❌

**Probleme:** Le template a correctement `RequestedWidth=450` mais le manifest racine `manifest-office.xml` utilise en dev/production n'a PAS cette propriete. Sans elle, Office utilise la valeur par defaut de 300px.

**Strategie:** Synchroniser `manifest-office.xml` avec le template. Ajouter `<RequestedWidth>450</RequestedWidth>` dans `<DefaultSettings>`. Verifier aussi que le script `scripts/generate-manifests.js` est utilise dans le workflow de deploiement (actuellement `generated-manifests/` est vide).

---

### GEN-M2 ✅ DONE — Bouton bug/feedback : traductions manquantes et taille excessive

**Fichiers concernes:**
- `frontend/src/components/settings/GeneralTab.vue` (lignes 120-131)
- `frontend/src/components/settings/FeedbackDialog.vue` (lignes 1-117)
- `frontend/src/i18n/locales/en.json` — cles manquantes
- `frontend/src/i18n/locales/fr.json` — cles manquantes

**Probleme:** 13 cles i18n manquantes : `reportBugOrFeedback`, `feedbackButtonText`, `feedbackTitle`, `feedbackSuccess`, `feedbackCategory`, `feedbackBug`, `feedbackFeature`, `feedbackOther`, `feedbackComment`, `feedbackPlaceholder`, `feedbackIncludeLogs`, `submitting`, `submit`. Le bouton affiche les cles brutes au lieu du texte traduit. Le bouton utilise `min-w-fit` ce qui le rend trop large.

**Strategie:** Ajouter les 13 cles dans `en.json` et `fr.json`. Remplacer `min-w-fit` par une largeur max sur le bouton dans `GeneralTab.vue`.

---

### XL-M1 ✅ DONE — Actions rapides Excel : comportement incoherent et sans tooltips

**Fichiers concernes:**
- `frontend/src/pages/HomePage.vue` (lignes 315-356) — definitions
- `frontend/src/composables/useAgentLoop.ts` (lignes 861-874) — execution mode `draft`
- `frontend/src/i18n/locales/en.json` / `fr.json` — tooltips manquants

**Probleme:** 4 des 5 actions rapides Excel utilisent `mode: 'draft'` (remplissent juste le champ chat) alors que Word et PowerPoint utilisent `mode: 'immediate'` avec `executeWithAgent: true` (execution directe). Les 5 tooltips sont manquants (`excelIngest_tooltip`, `excelAutoGraph_tooltip`, `excelExplain_tooltip`, `excelFormulaGenerator_tooltip`, `excelDataTrend_tooltip`).

**Strategie validee :**

| Action | Mode actuel | Nouveau mode | Justification |
|--------|------------|-------------|---------------|
| `excelIngest` (Analyser) | `immediate` ✅ | Garder `immediate` | Pas d'input necessaire |
| `excelAutoGraph` (Graphique auto) | `draft` ❌ | → `immediate` + `executeWithAgent: true` | Pas d'input — genere depuis selection |
| `excelExplain` (Expliquer) | `draft` ❌ | → `immediate` + `executeWithAgent: true` | Pas d'input — explique la selection |
| `excelFormulaGenerator` (Formule) | `draft` | Garder `draft` | **Input requis** — l'utilisateur decrit la formule |
| `excelDataTrend` (Tendances) | `draft` ❌ | → `immediate` + `executeWithAgent: true` | Pas d'input — analyse les tendances |

Ajouter les 5 tooltips manquants dans `en.json` et `fr.json`.

---

### XL-M2 ✅ DONE — Langue des formules Excel non utilisee partout

**Fichiers concernes:**
- `frontend/src/utils/excelTools.ts` (ligne 50-53) — `getExcelFormulaLanguage()`
- `frontend/src/composables/useAgentPrompts.ts` (lignes 28-30, 170)
- `frontend/src/utils/excelTools.ts` (lignes 1072-1240) — `applyConditionalFormatting` ❌
- `frontend/src/utils/excelTools.ts` (lignes 1463+) — `eval_officejs` ❌

**Probleme:** Le parametre "Langue des formules Excel" est utilise dans `setCellRange` et le prompt agent, MAIS pas dans :
1. `applyConditionalFormatting` — accepte `formula1`/`formula2` sans conversion locale
2. `eval_officejs` — pas de contexte de langue passe a l'environnement d'execution
3. Les instructions shell communes (`COMMON_SHELL_INSTRUCTIONS`)

**Strategie:** Injecter `getExcelFormulaLanguage()` dans `applyConditionalFormatting` pour utiliser `formulasLocal` si francais. Ajouter un commentaire/contexte dans `eval_officejs` pour que l'agent genere des formules dans la bonne langue. Mentionner la langue dans les shell instructions.

---

### PPT-M1 ✅ DONE — L'agent n'utilise pas les boites titre/corps du template

**Fichiers concernes:**
- `frontend/src/utils/powerpointTools.ts` (lignes 758-791) — `addSlide`
- `frontend/src/skills/powerpoint.skill.md` (lignes 99-105)

**Probleme:** `addSlide` cree une slide avec un layout mais ne peuple PAS les boites de texte du template (Titre, Corps, Sous-titre). L'agent doit ensuite appeler `getShapes()` pour decouvrir les IDs des shapes puis `insertContent` pour chacune — workflow non documente et rarement suivi.

**Strategie:** Modifier `addSlide` pour accepter des parametres optionnels `title` et `body`. Si fournis, le tool decouvre automatiquement les shapes du layout et y insere le contenu. Mettre a jour `powerpoint.skill.md` pour documenter ce workflow.

---

### PPT-M2 ✅ DONE — Notes d'orateur : l'action rapide ne les insere pas dans la slide

**Fichiers concernes:**
- `frontend/src/pages/HomePage.vue` (lignes 407-413)
- `frontend/src/utils/constant.ts` (lignes 261-276) — prompt speakerNotes
- `frontend/src/utils/powerpointTools.ts` (lignes 459-490) — `setSpeakerNotes`

**Probleme:** L'action rapide "Notes d'orateur" genere le texte dans le chat mais ne l'insere PAS dans les notes de la slide. Pourtant l'outil `setSpeakerNotes` existe et peut le faire.

**Strategie validee — Appel direct (rapide) :**

1. Modifier l'action rapide `speakerNotes` pour qu'elle genere les notes via LLM (prompt existant dans `constant.ts`).
2. Apres generation, appeler directement `setSpeakerNotes` (lignes 459-490 de `powerpointTools.ts`) avec le contenu genere, sans passer par la boucle agent.
3. Afficher un message de confirmation dans le chat une fois les notes inserees.

---

## ITEMS MINEURS

### PPT-L1 ✅ DONE — Action rapide "Impact" pas adaptee a PowerPoint

**Fichiers concernes:**
- `frontend/src/utils/constant.ts` (lignes 278-294) — prompt `punchify`
- `frontend/src/skills/powerpoint.skill.md` (lignes 56-74)

**Probleme:** Le prompt de l'action "Impact" est oriente copywriting general. Il ne mentionne pas : format puces, max 8-10 mots par puce, max 6-7 puces par slide.

**Strategie:** Mettre a jour le prompt `punchify` dans `constant.ts` pour ajouter des contraintes PowerPoint : format bullet points, concision, limites de puces. Aligner avec les bonnes pratiques documentees dans `powerpoint.skill.md`.

---

### GEN-L1 ✅ DONE — Checkboxes du chat : option "Formatage Word" potentiellement confuse

**Fichiers concernes:**
- `frontend/src/components/chat/ChatInput.vue` (lignes 108-136)
- `frontend/src/composables/useOfficeInsert.ts` (ligne 179)

**Probleme:** La checkbox "Formatage Word" ne controle PAS ce que l'IA genere mais seulement COMMENT le resultat est insere (HTML formate vs texte brut). Le nom prete a confusion. Les autres checkboxes ("Inclure le texte selectionne/cellules/diapo") sont claires et utiles.

**Strategie validee — Option A (renommer) :**

1. Renommer le label dans `en.json` : `useWordFormattingLabel` → "Insert with formatting"
2. Renommer le label dans `fr.json` : `useWordFormattingLabel` → "Inserer avec mise en forme"
3. Garder le comportement fonctionnel identique (controle du mode d'insertion HTML vs plain text).

---

### GEN-L2 ✅ DONE — Tooltips manquants sur les actions rapides (multi-apps)

**Fichiers concernes:**
- `frontend/src/i18n/locales/en.json`
- `frontend/src/i18n/locales/fr.json`

**Probleme:** Tooltips manquants pour :
- Excel : 5 tooltips (`excelIngest_tooltip`, `excelAutoGraph_tooltip`, `excelExplain_tooltip`, `excelFormulaGenerator_tooltip`, `excelDataTrend_tooltip`)
- Word : 1 tooltip (`summary_tooltip`)
- Outlook : 4 tooltips (`outlookProofread_tooltip`, `outlookConcise_tooltip`, `outlookExtract_tooltip`, `outlookReply_tooltip`)

**Strategie:** Ajouter les 10 cles manquantes dans `en.json` et `fr.json` avec des descriptions explicatives de chaque action.

---

## DEFERRED ITEMS (reconduits des revues precedentes)

- **IC2** — Containers run as root (low priority).
- **IH2** — Private IP in build arg.
- **IH3** — DuckDNS domain in example.
- **UM10** — PowerPoint HTML reconstruction (high complexity).

---

## Changelog

| Version | Date       | Changes                                                                                |
| ------- | ---------- | -------------------------------------------------------------------------------------- |
| v8.1    | 2026-03-08 | Implementation complete des 12 findings v8 (0 open). npx tsc → 0 erreur.               |
| v8.0    | 2026-03-08 | Nouvelle revue utilisateur — 12 findings (3 critiques, 6 majeurs, 3 mineurs).          |
| v7.2    | 2026-03-08 | Finalisation UX/UI, Code mort, Generalisation et Qualite de code (Axe 4 a 7). PR #178  |
| v7.1    | 2026-03-08 | Resolution des 5 critiques + bonus.                                                    |
| v7.0    | 2026-03-08 | Revue 7 axes (34 findings).                                                            |
