# KickOffice — Audit des Skills Agent Manquantes

> Date: 2026-02-16
> Objectif: Identifier toutes les skills manquantes pour que l'agent soit efficace dans les 4 outils Office

## Etat actuel

| Application    | Tools existants | Tools manquants identifies |
|----------------|-----------------|----------------------------|
| **Word**       | 39 tools        | 0 manquant (lot W1-W15 livre) |
| **Excel**      | 39 tools        | 0 manquant sur le lot E1-E18 |
| **PowerPoint** | 0 tools (!)     | ~8 a creer                 |
| **Outlook**    | 3 tools         | ~10 manquants              |

---

## 1. WORD (wordTools.ts) — 39 existants, 0 manquant sur le lot prioritaire

### Existant
- Texte: getSelectedText, getDocumentContent, insertText, replaceSelectedText, appendText, deleteText, selectText, findText, searchAndReplace
- Formatage inline: formatText, clearFormatting, setFontName, applyTaggedFormatting, getRangeInfo
- Structure: insertParagraph, insertList, insertTable, getTableInfo, insertPageBreak
- Navigation: insertBookmark, goToBookmark, insertContentControl
- Image: insertImage

### Lot Word integre (W1-W15)

| ID  | Skill               | Statut | Note implementation |
|-----|---------------------|--------|---------------------|
| W1  | setParagraphFormat  | Livre  | Formatage paragraphe sur selection (`alignment`, `lineSpacing`, `spaceBefore/After`, `leftIndent`, `firstLineIndent`) |
| W2  | insertHyperlink     | Livre  | Insertion de lien cliquable via `range.hyperlink` (+ `textToDisplay` optionnel) |
| W3  | getDocumentHtml     | Livre  | Retourne le HTML complet via `body.getHtml()` |
| W4  | modifyTableCell     | Livre  | Remplacement d'une cellule ciblee dans un tableau existant |
| W5  | addTableRow         | Livre  | Ajout de lignes dans un tableau (`addRows`) |
| W6  | addTableColumn      | Livre  | Ajout de colonnes dans un tableau (`addColumns`) |
| W7  | deleteTableRowColumn| Livre  | Suppression de lignes/colonnes (`deleteRows` / `deleteColumns`) |
| W8  | formatTableCell     | Livre  | Formatage de cellule (fond + police) |
| W9  | insertHeaderFooter  | Livre  | Insertion d'en-tete/pied (`getHeader` / `getFooter`) |
| W10 | insertFootnote      | Livre  | Insertion de note de bas de page (`insertFootnote`) |
| W11 | addComment          | Livre  | Ajout de commentaire de relecture (`insertComment`) |
| W12 | getComments         | Livre  | Lecture des commentaires (`getComments`) |
| W13 | setPageSetup        | Livre  | Reglage marges/orientation/papier (`pageSetup`) |
| W14 | getSpecificParagraph| Livre  | Lecture d'un paragraphe par index |
| W15 | insertSectionBreak  | Livre  | Insertion de saut de section (`SectionNext`) |

---

## 2. EXCEL (excelTools.ts) — 39 existants, 0 manquant sur le lot E1-E18

### Existant
- Data: getSelectedCells, setCellValue, getWorksheetData, getCellFormula, getDataFromSheet, copyRange
- Formules / validation: insertFormula, fillFormulaDown, addDataValidation, setNamedRange, getNamedRanges
- Formatage: formatRange (wrap/align/font/borders avances), setCellNumberFormat, autoFitColumns, setColumnWidth, setRowHeight, applyConditionalFormatting, getConditionalFormattingRules
- Structure: addWorksheet, renameWorksheet, activateWorksheet, deleteWorksheet, insertRow, insertColumn, deleteRow, deleteColumn, mergeCells, createTable
- Charts: createChart
- Filtres et volets: applyAutoFilter, removeAutoFilter, sortRange, freezePanes
- Liens / commentaires / protection: addHyperlink, addCellComment, protectWorksheet
- Utility: searchAndReplace, clearRange, getWorksheetInfo

### Lot Excel integre (E1-E18)

| ID  | Skill                  | Statut | Note implementation                                                                                 |
|-----|------------------------|--------|------------------------------------------------------------------------------------------------------|
| E1  | addDataValidation      | Livre  | Validation de donnees (listes, bornes numeriques, date, custom) via `range.dataValidation.rule`    |
| E2  | createTable            | Livre  | Conversion d'une plage en table structuree (`sheet.tables.add`, `name`, `style`)                    |
| E3  | copyRange              | Livre  | Copie valeurs/formules/format numerique entre plages                                                 |
| E4  | renameWorksheet        | Livre  | Renommage de feuille existante (`sheet.name = ...`)                                                  |
| E5  | deleteWorksheet        | Livre  | Suppression de feuille (`sheet.delete()`)                                                            |
| E6  | activateWorksheet      | Livre  | Activation/navigation entre feuilles (`sheet.activate()`)                                            |
| E7  | getDataFromSheet       | Livre  | Lecture de donnees d'une autre feuille sans activation                                               |
| E8  | freezePanes            | Livre  | Figer/defiger volets (rows/columns/at/unfreeze)                                                      |
| E9  | addHyperlink           | Livre  | Insertion de liens cliquables (`range.hyperlink`)                                                    |
| E10 | addCellComment         | Livre  | Ajout de commentaires de cellule (`workbook.comments.add`)                                           |
| E11 | wrapText (formatRange) | Livre  | `formatRange` etendu avec `wrapText`                                                                 |
| E12 | verticalAlignment      | Livre  | `formatRange` etendu avec `verticalAlignment`                                                        |
| E13 | fontName               | Livre  | `formatRange` etendu avec `fontName`                                                                 |
| E14 | removeAutoFilter       | Livre  | Retrait des filtres de feuille (`sheet.autoFilter.remove()`)                                         |
| E15 | protectWorksheet       | Livre  | Protection/deprotection de feuille (`sheet.protection.protect/unprotect`)                            |
| E16 | customizeBorders       | Livre  | `formatRange` etendu avec style/couleur/epaisseur de bordure globale et par cote                    |
| E17 | getNamedRanges         | Livre  | Lecture des plages nommees du classeur (`workbook.names.load(...)`)                                 |
| E18 | setNamedRange          | Livre  | Creation/mise a jour d'une plage nommee (`workbook.names.add`)                                      |

---

## 3. POWERPOINT (powerpointTools.ts) — 0 tools, ~8 a creer

### Existant
AUCUN tool agent. Seulement 3 fonctions helpers internes:
- `getPowerPointSelection()` — lire le texte selectionne
- `insertIntoPowerPoint()` — remplacer la selection
- `normalizePowerPointListText()` — normaliser les listes markdown

L'agent PowerPoint est en mode "prompt-only": il genere du texte, l'utilisateur doit inserer manuellement.

### A creer

| ID  | Skill                  | Cas d'usage bloquant                                                          | Implementation                                                                                 |
|-----|------------------------|-------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| P1  | getSelectedText        | L'agent ne peut pas lire la selection (le helper existe mais pas le tool)     | Wrapper agent autour de `getPowerPointSelection()` existant                                     |
| P2  | replaceSelectedText    | L'agent ne peut pas modifier la selection                                     | Wrapper agent autour de `insertIntoPowerPoint()` existant                                       |
| P3  | getSlideCount          | L'agent ne sait pas combien de slides existent                                | `PowerPoint.run` -> `presentation.slides.load('items')` -> `items.length`                       |
| P4  | getSlideContent        | Ne peut pas lire le contenu textuel d'une slide specifique                    | `PowerPoint.run` -> `slides.getItemAt(index)` -> `.shapes.load('items')` -> iterer textRange    |
| P5  | addSlide               | Ne peut pas ajouter de slides                                                 | `PowerPoint.run` -> `presentation.slides.add()` avec option de layout                           |
| P6  | setSlideNotes          | Ne peut pas ajouter les notes du presentateur                                 | `PowerPoint.run` -> slide.notesSlide (PowerPointApi 1.4+) ou Common API fallback               |
| P7  | insertTextBox          | Ne peut pas ajouter de contenu sur un slide                                   | `PowerPoint.run` -> `slide.shapes.addTextBox(text)` avec position/taille                        |
| P8  | insertImage            | Ne peut pas ajouter d'images sur les slides                                   | `PowerPoint.run` -> `slide.shapes.addImage(base64)` avec position/taille                        |

> **Note:** L'API PowerPoint.js a des limitations historiques. Les requirement sets recents (PowerPointApi 1.2+, 1.3+, 1.4+) ajoutent des capacites reelles. Verifier la compatibilite avec les versions Office ciblees.

---

## 4. OUTLOOK (outlookTools.ts) — 3 existants, ~10 manquants

### Existant
- getEmailBody — lire le corps entier
- getSelectedText — lire la selection
- setEmailBody — remplacer TOUT le corps (texte brut uniquement)

### Manquant

| ID  | Skill                  | Cas d'usage bloquant                                                          | Implementation                                                                                 |
|-----|------------------------|-------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| O1  | insertTextAtCursor     | setEmailBody remplace TOUT. Impossible d'inserer au curseur.                  | `mailbox.item.body.setSelectedDataAsync(text, { coercionType: Text })`                          |
| O2  | setEmailBodyHtml       | Seulement du texte brut. Pas de mails formates (gras, liens, listes).         | `mailbox.item.body.setAsync(html, { coercionType: Html })`                                      |
| O3  | getEmailSubject        | L'agent ne connait pas le sujet. Comment rediger une reponse pertinente?      | `mailbox.item.subject.getAsync(callback)` (read) / `.subject` (compose)                         |
| O4  | setEmailSubject        | Ne peut pas modifier le sujet en compose                                       | `mailbox.item.subject.setAsync(subject, callback)`                                               |
| O5  | getEmailRecipients     | Ne connait pas les destinataires. Impossible de personnaliser.                 | `mailbox.item.to.getAsync()` / `.cc.getAsync()` (compose) ou `.to` / `.cc` (read)               |
| O6  | addRecipient           | Ne peut pas ajouter de destinataires                                           | `mailbox.item.to.addAsync(recipients)` / `.cc.addAsync(...)` / `.bcc.addAsync(...)`              |
| O7  | getEmailSender         | Ne connait pas l'expediteur. Contexte perdu pour les reponses.                 | `mailbox.item.from` (read) / `mailbox.item.sender`                                              |
| O8  | getEmailDate           | Pas d'acces a la date de l'email                                               | `mailbox.item.dateTimeCreated` (read mode)                                                       |
| O9  | getAttachments         | Ne peut pas lister les pieces jointes                                          | `mailbox.item.attachments` (read) ou `mailbox.item.getAttachmentsAsync()`                        |
| O10 | insertHtmlAtCursor     | Inserer du contenu HTML formate au curseur                                     | `mailbox.item.body.setSelectedDataAsync(html, { coercionType: Html })`                           |

---

## 5. PRIORITES D'IMPLEMENTATION

### Priorite 0 — Termine (Word)

- Lot Word **W1 a W15** integre dans `wordTools.ts` (39 tools Word disponibles).

### Priorite 1 — Impact immediat (quick wins + bloquants critiques)

1. **P1+P2**: Wrapper les helpers PPT existants en tools agent
2. ✅ **E11+E12+E13**: Etendre `formatRange` avec wrapText/verticalAlignment/fontName
3. **O1**: `insertTextAtCursor` — arreter de tout ecraser
4. **O2**: `setEmailBodyHtml` — mails formates
5. **O3+O4**: get/set subject — contexte basique
6. ✅ **E4+E5+E6**: rename/delete/activate worksheet
7. ✅ **E1**: `addDataValidation` — listes deroulantes

### Priorite 2 — Cas d'usage professionnels

8. ✅ **E2**: `createTable` (ListObject Excel)
9. ✅ **E3**: `copyRange`
10. ✅ **E8**: `freezePanes`
11. **P3+P4**: getSlideCount/getSlideContent
12. **O5+O6+O7**: Recipients/Sender
13. ✅ **E14**: `removeAutoFilter`

### Priorite 3 — Features avancees

14. **P5+P6+P7+P8**: Creation slides/shapes/images PPT
15. ✅ **E9+E10**: Hyperlinks/Comments Excel
16. ✅ **E15**: Protection feuilles
17. ✅ **E16+E17+E18**: Borders avances / Named ranges
18. **O8+O9+O10**: Date/Attachments/HTML insert
