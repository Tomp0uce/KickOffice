# KickOffice — Audit des Skills Agent Manquantes

> Date: 2026-02-16
> Objectif: Identifier toutes les skills manquantes pour que l'agent soit efficace dans les 4 outils Office

## Etat actuel

| Application    | Tools existants | Tools manquants identifies |
|----------------|-----------------|----------------------------|
| **Word**       | 24 tools        | ~15 manquants              |
| **Excel**      | 25 tools        | ~18 manquants              |
| **PowerPoint** | 0 tools (!)     | ~8 a creer                 |
| **Outlook**    | 3 tools         | ~10 manquants              |

---

## 1. WORD (wordTools.ts) — 24 existants, ~15 manquants

### Existant
- Texte: getSelectedText, getDocumentContent, insertText, replaceSelectedText, appendText, deleteText, selectText, findText, searchAndReplace
- Formatage inline: formatText, clearFormatting, setFontName, applyTaggedFormatting, getRangeInfo
- Structure: insertParagraph, insertList, insertTable, getTableInfo, insertPageBreak
- Navigation: insertBookmark, goToBookmark, insertContentControl
- Image: insertImage

### Manquant

| ID  | Skill                  | Cas d'usage bloquant                                                            | Implementation                                                                                                        |
|-----|------------------------|---------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------|
| W1  | setParagraphFormat      | Centrer un titre, changer l'interligne, gerer les indentations                  | `selection.paragraphs` -> `.alignment`, `.lineSpacing`, `.spaceAfter`, `.spaceBefore`, `.leftIndent`, `.firstLineIndent` |
| W2  | insertHyperlink         | Inserer un lien cliquable dans un document                                      | `range.hyperlink = { address, textToDisplay }`                                                                         |
| W3  | getDocumentHtml         | Comprendre la structure reelle (titres, listes, gras...) pour analyser fidele   | `body.getHtml()` -> retourner le HTML                                                                                  |
| W4  | modifyTableCell         | Modifier une cellule d'un tableau existant                                      | `body.tables.getFirst().getCell(row, col).body.insertText(...)`                                                        |
| W5  | addTableRow             | Ajouter des lignes a un tableau existant                                        | `table.addRows(location, count, values)`                                                                               |
| W6  | addTableColumn          | Ajouter des colonnes a un tableau existant                                      | `table.addColumns(location, count, values)`                                                                            |
| W7  | deleteTableRowColumn    | Supprimer lignes/colonnes d'un tableau existant                                 | `table.deleteRows(index, count)` / `table.deleteColumns(index, count)`                                                 |
| W8  | formatTableCell         | Mettre en forme les cellules de tableau (fond, bordures)                        | `table.getCell(r,c)` -> `.shadingColor`, `.body.font.*`                                                                |
| W9  | insertHeaderFooter      | Creer en-tete/pied de page pour documents pro                                  | `document.sections.getFirst().getHeader(type)` / `.getFooter(type)` -> `.insertText(...)`                              |
| W10 | insertFootnote          | Notes de bas de page (academique/juridique)                                     | `range.insertFootnote(text)` (WordApi 1.5+)                                                                           |
| W11 | addComment              | Ajouter des commentaires de relecture                                           | `range.insertComment(text)` (WordApi 1.4+)                                                                            |
| W12 | getComments             | Lire les commentaires existants                                                 | `body.getComments()` -> `.load('items/content,items/authorName')`                                                      |
| W13 | setPageSetup            | Changer marges, orientation, taille de page                                     | `document.sections.getFirst().pageSetup.topMargin`, `.orientation`, `.paperSize`                                       |
| W14 | getSpecificParagraph    | Lire/cibler un paragraphe par index sans tout relire                            | `body.paragraphs.load('items')` -> `.items[index].load('text,style,font/*')`                                           |
| W15 | insertSectionBreak      | Changer orientation/marges en cours de document                                 | `range.insertBreak('SectionNext', location)`                                                                           |

---

## 2. EXCEL (excelTools.ts) — 25 existants, ~18 manquants

### Existant
- Data: getSelectedCells, setCellValue, getWorksheetData, getCellFormula
- Formules: insertFormula, fillFormulaDown
- Formatage: formatRange, setCellNumberFormat, autoFitColumns, setColumnWidth, setRowHeight, applyConditionalFormatting, getConditionalFormattingRules
- Structure: insertRow, insertColumn, deleteRow, deleteColumn, mergeCells, addWorksheet
- Charts: createChart
- Filtres: applyAutoFilter, sortRange
- Utility: searchAndReplace, clearRange, getWorksheetInfo

### Manquant

| ID  | Skill                  | Cas d'usage bloquant                                                           | Implementation                                                                                 |
|-----|------------------------|--------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| E1  | addDataValidation      | Creer des listes deroulantes ou regles de validation                           | `range.dataValidation.rule = { list: { source: "A,B,C" } }`                                    |
| E2  | createTable            | Convertir une plage en Table Excel structuree (ListObject)                     | `sheet.tables.add(range, hasHeaders)` -> `.name = ...` -> `.style = "TableStyleMedium2"`        |
| E3  | copyRange              | Copier une plage vers un autre emplacement                                     | `sourceRange.load('values,formulas,numberFormat')` -> `destRange.values = ...`                  |
| E4  | renameWorksheet        | Renommer une feuille existante                                                 | `sheet.name = newName`                                                                          |
| E5  | deleteWorksheet        | Supprimer une feuille                                                          | `sheet.delete()`                                                                                |
| E6  | activateWorksheet      | Naviguer entre les feuilles                                                    | `workbook.worksheets.getItem(name).activate()`                                                  |
| E7  | getDataFromSheet       | Lire les donnees d'une autre feuille sans la basculer                          | `workbook.worksheets.getItem(name).getUsedRange().load('values')`                               |
| E8  | freezePanes            | Figer les volets (en-tetes visibles)                                           | `sheet.freezePanes.freezeRows(count)` / `.freezeColumns(count)` / `.freezeAt(range)`            |
| E9  | addHyperlink           | Liens cliquables dans les cellules                                             | `range.hyperlink = { address, textToDisplay }`                                                   |
| E10 | addCellComment         | Commentaires/notes aux cellules                                                | `workbook.comments.add(range, text)` (ExcelApi 1.10+)                                           |
| E11 | wrapText (formatRange) | Retour a la ligne dans les cellules                                            | Ajouter `wrapText: boolean` dans formatRange -> `range.format.wrapText = true`                   |
| E12 | verticalAlignment (formatRange) | Alignement vertical (top/center/bottom)                               | Ajouter `verticalAlignment` dans formatRange -> `range.format.verticalAlignment = ...`           |
| E13 | fontName (formatRange) | Changer la police dans Excel                                                   | Ajouter `fontName: string` dans formatRange -> `range.format.font.name = fontName`               |
| E14 | removeAutoFilter       | Retirer les filtres appliques                                                  | `sheet.autoFilter.remove()`                                                                      |
| E15 | protectWorksheet       | Proteger/deproteger une feuille                                                | `sheet.protection.protect(options)` / `.unprotect(password)`                                     |
| E16 | customizeBorders       | Controle du style, epaisseur, couleur des bordures par cote                    | Etendre formatRange avec borderStyle/borderColor/borderWeight per-side                           |
| E17 | getNamedRanges         | Lire les plages nommees                                                        | `workbook.names.load('items/name,items/value')`                                                  |
| E18 | setNamedRange          | Creer des plages nommees pour formules complexes                               | `workbook.names.add(name, range)`                                                                |

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

### Priorite 1 — Impact immediat (quick wins + bloquants critiques)

1. **P1+P2**: Wrapper les helpers PPT existants en tools agent
2. **W1**: `setParagraphFormat` — centrer un titre est basique
3. **E11+E12+E13**: Etendre `formatRange` avec wrapText/verticalAlignment/fontName
4. **O1**: `insertTextAtCursor` — arreter de tout ecraser
5. **O2**: `setEmailBodyHtml` — mails formates
6. **O3+O4**: get/set subject — contexte basique
7. **E4+E5+E6**: rename/delete/activate worksheet
8. **W2**: `insertHyperlink`
9. **W4**: `modifyTableCell`
10. **E1**: `addDataValidation` — listes deroulantes

### Priorite 2 — Cas d'usage professionnels

11. **W8**: Headers/Footers
12. **W3**: `getDocumentHtml` — lecture structuree
13. **W5+W6+W7+W8**: Manipulation complete des tables Word
14. **E2**: `createTable` (ListObject Excel)
15. **E3**: `copyRange`
16. **E8**: `freezePanes`
17. **P3+P4**: getSlideCount/getSlideContent
18. **O5+O6+O7**: Recipients/Sender
19. **E14**: `removeAutoFilter`

### Priorite 3 — Features avancees

20. **P5+P6+P7+P8**: Creation slides/shapes/images PPT
21. **W10+W11+W12**: Footnotes/Comments Word
22. **W13**: Page setup
23. **E9+E10**: Hyperlinks/Comments Excel
24. **E15**: Protection feuilles
25. **E16+E17+E18**: Borders avances / Named ranges
26. **O8+O9+O10**: Date/Attachments/HTML insert
