# KickOffice — Audit des Skills Agent Manquantes

> Date: 2026-02-16  
> Objectif: Identifier les skills manquantes pour que l'agent soit efficace dans les 4 outils Office.

## État actuel

| Application    | Tools existants | Tools manquants identifiés |
|----------------|-----------------|----------------------------|
| **Word**       | 39 tools        | 0 manquant (lot W1-W15 livré) |
| **Excel**      | 39 tools        | 0 manquant (lot E1-E18 livré) |
| **PowerPoint** | 8 tools         | 0 manquant sur le lot P1-P8 |
| **Outlook**    | 3 tools         | ~10 manquants              |

---

## 1. WORD (wordTools.ts) — 39 existants

### Existant
- Texte: getSelectedText, getDocumentContent, insertText, replaceSelectedText, appendText, deleteText, selectText, findText, searchAndReplace
- Formatage inline: formatText, clearFormatting, setFontName, applyTaggedFormatting, getRangeInfo
- Structure: insertParagraph, insertList, insertTable, getTableInfo, insertPageBreak
- Navigation: insertBookmark, goToBookmark, insertContentControl
- Image: insertImage

### Lot Word intégré (W1-W15)
Statut: **Livré**.

---

## 2. EXCEL (excelTools.ts) — 39 existants

### Existant
- Data: getSelectedCells, setCellValue, getWorksheetData, getCellFormula, getDataFromSheet, copyRange
- Formules / validation: insertFormula, fillFormulaDown, addDataValidation, setNamedRange, getNamedRanges
- Formatage: formatRange, setCellNumberFormat, autoFitColumns, setColumnWidth, setRowHeight, applyConditionalFormatting, getConditionalFormattingRules
- Structure: addWorksheet, renameWorksheet, activateWorksheet, deleteWorksheet, insertRow, insertColumn, deleteRow, deleteColumn, mergeCells, createTable
- Charts: createChart
- Filtres et volets: applyAutoFilter, removeAutoFilter, sortRange, freezePanes
- Liens / commentaires / protection: addHyperlink, addCellComment, protectWorksheet
- Utility: searchAndReplace, clearRange, getWorksheetInfo

### Lot Excel intégré (E1-E18)
Statut: **Livré**.

---

## 3. POWERPOINT (powerpointTools.ts) — 8 existants

### Existant
- Helpers internes:
  - `getPowerPointSelection()`
  - `insertIntoPowerPoint()`
  - `normalizePowerPointListText()`
- Outils agent livrés:
  - `getSelectedText` (P1)
  - `replaceSelectedText` (P2)
  - `getSlideCount` (P3)
  - `getSlideContent` (P4)
  - `addSlide` (P5)
  - `setSlideNotes` (P6)
  - `insertTextBox` (P7)
  - `insertImage` (P8)

### Lot PowerPoint (P1-P8)

| ID  | Skill             | Statut | Note d'implémentation |
|-----|-------------------|--------|-----------------------|
| P1  | getSelectedText   | Livré  | Wrapper agent de `getPowerPointSelection()` |
| P2  | replaceSelectedText | Livré | Wrapper agent de `insertIntoPowerPoint()` |
| P3  | getSlideCount     | Livré  | `PowerPoint.run` + `presentation.slides.load('items')` |
| P4  | getSlideContent   | Livré  | Lecture des `shapes` textuelles d'une slide donnée |
| P5  | addSlide          | Livré  | `presentation.slides.add()` (+ layout optionnel) |
| P6  | setSlideNotes     | Livré  | `slide.notesSlide` (PowerPointApi 1.4+), message d'erreur si indisponible |
| P7  | insertTextBox     | Livré  | `slide.shapes.addTextBox(text)` avec position/taille |
| P8  | insertImage       | Livré  | `slide.shapes.addImage(base64)` avec position/taille |

> **Note compatibilité:** l'API PowerPoint.js dépend des requirement sets (`PowerPointApi 1.2+`, `1.3+`, `1.4+`). Un fallback explicite est renvoyé si le runtime ne supporte pas la capacité demandée.

---

## 4. OUTLOOK (outlookTools.ts) — 3 existants, ~10 manquants

### Existant
- getEmailBody — lire le corps entier
- getSelectedText — lire la sélection
- setEmailBody — remplacer tout le corps (texte brut)

### Manquant
- O1 insertTextAtCursor
- O2 setEmailBodyHtml
- O3 getEmailSubject
- O4 setEmailSubject
- O5 getEmailRecipients
- O6 addRecipient
- O7 getEmailSender
- O8 getEmailDate
- O9 getAttachments
- O10 insertHtmlAtCursor

---

## 5. Priorités d'implémentation

### Priorité 0 — Terminé
- ✅ Word lot W1-W15
- ✅ Excel lot E1-E18
- ✅ PowerPoint lot P1-P8

### Priorité 1 — Reste critique
1. **O1** `insertTextAtCursor`
2. **O2** `setEmailBodyHtml`
3. **O3+O4** `getEmailSubject` / `setEmailSubject`

### Priorité 2 — Cas d'usage professionnels
4. **O5+O6+O7** destinataires / expéditeur

### Priorité 3 — Features avancées
5. **O8+O9+O10** date / pièces jointes / insertion HTML au curseur
