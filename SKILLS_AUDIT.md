# KickOffice — Audit des Skills Agent Manquantes

> Date: 2026-02-16  
> Objectif: Identifier les skills manquantes pour que l'agent soit efficace dans les 4 outils Office.

## État actuel

| Application    | Tools existants | Tools manquants identifiés |
|----------------|-----------------|----------------------------|
| **Word**       | 39 tools        | 0 manquant (lot W1-W15 livré) |
| **Excel**      | 39 tools        | 0 manquant (lot E1-E18 livré) |
| **PowerPoint** | 8 tools         | 0 manquant sur le lot P1-P8 |
| **Outlook**    | 13 tools        | 0 manquant (lot O1-O10 livré) |

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

## 4. OUTLOOK (outlookTools.ts) — 13 existants

### Existant
- getEmailBody — lire le corps entier
- getSelectedText — lire la sélection
- setEmailBody — remplacer tout le corps (texte brut)
- insertTextAtCursor — insérer du texte brut au curseur
- setEmailBodyHtml — remplacer tout le corps en HTML
- getEmailSubject — lire l'objet du mail
- setEmailSubject — modifier l'objet en mode compose
- getEmailRecipients — lire To/Cc/Bcc
- addRecipient — ajouter un destinataire To/Cc/Bcc
- getEmailSender — lire l'expéditeur
- getEmailDate — lire la date de création du mail
- getAttachments — lister les pièces jointes
- insertHtmlAtCursor — insérer du HTML formaté au curseur

### Lot Outlook intégré (O1-O10)

| ID  | Skill              | Statut | Note d'implémentation |
|-----|--------------------|--------|-----------------------|
| O1  | insertTextAtCursor | Livré  | `mailbox.item.body.setSelectedDataAsync(text, { coercionType: Text })` |
| O2  | setEmailBodyHtml   | Livré  | `mailbox.item.body.setAsync(html, { coercionType: Html })` |
| O3  | getEmailSubject    | Livré  | `mailbox.item.subject.getAsync(...)` avec fallback `.subject` |
| O4  | setEmailSubject    | Livré  | `mailbox.item.subject.setAsync(subject, callback)` |
| O5  | getEmailRecipients | Livré  | `to/cc/bcc.getAsync()` + fallback lecture directe |
| O6  | addRecipient       | Livré  | `to/cc/bcc.addAsync(...)` avec normalisation des entrées |
| O7  | getEmailSender     | Livré  | `mailbox.item.from` / `mailbox.item.sender` |
| O8  | getEmailDate       | Livré  | `mailbox.item.dateTimeCreated` |
| O9  | getAttachments     | Livré  | `mailbox.item.getAttachmentsAsync()` + fallback `.attachments` |
| O10 | insertHtmlAtCursor | Livré  | `mailbox.item.body.setSelectedDataAsync(html, { coercionType: Html })` |

---

## 5. Priorités d'implémentation

### Priorité 0 — Terminé
- ✅ Word lot W1-W15
- ✅ Excel lot E1-E18
- ✅ PowerPoint lot P1-P8

### Priorité 1 — Reste critique
✅ Aucun blocage critique restant sur Outlook dans ce lot.

### Priorité 2 — Cas d'usage professionnels
✅ Couvert (destinataires / expéditeur).

### Priorité 3 — Features avancées
✅ Couvert (date / pièces jointes / insertion HTML au curseur).
