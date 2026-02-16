# KickOffice — Agent Skills Audit

> Date: 2026-02-16
> Objective: Identify missing tools for the agent across all 4 Office hosts, focusing on practical, high-value operations supported by the Office.js API.

## Current State

| Application    | Existing tools | Missing tools identified |
|----------------|---------------|--------------------------|
| **Word**       | 37 tools      | 4 proposed (styles, headings, doc properties) |
| **Excel**      | 39 tools      | 2 proposed (chart mgmt, named range deletion) |
| **PowerPoint** | 8 tools       | **11 proposed** (biggest gap — slide/shape management) |
| **Outlook**    | 13 tools      | 5 proposed (attachments, importance, HTML body) |
| **General**    | 2 tools       | 0 (sufficient for current scope) |

---

## 1. WORD (wordTools.ts) — 37 existing

### Existing (complete list)
- **Text R/W**: getSelectedText, getDocumentContent, getDocumentHtml, getSpecificParagraph, insertText, replaceSelectedText, appendText, deleteText, selectText, findText, searchAndReplace
- **Formatting**: formatText, clearFormatting, setFontName, applyTaggedFormatting, setParagraphFormat, getRangeInfo
- **Structure**: insertParagraph, insertList, insertTable, insertPageBreak, insertSectionBreak, getTableInfo, modifyTableCell, addTableRow, addTableColumn, deleteTableRowColumn, formatTableCell
- **Navigation**: insertBookmark, goToBookmark, insertContentControl
- **Media & Links**: insertImage, insertHyperlink
- **Headers/Footers/Notes**: insertHeaderFooter, insertFootnote
- **Comments**: addComment, getComments
- **Page/Doc**: setPageSetup, getDocumentProperties

### Previously delivered lots
- W1-W15: **Delivered** ✅

### Proposed additions (Priority 2)

| ID  | Skill | API | Use case |
|-----|-------|-----|----------|
| W16 | **applyStyle** | `range.style = "Heading 1"` | Apply named styles (Heading 1-6, Normal, Quote, etc.) to selection. Essential for structured documents, TOC generation, and professional formatting. |
| W17 | **getStyles** | `document.getStyles().load('items')` | List available styles in the document. Needed so the agent knows which styles exist before applying them. |
| W18 | **getHeadings** | Iterate paragraphs, filter by `paragraph.styleBuiltIn` for Heading types | Return document outline (heading hierarchy). Critical for understanding document structure, generating TOCs, and navigating large documents. |
| W19 | **setDocumentProperties** | `document.properties.title = "..."` etc. | Set core document properties (title, subject, author, keywords). Useful for metadata management. |

**Not proposed (intentionally excluded)**:
- Track changes (getRevisions, accept/reject): Requires Mailbox-level permissions and complex API surface. Low ROI for an assistant agent — users manage track changes manually.
- Insert TOC: No reliable API support in Word.js for programmatic TOC insertion.
- Form fields (dropdown, checkbox): Niche use case, not worth the complexity.

---

## 2. EXCEL (excelTools.ts) — 39 existing

### Existing (complete list)
- **Data R/W**: getSelectedCells, setCellValue, getWorksheetData, getCellFormula, getDataFromSheet, copyRange, clearRange, searchAndReplace
- **Formulas/Validation**: insertFormula, addDataValidation, setNamedRange, getNamedRanges
- **Formatting**: formatRange, setCellNumberFormat, autoFitColumns, setColumnWidth, setRowHeight, applyConditionalFormatting, getConditionalFormattingRules
- **Structure**: addWorksheet, renameWorksheet, activateWorksheet, deleteWorksheet, insertRow, insertColumn, deleteRow, deleteColumn, mergeCells, createTable
- **Charts**: createChart
- **Filters/Panes**: applyAutoFilter, removeAutoFilter, sortRange, freezePanes
- **Links/Comments/Protection**: addHyperlink, addCellComment, protectWorksheet
- **Info**: getWorksheetInfo

### Previously delivered lots
- E1-E18: **Delivered** ✅

### Proposed additions (Priority 3 — nice to have)

| ID  | Skill | API | Use case |
|-----|-------|-----|----------|
| E19 | **deleteNamedRange** | `workbook.names.getItem(name).delete()` | Remove a named range. Currently can create/list but not delete. |
| E20 | **getCharts** | `worksheet.charts.load('items')` | List all charts in the active worksheet with name, type, and position. Needed for the agent to understand existing charts before modifying. |

**Not proposed (intentionally excluded)**:
- Pivot tables: Very complex API surface (`worksheet.pivotTables`), high risk of errors, and most users create pivots manually via UI.
- Sparklines: Niche feature, low demand.
- Slicers: Requires pivot tables, compound complexity.

---

## 3. POWERPOINT (powerpointTools.ts) — 8 existing ⚠️ BIGGEST GAP

### Existing
- **Text**: getSelectedText (Common API), replaceSelectedText (Common API)
- **Slides**: getSlideCount (1.2+), getSlideContent (1.3+), addSlide (1.2+)
- **Notes**: setSlideNotes (1.4+)
- **Insertion**: insertTextBox (1.3+), insertImage (1.3+)

### Missing: the agent cannot delete, move, or modify existing content

PowerPoint has only 8 tools compared to Word's 37 and Excel's 39. The agent can read and insert, but it **cannot**:
- Delete a slide or shape
- Move or resize shapes
- Modify text in a specific shape (only the selected text via Common API)
- Read speaker notes (can only write them)
- Set fill/background colors
- Get an overview of all slides

### Proposed additions (Priority 1 — critical)

| ID  | Skill | API (min req set) | Use case |
|-----|-------|-------------------|----------|
| P9  | **deleteSlide** | `slide.delete()` (1.2+) | Remove a slide by number. Essential for reorganizing presentations. |
| P10 | **getSlideNotes** | `slide.notesSlide.textBody.text` (1.4+) | Read speaker notes from a slide. Currently can only write notes, not read. |
| P11 | **getShapes** | `slide.shapes.load('items')` → name, type, textFrame (1.3+) | List all shapes on a slide with their type, name, text, and position. Needed for the agent to understand slide layout before modifying. |
| P12 | **deleteShape** | `shape.delete()` (1.3+) | Remove a shape by name or index. Cannot currently clean up slides. |
| P13 | **setShapeText** | `shape.textFrame.textRange.text = "..."` (1.4+) | Set text of a specific shape by name/index (not just the selected text). Enables modifying any shape without user selection. |
| P14 | **formatShapeText** | `textRange.font.bold/italic/color/size` (1.4+) | Apply text formatting (bold, italic, font size, color) to shape text. Currently no formatting capability. |
| P15 | **setShapeFill** | `shape.fill.setSolidColor(hex)` (1.3+) | Set background fill color of a shape. Essential for visual design. |
| P16 | **insertShape** | `slide.shapes.addGeometricShape(type)` (1.3+) | Insert a geometric shape (rectangle, circle, arrow, etc.) with position and size. |
| P17 | **moveResizeShape** | `shape.left/top/width/height = ...` (1.3+) | Change position and/or dimensions of an existing shape. Essential for layout adjustments. |
| P18 | **getAllSlidesOverview** | Iterate slides → get title shapes + shape count (1.3+) | Return a summary of all slides (title text, number of shapes, has notes). Essential for navigating and understanding large presentations. |
| P19 | **duplicateSlide** | `presentation.insertSlidesFromBase64()` (1.2+) | Duplicate an existing slide. Common operation when building presentations from templates. |

> **API compatibility note**: All proposed tools use PowerPointApi 1.2-1.4 which are already used by existing tools. The `isPowerPointApiSupported()` helper already handles graceful fallbacks.

---

## 4. OUTLOOK (outlookTools.ts) — 13 existing

### Existing
- **Body R/W**: getEmailBody, getSelectedText, setEmailBody, insertTextAtCursor, setEmailBodyHtml, insertHtmlAtCursor
- **Subject**: getEmailSubject, setEmailSubject
- **Recipients**: getEmailRecipients, addRecipient
- **Metadata**: getEmailSender, getEmailDate, getAttachments

### Manifest scope
The current manifest (`manifest-outlook.template.xml`) supports `MessageReadCommandSurface` and `MessageComposeCommandSurface` only. Calendar/appointment extension points are **not declared**, so calendar tools are not feasible without manifest changes.

Permission: `ReadWriteMailbox` — sufficient for all proposed tools.

### Proposed additions (Priority 2)

| ID  | Skill | API | Use case |
|-----|-------|-----|----------|
| O11 | **addAttachment** | `item.addFileAttachmentAsync(uri, name)` | Attach a file from URL to the current email (compose). Commonly requested when drafting emails. |
| O12 | **removeAttachment** | `item.removeAttachmentAsync(id)` | Remove an attachment by ID (compose). Needed for agent to manage attachments it previously added. |
| O13 | **getEmailBodyHtml** | `body.getAsync(CoercionType.Html)` | Get the email body as HTML (current `getEmailBody` only returns plain text). Needed to preserve formatting when analyzing or modifying emails. |
| O14 | **getEmailImportance** | `item.importance` (read) | Read the importance/priority level (Normal, High, Low). Useful for triage and prioritization workflows. |
| O15 | **setEmailImportance** | `item.importance` setter (compose) | Set the importance level when composing. Agent can flag urgent emails automatically. |

**Not proposed (intentionally excluded)**:
- Calendar/meeting tools (createEvent, getEvents, etc.): Requires adding `AppointmentOrganizerCommandSurface` / `AppointmentAttendeeCommandSurface` extension points to the manifest. This is a separate feature that should be planned as its own work item with manifest changes.
- Folder operations (moveToFolder, createFolder): Not available via the Mailbox 1.1 API used by task pane add-ins.
- Send email: Too dangerous for an AI agent to send emails autonomously. User should always press Send manually.
- Contact/address book: Requires `ReadItem` permission at minimum and is complex to implement properly.

---

## 5. GENERAL (generalTools.ts) — 2 existing

### Existing
- getCurrentDate
- calculateMath (via mathjs)

### Assessment
Sufficient for current scope. Both tools are shared across all hosts. No additions proposed — the current date and math capabilities cover the general utility needs.

---

## Implementation Priority Summary

### Priority 1 — PowerPoint gap (P9-P19, 11 tools)
PowerPoint is severely underserved. With only 8 tools (vs 37-39 for Word/Excel), the agent cannot perform basic operations like deleting slides, modifying specific shapes, or reading notes. This limits the agent to simple text replacement and insertion.

**Estimated effort**: Medium — all use PowerPointApi 1.2-1.4, same patterns as existing tools.

### Priority 2 — Word styles + Outlook attachments (W16-W19 + O11-O15, 9 tools)
Word style management and Outlook attachment handling fill meaningful gaps in daily workflows.

**Estimated effort**: Low — well-documented APIs with straightforward implementation.

### Priority 3 — Excel extras (E19-E20, 2 tools)
Excel is already comprehensive at 39 tools. These are minor completions.

**Estimated effort**: Very low.

### Future consideration — Outlook Calendar
Adding calendar support would require:
1. New manifest extension points (`AppointmentOrganizerCommandSurface`, `AppointmentAttendeeCommandSurface`)
2. New calendar-specific tools (10-15 tools)
3. New agent prompt for calendar context
4. Testing across read/compose appointment modes

This should be planned as a separate feature, not added incrementally.
