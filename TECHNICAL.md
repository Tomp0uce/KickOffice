# Technical Reference — KickOffice

> This document contains implementation details extracted from the PRD. For feature requirements, see PRD.md.

---

## 1. Stack & Architecture

- **Frontend:** Vue 3 + TypeScript + Tailwind CSS v4 + Vite
- **Backend:** Node.js/Express SSE proxy
- **LLM Gateway:** LiteLLM (production) / OpenAI (testing)
- **Testing:** Vitest (unit) + Playwright (e2e)
- **Deployment:** Docker Compose on Synology DS416play

---

## 2. Backend API

### 2.1 Endpoints

| Endpoint             | Method | Description                                      |
| -------------------- | ------ | ------------------------------------------------ |
| `/api/chat`          | POST   | Streaming chat completion (SSE)                  |
| `/api/chat/sync`     | POST   | Synchronous chat completion                      |
| `/api/image`         | POST   | Image generation                                 |
| `/api/upload`        | POST   | File upload and processing                       |
| `/api/files`         | POST   | Forward file to LLM provider, returns `file_id`  |
| `/api/chart-extract` | POST   | Extract data points from chart image             |
| `/api/skill-creator` | POST   | Non-streaming LLM call for skill generation      |
| `/api/models`        | GET    | Get available models                             |
| `/health`            | GET    | Health check with timestamp and version          |

### 2.2 Rate Limiting

| Endpoint             | Window | Default Max  |
| -------------------- | ------ | ------------ |
| `/api/chat`          | 60s    | 20 requests  |
| `/api/image`         | 60s    | 5 requests   |
| `/api/upload`        | 60s    | 10 requests  |
| `/health`, `/models` | 60s    | 120 requests |

Skill Creator endpoint: 10 generations/hour/IP.

### 2.3 Timeouts

| Operation               | Duration   |
| ----------------------- | ---------- |
| Standard chat models    | 5 minutes  |
| Reasoning models        | 5 minutes  |
| Image generation        | 3 minutes  |
| Per-read operation      | 30 seconds |
| Outlook API calls       | 3 seconds  |
| Overall request timeout | 10 minutes |

Implementation constants (from `llmClient.js`):

```javascript
const TIMEOUTS = {
  CHAT_STANDARD: 300_000,   // 5 minutes
  CHAT_REASONING: 300_000,  // 5 minutes
  IMAGE: 180_000,           // 3 minutes
}
```

### 2.4 File Processing

**Supported formats with extraction:**

- PDF → text extraction via pdf-parse
- DOCX → text extraction via mammoth
- XLSX/XLS/CSV → all sheets to CSV format via xlsx
- TXT/MD → direct UTF-8 decoding
- Images (PNG, JPG, WEBP, GIF) → base64 encoding with data-URI

**Limits:**

- 10 MB per file (50 MB for `/api/files` provider upload)
- 600K characters after extraction (truncated with notification if exceeded)
- MIME type detection (not just declared type)

**Rate limit resilience**: When the upstream LLM provider returns a 429, the backend respects the `Retry-After` response header for retry timing. If all retries are exhausted, the client receives a `RATE_LIMITED` error code with a user-friendly message.

### 2.5 Streaming (SSE)

- Server-Sent Events format with delta updates
- `stream_options: { include_usage: true }` is set on all streaming requests
- Client disconnect detection with upstream cancellation
- Backpressure handling (waits for drain event)
- 30-second per-read timeout
- Token usage in final chunk
- Tool call deltas for agentic use

### 2.6 Request Retry Logic

The actual implementation in `llmClient.js` uses exponential backoff via `withRetry()`:

- **Exponential backoff:** `1s * 2^(attempt-1)`, capped at 8s (delays: 1s, 2s, 4s, 8s max)
- **3 total attempts** (2 retries) for all request types (chat, image)
- **Retries on:** HTTP 429 (rate limit) and 5xx (server errors), plus network-level failures
- **Respects `Retry-After` header:** Parsed as seconds (float) or HTTP-date; capped at 60s
- **Minimum retry floor:** 5s minimum on final rate-limit error to prevent hammering
- **Respects `AbortSignal`:** No retry after user cancellation
- **`RateLimitError` class:** Thrown when all retries exhausted on 429, carries `retryAfterMs`

---

## 3. Model Configuration

### 3.1 Model Tiers

Based on actual values from `config/models.js`:

| Tier | Default Model | Max Tokens | Context Window | Temperature | Special |
|------|--------------|------------|----------------|-------------|---------|
| Standard | gpt-5.2 | 32,000 | 400,000 | 0.7 | General purpose |
| Reasoning | gpt-5.2 | 65,000 | 400,000 | 1.0 | `reasoning_effort: high` |
| Image | gpt-image-1 | N/A | N/A | N/A | Image generation |

All values are configurable via environment variables:

| Environment Variable | Default |
|---|---|
| `MODEL_STANDARD` | `gpt-5.2` |
| `MODEL_STANDARD_MAX_TOKENS` | `32000` |
| `MODEL_STANDARD_CONTEXT_WINDOW` | `400000` |
| `MODEL_STANDARD_TEMPERATURE` | `0.7` |
| `MODEL_STANDARD_REASONING_EFFORT` | undefined |
| `MODEL_REASONING` | `gpt-5.2` |
| `MODEL_REASONING_MAX_TOKENS` | `65000` |
| `MODEL_REASONING_CONTEXT_WINDOW` | `400000` |
| `MODEL_REASONING_TEMPERATURE` | `1` |
| `MODEL_REASONING_EFFORT` | `high` |
| `MODEL_IMAGE` | `gpt-image-1` |
| `LLM_API_BASE_URL` | `https://litellm.kickmaker.net/v1` |
| `LLM_API_KEY` | required in production |

### 3.2 Model Detection

- **GPT-5.x** (`isGpt5Model`): Matches model IDs starting with `gpt-5` (case-insensitive)
  - Uses `max_completion_tokens` instead of `max_tokens`
  - Includes `reasoning_effort` parameter (when tier is not image)
  - Sampling temperature is NOT sent (GPT-5.x controls its own temperature)
- **ChatGPT models** (`isChatGptModel`): Matches model IDs starting with `chatgpt-` (case-insensitive)
  - Does not send `max_tokens` or `temperature` (no legacy parameter support)

**Request body construction** (`buildChatBody`):
- `stream_options: { include_usage: true }` added when streaming
- Empty `tool_calls` arrays stripped from assistant messages (Azure/LiteLLM rejects them)
- `tool_choice: 'auto'` set when tools are present, except for `gpt-5.2` models (Azure does not support explicit `tool_choice`)
- `reasoning_effort` only set for GPT-5.x models on non-image tiers

### 3.3 Tool Constraints

- **Maximum tools per request:** 128 (configurable via `MAX_TOOLS` env var)
- **Tool choice:** `'auto'` — model decides when to call tools (omitted for gpt-5.2)
- **Function name validation:** Regex `/^[a-zA-Z0-9_-]{1,64}$/`
- **Strict schema:** Optional boolean flag for strict JSON schema validation

---

## 4. Request Authentication

- **Required headers:** `X-User-Key`, `X-User-Email` sent with every request
- **Server-side validation:**
  - Email format validation (regex-based)
  - `X-User-Key` minimum length: 8 characters
  - Custom error messages for each validation failure
- **CSRF protection:** Token extracted from cookies, sent via `x-csrf-token` header
- **Credential retrieval:** Async fresh retrieval on every request

**Authentication Error Handling:**

| Error Type          | Displayed Message                                         |
| ------------------- | --------------------------------------------------------- |
| Missing credentials | "Please configure your credentials in Settings > Account" |
| 401 error           | "Authentication required"                                 |
| Timeout             | "Request timed out. The model took too long..."           |
| Network error       | "Connection error. Check your network..."                 |
| Rate limit          | "Too many requests — rate limit reached. Please wait..."  |
| Server error        | "Internal server error..."                                |

---

## 5. Security Implementation

- **CORS:** Allowed origins configured, credentials enabled
- **Helmet.js:** Security headers (HSTS in production)
- **Sensitive header redaction:** API keys never logged; headers `x-user-key`, `x-user-email`, `authorization`, `api-key` never logged
- **Trust proxy:** Correct client IP identification behind reverse proxy
- **JSON body limit:** 4MB maximum request body size
- **Header injection prevention:** `sanitizeHeaderValue()` strips `\r`, `\n`, and non-printable characters from forwarded headers
- **Credential encryption:** AES-GCM 256-bit with random unique key per installation and random 12-byte IV per operation

---

## 6. Tool Reference

### 6.1 Word Tools (34)

#### Document Reading (8)

| Tool                          | Description                                         | Parameters                            |
| ----------------------------- | --------------------------------------------------- | ------------------------------------- |
| getSelectedText               | Get currently selected text as plain text            | none                                  |
| getDocumentContent            | Get full document body as plain text                 | none                                  |
| getDocumentHtml               | Get full document as HTML with formatting            | none                                  |
| getDocumentProperties         | Get paragraph count, word count, character count     | none                                  |
| getSelectedTextWithFormatting | Get selection as Markdown with formatting preserved  | none                                  |
| getSpecificParagraph          | Read one paragraph by index                          | index (0-based)                       |
| findText                      | Search document and return match count               | searchText, matchCase, matchWholeWord |
| getComments                   | List all comments in document                        | none                                  |

#### Content Insertion (8)

| Tool               | Description                                                        | Key Parameters                                                                                  |
| ------------------ | ------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| insertContent      | **PREFERRED** — Add content with Markdown support                  | content, location (Start/End/Before/After/Replace), target (Selection/Body), preserveFormatting |
| searchAndReplace   | **PREFERRED for corrections** — Find and replace text              | searchText, replaceText, matchCase, matchWholeWord                                              |
| proposeRevision    | **PREFERRED for editing** — Diff-based revision with Track Changes | originalText, revisedText                                                                       |
| insertHyperlink    | Insert clickable link                                              | address, textToDisplay                                                                          |
| insertFootnote     | Add footnote at selection                                          | text                                                                                            |
| insertHeaderFooter | Add headers/footers                                                | headerText, footerText, location (Primary/FirstPage/EvenPages)                                  |
| insertSectionBreak | Insert section break                                               | breakType                                                                                       |
| addComment         | Add review comment bubble                                          | text, location                                                                                  |

#### Text Formatting (4)

| Tool                  | Description                            | Key Parameters                                                                                              |
| --------------------- | -------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| formatText            | Apply character formatting             | bold, italic, underline, fontSize, fontColor (hex), highlightColor                                          |
| applyStyle            | Apply Word built-in styles             | styleName (Normal, Heading 1-9, Title, Subtitle, Quote, etc.)                                               |
| setParagraphFormat    | Set paragraph formatting               | alignment, lineSpacing, spaceBefore, spaceAfter, leftIndent, firstLineIndent                                |
| applyTaggedFormatting | Convert inline tags to real formatting | tagName, fontName, fontSize, color, bold, italic, underline, strikethrough, allCaps, subscript, superscript |

#### Table Management (5)

| Tool                 | Description          | Key Parameters                                                                  |
| -------------------- | -------------------- | ------------------------------------------------------------------------------- |
| modifyTableCell      | Replace cell content  | row, column, text, tableIndex                                                   |
| addTableRow          | Add rows to table     | tableIndex, location (Before/After), count, values                              |
| addTableColumn       | Add columns to table  | tableIndex, location (Before/After), count, values                              |
| deleteTableRowColumn | Delete rows/columns   | tableIndex, rowIndex, columnIndex, deleteWhat                                   |
| formatTableCell      | Style table cells     | tableIndex, row, column, fillColor, fontName, fontSize, fontColor, bold, italic |

#### Document Structure (1)

| Tool         | Description     | Key Parameters                                                                                                 |
| ------------ | --------------- | -------------------------------------------------------------------------------------------------------------- |
| setPageSetup | Set page layout | marginTop, marginBottom, marginLeft, marginRight, orientation (Portrait/Landscape), pageSize (Letter/A4/Legal) |

#### Track Changes

- **Three strategies:** Token-based, sentence-based, block-based with automatic fallback
- **Statistics returned:** Count of insertions, deletions, and unchanged items

#### Custom Code Execution (eval_wordjs)

- Last resort — only when no dedicated tool exists
- Code validated before execution (load/sync patterns, try/catch)
- Sandboxed via SES `Compartment` (Word namespace only)
- Required pattern: `load()` → `sync()` → access → `sync()`

### 6.2 Excel Tools (28)

#### Data Reading (10)

| Tool                          | Description                                          | Returns                                                           |
| ----------------------------- | ---------------------------------------------------- | ----------------------------------------------------------------- |
| getSelectedCells              | Get values, address, dimensions of selection         | JSON with address, rowCount, columnCount, values (2D array)       |
| getWorksheetData              | Get all data from used range                         | values, address, rowCount, columnCount                            |
| getWorksheetInfo              | Get workbook structure                               | activeName, position, usedRange, totalSheets, sheetNames          |
| getNamedRanges                | List all named ranges                                | names and formulas                                                |
| getConditionalFormattingRules | List conditional formatting rules on a range         | rules with type, conditions, format                               |
| findData                      | Search values workbook-wide with optional pagination | matches, totalFound, offset, hasMore, nextOffset                  |
| getAllObjects                  | List charts and pivot tables                         | object details                                                    |
| getRangeAsCsv                 | Export a range as compact CSV text                   | CSV rows (2-3x fewer tokens than JSON), with truncation indicator |
| screenshotRange               | Capture a range as a PNG image                       | Image injected into vision context for AI visual verification     |
| detectDataHeaders             | Detect column/row headers in a range                 | hasColumnHeaders, hasRowHeaders, suggestedHasHeaders, suggestedSeriesBy, detected labels |

#### Writing and Editing (14)

| Tool                    | Description                                             | Key Parameters                                                                      |
| ----------------------- | ------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| setCellRange            | **PREFERRED** — Write values, formulas, or formatting   | address, sheetName, values (2D array), formulas (2D array), formatting, copyToRange |
| clearRange              | Clear contents or formatting                            | address, clearContents, clearFormatting                                              |
| modifyStructure         | Insert/delete/hide/unhide rows/columns, freeze panes   | sheetName, operation, dimension, reference, count                                   |
| modifyWorkbookStructure | Create, delete, rename, or duplicate worksheets         | operation, sheetName, newName, tabColor                                             |
| addWorksheet            | Add a new worksheet to the workbook                     | sheetName, position                                                                 |
| createTable             | Convert range to structured table                       | address, hasHeaders, tableName, style                                               |
| sortRange               | Sort data by column                                     | columnIndex, ascending, hasHeaders                                                  |
| searchAndReplace        | Find and replace values in a sheet or range             | searchText, replaceText, sheetName, matchCase                                       |
| setNamedRange           | Create or update a named range                          | name, address, sheetName                                                            |
| protectWorksheet        | Protect or unprotect a worksheet                        | sheetName, protect, password                                                        |
| importCsvToSheet        | Import CSV text into a sheet with type coercion         | csvText, sheetName, startAddress, hasHeaders                                        |
| imageToSheet            | Convert an uploaded image to pixel art using cell colors | imageId, sheetName, startAddress, width, height                                     |
| extract_chart_data      | Extract data points from a chart image (vision-based)   | imageId                                                                             |
| clearAgentHighlights    | Remove all AI-applied highlight colors from workbook    | none                                                                                |

#### Formatting (2)

| Tool                       | Description                    | Key Parameters                                                                                                                   |
| -------------------------- | ------------------------------ | -------------------------------------------------------------------------------------------------------------------------------- |
| formatRange                | Apply comprehensive formatting | address, fillColor, fontColor, bold, italic, fontSize, fontName, borders, alignment, wrapText, borderStyle/Color/Weight per edge |
| applyConditionalFormatting | Add conditional format rules   | address, rule type, conditions, format                                                                                           |

#### Custom Code Execution (1)

| Tool          | Description                                            | Key Parameters     |
| ------------- | ------------------------------------------------------ | ------------------ |
| eval_officejs | Execute custom Office.js code in an Excel.run context  | code, explanation  |

#### Conditional Formatting Types

| Type           | Description                                                 |
| -------------- | ----------------------------------------------------------- |
| Cell value     | Comparison (equal, not equal, greater, less, between...)    |
| Text match     | Contains, starts with, ends with                            |
| Custom formula | Formatting based on formula                                 |
| Color scale    | Color gradient min→max                                      |
| Data bars      | Proportional visual bars                                    |
| Icon sets      | Icons based on thresholds (traffic lights, arrows, symbols) |

#### Chart Types

Column (Clustered, Stacked), Line (simple, with markers), Pie, Bar (Clustered), Area, Doughnut, XY Scatter.

Charts/Pivots managed via `manageObject` tool: create from data range with anchor positioning, set title and dimensions, update type/source/title, delete existing.

**Header auto-detection:** Before creating a chart, the agent calls `detectDataHeaders` on the source range to determine whether the first row/column contains labels. The result drives `hasHeaders` and `seriesBy` parameters.

### 6.3 PowerPoint Tools (24)

#### Presentation Reading (9)

| Tool                 | Description                                        | Parameters                                      |
| -------------------- | -------------------------------------------------- | ----------------------------------------------- |
| getSelectedText      | Get currently selected text                        | none                                            |
| getSlideContent      | Get all text from a specific slide                 | slideNumber (1-based)                           |
| getShapes            | Get all shapes with properties                     | slideNumber (1-based)                           |
| getAllSlidesOverview  | Get text overview of entire presentation           | none                                            |
| getCurrentSlideIndex | Get the index of the slide currently viewed        | none — returns 1-based slide number             |
| getSpeakerNotes      | Get speaker notes for a slide                      | slideNumber (1-based)                           |
| screenshotSlide      | Capture a slide as PNG for visual verification     | slideNumber (1-based, optional)                 |
| verifySlides         | Detect shape overflows and overlaps on all slides  | none — returns a text report of issues          |
| searchIcons          | Search the Iconify icon library by keyword         | query, limit, prefix (icon set filter)          |

#### Content Insertion and Modification (15)

| Tool                              | Description                                                       | Key Parameters                                                    |
| --------------------------------- | ----------------------------------------------------------------- | ----------------------------------------------------------------- |
| insertContent                     | **PREFERRED** — Add/replace content with Markdown                 | content, slideNumber, shapeIdOrName                               |
| proposeShapeTextRevision          | Modify shape text with diff tracking                              | slideNumber, shapeIdOrName, revisedText                           |
| replaceShapeParagraphs            | Replace paragraphs in a shape while preserving formatting         | slideNumber, shapeIdOrName, paragraphs                            |
| searchAndReplaceInShape           | Find and replace text in a specific shape                         | slideNumber, shapeIdOrName, searchText, replaceText               |
| searchAndFormatInPresentation     | Search text across all slides and apply formatting                | searchText, format                                                |
| replaceSelectedText               | Replace the currently selected text                               | text                                                              |
| setSpeakerNotes                   | Set speaker notes for a slide                                     | slideNumber, notes                                                |
| addSlide                          | Add new slide — picks the best layout from the slide master       | layout (Blank, Title, TitleAndContent...)                         |
| deleteSlide                       | Delete slide by number                                            | slideNumber (1-based)                                             |
| duplicateSlide                    | Copy a slide; insert duplicate after the original                 | slideNumber (1-based)                                             |
| reorderSlide                      | Move a slide to a different position                              | slideNumber, targetPosition (1-based)                             |
| insertIcon                        | Insert an Iconify icon as an image on a slide                     | iconId ("prefix:name"), slideNumber, left, top, width, height, color |
| insertImageOnSlide                | Insert an uploaded image on a slide                               | imageId, slideNumber, left, top, width, height                    |
| eval_powerpointjs                 | Execute custom Office.js code in a PowerPoint.run context         | code, explanation                                                 |
| editSlideXml                      | Directly edit slide OOXML for operations beyond the API           | slideNumber, code (JS with JSZip access), explanation             |

#### Icon Library

200,000+ icons across 150+ open-source icon sets (Material Design, Fluent UI, Feather, Bootstrap, Heroicons). Icons rendered as SVG images, positionable and sizable. Color settable at insertion time.

#### OOXML Direct Editing

`editSlideXml` provides an escape hatch for operations the Office.js API cannot express (chart XML manipulation, SmartArt, complex animations). The agent receives access to the slide's ZIP archive and can modify the underlying XML directly.

### 6.4 Outlook Tools (9)

#### Email Reading (4)

| Tool               | Description                                | Returns                                |
| ------------------ | ------------------------------------------ | -------------------------------------- |
| getEmailBody       | Get full email body (read or compose mode) | Text with automatic image preservation |
| getEmailSubject    | Get email subject                          | Subject line                           |
| getEmailRecipients | Get To, Cc, Bcc recipients                | JSON with arrays                       |
| getEmailSender     | Get sender info                            | JSON with displayName, emailAddress    |

#### Email Writing (5)

| Tool            | Description                                              | Key Parameters                                      |
| --------------- | -------------------------------------------------------- | --------------------------------------------------- |
| writeEmailBody  | **PREFERRED** — Modify email body                        | content, mode (Append/Insert/Replace), diffTracking |
| setEmailSubject | Set email subject                                        | subject                                             |
| addRecipient    | Add recipients                                           | field (to/cc/bcc), recipients (comma-separated)     |
| addAttachment   | Add a file attachment to the email                       | url, attachmentType, attachmentName                 |
| eval_outlookjs  | Execute custom Office.js code in the Outlook mailbox context | code, explanation                               |

#### Email Body Writing Modes

| Mode                 | Description                          | Use Case                            |
| -------------------- | ------------------------------------ | ----------------------------------- |
| **Append (DEFAULT)** | Adds at end, preserves history       | Replies, forwards — ALWAYS use this |
| **Insert**           | Inserts at cursor with optional diff | Specific text replacement in draft  |
| **Replace**          | Replaces entire body                 | Brand new emails ONLY               |

### 6.5 General Tools (6)

| Tool           | Category | Description                              |
| -------------- | -------- | ---------------------------------------- |
| getCurrentDate | read     | Get current date/time in various formats |
| calculateMath  | write    | Evaluate mathematical expressions safely |
| executeBash    | write    | Execute bash commands in sandboxed VFS   |
| vfsWriteFile   | write    | Write files to virtual filesystem        |
| vfsReadFile    | read     | Read files from virtual filesystem       |
| vfsListFiles   | read     | List files in VFS uploads directory      |

### 6.6 Tool Decision Trees

**Word Content Modification:**

```
Simple word/phrase replacement → searchAndReplace
Rewriting sentences/paragraphs → proposeRevision (preserves formatting)
Adding new content → insertContent
```

**Excel Data Operations:**

```
Write data → setCellRange (ALWAYS preferred)
Read large data → getRangeAsCsv (token-efficient)
Search paginated → findData with maxResults + offset
Format data → formatRange
Create table → createTable
Charts/Pivots → manageObject
Sheet management → modifyWorkbookStructure (create/delete/rename/duplicate)
Visual check → screenshotRange (after formatting changes)
Advanced → eval_officejs
```

**PowerPoint Workflow:**

```
1. getAllSlidesOverview (understand structure)
2. getShapes(slideNumber) (discover shape IDs)
3. insertContent or proposeShapeTextRevision (modify text)
4. screenshotSlide (verify visual result)
5. verifySlides (check for overflows/overlaps)
```

**PowerPoint Icons:**

```
1. searchIcons (find icon by keyword, optionally filter by set)
2. insertIcon (insert by "prefix:name" on target slide)
```

**PowerPoint Advanced (OOXML):**

```
editSlideXml → when no dedicated tool covers the operation
(charts, SmartArt, animations, master layouts)
```

**Outlook Email:**

```
Reply/Forward → ALWAYS mode "Append"
New email → Can use mode "Replace"
```

---

## 7. UI Implementation Details

### 7.1 Color Palette

| Element              | Light Mode | Dark Mode |
| -------------------- | ---------- | --------- |
| Primary text         | #1d1d1f    | #f5f5f7   |
| Secondary text       | #6e6e73    | #a1a1a6   |
| Primary background   | #ffffff    | #000000   |
| Secondary background | #f5f5f7    | #1c1c1e   |
| Accent               | #33abc6    | #33abc6   |
| Success              | #34c759    | #34c759   |
| Warning              | #f1930f    | #f1930f   |
| Danger               | #ff3b30    | #ff3b30   |

### 7.2 Accessibility (ARIA)

- `aria-live="polite"` on message list and status indicators
- `role="log"` on chat container
- `role="status"` on activity indicators
- `role="tab"` and `role="tabpanel"` on settings tabs
- `aria-label` on all icon buttons
- `aria-selected` on tabs
- SR-only announcement div for screen readers

#### Motion Preferences

- All animations disabled if `prefers-reduced-motion: reduce`

### 7.3 Keyboard Navigation

- **Tab:** Navigate between elements
- **Enter:** Send message (in textarea)
- **Shift+Enter:** New line in message
- **Escape:** Close dropdowns
- Visible focus ring with accent color

### 7.4 Responsive Design

- **Responsive grid:** 2-column layout on medium+ screens
- **Max-width controls:** Text truncation with ellipsis
- **Flex layouts:** Responsive wrapping
- **Touch-friendly:** Minimum 7px button height
- **Dropdown positioning:** Auto-adjusts based on viewport space

### 7.5 Animations

- **Draft focus glow:** 3-iteration pulse animation on input focus
- **Spinner dots:** Animated ellipsis during generation
- **Streaming dots:** Three animated dots in thinking blocks
- **Button hover:** Slight upward translate with shadow transition
- **Thinking block:** Collapsible with chevron rotation animation

---

## 8. Logging & Monitoring

### 8.1 Server-Side Logging

- **Morgan HTTP logger:** Method, URL, status, content-length, response time
- **Custom system logger:** INFO, ERROR, WARN, DEBUG levels
- **Rotating file stream:** 10MB per file, 30 max files, gzip compression, daily rotation
- **Log location:** `/logs/kickoffice.log`

### 8.2 Request Tracking

- Unique UUID per request in `res.locals.reqId`
- Verbose logging mode (`VERBOSE_LOGGING=true`) for debugging

### 8.3 Sensitive Data Protection

- API keys redacted from logs
- Sensitive headers (`x-user-key`, `x-user-email`, `authorization`, `api-key`) never logged
- Base64 image payloads sanitized via `sanitizePayloadForLogs` before logging

---

## 9. Graceful Shutdown

- **Signal handling:** SIGTERM and SIGINT
- **Process:** Stop accepting connections → wait for in-flight requests → close all connections
- **Force exit:** After 30 seconds if still hanging
- **Logging:** Shutdown signals and completion status logged

---

## 10. Host-Specific Technical Rules

### 10.1 Word

- **Track Changes:** 3 strategies (token, sentence, block) with auto-fallback
- **eval_wordjs:** SES `Compartment` sandbox, load/sync patterns required
- **2D array conventions** for table operations (values passed as 2D arrays)
- **Modification history:** Native Track Changes — AI modifications visible in Review pane
- **Undo:** Ctrl+Z works for AI modifications (grouped as single action via Office.js `run()`)

### 10.2 Excel

- **2D arrays required** for values and formulas
- **Dimensions must match** range
- **No iteration modification** — never modify cells while iterating
- **Use `getUsedRange()`** for data bounds
- **Formula localization:** English uses comma (e.g., `=SUM(A1,B1)`), French uses semicolon (e.g., `=SOMME(A1;B1)`)
- **Modification history:** Comment-based tracking — "Modified by AI. Old value: [X]"
- **Paginated search:** `findData` supports `maxResults` and `offset` parameters with `nextOffset`
- **CSV export:** `getRangeAsCsv` preferred over JSON for large ranges (significant token savings)
- **Visual verification:** `screenshotRange` for visual inspection without reading raw cell data

### 10.3 PowerPoint

- **Slide numbers:** 1-based in UI, 0-indexed in code arrays
- **No host-specific `run()` context** — uses `Office.context.document` with `CoercionType.Text`
- **Shape discovery workflow:** `getAllSlidesOverview` → `getShapes` → `insertContent`
- **No Track Changes:** Must use Speaker Notes logging or slide duplication for modification visibility
- **Active slide detection:** `getCurrentSlideIndex` returns the currently viewed slide
- **Slide master modification:** Limited via standard API; advanced via `editSlideXml` (OOXML)

### 10.4 Outlook

- **Callback pattern** (not promises) — must wrap in Promise
- **3-second timeout** on all operations
- **Compose mode only** for write operations
- **Email history protection:** Auto-detection of replies/forwards, defaults to Append mode, `{{PRESERVE_N}}` placeholders for embedded images
- **Reply language:** ALWAYS reply in same language as original email
- **Constraints:** Cannot read inbox or other emails; no calendar integration; no email attachment analysis
