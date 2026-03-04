# Product Requirements Document (PRD): KickOffice

## 1. Product Vision & Target Audience

**KickOffice** is an enterprise-grade AI assistant natively integrated into the Microsoft Office suite (Word, Excel, PowerPoint, Outlook). Its goal is to boost productivity, automate repetitive tasks, and assist in complex document/data manipulation while keeping all data flows secure and internal.

**Target Audience:**
All employees within the organization, specifically tailored to handle the diverse needs of:
- Engineers (Hardware, Software, Firmware, Mechanical)
- Project Managers
- Accounting & Finance
- Sales Representatives
- Administrative Services

---

## 2. Deployment, Quotas & Telemetry

- **Deployment:** Internal distribution. Users download/install the add-in via a manifest file hosted on the company's internal SharePoint or local server.
- **Monetization & Quotas:** No user-facing subscriptions or internal billing. Quotas, rate limiting, and access control are centrally managed by the internal LLM gateway (LiteLLM).
- **Telemetry & DLP:** The add-in itself does not log telemetry or block sensitive data. All AI telemetry, auditing, and DLP filtering are strictly delegated to the internal LiteLLM gateway.

---

## 3. Cross-Application Features (All Apps)

### 3.1 Chat Interface

#### Visual Components
- **Chat header**: KickOffice logo, subtitle, status indicator (3 dots)
- **Input area**: Auto-expanding text field with implicit character counting
- **Message list**: Messages displayed with timestamps (HH:MM format), clear user/assistant separation
- **Stats bar**: Input/output tokens (with K/M formatting), context window percentage, active model name, shell activity indicator

#### Session Management
- **Multiple sessions**: Users can maintain several independent conversations
- **Session switching**: Dropdown listing all sessions with message count
- **New chat**: Button to start a fresh conversation
- **Session deletion**: Ability to delete a session (disabled if only one remains)
- **Auto-naming**: Sessions automatically named with date/time of first message (DD/MM/YY HH:mm format)
- **Persistence**: Sessions stored in IndexedDB, persist between restarts
- **Per-host isolation**: Each Office app (Word, Excel, PowerPoint, Outlook) has its own sessions

#### Model Selection
- **Tier selector**: Dropdown to choose between model tiers (Standard, Reasoning, Image)
- **Active model display**: Visual indication of current model in stats bar
- **Placeholder adaptation**: Input placeholder changes based on selected model type (e.g., "Describe Image" for image tier)
- **Persistence**: Model selection saved in localStorage

#### File Attachments
- **File upload**: Up to 3 files per message
- **Supported formats**: PDF, DOCX, XLSX, XLS, CSV, TXT, MD, PNG, JPEG, JPG, WEBP, GIF
- **Size limit**: 10 MB per file
- **Content limit**: 100K characters after text extraction
- **Drag-and-drop**: Support with visual feedback (ring border on dragover)
- **Preview**: Attached files shown as badges with X button for removal
- **Duplicate detection**: Prevents adding same file twice by name
- **Error messages**: Specific notifications for oversized files, unsupported formats, or max files reached

#### Context Options
- **Include selection**: Checkbox to automatically include selected text from document
- **Word formatting**: Checkbox (Word only) to enable rich formatting in responses

### 3.2 Quick Actions

Contextual buttons at the top of the interface for one-click actions:
- **Tooltips**: Each button shows a description on hover
- **Disabled state**: Buttons grayed out during loading
- **Saved prompts dropdown**: Quick access to custom prompts
- **Host-specific actions**: Different actions per Office app

### 3.3 Message Display

#### User Messages
- Plain text display with timestamp
- Indication of attached files

#### Assistant Messages
- **Markdown rendering**: Full support (bold, italic, lists, code blocks, tables)
- **Thinking blocks**: Collapsible `<think>` sections showing AI reasoning process with brain icon
- **Tool calls**: Expandable blocks showing tool name, arguments (JSON), status, and results
- **Generated images**: Inline display of AI-created images
- **Streaming animation**: Three animated dots during response generation
- **DOMPurify sanitization**: XSS prevention on all rendered content

#### Tool Call Visualization
- **Running status**: Animated spinner (Loader2 icon)
- **Success status**: Green checkmark
- **Error status**: Red X with error message in red text
- **Arguments display**: JSON-formatted with syntax highlighting
- **Expandable/collapsible**: Click to show/hide details

#### Message Actions
- **Replace selection**: Insert message content replacing current selection
- **Append to selection**: Add content at the end of selection
- **Copy to clipboard**: Copy text with fallback for Office WebView environments

### 3.4 Image Generation

- **Dedicated model tier**: Uses specific image model (gpt-image-1)
- **Size options**: 1024x1024, 1024x1536, 1536x1024
- **Quality levels**: low, medium, high, auto
- **Multiple images**: Can generate 1-4 images per request
- **Prompt limit**: Maximum 4000 characters
- **Direct insertion**: Images inserted directly into document
- **Clipboard fallback**: If insertion fails, image copied to clipboard with notification
- **Extended timeout**: 3-minute timeout for generation

### 3.5 Autonomous Agent

- **Agentic loop**: AI can execute multiple tools in sequence for complex tasks
- **Iteration limit**: Configurable from 1 to 100 (default: 25)
- **Loop detection**: Sliding window signature comparison prevents infinite repetitive calls
- **User stop**: Button to stop agent at any time
- **Status indicator**: Real-time display of current action with emojis (⏳ analyzing, 🛠️ tools, 📖 reading, 🎨 formatting)
- **Tool selection**: Agent autonomously chooses appropriate tools for each task

### 3.6 Stats Bar Details

- **Input tokens (↑)**: Tokens sent to model, auto-formatted (K for thousands, M for millions)
- **Output tokens (↓)**: Tokens received from model
- **Context percentage**: Visual gauge of context window usage
- **Context gauge colors**:
  - Normal: Default color
  - Warning (80%+): Orange
  - Critical (100%): Red
- **Model name**: Currently selected model identifier
- **Activity pulse**: Red animated dot with Terminal icon when agent is executing
- **Hover tooltips**: Detailed token counts on hover

### 3.7 Virtual File System (VFS)

- **Sandboxed filesystem**: In-memory virtual filesystem for agent use
- **User uploads**: Files stored in `/home/user/uploads/`
- **Custom scripts**: Reusable bash scripts in `/home/user/scripts/`
- **Session persistence**: VFS state saved and restored with each session
- **Available commands**: ls, cat, grep, find, awk, sed, sort, uniq, wc, cut, head, tail, base64

---

## 4. Authentication & Security

### 4.1 Credential Management

#### Configuration
- **LiteLLM API key**: Masked input field (password type)
- **User email**: Input field for identification
- **Status indicator**: Colored badge (green = configured, yellow = missing)
- **Portal link**: Redirect to key generation portal (getkey.ai.kickmaker.net)

#### Credential Persistence
- **Remember me**: "Remember credentials" checkbox
- **Encrypted storage**: AES-GCM 256-bit encryption for saved credentials
- **Encryption key**: Random unique key generated per installation
- **Initialization vector**: Random 12-byte IV for each encryption operation
- **Auto-migration**: Detection and re-encryption of old unencrypted credentials
- **Session storage**: If remember disabled, credentials in sessionStorage (cleared on close)
- **Corrupted data recovery**: Graceful handling of decryption failures

#### Logout
- Manual credential deletion in settings
- Associated encryption keys cleared

### 4.2 Request Authentication

- **Required headers**: X-User-Key, X-User-Email sent with every request
- **Server-side validation**:
  - Email format validation (regex-based)
  - X-User-Key minimum length (8 characters)
  - Custom error messages for each validation failure
- **CSRF protection**: Token extracted from cookies, sent via x-csrf-token header
- **Credential retrieval**: Async fresh retrieval on every request

### 4.3 Authentication Error Handling

| Error Type | Displayed Message |
|------------|-------------------|
| Missing credentials | "Please configure your credentials in Settings > Account" |
| 401 error | "Authentication required" |
| Timeout | "Request timed out. The model took too long..." |
| Network error | "Connection error. Check your network..." |
| Rate limit | "Too many requests — rate limit reached. Please wait..." |
| Server error | "Internal server error..." |

### 4.4 Security Features

- **CORS**: Allowed origins configured, credentials enabled
- **Helmet.js**: Security headers (HSTS in production)
- **Sensitive header redaction**: API keys never logged
- **Trust proxy**: Correct client IP identification behind reverse proxy
- **JSON body limit**: 4MB maximum request body size

---

## 5. Settings Page

### 5.1 Account Tab
- API key field (masked with show/hide toggle)
- Email field
- Remember credentials checkbox with description
- Credential status badge (green/yellow)
- Link to configuration portal

### 5.2 General Tab
- **Interface language**: FR/EN selector with persistence
- **Dark mode**: On/off toggle with theme persistence
- **User profile**:
  - First and last name
  - Gender (Unspecified, Female, Male, Non-binary)
- **Agent max iterations**: Slider 1-100 (default 25)
- **Backend status**: Real-time indicator (green online, red offline)
- **Application version**
- **Available models list** (read-only display)

### 5.3 Custom Prompts Tab
- List of all saved prompts
- New prompt add button
- Inline editing (system prompt + user prompt)
- Deletion (except last one)
- Prompt preview with truncated display

### 5.4 Built-in Prompts Tab
- Default prompts per application (Translate, Polish, Academic, Summarize, Proofread...)
- Ability to edit each built-in prompt
- Reset to default button
- System and user prompt preview

### 5.5 Tools Tab
- Complete tool list with checkboxes (67 tools total)
- Enable/disable per tool
- Description for each tool
- Preference persistence per Office host
- Auto-migration when new tools added

---

## 6. Microsoft Word

### 6.1 Document Reading Tools

| Tool | Description | Parameters |
|------|-------------|------------|
| getSelectedText | Get currently selected text as plain text | none |
| getDocumentContent | Get full document body as plain text | none |
| getDocumentHtml | Get full document as HTML with formatting | none |
| getDocumentProperties | Get paragraph count, word count, character count | none |
| getSelectedTextWithFormatting | Get selection as Markdown with formatting preserved | none |
| getSpecificParagraph | Read one paragraph by index | index (0-based) |
| findText | Search document and return match count | searchText, matchCase, matchWholeWord |
| getComments | List all comments in document | none |

### 6.2 Content Insertion Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| insertContent | **PREFERRED** - Add content with Markdown support | content, location (Start/End/Before/After/Replace), target (Selection/Body), preserveFormatting |
| searchAndReplace | **PREFERRED for corrections** - Find and replace text | searchText, replaceText, matchCase, matchWholeWord |
| proposeRevision | **PREFERRED for editing** - Diff-based revision with Track Changes | originalText, revisedText |
| insertHyperlink | Insert clickable link | address, textToDisplay |
| insertFootnote | Add footnote at selection | text |
| insertHeaderFooter | Add headers/footers | headerText, footerText, location (Primary/FirstPage/EvenPages) |
| insertSectionBreak | Insert section break | breakType |
| addComment | Add review comment bubble | text, location |

### 6.3 Text Formatting Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| formatText | Apply character formatting | bold, italic, underline, fontSize, fontColor (hex), highlightColor |
| applyStyle | Apply Word built-in styles | styleName (Normal, Heading 1-9, Title, Subtitle, Quote, etc.) |
| setParagraphFormat | Set paragraph formatting | alignment, lineSpacing, spaceBefore, spaceAfter, leftIndent, firstLineIndent |
| applyTaggedFormatting | Convert inline tags to real formatting | tagName, fontName, fontSize, color, bold, italic, underline, strikethrough, allCaps, subscript, superscript |

### 6.4 Table Management Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| modifyTableCell | Replace cell content | row, column, text, tableIndex |
| addTableRow | Add rows to table | tableIndex, location (Before/After), count, values |
| addTableColumn | Add columns to table | tableIndex, location (Before/After), count, values |
| deleteTableRowColumn | Delete rows/columns | tableIndex, rowIndex, columnIndex, deleteWhat |
| formatTableCell | Style table cells | tableIndex, row, column, fillColor, fontName, fontSize, fontColor, bold, italic |

### 6.5 Document Structure Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| setPageSetup | Set page layout | marginTop, marginBottom, marginLeft, marginRight, orientation (Portrait/Landscape), pageSize (Letter/A4/Legal) |

### 6.6 Track Changes (CRITICAL REQUIREMENT)

- **Mandatory activation**: AI MUST use native "Track Changes" feature when modifying existing text
- **Word-by-word diff**: Smart diff algorithm preserving formatting on unmodified portions
- **Three strategies**: Token-based, sentence-based, block-based with automatic fallback
- **Statistics returned**: Count of insertions, deletions, and unchanged items
- **User review**: User can accept/reject each modification individually via Word's native UI

### 6.7 Modification History (Word-specific)

- **Native Track Changes**: AI modifications visible in Word's Review pane
- **Accept/Reject workflow**: Standard Word revision review process
- **Undo support**: Native Ctrl+Z works for AI modifications (grouped as single action)

### 6.8 Word Quick Actions

| Action | Description |
|--------|-------------|
| Proofread | Grammar and spelling correction with Track Changes |
| Translate | FR↔EN translation of selection |
| Polish | Style and clarity improvement |
| Academic | Formalization for academic writing |
| Summarize | Concise summary generation |

### 6.9 Custom Code Execution (eval_wordjs)

- **Last resort**: Only when no dedicated tool exists
- **Code validation**: Validated before execution (load/sync patterns, try/catch)
- **Sandboxed**: Word namespace only
- **Required pattern**: load() → sync() → access → sync()

---

## 7. Microsoft Excel

### 7.1 Data Reading Tools

| Tool | Description | Returns |
|------|-------------|---------|
| getSelectedCells | Get values, address, dimensions of selection | JSON with address, rowCount, columnCount, values (2D array) |
| getWorksheetData | Get all data from used range | values, address, rowCount, columnCount |
| getWorksheetInfo | Get workbook structure | activeName, position, usedRange, totalSheets, sheetNames |
| getDataFromSheet | Read from any sheet by name | data from specified sheet |
| getNamedRanges | List all named ranges | names and formulas |
| findData | Search values workbook-wide | matches with locations |
| getAllObjects | List charts and pivot tables | object details |

### 7.2 Writing and Editing Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| setCellRange | **PREFERRED** - Write values, formulas, or formatting | address, sheetName, values (2D array), formulas (2D array), formatting, copyToRange |
| clearRange | Clear contents or formatting | address, clearContents, clearFormatting |
| modifyStructure | Insert/delete rows/columns, freeze panes | sheetName, operation, rowIndex, columnIndex |
| addWorksheet | Create new worksheet | sheetName |
| createTable | Convert range to structured table | address, hasHeaders, tableName, style |
| sortRange | Sort data by column | columnIndex, ascending, hasHeaders |

### 7.3 Formatting Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| formatRange | Apply comprehensive formatting | address, fillColor, fontColor, bold, italic, fontSize, fontName, borders, alignment, wrapText, borderStyle/Color/Weight per edge |
| applyConditionalFormatting | Add conditional format rules | address, rule type, conditions, format |

### 7.4 Conditional Formatting Types

| Type | Description |
|------|-------------|
| Cell value | Comparison (equal, not equal, greater, less, between...) |
| Text match | Contains, starts with, ends with |
| Custom formula | Formatting based on formula |
| Color scale | Color gradient min→max |
| Data bars | Proportional visual bars |
| Icon sets | Icons based on thresholds (traffic lights, arrows, symbols) |

### 7.5 Chart Management

**Supported Chart Types:**
- Column (Clustered, Stacked)
- Line (simple, with markers)
- Pie
- Bar (Clustered)
- Area
- Doughnut
- XY Scatter

**Features via manageObject tool:**
- Create from data range with anchor positioning
- Set title and dimensions
- Update type, source, or title
- Delete existing charts

### 7.6 Pivot Table Management

- Create from data range
- Custom naming
- Controlled positioning
- Delete existing pivot tables

### 7.7 Modification History (Excel-specific)

- **Comment-based tracking**: When modifying a cell, AI MUST insert a comment containing previous value
- **Comment format**: "Modified by AI. Old value: [X]"
- **Native undo**: Ctrl+Z works for AI modifications

### 7.8 Excel Quick Actions

| Action | Mode | Description |
|--------|------|-------------|
| Clean | Immediate | Detect and fix data quality issues |
| Beautify | Immediate | Apply professional formatting |
| Formula | Draft | Generate Excel formulas as needed |
| Transform | Draft | Restructure data (transpose, pivot...) |
| Highlight | Draft | Apply visual emphasis |

### 7.9 Formula Localization

- **English**: Comma separators (e.g., `=SUM(A1,B1)`)
- **French**: Semicolon separators (e.g., `=SOMME(A1;B1)`)
- User preference in settings
- Auto-conversion based on Excel locale

### 7.10 Excel-Specific Rules

- **2D arrays required**: Values and formulas MUST be 2D arrays
- **Dimensions must match**: Array dimensions MUST match range dimensions
- **No iteration modification**: Never modify cells while iterating
- **Use getUsedRange()**: To find data bounds before operations

---

## 8. Microsoft PowerPoint

### 8.1 Presentation Reading Tools

| Tool | Description | Parameters |
|------|-------------|------------|
| getSelectedText | Get currently selected text | none |
| getSlideContent | Get all text from a specific slide | slideNumber (1-based) |
| getShapes | Get all shapes with properties | slideNumber (1-based) |
| getAllSlidesOverview | Get text overview of entire presentation | none |

### 8.2 Content Insertion Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| insertContent | **PREFERRED** - Add/replace content with Markdown | content, slideNumber, shapeIdOrName |
| proposeShapeTextRevision | Modify shape text with diff tracking | slideNumber, shapeIdOrName, revisedText |
| addSlide | Add new slide | layout (Blank, Title, TitleAndContent...) |
| deleteSlide | Delete slide by number | slideNumber (1-based) |

### 8.3 Modification History (PowerPoint-specific)

PowerPoint has NO native Track Changes. Two workarounds:

1. **Speaker Notes logging**: AI logs modifications in the Speaker Notes section at the bottom of the slide
2. **Slide duplication**: For major structural changes, duplicate the original slide to allow before/after comparison

### 8.4 PowerPoint Quick Actions

| Action | Mode | Description |
|--------|------|-------------|
| Proofread | Immediate | Grammar/spelling correction |
| Translate | Immediate | FR↔EN translation |
| Notes | Immediate | Generate speaker notes (<100 words) |
| Impact | Immediate | Steve Jobs-style rewrite (punch, hook) |
| Visual | Immediate | Generate detailed AI image prompts |

### 8.5 Slide Style Rules

- Maximum 8-10 words per bullet point
- Maximum 6-7 bullets per slide
- Active voice, present tense preferred

### 8.6 PowerPoint-Specific Rules

- **Slide numbers**: 1-based in UI, 0-indexed in code arrays
- **Shape discovery workflow**: getAllSlidesOverview → getShapes → insertContent
- **No Track Changes**: Must use workarounds for modification visibility

### 8.7 Constraints

- **Out of scope**: No Slide Master modification

---

## 9. Microsoft Outlook

### 9.1 Email Reading Tools

| Tool | Description | Returns |
|------|-------------|---------|
| getEmailBody | Get full email body (read or compose mode) | Text with automatic image preservation |
| getEmailSubject | Get email subject | Subject line |
| getEmailRecipients | Get To, Cc, Bcc recipients | JSON with arrays |
| getEmailSender | Get sender info | JSON with displayName, emailAddress |

### 9.2 Email Writing Tools

| Tool | Description | Key Parameters |
|------|-------------|----------------|
| writeEmailBody | **PREFERRED** - Modify email body | content, mode (Append/Insert/Replace), diffTracking |
| setEmailSubject | Set email subject | subject |
| addRecipient | Add recipients | field (to/cc/bcc), recipients (comma-separated) |

### 9.3 Email Body Writing Modes

| Mode | Description | Use Case |
|------|-------------|----------|
| **Append (DEFAULT)** | Adds at end, preserves history | Replies, forwards - ALWAYS use this |
| **Insert** | Inserts at cursor with optional diff | Specific text replacement in draft |
| **Replace** | Replaces entire body | Brand new emails ONLY |

### 9.4 Email History Protection (CRITICAL)

- **Automatic protection**: Prevents accidental deletion of email threads
- **Auto-detection**: Recognizes replies/forwards by headers (From:, Sent:, To:, Subject:, quoted text >)
- **Safe default mode**: ALWAYS uses Append mode by default
- **Visual alerts**: Warning if Replace mode might erase history
- **Agent instructions**: Explicit instructions to NEVER delete email history
- **Image preservation**: Uses {{PRESERVE_N}} placeholders for embedded images

### 9.5 Outlook Quick Actions

| Action | Mode | Description |
|--------|------|-------------|
| **Proofread** | Immediate | Grammar and spelling correction |
| **Translate & Formalize** | Immediate | FR↔EN translation with professional formalization |
| **Concise** | Immediate | Condense for maximum readability, bullet points if multiple ideas |
| **Extract Tasks** | Immediate | Identify actions, follow-ups, next steps with owners and deadlines |
| **Smart Reply** | Smart-reply | Assisted drafting with contextual analysis |

### 9.6 Smart Reply Mode

#### Automatic Analysis
- Detect original email language (replies in same language)
- Determine formality level (Formal/Semi-formal/Casual)
- Analyze appropriate length based on original
- Identify key points to address
- Infer sender relationship from communication style

#### Smart Generation
- Precise tone and formality matching
- Appropriate greetings/sign-offs
- Proportional length calibration
- Asks user: "Briefly describe what you want to reply"

### 9.7 Outlook-Specific Rules

- **Reply language**: ALWAYS reply in same language as original email
- **Callbacks not promises**: Outlook uses callback pattern, must wrap in Promise
- **3-second timeout**: All Outlook operations have timeout protection
- **Compose mode only**: Write operations only work in compose mode

### 9.8 Constraints

- AI CANNOT read entire inbox or other emails than currently open one
- No calendar integration (removed from scope)
- No email attachment analysis (not planned for now)

---

## 10. User Interface & Experience

### 10.1 Theme and Appearance

#### Dark/Light Mode
- Toggle in settings with persistence
- Complete color palette for each mode
- CSS custom properties for all components
- Smooth transitions between themes

#### Color Palette

| Element | Light Mode | Dark Mode |
|---------|-----------|-----------|
| Primary text | #1d1d1f | #f5f5f7 |
| Secondary text | #6e6e73 | #a1a1a6 |
| Primary background | #ffffff | #000000 |
| Secondary background | #f5f5f7 | #1c1c1e |
| Accent | #33abc6 | #33abc6 |
| Success | #34c759 | #34c759 |
| Warning | #f1930f | #f1930f |
| Danger | #ff3b30 | #ff3b30 |

### 10.2 Toast Notifications

| Type | Color | Icon | Usage |
|------|-------|------|-------|
| Success | Green | CheckCircle | Operation succeeded |
| Error | Red | AlertCircle | Operation failed |
| Info | Blue | Info | Neutral information |
| Warning | Yellow/Amber | AlertTriangle | Warning (context limit, truncation) |

**Features:**
- Auto-close with animated progress bar
- Slide animation from right
- Configurable duration (default 3000ms)
- Fixed top-right positioning (z-index 9999)
- Backdrop blur effect

### 10.3 Localization

#### Supported Languages
- **French** (default)
- **English**

#### Translated Elements (400+ keys)
- Navigation and buttons
- Error messages
- Tool descriptions (100+ per application)
- Quick action tooltips
- Agent status messages
- All settings labels

### 10.4 Accessibility

#### ARIA Support
- `aria-live="polite"` on message list and status indicators
- `role="log"` on chat container
- `role="status"` on activity indicators
- `role="tab"` and `role="tabpanel"` on settings tabs
- `aria-label` on all icon buttons
- `aria-selected` on tabs
- SR-only announcement div for screen readers

#### Keyboard Navigation
- **Tab**: Navigate between elements
- **Enter**: Send message (in textarea)
- **Shift+Enter**: New line in message
- **Escape**: Close dropdowns
- Visible focus ring with accent color

#### Motion Preferences
- All animations disabled if `prefers-reduced-motion: reduce`

### 10.5 Responsive Design

- **Responsive grid**: 2-column layout on medium+ screens
- **Max-width controls**: Text truncation with ellipsis
- **Flex layouts**: Responsive wrapping
- **Touch-friendly**: Minimum 7px button height
- **Dropdown positioning**: Auto-adjusts based on viewport space

### 10.6 Loading States and Animations

- **Draft focus glow**: 3-iteration pulse animation on input focus
- **Spinner dots**: Animated ellipsis during generation
- **Streaming dots**: Three animated dots in thinking blocks
- **Button hover**: Slight upward translate with shadow transition
- **Thinking block**: Collapsible with chevron rotation animation

---

## 11. Context Management & Truncation

### 11.1 Context Budget

- **Total budget**: 100,000 characters
- **Model**: GPT-5.1 context window
- **Includes**: System prompts, conversation history, document content, tool messages

### 11.2 Stats Bar Visual Feedback

- **Normal**: Default color gauge
- **Warning (80%+)**: Orange gauge color
- **Critical (100%)**: Red gauge color

### 11.3 Automatic Truncation Behavior

When context limit is reached:
1. System does NOT block the action
2. Automatically truncates payload (oldest messages first for email threads, or document content)
3. Immediately displays **Warning Toast**: "Context limit reached. Only the current selection / first X pages were sent to the AI."

### 11.4 Email Thread Truncation

- **Strategy**: Relies on global 100K character budget, not fixed email count
- **Behavior**: As thread grows and pushes gauge to 100%, system automatically:
  1. Parses the thread
  2. Truncates only oldest messages to fit budget
  3. Triggers **Info/Warning Toast**: "Long thread detected: older historical messages were excluded to process your request."

### 11.5 Large Document Handling

- Same truncation strategy for Word documents, Excel workbooks, PowerPoint presentations
- Always keeps most recent/relevant content
- User notified via toast when truncation occurs

---

## 12. Undo & Modification Tracking

### 12.1 Undo Operations

- **Native Office Undo**: Rely exclusively on Ctrl+Z / Cmd+Z and Office ribbon undo button
- **No custom undo button**: Office.js groups batched script executions (run()) into single undoable actions
- **Behavior**: All AI modifications can be undone as a single action via native undo

### 12.2 Modification History by Application

| Application | Method |
|-------------|--------|
| **Word** | Native Track Changes feature - AI modifications visible in Review pane |
| **Excel** | AI inserts comments on modified cells: "Modified by AI. Old value: [X]" |
| **PowerPoint** | AI logs changes in Speaker Notes, or duplicates slides for before/after comparison |
| **Outlook** | Email history preserved via Append mode; no special tracking needed |

---

## 13. Backend API

### 13.1 Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/chat` | POST | Streaming chat completion (SSE) |
| `/api/chat/sync` | POST | Synchronous chat completion |
| `/api/image` | POST | Image generation |
| `/api/upload` | POST | File upload and processing |
| `/api/models` | GET | Get available models |
| `/health` | GET | Health check with timestamp and version |

### 13.2 Rate Limiting

| Endpoint | Window | Default Max |
|----------|--------|-------------|
| `/api/chat` | 60s | 20 requests |
| `/api/image` | 60s | 5 requests |
| `/api/upload` | 60s | 10 requests |
| `/health`, `/models` | 60s | 120 requests |

### 13.3 Timeouts

| Operation | Duration |
|-----------|----------|
| Standard chat models | 2 minutes |
| Reasoning models | 5 minutes |
| Image generation | 3 minutes |
| Per-read operation | 30 seconds |
| Outlook API calls | 3 seconds |
| Overall request timeout | 10 minutes |

### 13.4 File Processing

**Supported formats with extraction:**
- PDF → text extraction via pdf-parse
- DOCX → text extraction via mammoth
- XLSX/XLS/CSV → all sheets to CSV format via xlsx
- TXT/MD → direct UTF-8 decoding
- Images (PNG, JPG, WEBP, GIF) → base64 encoding with data-URI

**Limits:**
- 10 MB per file
- 100K characters after extraction (truncated with notification)
- MIME type detection (not just declared type)

### 13.5 Streaming (SSE)

- Server-Sent Events format with delta updates
- Client disconnect detection with upstream cancellation
- Backpressure handling (waits for drain event)
- 30-second per-read timeout
- Token usage in final chunk
- Tool call deltas for agentic use

### 13.6 Request Retry Logic

- **Network errors and timeouts only**: Automatic retry
- **POST requests**: Max 1 retry
- **GET requests**: Up to 2 retries
- **Delays**: 1.5s, then 4s
- **Respects AbortSignal**: No retry after user cancellation

---

## 14. Model Configuration

### 14.1 Model Tiers

| Tier | Default Model | Max Tokens | Temperature | Special |
|------|--------------|------------|-------------|---------|
| Standard | gpt-5.1 | 4,096 | 0.7 | General purpose |
| Reasoning | gpt-5.1 | 8,192 | 1.0 | reasoning_effort parameter |
| Image | gpt-image-1 | N/A | N/A | Image generation |

### 14.2 Model Detection

- **GPT-5.x**: Uses `max_completion_tokens` instead of `max_tokens`, includes `reasoning_effort`
- **ChatGPT models**: Different parameter handling for legacy compatibility

### 14.3 Tool Constraints

- **Maximum tools per request**: 128 (configurable via MAX_TOOLS)
- **Tool choice**: 'auto' - model decides when to call tools
- **Function name validation**: Regex `/^[a-zA-Z0-9_-]{1,64}$/`
- **Strict schema**: Optional boolean flag for strict JSON schema validation

---

## 15. Platform Support

### 15.1 Supported Platforms

| Platform | Supported |
|----------|-----------|
| Office 365 Desktop (Windows) | Yes |
| Office 365 Online (Web) | Yes |
| Office Mobile (iOS/Android) | No |
| Older Office versions | No |

### 15.2 Out of Scope

| Feature | Status |
|---------|--------|
| VBA/Macros interaction | Not planned for now |
| Password-protected documents | Not planned for now |
| Template creation/modification | No |
| Co-authoring support | No |
| Third-party add-in compatibility | Not managed |
| OneDrive/SharePoint special handling | Treated as regular documents |
| New LLM models | No additions planned |
| New Office apps (Teams, OneNote, Visio) | No |
| Extended agent mode (web search, databases) | Not planned for now |
| Credential expiration/refresh | No |
| Audit trail | No |
| Sensitive data masking | No |
| Session sharing | No |
| Team-shared prompts | No |
| Email attachment analysis | Not planned for now |
| Calendar integration | Removed from scope |

---

## 16. Tool Summary

### 16.1 Tool Count by Application

| Application | Read | Write | Format | Eval | Total |
|-------------|------|-------|--------|------|-------|
| Word | 8 | 13 | 5 | 1 | 27 |
| Excel | 7 | 10 | 1 | 1 | 18 (with chart tool) |
| PowerPoint | 4 | 5 | 0 | 1 | 9 |
| Outlook | 4 | 4 | 0 | 1 | 8 |
| General | 3 | 3 | 0 | 0 | 6 |
| **Total** | **26** | **35** | **6** | **4** | **67** |

### 16.2 General Tools (All Apps)

| Tool | Category | Description |
|------|----------|-------------|
| getCurrentDate | read | Get current date/time in various formats |
| calculateMath | write | Evaluate mathematical expressions safely |
| executeBash | write | Execute bash commands in sandboxed VFS |
| vfsWriteFile | write | Write files to virtual filesystem |
| vfsReadFile | read | Read files from virtual filesystem |
| vfsListFiles | read | List files in VFS uploads directory |

### 16.3 Tool Decision Trees

**Word Content Modification:**
```
Simple word/phrase replacement → searchAndReplace
Rewriting sentences/paragraphs → proposeRevision (preserves formatting)
Adding new content → insertContent
```

**Excel Data Operations:**
```
Write data → setCellRange (ALWAYS preferred)
Format data → formatRange
Create table → createTable
Charts/Pivots → manageObject
Advanced → eval_officejs
```

**PowerPoint Workflow:**
```
1. getAllSlidesOverview (understand structure)
2. getShapes(slideNumber) (discover shape IDs)
3. insertContent or proposeShapeTextRevision (modify)
```

**Outlook Email:**
```
Reply/Forward → ALWAYS mode "Append"
New email → Can use mode "Replace"
```

---

## 17. Logging & Monitoring

### 17.1 Server-Side Logging

- **Morgan HTTP logger**: Method, URL, status, content-length, response time
- **Custom system logger**: INFO, ERROR, WARN, DEBUG levels
- **Rotating file stream**: 10MB per file, 30 max files, gzip compression, daily rotation
- **Log location**: `/logs/kickoffice.log`

### 17.2 Request Tracking

- Unique UUID per request in `res.locals.reqId`
- Verbose logging mode (VERBOSE_LOGGING=true) for debugging

### 17.3 Sensitive Data Protection

- API keys redacted from logs
- Sensitive headers (x-user-key, x-user-email, authorization, api-key) never logged

---

## 18. Graceful Shutdown

- **Signal handling**: SIGTERM and SIGINT
- **Process**: Stop accepting connections → wait for in-flight requests → close all connections
- **Force exit**: After 30 seconds if still hanging
- **Logging**: Shutdown signals and completion status logged

---

## 19. Open Questions

### Remaining Implementation Details

1. **Comment format for Excel**: Should the "Modified by AI" comment include timestamp or user info?

2. **PowerPoint Speaker Notes format**: What exact format should AI use when logging modifications to notes?

3. **Truncation notification wording**: Final wording for context truncation toast messages in both FR and EN?

4. **Stats bar gauge thresholds**: Confirm 80% for orange warning, 100% for red critical?

5. **Email thread detection regex**: Exact patterns used to detect reply/forward headers across different email clients?
