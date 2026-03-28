# Product Requirements Document (PRD): KickOffice

> For technical implementation details (API spec, tool parameters, model config, logging), see [TECHNICAL.md](TECHNICAL.md).

---

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
- **Telemetry & DLP:** The add-in itself does not log telemetry or block sensitive data. All AI telemetry, auditing, and DLP filtering are strictly delegated to the internal LLM gateway.

---

## 3. Cross-Application Features (All Apps)

### 3.1 Chat Interface

#### Visual Components

- **Chat header**: KickOffice logo, subtitle, settings button
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
- **Session Persistence**: Uploaded files (extracted text) and images stay available across the entire conversation. File context is fully restored when switching to a previous session or reloading the add-in.
- **File Attachment Badges**: Attached document names are displayed inline in the user's message bubble after sending, giving visual confirmation of what was uploaded.
- **Large File Support**: Extended LLM processing timeout (5 minutes for standard requests) to handle large documents without timeout errors. Files can optionally be forwarded to the LLM provider's storage to avoid re-transmitting content on every message.
- **Log Sanitization**: Automated protection against log file saturation by truncating large data payloads before they reach server logs.
- **Duplicate detection**: Prevents adding same file twice by name
- **Error messages**: Specific notifications for oversized files, unsupported formats, or max files reached

### 3.2 Quick Actions

Contextual buttons at the top of the interface for one-click actions:

- **Tooltips**: Each button shows a description on hover
- **Disabled state**: Buttons grayed out during loading
- **User Skills dropdown**: One-click access to user-created skills; executes immediately on selected text
- **Create skill button (`+`)**: Opens the Skill Creator modal
- **Host-specific actions**: Different built-in actions per Office app

### 3.3 Message Display

#### User Messages

- Plain text display with timestamp
- Attached document names displayed as pill badges below the message content (PDF, DOCX, XLSX...)

#### Assistant Messages

- **Markdown rendering**: Full support (bold, italic, lists, code blocks, tables)
- **Thinking blocks**: Collapsible sections showing AI reasoning process with brain icon
- **Tool calls**: Expandable blocks showing tool name, arguments, status, and results
- **Generated images**: Inline display of AI-created images
- **Streaming animation**: Three animated dots during response generation
- **Content sanitization**: XSS prevention on all rendered content

#### Tool Call Visualization

- **Running status**: Animated spinner
- **Success status**: Green checkmark
- **Error status**: Red X with error message
- **Arguments display**: JSON-formatted with syntax highlighting
- **Expandable/collapsible**: Click to show/hide details

#### Message Actions

- **Replace selection**: Insert message content replacing current selection
- **Append to selection**: Add content at the end of selection
- **Copy to clipboard**: Copy text with fallback for Office WebView environments

### 3.4 Image Generation

- **Dedicated model tier**: Uses specific image model (gpt-image-1)
- **Default size**: 1792x1024 (landscape, suited to slides and documents)
- **Size options**: 1024x1024, 1024x1536, 1536x1024, 1792x1024
- **Quality levels**: low, medium, high, auto
- **Multiple images**: Can generate 1-4 images per request
- **Prompt limit**: Maximum 4000 characters
- **Framing**: A framing instruction is automatically prepended to every prompt to prevent subjects from being cropped at image edges
- **Direct insertion**: Images inserted directly into document
- **Clipboard fallback**: If insertion fails, image copied to clipboard with notification
- **Extended timeout**: 3-minute timeout for generation

### 3.5 Autonomous Agent

- **Agentic loop**: AI can execute multiple tools in sequence for complex tasks
- **Iteration limit**: Configurable from 1 to 100 (default: 25)
- **Loop detection**: Sliding window signature comparison prevents infinite repetitive calls
- **User stop**: Button to stop agent at any time
- **Status indicator**: Real-time display of current action. Long action labels wrap over up to 2 lines instead of being truncated.
- **Tool selection**: Agent autonomously chooses appropriate tools for each task

### 3.6 Stats Bar Details

- **Input tokens**: Tokens sent to model, auto-formatted (K for thousands, M for millions)
- **Output tokens**: Tokens received from model
- **Context percentage**: Visual gauge of context window usage
- **Context gauge colors**:
  - Normal: Default color
  - Warning (70%+): Orange
  - Critical (90%+): Red
- **Model name**: Currently selected model identifier
- **Activity pulse**: Red animated dot with Terminal icon when agent is executing
- **Hover tooltips**: Detailed token counts on hover

### 3.8 Skill Creator

LLM-assisted modal for creating user skills without writing code:

- **Step 1 -- Describe**: Free-text description + host selector (Ctrl+Enter to submit)
- **Step 2 -- Generate**: Non-streaming LLM call produces a draft skill; the system prompt embeds full tool knowledge per host, execution mode guidance, and Theory of Mind writing principles
- **Step 3 -- Review**: Editable fields (name, description, host, execution mode, icon) + markdown textarea for skill content; Regenerate button returns to Step 1
- **Step 4 -- Test**: Executes the draft skill on currently selected text in the Office document; result appears in chat; user can return to review
- **Save**: Adds skill to localStorage registry; appears immediately in Quick Actions Bar dropdown
- **Rate limit**: 10 generations/hour/IP server-side

### 3.9 Virtual File System (VFS)

- **Sandboxed filesystem**: In-memory virtual filesystem for agent use
- **User uploads**: Files stored in a dedicated uploads directory
- **Custom scripts**: Reusable bash scripts in a user scripts directory
- **Session persistence**: VFS state saved and restored with each session
- **Available commands**: Standard Unix commands (ls, cat, grep, find, awk, sed, sort, uniq, wc, cut, head, tail, base64)

### 3.10 Document Context Injection

Uploaded files persist as a session-level knowledge base accessible to the agent across all subsequent messages.

- **Persistent file context**: Files uploaded in earlier messages remain available for the entire session — the user can reference them at any point (e.g., "based on the Q4 report I uploaded earlier")
- **VFS as shared memory**: Files written to the VFS by the agent (scripts, exports, intermediate data) are also available as context for later tasks within the same session
- **Cross-message references**: The agent can correlate information across multiple uploaded documents in a single conversation (e.g., "compare the figures in budget.xlsx with the narrative in report.docx")

### 3.11 AI Audit Trail

Every AI tool execution is logged for traceability and compliance (EU AI Act, August 2026).

- **Action logging**: Each tool call records the tool name, parameters, timestamp, user ID, and host application
- **Before/after state**: For write operations, the previous value is captured alongside the new value (e.g., cell content before/after, replaced text)
- **Reasoning context**: The agent's reasoning or explanation for each action is preserved
- **Export**: Audit logs can be exported for compliance review
- **Retention**: Logs are retained server-side with configurable retention period
- **Non-intrusive**: Logging does not affect user experience or add latency to operations

### 3.12 Shared Skill Library

Team-wide skills managed server-side, complementing the local user skills.

- **Server-hosted skills**: An administrator can publish `.skill.md` files to a shared directory on the backend
- **Automatic merge**: The frontend merges local user skills with shared team skills in the Quick Actions dropdown — shared skills are visually distinguished (e.g., "Team" badge)
- **Read-only for users**: Shared skills cannot be edited or deleted by end users; only admins manage them
- **Override**: If a user creates a local skill with the same name, the local version takes precedence
- **No sync complexity**: Shared skills are fetched on session init via a lightweight API call; no real-time sync needed

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

Every request includes user credentials and CSRF protection, validated server-side.

### 4.3 Authentication Error Handling

| Error Type          | Displayed Message                                         |
| ------------------- | --------------------------------------------------------- |
| Missing credentials | "Please configure your credentials in Settings > Account" |
| 401 error           | "Authentication required"                                 |
| Timeout             | "Request timed out. The model took too long..."           |
| Network error       | "Connection error. Check your network..."                 |
| Rate limit          | "Too many requests -- rate limit reached. Please wait..." |
| Server error        | "Internal server error..."                                |

For security implementation details (CORS, Helmet, headers), see [TECHNICAL.md](TECHNICAL.md).

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

### 5.3 User Skills Library Tab

Replaces the former Custom Prompts and Built-in Prompts tabs.

- List of all user-created skills with name, host badge, execution mode badge, and description preview
- Filter chips by host (Word / Excel / PowerPoint / Outlook / All)
- Inline edit form per skill: name, description, host, execution mode, icon (Lucide name), markdown content
- Export individual skill as `.skill.md` file (with YAML frontmatter)
- Import skill from `.skill.md` file (generates new id to avoid collisions)
- Create skill via LLM-assisted modal (see 3.8)
- Automatic migration banner on first launch if legacy custom prompts exist in localStorage

### 5.4 Tools Tab

- Complete tool list with checkboxes (101 tools total)
- Enable/disable per tool
- Description for each tool
- Preference persistence per Office host
- Auto-migration when new tools added

---

## 6. Microsoft Word

### 6.1 Document Reading (8 tools)

getSelectedText, getDocumentContent, getDocumentHtml, getDocumentProperties, getSelectedTextWithFormatting, getSpecificParagraph, findText, getComments

### 6.2 Content Insertion (8 tools)

insertContent *(preferred)*, searchAndReplace *(preferred for corrections)*, proposeRevision *(preferred for editing)*, insertHyperlink, insertFootnote, insertHeaderFooter, insertSectionBreak, addComment

### 6.3 Text Formatting (4 tools)

formatText, applyStyle, setParagraphFormat, applyTaggedFormatting

### 6.4 Table Management (5 tools)

modifyTableCell, addTableRow, addTableColumn, deleteTableRowColumn, formatTableCell

### 6.5 Document Structure (1 tool)

setPageSetup

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

| Action    | Description                                        |
| --------- | -------------------------------------------------- |
| Proofread | Grammar and spelling correction with Track Changes |
| Translate | FR-EN translation of selection                     |
| Polish    | Style and clarity improvement                      |
| Summarize | Concise summary generation                         |

### 6.9 Custom Code Execution (eval_wordjs)

- **Last resort**: Only when no dedicated tool exists
- **Code validation**: Validated before execution
- **Sandboxed**: Word namespace only

---

## 7. Microsoft Excel

### 7.1 Data Reading (10 tools)

getSelectedCells, getWorksheetData, getWorksheetInfo, getNamedRanges, getConditionalFormattingRules, findData, getAllObjects, getRangeAsCsv, screenshotRange, detectDataHeaders

### 7.2 Writing and Editing (15 tools)

setCellRange *(preferred)*, clearRange, modifyStructure, modifyWorkbookStructure, addWorksheet, createTable, sortRange, searchAndReplace, setNamedRange, protectWorksheet, importCsvToSheet, imageToSheet, extract_chart_data, clearAgentHighlights, eval_officejs

### 7.3 Formatting (2 tools)

formatRange, applyConditionalFormatting

### 7.4 Conditional Formatting Types

| Type           | Description                                                 |
| -------------- | ----------------------------------------------------------- |
| Cell value     | Comparison (equal, not equal, greater, less, between...)    |
| Text match     | Contains, starts with, ends with                            |
| Custom formula | Formatting based on formula                                 |
| Color scale    | Color gradient min to max                                   |
| Data bars      | Proportional visual bars                                    |
| Icon sets      | Icons based on thresholds (traffic lights, arrows, symbols) |

### 7.5 Chart Management

**Supported Chart Types:**

- Column (Clustered, Stacked)
- Line (simple, with markers)
- Pie
- Bar (Clustered)
- Area
- Doughnut
- XY Scatter

**Features:**

- Create from data range with anchor positioning
- Set title and dimensions
- Update type, source, or title
- Delete existing charts

**Header auto-detection**: Before creating a chart, the agent calls detectDataHeaders on the source range to determine whether the first row/column contains labels, ensuring axis labels are not plotted as data series.

### 7.6 Pivot Table Management

- Create from data range
- Custom naming
- Controlled positioning
- Delete existing pivot tables

### 7.7 Large Workbook Handling

- **Paginated search**: findData supports pagination. On large workbooks, the agent can retrieve results page by page, avoiding context window overflow.
- **CSV export**: getRangeAsCsv is preferred over JSON for data analysis tasks on large ranges (significant token savings).
- **Visual verification**: screenshotRange allows the agent to visually inspect formatting and chart rendering without reading raw cell data.

### 7.8 Modification History (Excel-specific)

- **Comment-based tracking**: When modifying a cell, AI MUST insert a comment containing previous value
- **Comment format**: "Modified by AI. Old value: [X]"
- **Native undo**: Ctrl+Z works for AI modifications

### 7.9 Excel Quick Actions

| Action            | Mode        | Description                                                                                                          |
| ----------------- | ----------- | -------------------------------------------------------------------------------------------------------------------- |
| Smart Ingestion   | Immediate   | Intelligently converts raw pasted data (CSV) into a proper Excel table, silently fixing locale issues (dots vs commas). |
| Auto-Graph        | Immediate   | Visual analysis partner. Suggests charts to highlight data, generates new columns if visually identified, and creates the chart with legends and title. |
| Explain Formula   | Immediate   | Explains in natural language the logic behind any cell containing a complex formula.                                   |
| Formula Generator | Interactive | Guides the user in writing a formula prompt with an adapted system prompt instead of a blocking step-by-step system.  |
| Data Trend        | Immediate   | Analyzes data to identify trends and suggests actions to better highlight certain patterns.                            |
| Dashboard         | Agent       | Conversational dashboard builder: analyzes the data, proposes relevant KPIs and visualizations, then creates charts, formatting, and layout automatically. |

### 7.10 Formula Localization

- **English**: Comma separators (e.g., =SUM(A1,B1))
- **French**: Semicolon separators (e.g., =SOMME(A1;B1))
- User preference in settings
- Auto-conversion based on Excel locale

### 7.11 Custom Function: =KICKOFFICE() *(planned)*

An in-cell AI function that brings LLM capabilities directly into Excel formulas, similar to =COPILOT() or =CLAUDE().

- **Syntax**: `=KICKOFFICE(prompt, [cell_reference], ...)` — natural language prompt with optional cell context
- **Drag & fill**: Works like any native formula — drag across rows/columns to apply the same prompt to different data
- **Use cases**: Text classification, entity extraction, summarization, translation, data enrichment, sample data generation
- **Composable**: Can be combined with standard Excel functions (IF, SWITCH, CONCATENATE)
- **Caching**: Results are cached to avoid redundant API calls when the workbook recalculates
- **Rate limiting**: Batch execution with configurable concurrency to respect API rate limits
- **Auto-update**: Results refresh when referenced cells change (like any formula dependency)

### 7.12 Cross-Host Pipeline *(roadmap)*

Future capability to pass data between Office hosts via the VFS (e.g., export Excel data → generate PowerPoint charts, or extract Word content → draft Outlook email). Currently each host operates in isolation; the VFS and Document Context Injection (§3.10) lay the groundwork for cross-host workflows.

---

## 8. Microsoft PowerPoint

### 8.1 Presentation Reading (9 tools)

getSelectedText, getSlideContent, getShapes, getAllSlidesOverview, getCurrentSlideIndex, getSpeakerNotes, screenshotSlide, verifySlides, searchIcons

### 8.2 Content Insertion and Modification (15 tools)

insertContent *(preferred)*, proposeShapeTextRevision, replaceShapeParagraphs, searchAndReplaceInShape, searchAndFormatInPresentation, replaceSelectedText, setSpeakerNotes, addSlide, deleteSlide, duplicateSlide, reorderSlide, insertIcon, insertImageOnSlide, eval_powerpointjs, editSlideXml

### 8.3 Visual Verification Workflow

After making visual changes (layout, images, icons, positioning), the agent should:
1. Call screenshotSlide to capture the modified slide
2. Analyze the screenshot image to verify the result visually
3. Call verifySlides to programmatically check for overflows and overlaps
4. Iterate if issues are detected

### 8.4 Icon Library Integration

The agent can insert professional icons from the Iconify library (200,000+ icons across 150+ open-source icon sets including Material Design, Fluent UI, Feather, Bootstrap, Heroicons):
1. Search by keyword (and optionally filter by icon set) to discover available icons
2. Insert the chosen icon on a slide with custom position, size, and color

### 8.5 OOXML Direct Editing

editSlideXml provides an escape hatch for operations that the Office.js API cannot express (e.g., chart manipulation, SmartArt, complex animations). This is an advanced tool intended for cases where no dedicated tool covers the need.

### 8.6 Modification History (PowerPoint-specific)

PowerPoint has NO native Track Changes. Two workarounds:

1. **Speaker Notes logging**: AI logs modifications in the Speaker Notes section at the bottom of the slide
2. **Slide duplication**: For major structural changes, duplicate the original slide to allow before/after comparison

### 8.7 Brand-Aware Generation

When generating or modifying slide content, the agent reads the active presentation's slide master to extract brand constraints (color scheme, fonts, available layouts). These constraints are injected into the system prompt so that all generated content respects the organization's visual identity without requiring manual brand guideline uploads.

- **Automatic extraction**: Colors, font families, and layout names from the active slide master
- **Layout selection**: addSlide picks the most appropriate layout from the master (already implemented)
- **Consistency enforcement**: Generated text, charts, and shapes use brand-approved styles
- **No manual setup**: Works out of the box with any branded template

### 8.8 Presentation from Document *(planned)*

Generate a complete PowerPoint presentation from an uploaded Word document, PDF, or structured text.

- **Source formats**: Word (.docx), PDF, plain text, or markdown
- **Auto-structuring**: The agent analyzes the source document's structure (headings, sections, key points) and maps it to slides
- **Full deck generation**: Title slide, content slides with appropriate layouts, speaker notes, and closing slide
- **Image suggestions**: The agent suggests or generates relevant visuals for key slides
- **Iterative refinement**: After initial generation, the user can refine individual slides via chat

### 8.9 PowerPoint Quick Actions

| Action         | Mode         | Description                                                                                 |
| -------------- | ------------ | ------------------------------------------------------------------------------------------- |
| Proofread      | Immediate    | Grammar and spelling correction                                                             |
| Translate      | Immediate    | FR-EN translation                                                                           |
| Review         | Immediate    | Visual feedback on active slide via screenshot analysis                                     |
| Punchify       | Agent (auto) | Reads the active slide via agent tools, rewrites all text shapes to be punchy and concise   |
| Visual         | Immediate    | Generate an AI illustration that visually represents the slide content (two-step: LLM creates image description, then image model renders) |
| From Document  | Agent        | Generate a full presentation from an uploaded Word/PDF document (see §8.8)                  |

### 8.10 Slide Style Rules

- Maximum 8-10 words per bullet point
- Maximum 6-7 bullets per slide
- Active voice, present tense preferred

### 8.11 Constraints

- **Slide Master modification**: Limited via standard API; advanced modifications are possible via editSlideXml (OOXML) for expert use cases

---

## 9. Microsoft Outlook

### 9.1 Email Reading (4 tools)

getEmailBody, getEmailSubject, getEmailRecipients, getEmailSender

### 9.2 Email Writing (5 tools)

writeEmailBody *(preferred)*, setEmailSubject, addRecipient, addAttachment, eval_outlookjs

### 9.3 Email Body Writing Modes

| Mode                 | Description                          | Use Case                            |
| -------------------- | ------------------------------------ | ----------------------------------- |
| **Append (DEFAULT)** | Adds at end, preserves history       | Replies, forwards - ALWAYS use this |
| **Insert**           | Inserts at cursor with optional diff | Specific text replacement in draft  |
| **Replace**          | Replaces entire body                 | Brand new emails ONLY               |

### 9.4 Email History Protection (CRITICAL)

- **Automatic protection**: Prevents accidental deletion of email threads
- **Auto-detection**: Recognizes replies/forwards by standard email headers and quoted text
- **Safe default mode**: ALWAYS uses Append mode by default
- **Visual alerts**: Warning if Replace mode might erase history
- **Agent instructions**: Explicit instructions to NEVER delete email history
- **Image preservation**: Embedded images are preserved via placeholders

### 9.5 Outlook Quick Actions

| Action          | Mode        | Description                                              |
| --------------- | ----------- | -------------------------------------------------------- |
| Proofread       | Immediate   | Grammar and spelling correction                          |
| Translate       | Immediate   | FR-EN translation                                        |
| Extract Tasks   | Immediate   | Identify actions, follow-ups, next steps                 |
| Smart Reply     | Smart-reply | Assisted drafting with contextual analysis               |
| Meeting Minutes | Immediate   | Generate structured meeting minutes from email thread    |
| Coach           | Immediate   | Analyze draft email for tone, clarity, and communication strategy — suggests improvements beyond grammar (e.g., "too blunt", "missing call to action", "consider a softer opening") |

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

### 9.7 Constraints

- AI CANNOT read entire inbox or other emails than currently open one
- No calendar integration (removed from scope)
- No email attachment analysis (not planned for now)

---

## 10. User Interface & Experience

### 10.1 Theme and Appearance

- Toggle between dark and light mode in settings with persistence
- Complete color palette for each mode with smooth transitions
- CSS custom properties for consistent theming across all components

For theme color specifications, see [TECHNICAL.md](TECHNICAL.md).

### 10.2 Toast Notifications

| Type    | Color        | Icon          | Usage                               |
| ------- | ------------ | ------------- | ----------------------------------- |
| Success | Green        | CheckCircle   | Operation succeeded                 |
| Error   | Red          | AlertCircle   | Operation failed                    |
| Info    | Blue         | Info          | Neutral information                 |
| Warning | Yellow/Amber | AlertTriangle | Warning (context limit, truncation) |

**Features:**

- Auto-close with animated progress bar
- Slide animation from right
- Configurable duration (default 3000ms)
- Fixed top-right positioning
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

KickOffice follows WCAG guidelines with full ARIA landmark and live-region support. All interactive elements are keyboard-navigable (Tab, Enter, Shift+Enter, Escape), with visible focus indicators. Animations respect the user's prefers-reduced-motion setting.

For detailed ARIA attributes and responsive design specifications, see [TECHNICAL.md](TECHNICAL.md).

---

## 11. Context Management & Truncation

### 11.1 Context Budget

- **Total budget**: Adapts to the active model's context window (currently ~1.2M characters for GPT-5.2)
- **Includes**: System prompts, conversation history, document content, tool messages

### 11.2 Stats Bar Visual Feedback

- **Normal**: Default color gauge
- **Warning (70%+)**: Orange gauge color
- **Critical (90%+)**: Red gauge color

### 11.3 Automatic Truncation Behavior

When context limit is reached:

1. System does NOT block the action
2. Automatically truncates payload (oldest messages first for email threads, or document content)
3. Immediately displays **Warning Toast**: "Context limit reached. Only the current selection / first X pages were sent to the AI."

### 11.4 Email Thread Truncation

- **Strategy**: Relies on global character budget, not fixed email count
- **Behavior**: As thread grows and pushes gauge to critical, system automatically:
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
- **No custom undo button**: Office.js groups batched script executions into single undoable actions
- **Behavior**: All AI modifications can be undone as a single action via native undo

### 12.2 Modification History by Application

| Application    | Method                                                                             |
| -------------- | ---------------------------------------------------------------------------------- |
| **Word**       | Native Track Changes feature - AI modifications visible in Review pane             |
| **Excel**      | AI inserts comments on modified cells: "Modified by AI. Old value: [X]"            |
| **PowerPoint** | AI logs changes in Speaker Notes, or duplicates slides for before/after comparison |
| **Outlook**    | Email history preserved via Append mode; no special tracking needed                |

---

## 13. Platform Support

### 13.1 Supported Platforms

| Platform                     | Supported |
| ---------------------------- | --------- |
| Office 365 Desktop (Windows) | Yes       |
| Office 365 Online (Web)      | Yes       |
| Office Mobile (iOS/Android)  | No        |
| Older Office versions        | No        |

### 13.2 Out of Scope

| Feature                                     | Status                       |
| ------------------------------------------- | ---------------------------- |
| VBA/Macros interaction                      | Not planned for now          |
| Password-protected documents                | Not planned for now          |
| Template creation/modification              | No                           |
| Co-authoring support                        | No                           |
| Third-party add-in compatibility            | Not managed                  |
| OneDrive/SharePoint special handling        | Treated as regular documents |
| New LLM models                              | No additions planned         |
| New Office apps (Teams, OneNote, Visio)     | No                           |
| Extended agent mode (web search, databases) | Web search/fetch deferred -- planned for future release |
| Credential expiration/refresh               | No                           |
| Audit trail                                 | Planned -- see §3.11 AI Audit Trail |
| Sensitive data masking                      | No                           |
| Session sharing                             | No                           |
| Team-shared prompts (server-side)           | Planned -- see §3.12 Shared Skill Library |
| Email attachment analysis                   | Not planned for now          |
| Calendar integration                        | Removed from scope           |

---

## 14. Tool Summary

### 14.1 Tool Count by Application

| Application    | Total   | Notable capabilities                                                        |
| -------------- | ------- | --------------------------------------------------------------------------- |
| **Word**       | 34      | Format-preserving edits, Track Changes, tables, comments, diff tracking     |
| **Excel**      | 28      | Charts, formulas, screenshots, CSV export, workbook structure management, header detection, pixel art, chart digitizer |
| **PowerPoint** | 24      | Slides, shapes, screenshots, layout verification, icon library, OOXML edit, speaker notes, slide reorder, brand-aware generation |
| **Outlook**    | 9       | Email compose/read, smart reply, email history protection, attachments      |
| **General**    | 6       | VFS, math, date, file operations                                            |
| **Total**      | **101** |                                                                             |

### 14.2 General Tools (All Apps)

| Tool           | Description                              |
| -------------- | ---------------------------------------- |
| getCurrentDate | Get current date/time in various formats |
| calculateMath  | Evaluate mathematical expressions safely |
| executeBash    | Execute bash commands in sandboxed VFS   |
| vfsWriteFile   | Write files to virtual filesystem        |
| vfsReadFile    | Read files from virtual filesystem       |
| vfsListFiles   | List files in VFS uploads directory      |

For tool decision trees and parameter specifications, see [TECHNICAL.md](TECHNICAL.md).

---

## 15. Open Questions

### Remaining Implementation Details

1. **Comment format for Excel**: Should the "Modified by AI" comment include timestamp or user info?

2. **PowerPoint Speaker Notes format**: What exact format should AI use when logging modifications to notes?

3. **Truncation notification wording**: Final wording for context truncation toast messages in both FR and EN?

4. **Stats bar gauge thresholds**: Confirm 70% for orange warning, 90% for red critical?
