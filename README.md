# KickOffice

AI-powered add-in for Microsoft Office applications. Provides a chat interface, document manipulation agent, and quick AI actions (translate, polish, summarize, email reply, presentation helpers, etc.) directly inside Word, Excel, PowerPoint, and Outlook.

Built for **professional environments**: all LLM traffic goes through a controlled backend server (no API keys on the client), and no data is sent to third-party services.

Based on the [WordGPT Plus](https://github.com/Kuingsmile/word-GPT-Plus) open-source project, heavily modified for enterprise use.

Also based on [excel-ai-assistant](https://github.com/ilberpy/excel-ai-assistant) (MIT License), with additional adaptations for KickOffice.

---

## Architecture

```
┌──────────────────────┐     ┌──────────────────────┐     ┌──────────────────┐
│  Office Add-in       │     │  KickOffice Backend   │     │  LLM API         │
│  (Vue 3 + Vite)      │────>│  (Express.js)         │────>│  (OpenAI /       │
│  Port 3002           │     │  Port 3003            │     │   LiteLLM)       │
│                      │<────│                       │<────│                  │
└──────────────────────┘     └──────────────────────┘     └──────────────────┘
        │                              │
        │  Office.js API               │  Health check
        ▼                              │  Model config
   Word / Excel / PowerPoint /            │  API key storage
   PowerPoint                          ▼
                                  .env file
```

- **Frontend**: Vue 3 task pane add-in loaded inside Office apps. Handles UI, chat, agent tool execution (Word API calls run locally in the browser).
- **Backend**: Express.js proxy server. Holds all secrets (API keys), exposes `/api/chat` (streaming), `/api/chat/sync` (agent mode with tools), `/api/image`, `/api/models`, and `/health`.
- **LLM API**: Any OpenAI-compatible endpoint. For testing: OpenAI API directly. For production: Azure-hosted LiteLLM proxy with dedicated endpoints.

---

## Model Tiers

Models are configured **server-side only** (in `backend/.env`). Users cannot add or modify models. Four tiers:

| Tier | Purpose | Default Model | Use Case |
|------|---------|---------------|----------|
| `nano` | Basic tasks | `gpt-4.1-nano` | Quick answers, simple formatting |
| `standard` | Normal tasks | `gpt-4.1` | Chat, writing, analysis |
| `reasoning` | Complex tasks | `o3` | Multi-step reasoning, planning |
| `image` | Image generation | `gpt-image-1` | Generate images |

---

## Deployment (Docker on Synology NAS)

### Prerequisites
- Synology NAS with Container Manager (Docker)
- IP: `192.168.50.10` (configurable)
- OpenAI API key (for testing)

### Steps

1. **Clone the repository** on the NAS or copy the project to `/volume1/docker/kickoffice/`

2. **Create environment files**:
   ```bash
   # Root .env – server IP and ports (used by docker-compose + manifest generation)
   cp .env.example .env
   # Edit .env if your NAS IP or ports differ from the defaults

   # Backend .env – LLM API key and model config
   cp backend/.env.example backend/.env
   # Edit backend/.env and set your LLM_API_KEY
   ```

3. **Build and start**:
   ```bash
   docker compose up -d --build
   ```
   > The `manifest-gen` init service (Node.js) automatically generates two
   > manifests from the templates in `manifests-templates/`:
   > - `manifest-office.xml` — Word, Excel, PowerPoint (TaskPaneApp)
   > - `manifest-outlook.xml` — Outlook (MailApp)
   >
   > URLs are built from `SERVER_IP`, `FRONTEND_PORT`, and `BACKEND_PORT`
   > defined in the root `.env`. No manual URL editing is required.
   >
   > Optional cleanup (once generation is done):
   > ```bash
   > docker compose rm -f manifest-gen
   > ```
   > This removes the stopped init container from status views; it will be recreated automatically on the next `docker compose up`.

4. **Verify**:
   - Backend health: `curl http://192.168.50.10:3003/health`
   - Frontend: open `http://192.168.50.10:3002` in a browser
   - Models endpoint: `curl http://192.168.50.10:3003/api/models`

5. **Install the Office add-ins**:
   - Sideload `manifest-office.xml` in Word / Excel / PowerPoint via Insert > My Add-ins > Upload My Add-in
   - Sideload `manifest-outlook.xml` in Outlook via the same dialog
   - Or configure a shared catalog in Trust Center Settings

### Docker Compose Details

| Container | Port | Image | Notes |
|-----------|------|-------|-------|
| `kickoffice-manifest-gen` | — | Node 18 Alpine (init) | Generates manifests, then exits (can be removed with `docker compose rm -f manifest-gen`) |
| `kickoffice-backend` | 3003 | Node 22 Alpine | |
| `kickoffice-frontend` | 3002 | Nginx Alpine (serving built Vue app) | |

Both containers use `PUID=1026` / `PGID=100` for Synology compatibility.

The backend container has a built-in health check (`/health` endpoint, every 30s).

---

## Project Structure

```
KickOffice/
├── backend/
│   ├── Dockerfile
│   ├── .env.example          # Model config + API keys (copy to .env)
│   ├── package.json
│   └── src/
│       └── server.js         # Express server: proxy, health, models
├── frontend/
│   ├── Dockerfile            # Multi-stage: build + nginx
│   ├── nginx.conf
│   ├── package.json
│   ├── vite.config.js
│   ├── index.html
│   └── src/
│       ├── main.ts           # Office.js init + Vue app mount
│       ├── App.vue
│       ├── api/
│       │   ├── backend.ts    # HTTP client for backend API
│       │   └── common.ts     # Word document insertion helpers
│       ├── components/       # Reusable UI components
│       ├── i18n/
│       │   └── locales/
│       │       ├── en.json
│       │       └── fr.json
│       ├── pages/
│       │   ├── HomePage.vue      # Main chat + agent + image + quick actions (Word/Excel/PowerPoint/Outlook)
│       │   └── SettingsPage.vue  # Settings (language, prompts, tools)
│       ├── router/
│       ├── types/
│       └── utils/
│           ├── constant.ts       # Built-in prompts (Word, Excel, PowerPoint, Outlook)
│           ├── enum.ts           # localStorage keys
│           ├── generalTools.ts   # Date + Math tools (for agent)
│           ├── excelTools.ts     # Excel API tools (for agent)
│           ├── hostDetection.ts  # Host detection helpers (isWord, isExcel, isPowerPoint, isOutlook)
│           ├── outlookTools.ts   # Outlook API tools (for agent)
│           ├── powerpointTools.ts # PowerPoint Common API helpers
│           ├── wordFormatter.ts  # Markdown-to-Word formatting
│           ├── wordTools.ts      # Word API tools (for agent)
│           ├── common.ts         # Option lists
│           └── message.ts        # Toast notifications
├── manifests-templates/
│   ├── manifest-office.template.xml    # TaskPaneApp template (Word/Excel/PowerPoint)
│   └── manifest-outlook.template.xml   # MailApp template (Outlook)
├── scripts/
│   └── generate-manifests.js           # Node.js script: templates → manifests
├── .env.example              # Root env vars: SERVER_IP, ports (copy to .env)
├── docker-compose.yml
├── manifest-office.xml       # Generated at docker-compose up (do not edit by hand)
├── manifest-outlook.xml      # Generated at docker-compose up (do not edit by hand)
└── README.md
```

---

## Implementation Status

### Core Infrastructure
- [x] Backend Express server with CORS and JSON parsing
- [x] LLM API proxy (streaming + synchronous)
- [x] Image generation proxy endpoint
- [x] Health check endpoint (`GET /health`)
- [x] Model configuration via `.env` (4 tiers: nano, standard, reasoning, image)
- [x] Models endpoint (`GET /api/models`) - exposes labels only, no secrets
- [x] Docker Compose for Synology NAS (ports 3002/3003, PUID/PGID)
- [x] Backend Dockerfile with health check
- [x] Frontend Dockerfile (multi-stage build + nginx)
- [x] Office add-in manifests: TaskPaneApp (Word/Excel/PowerPoint) + MailApp (Outlook)

### Frontend - Chat Interface
- [x] Chat UI with message history (user/assistant bubbles)
- [x] Streaming responses (SSE parsing)
- [x] Model tier selector (dropdown from backend-provided list)
- [x] New chat / clear history
- [x] Stop generation button
- [x] Copy to clipboard
- [x] Insert into document (replace / append)
- [x] Include selected text from Word in messages
- [x] Word formatting toggle (markdown-to-Word conversion)
- [x] `<think>` tag parsing (collapsible reasoning display)
- [x] Backend online/offline indicator with auto-reconnect check
- [x] Image generation mode (UI + backend integration)

### Frontend - Agent Mode
- [x] Ask mode / Agent mode toggle
- [x] Agent loop: sends tools to LLM via `/api/chat/sync`, executes locally, loops
- [x] OpenAI function-calling format for tool definitions
- [x] Tool execution status display in chat
- [x] Max iterations limit (configurable)
- [x] 24 Word tools: getSelectedText, insertText, replaceSelectedText, appendText, insertParagraph, formatText, searchAndReplace, getDocumentContent, getDocumentProperties, insertTable, insertList, deleteText, clearFormatting, setFontName, insertPageBreak, getRangeInfo, selectText, insertImage, getTableInfo, insertBookmark, goToBookmark, insertContentControl, findText, applyTaggedFormatting
- [x] 25 Excel tools: getSelectedCells, setCellValue, getWorksheetData, insertFormula, fillFormulaDown, createChart, formatRange, sortRange, applyAutoFilter, getWorksheetInfo, insertRow, insertColumn, deleteRow, deleteColumn, mergeCells, setCellNumberFormat, clearRange, getCellFormula, searchAndReplace, autoFitColumns, addWorksheet, setColumnWidth, setRowHeight, applyConditionalFormatting, getConditionalFormattingRules
- [x] 3 Outlook tools: getEmailBody, getSelectedText, setEmailBody
- [x] 2 General tools: getCurrentDate, calculateMath

### Frontend - Quick Actions (Word)
- [x] Translate (with target language)
- [x] Polish / Rewrite
- [x] Academic rewriting
- [x] Summary
- [x] Grammar check
- [x] Customizable built-in prompts (editable in settings)

### Frontend - Quick Actions (Excel)
- [x] Clean (normalize whitespace, trim, remove duplicates)
- [x] Beautify (format cell content for readability)
- [x] Formula (generate an Excel formula from a natural-language description)
- [x] Transform (restructure or transpose data in the selection)
- [x] Highlight (suggest conditional formatting rules for the selection)

### Frontend - Quick Actions (Outlook)
- [x] Smart Reply (pre-fills prompt, user completes intent, sends with email context)
- [x] Formalize (transforms draft into professional email)
- [x] Concise (reduces text by 30-50% while keeping key info)
- [x] Proofread (grammar and spelling correction only, preserves style)
- [x] Extract Tasks (extracts summary, key points, and required actions from email)

### Frontend - Quick Actions (PowerPoint)
- [x] Bullets (convert selected text to concise bullet-point list)
- [x] Speaker Notes (generate conversational presenter notes from slide content)
- [x] Impact / Punchify (rewrite text in punchy headline/marketing style)
- [x] Shrink (reduce text length by ~30% while preserving key info)
- [x] Visual (draft mode: generate an image prompt for slide visuals)

### Frontend - PowerPoint Support
- [x] PowerPoint host detection (`isPowerPoint()`)
- [x] Manifest `<Host xsi:type="Presentation">` with `PrimaryCommandSurface` in TabHome
- [x] Text selection via Common API (`getSelectedDataAsync` with `CoercionType.Text`)
- [x] Text insertion via Common API (`setSelectedDataAsync`)
- [x] PowerPoint-specific agent prompt (slide-first, concise, visual-oriented)
- [x] PowerPoint-specific built-in prompts (customizable)

### Frontend - Settings
- [x] UI language selector (French / English)
- [x] Reply language selector
- [x] Agent max iterations setting
- [x] User profile settings (first name, last name, gender)
- [x] Backend status display
- [x] Configured models display (read-only)
- [x] Custom prompts management (add/edit/delete)
- [x] Built-in prompts editor (with reset)
- [x] Tool enable/disable toggles

### Internationalization
- [x] i18n framework (vue-i18n)
- [x] 2 UI locales: English, French
- [x] 13 reply languages selectable by the user: English, Spanish, French, German, Italian, Portuguese, Simplified Chinese, Japanese, Korean, Dutch, Polish, Arabic, Russian

### Security
- [x] API keys stored server-side only (never sent to client)
- [x] CORS restricted to frontend origin
- [x] No third-party web search or web fetch (removed)
- [x] No user-configurable API endpoints or models
- [ ] HTTPS/TLS configuration (needed for production and Office add-in requirement)
- [ ] Authentication / user login system
- [ ] Rate limiting on backend
- [ ] Request logging / audit trail

### Frontend - Outlook Support
- [x] Outlook host detection (`isOutlook()`)
- [x] Manifest extension points: `MessageReadCommandSurface` + `MessageComposeCommandSurface`
- [x] Asynchronous email body retrieval (`body.getAsync`)
- [x] Selected text retrieval in compose mode (`getSelectedDataAsync`)
- [x] Email body insertion in compose mode (`body.setAsync`)
- [x] Outlook-specific standard and agent prompts
- [x] `ReadWriteMailbox` permission

### Not Yet Implemented
- [ ] Conversation history persistence (currently in-memory only, lost on page reload)
- [ ] User authentication and authorization
- [ ] HTTPS/TLS (required for production Office add-in sideloading)
- [ ] Azure deployment configuration (production server)
- [ ] LiteLLM integration configuration (production LLM endpoints)
- [ ] Custom logo/branding assets (user mentioned they have a logo - needs to replace placeholder icons)
- [ ] Web search capability (disabled for now, could be re-enabled via backend proxy)
- [ ] Chat export (save conversation to file)
- [ ] Token usage tracking / cost monitoring
- [ ] Admin dashboard for model configuration (currently .env only)
- [ ] Multi-user support / user preferences stored server-side
- [ ] Offline mode / graceful degradation when backend is down

---

## Development

### Backend (local)
```bash
cd backend
cp .env.example .env   # Fill in LLM_API_KEY
npm install
npm run dev            # Starts on port 3003 with --watch
```

### Frontend (local)
```bash
cd frontend
npm install
npm run dev            # Starts on port 3002 with HMR
```

### Environment Variables

#### Root (`.env`)
| Variable | Description | Default |
|----------|-------------|---------|
| `SERVER_IP` | Host machine IP address | `192.168.50.10` |
| `FRONTEND_PORT` | Published port for the frontend | `3002` |
| `BACKEND_PORT` | Published port for the backend | `3003` |

#### Backend (`backend/.env`)
| Variable | Description | Default |
|----------|-------------|---------|
| `PORT` | Backend port | `3003` |
| `FRONTEND_URL` | Allowed CORS origin | `http://192.168.50.10:3002` |
| `LLM_API_BASE_URL` | OpenAI-compatible API base URL | `https://api.openai.com/v1` |
| `LLM_API_KEY` | API key for LLM provider | (required) |
| `MODEL_NANO` | Model ID for basic tasks | `gpt-4.1-nano` |
| `MODEL_STANDARD` | Model ID for standard tasks | `gpt-4.1` |
| `MODEL_REASONING` | Model ID for complex tasks | `o3` |
| `MODEL_IMAGE` | Model ID for image generation | `gpt-image-1` |

#### Frontend (`frontend/.env`)
| Variable | Description | Default |
|----------|-------------|---------|
| `VITE_BACKEND_URL` | Backend URL (build-time) | `http://localhost:3003` |

---

## Production Deployment (Azure)

For production, the architecture stays the same but:

1. **Server**: Azure VM or App Service instead of Synology NAS
2. **LLM**: LiteLLM proxy (OpenAI-compatible format) instead of direct OpenAI API
3. **TLS**: HTTPS required for Office add-in (configure nginx with certificates or use Azure Front Door)
4. **Manifest**: `manifest-office.xml` and `manifest-outlook.xml` are auto-generated at `docker compose up`. You can remove the stopped init container with `docker compose rm -f manifest-gen`; it will be recreated on next startup. Update `SERVER_IP` / ports in the root `.env` to change all URLs
5. **Auth**: Add authentication middleware to the backend

Update `backend/.env`:
```env
LLM_API_BASE_URL=https://your-litellm-proxy.azure.com/v1
LLM_API_KEY=your-litellm-key
FRONTEND_URL=https://kickoffice.yourdomain.com
```

---

## Credits

### Based on [WordGPT Plus](https://github.com/Kuingsmile/word-GPT-Plus) by Kuingsmile (MIT License)

The following was directly reused or adapted from WordGPT Plus:

- **`wordFormatter.ts`** — Markdown-to-Word conversion engine: parses headings, bold, italic, code blocks, and lists, then applies Word built-in styles via `Word.run()`
- **`api/common.ts`** — Document insertion logic: replace / append / newLine insertion modes using the Word.js API
- **Chat UI architecture** — Vue 3 task pane structure, message history (user/assistant bubbles), streaming SSE parsing, stop-generation button
- **Built-in prompt structure** (`constant.ts`) — `buildInPrompt` pattern with translate, polish, academic rewrite, summary, and grammar-check prompts
- **Settings page architecture** (`SettingsPage.vue`) — custom prompt management (add/edit/delete), built-in prompt editor with per-prompt reset
- **i18n framework** — `vue-i18n` integration and locale file structure

Removed from WordGPT Plus for KickOffice:
- Multi-provider LLM support (OpenAI, Azure, Gemini, etc.) → single controlled backend endpoint
- User-side API key and endpoint configuration → admin-only via `backend/.env`
- Web search / web fetch features (privacy)

Added or extended for KickOffice:
- Backend Express proxy server (API keys never reach the client)
- Docker deployment for Synology NAS (PUID/PGID, health check, multi-stage frontend build)
- Extended from Word-only to Word + Excel + PowerPoint + Outlook
- Model tier system (nano / standard / reasoning / image), configured server-side only
- French translations added to the i18n locale files
- Image generation support

---

### Based on [excel-ai-assistant](https://github.com/ilberpy/excel-ai-assistant) by ilberpy (MIT License)

The following was directly reused or adapted from excel-ai-assistant:

- **Tool definition schema** — the `{ name, description, inputSchema, execute }` pattern (originally a LangChain tool format) was adapted to the OpenAI function-calling JSON schema format and applied to all tool sets (Word, Excel, general)
- **Excel tool set** (`excelTools.ts`) — tool names, descriptions, and parameter schemas for Excel operations (getSelectedCells, setCellValue, getWorksheetData, insertFormula, createChart, formatRange, sortRange, applyAutoFilter, etc.) were derived from excel-ai-assistant's tool catalogue; implementations were rewritten using `Excel.run()` directly
- **Agent loop pattern** — the core loop (send tools to LLM → detect `tool_calls` in response → execute locally → feed result back → repeat) was adapted from excel-ai-assistant's LangChain agent to a custom TypeScript implementation
- **Formula language localization** — concept of detecting the user's configured formula language (`getExcelFormulaLanguage()`, fr/en) to match locale-specific function names

Removed from excel-ai-assistant for KickOffice:
- LangChain dependency entirely → direct OpenAI-compatible API calls
- Python/server-side Excel bindings → replaced by client-side `Excel.run()` (Office.js)
- Standalone script architecture → integrated into the unified Office add-in frontend
