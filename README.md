# KickOffice

AI-powered Microsoft Office add-in for Word, Excel, PowerPoint, and Outlook. Features a chat interface, autonomous document agent with 100 specialized tools, image generation, and quick AI actions—all running through a secure backend proxy.

**Built for enterprise environments**: API keys never reach the client, all LLM traffic flows through a controlled backend, and no data is sent to third-party services.

---

## Features

- **Chat Interface** — Converse with AI directly within Office apps
- **Autonomous Agent** — 100 tools for document manipulation, data analysis, and automation
- **Quick Actions** — One-click translate, polish, summarize, generate formulas, and more
- **Image Generation** — Create and insert AI-generated images into documents
- **Native Track Changes** — Word `proposeRevision` and `proposeDocumentRevision` generate real `<w:ins>/<w:del>` OOXML markup via docx-redline-js; users accept/reject in Word's Review pane. Bulk-accept/reject via `acceptAiChanges`/`rejectAiChanges` tools or the "Valider" UI button
- **Multi-Host Support** — Word (34 tools), Excel (27 tools), PowerPoint (24 tools), Outlook (9 tools), General (6 tools)
- **Skill System** — 17 Quick Action skill files + 5 host skill files define agent behavior in Markdown
- **Context Management** — Automatic context window compression: older tool results are truncated, recent iterations kept in full
- **Secure Sandbox** — SES-based execution environment for safe dynamic code
- **File Analysis** — Upload and analyze PDF, DOCX, XLSX, CSV documents (up to 50 MB)
- **Session Persistence** — Uploaded files and images stay in context across the conversation and are restored on session switch
- **Large File Support** — Extended 5-minute LLM timeout; files optionally forwarded via `/v1/files` API to avoid re-sending content on every message
- **Internationalization** — 2 UI languages (EN/FR), 13 reply languages
- **Reverse Proxy Support** — Compatible with Synology/nginx reverse proxies
- **Stream Error Recovery** — Retry button on interrupted responses; SSE parse failures and upstream errors logged and surfaced to the user
- **Request Correlation** — `X-Request-Id` header shared between frontend and backend for end-to-end log tracing
- **Frontend Log Forwarding** — warn/error log entries flushed to the backend `/api/logs` endpoint every 30 s for centralized monitoring

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
   Word / Excel /                      │  API key storage
   PowerPoint / Outlook                ▼
                                  .env file
```

### Frontend

Vue 3 task pane loaded inside Office apps via `createMemoryHistory` router (avoids URL conflicts with Office iframe). Composable-based architecture: `useAgentLoop`, `useQuickActions`, `useSessionFiles`, `useOfficeInsert`, `useImageActions`, and more.

### Backend

Express.js proxy server. Holds all secrets (API keys), validates requests, rate-limits by IP, and exposes:

- `POST /api/chat` — Streaming chat with SSE
- `POST /api/chat/sync` — Synchronous chat for agent tool loops
- `POST /api/image` — Image generation
- `POST /api/upload` — File processing (PDF, DOCX, XLSX, CSV, images)
- `POST /api/chart-extract` — Chart image data extraction (pixel color analysis)
- `POST /api/files` — Proxy: upload file to LLM provider's `/v1/files` endpoint, returns `file_id`
- `GET /api/icons/search` — Icon search proxy (Iconify API)
- `GET /api/icons/svg/:prefix/:name` — SVG icon fetch with optional color
- `GET /api/models` — Available model tiers
- `POST /api/feedback` — User feedback submission with log export
- `POST /api/logs` — Frontend log aggregation endpoint
- `GET /health` — Health check

### LLM API

Any OpenAI-compatible endpoint. For testing: OpenAI API directly. For production: Azure-hosted LiteLLM proxy.

---

## Model Tiers

Models are configured **server-side only** in `backend/.env`. Three tiers:

| Tier        | Purpose          | Default Model                       | Use Case                       |
| ----------- | ---------------- | ----------------------------------- | ------------------------------ |
| `standard`  | Normal tasks     | `gpt-5.1`                           | Chat, writing, analysis        |
| `reasoning` | Complex tasks    | `gpt-5.1` + `reasoning_effort=high` | Multi-step reasoning, planning |
| `image`     | Image generation | `gpt-image-1`                       | Generate images                |

---

## Agent Stability System

KickOffice implements three complementary systems for reliable Office.js code execution:

### 1. Skills System (Defensive Prompting + User Customization)

Office.js best practices automatically injected into agent prompts via `.skill.md` files.

All skill files share a unified YAML frontmatter format (`name`, `description`, `host`, `executionMode`, `icon`, `actionKey`), parsed at build time by `skillParser.ts`. This enables a metadata-driven registry and powers the User Skills feature.

**Built-in skills:**
- **5 host skills**: `common.skill.md` (universal) + Word / Excel / PowerPoint / Outlook
- **24 Quick Action skills**: bullets, punchify, review, translate, formalize, concise, proofread, polish, academic, summary, word-translate, word-proofread, word-review, ppt-translate, ppt-proofread, ingest, autograph, explain-excel, formula-generator, data-trend, chart-digitizer, pixel-art, extract, reply

**User Skills (new):** Users can create custom skills directly in the add-in via a 4-step LLM-assisted creator (describe in natural language → LLM generates a full `.skill.md` → review/edit → test on selected text → save). User skills appear in the quick actions bar and execute with the same pipeline as built-in skills. Skills are stored in `localStorage` and can be exported/imported as `.skill.md` files for sharing.

### 2. Code Validator (Pre-Execution Safety)

All `eval_*` tools validate code before execution:

- **Blocked**: Missing `sync()`, missing `load()`, wrong namespace, infinite loops, `eval()`/`new Function()`
- **Warnings**: Missing try/catch, excessive sync calls, incorrect array formats

### 3. Track Changes Integration (Format Preservation)

Native Word revision markup via `docx-redline-js`:

- **`proposeRevision`**: Applies Track Changes to the current selection — `w:ins`/`w:del` OOXML injected via disable-TC → insertOoxml → restore-TC pattern
- **`proposeDocumentRevision`**: Same chirurgical diff, document-wide — matches paragraphs by text, applies redlines paragraph by paragraph without requiring a selection
- **`editDocumentXml`**: Direct OOXML manipulation for formatting preservation (fonts, colors, styles)
- **Configurable author**: Track Changes attributed to "KickOffice AI" (customizable in Settings)

---

## Tool Summary

| Host           | Tools  | Highlights                                                                                    |
| -------------- | ------ | --------------------------------------------------------------------------------------------- |
| **Word**       | 34     | `proposeRevision`, `proposeDocumentRevision`, `editDocumentXml`, `insertOoxml`, `acceptAiChanges`, `rejectAiChanges`, `addComment`, `getComments`, `getDocumentOoxml`, `eval_wordjs`, Track Changes |
| **Excel**      | 27     | `eval_officejs`, formulas, charts (incl. Waterfall/Treemap/Funnel), screenshots, CSV, pixel art, header detection |
| **PowerPoint** | 24     | `editSlideXml`, `reorderSlide`, `getSpeakerNotes`, `setSpeakerNotes`, slides, shapes, screenshots, icons (Iconify), `verifySlides` |
| **Outlook**    | 9      | `eval_outlookjs`, `addAttachment`, email body/subject, recipients, rich content preservation  |
| **General**    | 6      | `executeBash` (VFS), `calculateMath`, `getCurrentDate`, file operations                       |
| **Total**      | **100** |                                                                                              |

---

## Quick Actions

### Word

Translate, Polish, Academic, Summary, Proofread, Formalize, Concise

### Excel

Clean (Ingest), Beautify (Autograph), Formula (Formula Generator), Data Trend, Explain Excel

### PowerPoint

Bullets, Speaker Notes (Review), Impact (Punchify), Visual

### Outlook

Smart Reply, Formalize, Concise, Proofread, Extract Tasks

---

## Deployment (Docker)

### Prerequisites

- Docker with Compose
- OpenAI API key (or LiteLLM proxy)

### Steps

1. **Clone and configure**:

   ```bash
   git clone https://github.com/your-org/kickoffice.git
   cd kickoffice
   cp .env.example .env
   cp backend/.env.example backend/.env
   # Edit backend/.env and set LLM_API_KEY
   ```

2. **Build and start**:

   ```bash
   docker compose up -d --build
   ```

3. **Verify**:
   - Backend: `curl http://localhost:3003/health`
   - Frontend: Open `http://localhost:3002`
   - Models: `curl http://localhost:3003/api/models`

4. **Install Office add-ins**:
   - Sideload `manifest-office.xml` in Word/Excel/PowerPoint
   - Sideload `manifest-outlook.xml` in Outlook

### Docker Services

| Container                 | Port | Description                                               |
| ------------------------- | ---- | --------------------------------------------------------- |
| `kickoffice-manifest-gen` | —    | Generates manifests from templates (init, can be removed) |
| `kickoffice-backend`      | 3003 | Express.js API server (non-root, node:22-slim)            |
| `kickoffice-frontend`     | 3002 | nginx-unprivileged serving Vue app                        |

> **Note**: All containers run as non-root users. Debian-based images required (`node:22-slim`, `nginxinc/nginx-unprivileged`) — Alpine is incompatible with older Intel Celeron hardware (Synology DS416play).

---

## Project Structure

```
KickOffice/
├── backend/                    # Express.js API server
│   └── src/
│       ├── server.js           # Entry point
│       ├── config/             # env.js, models.js, limits.js
│       ├── middleware/         # auth.js, validate.js + validators/
│       ├── routes/             # chat, image, upload, files, icons, models, health, logs, feedback
│       ├── services/           # llmClient.js, plotDigitizerService.js, imageStore.js
│       └── utils/              # http.js, logger.js, toolUsageLogger.js
├── frontend/                   # Vue 3 + TypeScript
│   └── src/
│       ├── api/                # backend.ts (HTTP client)
│       ├── components/         # Chat UI, settings tabs
│       ├── composables/        # 17 composables: useAgentLoop, useQuickActions, useSessionFiles,
│       │                       # useDocumentUndo, useSessionDB, useSessionManager, etc.
│       ├── constants/          # limits.ts (centralized magic numbers)
│       ├── i18n/               # en.json, fr.json
│       ├── pages/              # HomePage, SettingsPage
│       ├── router/             # Memory history router (Office iframe compatible)
│       ├── skills/             # 5 host skills + 17 Quick Action skills
│       ├── types/              # TypeScript definitions
│       └── utils/              # Tools (word, excel, ppt, outlook, general),
│                               # tokenManager, wordDiffUtils, toolProviderRegistry, etc.
├── manifests-templates/        # XML templates for Office add-ins
├── scripts/                    # generate-manifests.js
├── docker-compose.yml
└── .env.example
```

---

## Development

### Backend

```bash
cd backend
cp .env.example .env  # Set LLM_API_KEY
npm install
npm run dev           # Port 3003 with --watch
```

### Frontend

```bash
cd frontend
npm install
npm run dev           # Port 3002 with HMR
npm run build         # Production build
```

---

## Environment Variables

### Root (`.env`)

| Variable        | Description     | Default       |
| --------------- | --------------- | ------------- |
| `SERVER_IP`     | Host machine IP | `localhost`   |
| `FRONTEND_PORT` | Frontend port   | `3002`        |
| `BACKEND_PORT`  | Backend port    | `3003`        |

### Backend (`backend/.env`)

| Variable                 | Description                | Default                     |
| ------------------------ | -------------------------- | --------------------------- |
| `LLM_API_KEY`            | API key for LLM provider   | (required)                  |
| `LLM_API_BASE_URL`       | OpenAI-compatible base URL | `https://api.openai.com/v1` |
| `MODEL_STANDARD`         | Standard model ID          | `gpt-5.1`                   |
| `MODEL_REASONING`        | Reasoning model ID         | `gpt-5.1`                   |
| `MODEL_REASONING_EFFORT` | Reasoning effort level     | `high`                      |
| `MODEL_IMAGE`            | Image model ID             | `gpt-image-1`               |

---

## Security

- **API keys server-side only** — Never sent to client
- **CORS restricted** — Frontend origin only
- **Rate limiting** — IP-based on chat, image, and upload endpoints; 5 s minimum floor on rate-limit retry
- **Credential encryption** — Web Crypto API (AES-GCM 256-bit) for stored credentials
- **Non-root containers** — Both backend and frontend run as non-root users
- **SES sandbox** — Safe dynamic code execution with host isolation
- **Code validation** — Pre-execution checks for Office.js patterns
- **Helmet headers** — HSTS, X-Frame-Options, X-Content-Type-Options
- **DOMPurify** — XSS protection with strict allowlists; custom color syntax validated against a strict allowlist
- **Structured logging** — All errors/warnings routed through logService (not raw console); forwarded to backend
- **Request correlation** — `X-Request-Id` header links frontend and backend log entries
- **No third-party services** — Privacy-first, no telemetry

---

## Credits & Inspirations

### ⭐ [Office Agents](https://github.com/nicepkg/office-agents) (MIT License) — Major inspiration

- Excel range screenshot with row/column header overlay (A, B, C… / 1, 2, 3…) for improved vision accuracy
- Enriched Office.js error feedback (`debugInfo.errorLocation`, `statement`, `surroundingStatements`)
- Static mutation tracker — detects write operations in `eval_*` code, returns `hasMutated` flags
- VFS injection in SES sandbox (`btoa`, `atob`, `readFile`, `readFileBuffer`, `writeFile`)
- CSV-to-sheet / sheet-to-CSV bash commands with type coercion
- Image-to-sheet pixel art (downsampling, run-length encoding, batched color assignments)
- Word OOXML extraction with body-child summaries, referenced styles, numbering definitions
- Word document metadata enrichment (run-level formatting, heading outline, content control info)
- Excel formula search in `findData`

### [word-GPT-Plus](https://github.com/Kuingsmile/word-GPT-Plus) (MIT License)

- Office task pane chat UI architecture
- SSE streaming response handling
- Settings page structure
- i18n framework (multi-language UI + reply language selection)
- Quick action prompt system

### [excel-ai-assistant](https://github.com/ilberpy/excel-ai-assistant) (MIT License)

- Tool definition schema pattern
- Excel tool set design and agent loop
- Formula locale switching (French / English)
- Chart creation and data analysis flows

### [docx-redline-js](https://github.com/AnsonLai/docx-redline-js) (MIT License)

- Native Word Track Changes engine (`<w:ins>` / `<w:del>` OOXML)
- Configurable revision author
- Formatting preservation in diffs
- Zero-dependency implementation

### [Gemini AI for Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting) (MIT License)

- TC disable → `insertOoxml` → TC restore pattern to prevent double-tracking

### [Redink](https://github.com/LawDigital/redink) (MIT License)

- Multi-host Office add-in design (Word, Excel, Outlook)
- AI-powered inline review and commenting approach
- Document revision workflow patterns

### [Iconify](https://iconify.design) (MIT License — API free to use)

- Icon search and SVG delivery (`searchIcons` / `insertIcon` PowerPoint tools, 200,000+ icons)

### [JSZip](https://stuk.github.io/jszip/) (MIT License)

- PPTX ZIP manipulation for `editSlideXml`

---

## License

This project is proprietary software. Third-party dependencies retain their original licenses (MIT, Apache 2.0, etc.).
