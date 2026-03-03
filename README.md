# KickOffice

AI-powered Microsoft Office add-in for Word, Excel, PowerPoint, and Outlook. Features a chat interface, autonomous document agent with 129 specialized tools, image generation, and quick AI actions—all running through a secure backend proxy.

**Built for enterprise environments**: API keys never reach the client, all LLM traffic flows through a controlled backend, and no data is sent to third-party services.

---

## Features

- **Chat Interface** — Converse with AI directly within Office apps
- **Autonomous Agent** — 129 tools for document manipulation, data analysis, and automation
- **Quick Actions** — One-click translate, polish, summarize, generate formulas, and more
- **Image Generation** — Create and insert AI-generated images into documents
- **Format Preservation** — Word-level diffing preserves formatting when editing text
- **Multi-Host Support** — Word (41 tools), Excel (45 tools), PowerPoint (16 tools), Outlook (14 tools)
- **Secure Sandbox** — SES-based execution environment for safe dynamic code
- **File Analysis** — Upload and analyze PDF, DOCX, XLSX, CSV documents
- **Internationalization** — 2 UI languages (EN/FR), 13 reply languages
- **Reverse Proxy Support** — Compatible with Synology/nginx reverse proxies
- **Message Timestamps** — Chat messages display creation time for better context

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
Vue 3 task pane loaded inside Office apps. Handles UI, chat, and agent tool execution (Office.js API calls run locally in the browser).

### Backend
Express.js proxy server. Holds all secrets (API keys), validates requests, rate-limits by IP, and exposes:
- `POST /api/chat` — Streaming chat with SSE
- `POST /api/chat/sync` — Synchronous chat for agent tool loops
- `POST /api/image` — Image generation
- `POST /api/upload` — File processing (PDF, DOCX, XLSX, CSV)
- `GET /api/models` — Available model tiers
- `GET /health` — Health check

### LLM API
Any OpenAI-compatible endpoint. For testing: OpenAI API directly. For production: Azure-hosted LiteLLM proxy.

---

## Model Tiers

Models are configured **server-side only** in `backend/.env`. Three tiers:

| Tier | Purpose | Default Model | Use Case |
|------|---------|---------------|----------|
| `standard` | Normal tasks | `gpt-5.1` | Chat, writing, analysis |
| `reasoning` | Complex tasks | `gpt-5.1` + `reasoning_effort=high` | Multi-step reasoning, planning |
| `image` | Image generation | `gpt-image-1` | Generate images |

---

## Agent Stability System

KickOffice implements three complementary systems for reliable Office.js code execution:

### 1. Skills System (Defensive Prompting)
Office.js best practices automatically injected into agent prompts:
- **THE PROXY PATTERN**: Explains Office.js object lifecycle (proxy → load → sync → access)
- **5 Critical Rules**: Always load() before reading, always sync() after writing, use try/catch, check empty selections, prefer dedicated tools
- **Host-Specific Guidance**: Word, Excel, PowerPoint, Outlook patterns

### 2. Code Validator (Pre-Execution Safety)
All `eval_*` tools validate code before execution:
- **Blocked**: Missing sync(), missing load(), wrong namespace, infinite loops, eval()/new Function()
- **Warnings**: Missing try/catch, excessive sync calls, incorrect array formats

### 3. Diffing Integration (Format Preservation)
Word-level surgical editing via `office-word-diff` library:
- **Word `proposeRevision`**: Applies only insertions/deletions, preserving formatting on unchanged text
- **PowerPoint `proposeShapeTextRevision`**: Diff statistics with full replacement
- **Cascading strategies**: Token → Sentence → Block fallback

---

## Tool Summary

| Host | Tools | Highlights |
|------|-------|------------|
| **Word** | 41 | `proposeRevision` (format-preserving edits), `eval_wordjs`, tables, comments, Track Changes |
| **Excel** | 45 | `eval_officejs`, formulas, charts, conditional formatting, data validation |
| **PowerPoint** | 16 | `proposeShapeTextRevision`, slides, shapes, speaker notes, images |
| **Outlook** | 14 | `eval_outlookjs`, email body/subject, recipients, attachments |
| **General** | 6 | `executeBash` (VFS), `calculateMath`, `getCurrentDate`, file operations |
| **Total** | **129** | |

---

## Quick Actions

### Word
Translate, Polish, Academic, Summary, Grammar Check

### Excel
Clean, Beautify, Formula, Transform, Highlight

### PowerPoint
Bullets, Speaker Notes, Impact, Shrink, Visual

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

| Container | Port | Description |
|-----------|------|-------------|
| `kickoffice-manifest-gen` | — | Generates manifests from templates (init, can be removed) |
| `kickoffice-backend` | 3003 | Express.js API server with health check |
| `kickoffice-frontend` | 3002 | Nginx serving Vue app |

---

## Project Structure

```
KickOffice/
├── backend/                    # Express.js API server
│   └── src/
│       ├── server.js           # Entry point
│       ├── config/             # env.js, models.js
│       ├── middleware/         # auth.js, validate.js
│       ├── routes/             # chat, image, upload, models, health
│       ├── services/           # llmClient.js
│       └── utils/              # http.js, logger.js
├── frontend/                   # Vue 3 + TypeScript
│   └── src/
│       ├── api/                # backend.ts (HTTP client)
│       ├── components/         # Chat UI, settings components
│       ├── composables/        # useAgentLoop, useImageActions, etc.
│       ├── i18n/               # en.json, fr.json
│       ├── pages/              # HomePage, SettingsPage
│       ├── skills/             # Office.js best practices (5 files)
│       ├── types/              # TypeScript definitions
│       └── utils/              # Tools (word, excel, ppt, outlook), validators
├── office-word-diff/           # Word diffing library (Apache 2.0)
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
```

### Testing
```bash
cd frontend
npm run test:e2e      # Playwright tests
```

---

## Environment Variables

### Root (`.env`)
| Variable | Description | Default |
|----------|-------------|---------|
| `SERVER_IP` | Host machine IP | `192.168.50.10` |
| `FRONTEND_PORT` | Frontend port | `3002` |
| `BACKEND_PORT` | Backend port | `3003` |

### Backend (`backend/.env`)
| Variable | Description | Default |
|----------|-------------|---------|
| `LLM_API_KEY` | API key for LLM provider | (required) |
| `LLM_API_BASE_URL` | OpenAI-compatible base URL | `https://api.openai.com/v1` |
| `MODEL_STANDARD` | Standard model ID | `gpt-5.1` |
| `MODEL_REASONING` | Reasoning model ID | `gpt-5.1` |
| `MODEL_REASONING_EFFORT` | Reasoning effort level | `high` |
| `MODEL_IMAGE` | Image model ID | `gpt-image-1` |

---

## Security

- **API keys server-side only** — Never sent to client
- **CORS restricted** — Frontend origin only
- **Rate limiting** — IP-based on chat, image, and upload endpoints
- **Credential encryption** — Web Crypto API (AES-GCM 256-bit) for stored credentials
- **CSRF protection** — Origin validation for state-changing requests
- **Stream abort handling** — Proper cleanup and timeout for streaming connections
- **SES sandbox** — Safe dynamic code execution with host isolation
- **Code validation** — Pre-execution checks for Office.js patterns
- **Helmet headers** — HSTS, X-Frame-Options, X-Content-Type-Options
- **DOMPurify** — XSS protection with strict allowlists
- **Safe JSON handling** — Depth validation and circular reference detection
- **No third-party services** — Privacy-first, no telemetry

---

## Credits & Inspirations

KickOffice builds upon several excellent open-source projects:

### [word-GPT-Plus](https://github.com/Kuingsmile/word-GPT-Plus) (MIT License)
The original foundation for the Word add-in architecture. Directly reused or adapted:
- **`wordFormatter.ts`** — Markdown-to-Word conversion engine
- **Chat UI architecture** — Vue 3 task pane, message bubbles, SSE streaming
- **Built-in prompt structure** — Translate, polish, academic, summary patterns
- **Settings page architecture** — Custom prompt management
- **i18n framework** — vue-i18n integration

### [excel-ai-assistant](https://github.com/ilberpy/excel-ai-assistant) (MIT License)
Inspired the Excel tooling and agent loop pattern:
- **Tool definition schema** — `{ name, description, inputSchema, execute }` pattern
- **Excel tool set** — Tool names, descriptions, parameter schemas
- **Agent loop pattern** — Send tools → detect tool_calls → execute → loop
- **Formula localization** — Locale-specific function names (en/fr)

### [office-word-diff](https://github.com/yuch85/office-word-diff) (Apache 2.0)
Integrated as a local package for format-preserving text editing:
- **Word-level diffing** — Token mapping with formatting preservation
- **Track Changes integration** — Native Word revision tracking
- **Cascading strategies** — Token → Sentence → Block fallback
- **diff-match-patch extension** — Google's algorithm with word-mode

### [Redink](https://github.com/LawDigital/redink) (MIT License)
Conceptual inspiration for document comparison and revision workflows.

---

## License

This project is proprietary software. The integrated `office-word-diff` library is licensed under Apache 2.0. Third-party dependencies retain their original licenses.

---

## Known Issues

See [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) for the complete audit with issues organized by severity.

**Critical**:
- Agent max iterations silently capped at 10
- `.env.example` contains invalid `reasoning_effort=none`
- Quick actions bypass loading/abort state

See the design review for full details and recommendations.
