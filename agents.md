# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope
- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

## 2) Product and Architecture Snapshot
KickOffice is a Microsoft Office add-in with:
- `frontend/` (Vue 3 + Vite + TypeScript): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express, modular): secure proxy for model calls and model configuration exposure.
- `manifests-templates/`: manifest templates used to generate runtime manifests.
- `scripts/generate-manifests.js`: script that generates `manifest-office.xml` and `manifest-outlook.xml` from root `.env`.

### Frontend architecture (post-refactor)
The frontend follows a composable-based architecture:
- **Pages**: `HomePage.vue` (265 lines, orchestration only), `SettingsPage.vue` (settings UI)
- **Components**: `chat/ChatHeader.vue`, `chat/ChatInput.vue`, `chat/ChatMessageList.vue`, `chat/QuickActionsBar.vue` + generic components (`CustomButton`, `CustomInput`, `Message`, `SettingCard`, `SingleSelect`)
- **Composables**: `useAgentLoop.ts` (agent execution + prompts), `useImageActions.ts` (image generation + insertion), `useOfficeInsert.ts` (document insertion + clipboard)
- **Utils**: `wordTools.ts` (39 tools), `excelTools.ts` (39 tools), `powerpointTools.ts` (8 tools), `outlookTools.ts` (13 tools), `generalTools.ts` (2 tools), `wordFormatter.ts` (markdown-to-Word), `constant.ts` (built-in prompts), `hostDetection.ts`, `message.ts`, `common.ts`, `enum.ts`, `savedPrompts.ts`, `officeOutlook.ts`

### Backend architecture (post-refactor)
```
backend/src/
├── server.js              # Entry point: middleware setup, route mounting (80 lines)
├── config/
│   ├── env.js             # Environment variable loading (PORT, FRONTEND_URL, rate limits)
│   └── models.js          # Model tier config, buildChatBody(), isGpt5Model(), isChatGptModel()
├── middleware/
│   ├── auth.js            # ensureLlmApiKey — checks API key is configured
│   └── validate.js        # validateTemperature, validateMaxTokens, validateTools, validateImagePayload
├── routes/
│   ├── chat.js            # POST /api/chat (streaming), POST /api/chat/sync (agent tool loop)
│   ├── health.js          # GET /health
│   ├── image.js           # POST /api/image
│   └── models.js          # GET /api/models
└── utils/
    └── http.js            # fetchWithTimeout, logAndRespond
```

Manifest outputs:
- `manifest-office.xml`: Word + Excel + PowerPoint TaskPane add-in.
- `manifest-outlook.xml`: Outlook Mail add-in.

Important manifest rule:
- Do not hand-edit generated manifest files when URL/host values change. Update templates and/or `.env`, then regenerate.

Backend API surface:
- `POST /api/chat` (streaming) — uses `buildChatBody` with `stream: true`, no tools
- `POST /api/chat/sync` (agent tool loop) — uses `buildChatBody` with `stream: false`, tools + tool_choice
- `POST /api/image`
- `GET /api/models`
- `GET /health`

Operational backend behavior to preserve:
- IP rate limiting is enabled on `/api/chat*` and `/api/image`.
- Server logs include request logging (`morgan`) and sanitized API error responses.
- Upstream provider errors are logged server-side; clients receive generic/safe error messages (502).
- Helmet security headers are enabled (CSP and COEP disabled for Office add-in compatibility).

## 3) Working Principles
1. **Preserve behavior unless explicitly requested**.
2. **Prefer minimal, localized diffs** over broad rewrites.
3. **Keep frontend and backend contracts aligned**.
4. **Do not introduce secrets in frontend code**.
5. **Update documentation only where needed** (no unnecessary restructuring).

## 4) API Contract Rules (Important)

### Image responses
Do not assume a single image payload format. Compatible providers may return:
- `data[0].b64_json`
- `data[0].url`

Frontend changes touching image flow must keep support for both payload styles unless a migration plan says otherwise.

### Chat responses
- Streaming path (`/api/chat`) consumes SSE-like `data:` lines with buffer-based parsing (`backend.ts:131-139`).
- Sync path (`/api/chat/sync`) expects OpenAI-compatible `choices/message/tool_calls` structures.

### Model parameter compatibility
- Keep compatibility with OpenAI-compatible providers that may differ on parameter support.
- `ChatGPT-*` model IDs do not support `temperature` / token-limit parameters in current backend logic (`isChatGptModel` check).
- GPT-5 models use `max_completion_tokens` while non-GPT-5 chat models use `max_tokens` (`isGpt5Model` check).
- **CRITICAL**: `reasoning_effort` parameter behavior varies by model. The value `'none'` is NOT a valid OpenAI API value and causes empty responses when used with tools. Valid values are `'low'`, `'medium'`, `'high'`. When reasoning is not needed, **omit the parameter entirely** rather than sending `'none'`. See DESIGN_REVIEW.md C7 for details.
- If model plumbing is changed, update backend validation + request shaping + frontend expectations together.

Any contract change should update both backend and frontend in the same change set.

## 5) Frontend Editing Guidelines
- Keep type names explicit and stable.
- Avoid silent failures in user-facing flows; surface understandable errors.
- Preserve i18n behavior and existing translation keys when possible.
- **All user-visible strings must use `t()` / i18n keys** — no hardcoded French or English strings in components or composables.
- For Office tool changes, ensure host-specific behavior remains correct (Word vs Excel vs PowerPoint vs Outlook).
- Host capability gaps must be handled explicitly in UX (clear message instead of degraded/broken fallback behavior).
- Image insertion: Word uses `insertInlinePictureFromBase64`, PowerPoint uses `slide.shapes.addImage`, Excel shows an info message (not supported). Never fall back to copying base64 text to clipboard.
- When modifying the agent loop (`useAgentLoop.ts`), ensure:
  - Abort signal is checked at loop start and passed to `chatSync`
  - Empty model responses are detected and surfaced to the user
  - Tool definitions are correctly gathered for the current host

Current host/tool landscape (keep in mind for tool/agent changes):
- Word tools: 39 tools (high-count set) + 2 general tools used in agent mode.
- Excel tools: 39 tools (high-count set) + 2 general tools used in agent mode.
- PowerPoint tools: 8 tools (focused set for slide editing and speaker notes) + 2 general tools.
- Outlook tools: 13 tools (mail compose/read helpers) + 2 general tools.

**Known dead code**: The Settings "Tools" tab saves enabled/disabled tool preferences to `localStorage('enabledTools')`, but `useAgentLoop.ts` does not read this value. All tools are always included. If wiring this up, filter tools in `useAgentLoop.ts:151-153` before constructing the tools array.

## 6) Backend Editing Guidelines
- Keep proxy logic provider-agnostic for OpenAI-compatible endpoints.
- Log upstream provider errors on the server, but return sanitized client-facing messages (no raw upstream `details`).
- Do not leak API keys or environment secrets in logs or responses.
- Keep input validation strict for: `messages`, `tools`, `temperature`, `maxTokens`, image prompt/size/quality/count. All validation is in `middleware/validate.js`.
- Preserve timeout behavior per endpoint/tier (120s for standard, 300s for reasoning).
- Use `logAndRespond()` from `utils/http.js` for all error responses — never use bare `res.status().json()`.
- **`buildChatBody` in `config/models.js` is the single source of truth** for request shaping. When changing model parameters, update this function and ensure both streaming and sync paths are tested.
- **Do not send `reasoning_effort: 'none'` to the API**. Either omit the parameter or use a valid value (`low`, `medium`, `high`).

## 7) Documentation Guidelines
- **Language**: All documentation files (`.md`) must be written in **English**, not French. This includes `README.md`, `DESIGN_REVIEW.md`, `UX_REVIEW.md`, `SKILLS_AUDIT.md`, `agents.md`, and any future documentation files.
- When updating `README.md`:
  - Prefer **incremental updates** to existing sections.
  - Only modify outdated statements.
  - Keep deployment and environment variable sections accurate.
  - Reflect implemented capabilities without speculative roadmap edits.
  - **Keep the model tiers table in sync** with `backend/src/config/models.js` and `backend/.env.example`.

## 8) PowerPoint Agent
- **Persona**: Expert in visual communication and public speaking.
- **Style**: Concise, punchy, structured (bullet points), slide-oriented.
- **System prompt basis**: "Tu es un expert en présentations PowerPoint. Ton but est d'aider l'utilisateur à créer des diapositives claires et percutantes. Tu privilégies les listes à puces, les phrases courtes et le style direct. Tu peux aussi rédiger des notes pour l'orateur qui sont, à l'inverse, conversationnelles et engageantes."
- **API layer**: Uses the Office Common API (`Office.context.document.getSelectedDataAsync` / `setSelectedDataAsync`) with `CoercionType.Text`. Unlike Word or Excel, PowerPoint has no host-specific `run()` context. Text interaction is limited to the active text selection within a shape.
- **Quick actions**: Bullets, Speaker Notes, Impact (Punchify), Shrink, Visual (draft mode).

## 9) Known Issues to Watch For
- **C7 (CRITICAL)**: Chat in Word is broken. `reasoning_effort: 'none'` sent with tools causes empty model responses. Fix: omit `reasoning_effort` when value is `'none'` in `buildChatBody`. See DESIGN_REVIEW.md.
- **H6**: Agent loop exits silently on empty model response. The "⏳ Analyse de la demande..." placeholder stays with no error. Fix: add empty response detection after the loop.
- **H7**: Tool enable/disable toggles in Settings are dead code. `useAgentLoop.ts` ignores `localStorage('enabledTools')`.
- **M6**: Hardcoded French strings in `useAgentLoop.ts:157,204` — should use `t()` i18n keys.
- **M7**: README model tiers table is stale (shows 4 tiers with old model IDs, actual config has 3 tiers with GPT-5.2).

## 10) Validation Checklist Before Commit
- Run at least one relevant build/check command (frontend and/or backend depending on change).
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- Keep commit message clear and scoped.
- If changing templates, manifest generation, ports, or host URLs: regenerate manifests and verify both outputs are updated as expected.
- If changing model parameters in `config/models.js`: test both streaming (quick actions) AND sync (chat) paths.
- If changing tool definitions: verify the tool count stays under `MAX_TOOLS` (default 128).

## 11) Commit/PR Quality Bar
- Commit title should describe user-facing impact.
- PR summary should include:
  - what changed,
  - why it changed,
  - how it was validated,
  - any compatibility notes.
