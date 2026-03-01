# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope

- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

### Companion documents

| File                                   | Purpose                                                                                                     |
| -------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) | Code audit v3 — 162 issues across backend, frontend, infra (13 CRITICAL, 34 HIGH, 58 MEDIUM, 28 LOW, 29 dead code) |
| [UX_REVIEW.md](./UX_REVIEW.md)         | User experience issues — open items by priority (HIGH/MEDIUM/LOW)                                           |
| [SKILLS_AUDIT.md](./SKILLS_AUDIT.md)   | Tool set audit — current tools per host + proposed additions                                                |
| [REPORT-openexcel-kickoffice-comparison.md](./REPORT-openexcel-kickoffice-comparison.md) | OpenExcel vs KickOffice feature comparison — chat UX, conversation management, tool status, stats bar |

## 2) Product and Architecture Snapshot

KickOffice is a Microsoft Office add-in with:

- `frontend/` (Vue 3 + Vite + TypeScript): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express, modular): secure proxy for model calls and model configuration exposure.
- `manifests-templates/`: manifest templates used to generate runtime manifests.
- `scripts/generate-manifests.js`: script that generates `manifest-office.xml` and `manifest-outlook.xml` from root `.env`.

### Frontend architecture (post-refactor)

The frontend follows a composable-based architecture:

- **Pages**: `HomePage.vue` (orchestration, history persistence per host), `SettingsPage.vue` (settings UI, feature toggles)
- **Components**: `chat/ChatHeader.vue`, `chat/ChatInput.vue`, `chat/ChatMessageList.vue`, `chat/QuickActionsBar.vue` + generic components (`CustomButton`, `CustomInput`, `Message`, `SettingCard`, `SingleSelect`)
- **Composables**: `useAgentLoop.ts` (agent execution loop with recursive context gathering and dynamic tool filtering), `useAgentPrompts.ts` (prompts array generation), `useOfficeSelection.ts` (Office API text selection), `useImageActions.ts` (image processing), `useOfficeInsert.ts` (document insertion + clipboard)
- **Utils**: `tokenManager.ts` (context pruning and LLM token budget management), `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`, `generalTools.ts`, `wordFormatter.ts`, `constant.ts` (translation-aware built-in prompts), `hostDetection.ts`, `message.ts`, `common.ts`, `enum.ts`, `savedPrompts.ts`, `toolStorage.ts`, `officeOutlook.ts`, `officeAction.ts`, `markdown.ts`

### Backend architecture (post-refactor)

```
backend/src/
├── server.js              # Entry point: middleware setup, route mounting
├── config/
│   ├── env.js             # Environment variable loading (PORT, FRONTEND_URL, rate limits)
│   └── models.js          # Model tier config, buildChatBody(), isGpt5Model(), isChatGptModel()
├── middleware/
│   ├── auth.js            # ensureLlmApiKey, ensureUserCredentials — API key + user credential checks
│   └── validate.js        # validateChatRequest, validateTemperature, validateMaxTokens, validateTools, validateImagePayload
├── routes/
│   ├── chat.js            # POST /api/chat (streaming, supports tools), POST /api/chat/sync (synchronous fallback)
│   ├── health.js          # GET /health
│   ├── image.js           # POST /api/image
│   ├── models.js          # GET /api/models
│   └── upload.js          # POST /api/upload (file processing: PDF, DOCX, XLSX, CSV)
├── services/
│   └── llmClient.js       # Centralized LLM API client (chatCompletion, imageGeneration, handleErrorResponse)
└── utils/
    └── http.js            # fetchWithTimeout, logAndRespond, sanitizeErrorText
```

Manifest outputs:

- `manifest-office.xml`: Word + Excel + PowerPoint TaskPane add-in.
- `manifest-outlook.xml`: Outlook Mail add-in.

Important manifest rule:

- Do not hand-edit generated manifest files when URL/host values change. Update templates and/or `.env`, then regenerate.

Backend API surface:

- `POST /api/chat` (streaming) — uses `buildChatBody` with `stream: true`, supports tools via `onToolCallDelta` SSE deltas
- `POST /api/chat/sync` (synchronous) — uses `buildChatBody` with `stream: false`, tools + tool_choice (kept as fallback; primary agent loop now uses streaming)
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

- Streaming path (`/api/chat`) consumes SSE-like `data:` lines with buffer-based parsing (`backend.ts:131-139`). The agent loop primarily relies on streaming (even for tool calls).
- Sync path (`/api/chat/sync`) expects OpenAI-compatible `choices/message/tool_calls` structures. Reserved for deeply nested sequential synchronous logic if required.

### Model parameter compatibility

- Keep compatibility with OpenAI-compatible providers that may differ on parameter support.
- `ChatGPT-*` model IDs do not support `temperature` / token-limit parameters in current backend logic (`isChatGptModel` check).
- GPT-5 models use `max_completion_tokens` while non-GPT-5 chat models use `max_tokens` (`isGpt5Model` check).
- **CRITICAL**: `reasoning_effort` parameter behavior varies by model. The value `'none'` is NOT a valid OpenAI API value and causes empty responses when used with tools. Valid values are `'low'`, `'medium'`, `'high'`. When reasoning is not needed, **omit the parameter entirely** rather than sending `'none'`. See DESIGN_REVIEW.md v2 C2 for the `.env.example` issue.
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

- Word tools: 40 tools (high-count set, including `eval_wordjs` via SES sandbox) + 2 general tools.
- Excel tools: 45 tools (high-count set, including `eval_officejs` and OpenExcel ports) + 2 general tools.
- PowerPoint tools: 15 tools (slides, shapes, notes, modify, and `eval_powerpointjs`) + 2 general tools.
- Outlook tools: 14 tools (mail compose/read helpers and `eval_outlookjs`) + 2 general tools.

## 6) Backend Editing Guidelines

- Keep proxy logic provider-agnostic for OpenAI-compatible endpoints.
- Log upstream provider errors on the server, but return sanitized client-facing messages (no raw upstream `details`).
- Do not leak API keys or environment secrets in logs or responses.
- Keep input validation strict for: `messages`, `tools`, `temperature`, `maxTokens`, image prompt/size/quality/count. All validation is in `middleware/validate.js`.
- Preserve timeout behavior per endpoint/tier (120s for standard, 300s for reasoning).
- Use `logAndRespond()` from `utils/http.js` for all error responses — never use bare `res.status().json()`.
- **`buildChatBody` in `config/models.js` is the single source of truth** for request shaping. When changing model parameters, update this function and ensure both streaming and sync paths are tested.
- **Do not send `reasoning_effort: 'none'` to the API**. Either omit the parameter or use a valid value (`low`, `medium`, `high`).
- **CI Pipelines**: The `.github/workflows/bump-version.yml` action relies on `permissions: contents: write` to allow the GitHub Actions bot to push version bumps back to the repository. No manual configuration in the repository settings is needed as the workflow file overrides the default read-only token permissions explicitly.

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

For the most up-to-date list of known issues, architectural gaps, and proposed fixes, always refer to the [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) (v2) file. This document serves as the single source of truth for technical debt and active blocking issues. Key issues to be aware of:

- **C1**: Agent max iterations setting is silently capped at 10 regardless of user setting.
- **C2**: `.env.example` contains invalid `reasoning_effort=none` which breaks tool use with GPT-5 models.
- **C3**: Quick actions bypass loading/abort state, preventing stop and risking history corruption.

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

## 12) Pull Request Workflow

When the user asks to create a Pull Request, **always follow this exact sequence**:

1. **Update `README.md`**: Add the newly developed features directly to the README. Be concise, list the features clearly, and avoid detailing the underlying code implementation.
2. **Update `CHANGELOG.md`**: Add an entry for the new version/changes. Use a standard changelog format. Clearly outline the user-facing changes (features, fixes, improvements) without deep technical implementation details. Include the link to the PR if known, or mention the PR branch/name.
3. **Stage & Commit**:
   - Use `git add .` to stage changes.
   - Use `git commit -m "feat: your concise commit message"`
   - **Important**: If chaining commands, use PowerShell syntax `;` (e.g., `git add .; git commit -m "..."`), NOT `&&`.
4. **Push to Remote**: `git push origin <branch_name>`
5. **Draft the PR**:
   - Create a temporary markdown file (e.g., `.github/pr_body.md`) containing the detailed PR body outlining the changes. This avoids multi-line string escaping errors in PowerShell.
   - Run the GitHub CLI command: `gh pr create --title "feat: Your PR Title" --body-file .github/pr_body.md`
6. **Notify the User**: Confirm the PR creation and provide the generated link.
