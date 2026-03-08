# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope

- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

### Companion documents

| File                                                                                     | Purpose                                                                                                                                                                              |
| ---------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [DESIGN_REVIEW.md](./DESIGN_REVIEW.md)                                                   | Code audit — issues by severity (CRITICAL/HIGH/MEDIUM/LOW) with status tracking                                                                                                      |
| [UX_REVIEW.md](./UX_REVIEW.md)                                                           | User experience issues — open items by priority                                                                                                                                      |
| [SKILLS_AUDIT.md](./SKILLS_AUDIT.md)                                                     | Tool set audit — current tools per host + proposed additions                                                                                                                         |
| [AGENT_MODE_ANALYSIS.md](./AGENT_MODE_ANALYSIS.md)                                       | Agent execution mode analysis — streaming vs sync performance comparison                                                                                                             |
| [INTEGRATION_PLAN.md](./INTEGRATION_PLAN.md)                                             | Technical integration roadmap and implementation strategy                                                                                                                            |
| [REPORT-openexcel-kickoffice-comparison.md](./REPORT-openexcel-kickoffice-comparison.md) | OpenExcel vs KickOffice feature comparison — chat UX, conversation management, tool status, stats bar                                                                                |
| [PRD.md](./PRD.md)                                                                       | Product Requirements Document — Single source of truth for product features, user personas, acceptance criteria, and business logic. **Read this before implementing new features.** |

## 2) Product and Architecture Snapshot

KickOffice is a Microsoft Office add-in with:

- `frontend/` (Vue 3 + Vite + TypeScript): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express, modular): secure proxy for model calls and model configuration exposure.
- `manifests-templates/`: manifest templates used to generate runtime manifests.
- `scripts/generate-manifests.js`: script that generates `manifest-office.xml` and `manifest-outlook.xml` from root `.env`.

### Docker & Deployment Constraints

**CRITICAL Hardware Compatibility Requirement**:

- **MUST use `node:22-slim`** (Debian-based, glibc) for all Node.js containers
- **MUST use `nginx:stable`** (Debian-based) for frontend serving
- **DO NOT use Alpine Linux images** (`node:*-alpine`, `nginx:alpine`) — incompatible with older Intel Celeron processors (Synology DS416play) due to musl libc + AVX instruction issues

See DESIGN_REVIEW.md §C0c for full context. This constraint is already enforced in `backend/Dockerfile`, `frontend/Dockerfile`, and `docker-compose.yml`.

### Frontend architecture (post-refactor)

The frontend follows a composable-based architecture:

- **Pages**: `HomePage.vue` (now minimal orchestration, delegates logic to `useHomePage.ts`), `SettingsPage.vue` (decomposed into tab components)
- **Components**: `chat/ChatHeader.vue`, `chat/ChatInput.vue`, `chat/ChatMessageList.vue`, `chat/QuickActionsBar.vue` + generic components (`CustomButton`, `CustomInput`, `Message`, `SettingCard`, `SingleSelect`) + Settings tabs (`AccountTab`, `GeneralTab`, etc.)
- **Composables**: `useHomePage.ts` (orchestrates homepage state, session switching, scroll management), `useAgentLoop.ts` (modularized agent loop), `useAgentStream`, `useToolExecutor`, `useLoopDetection`, `useAgentPrompts.ts`, `useOfficeSelection.ts` (Office API text selection), `useImageActions.ts` (image processing), `useOfficeInsert.ts` (document insertion + clipboard), `useHealthCheck.ts`
- **Constants**: `limits.ts` (centralized magic numbers: upload sizes, timeouts, buffer sizes, icon sizes)
- **Utils**: `tokenManager.ts` (context pruning), `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`, `generalTools.ts`, `wordFormatter.ts`, `hostDetection.ts`, `message.ts`, `common.ts` (normalizeLineEndings), `enum.ts`, `officeCodeValidator.ts` (pre-execution safety), `markdown.ts` (deduplicated MarkdownIt config)

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
- **CRITICAL**: `reasoning_effort` parameter behavior varies by model. The value `'none'` is NOT a valid OpenAI API value and causes empty responses when used with tools. Valid values are `'low'`, `'medium'`, `'high'`. When reasoning is not needed, **omit the parameter entirely** rather than sending `'none'`. See DESIGN_REVIEW.md for historical context.
- If model plumbing is changed, update backend validation + request shaping + frontend expectations together.

Any contract change should update both backend and frontend in the same change set.

## 5) Frontend Editing Guidelines

- Keep type names explicit and stable.
- Avoid silent failures in user-facing flows; surface understandable errors.
- Preserve i18n behavior and existing translation keys when possible.
- **All user-visible strings must use `t()` / i18n keys** — no hardcoded French or English strings in components or composables.
- **Naming Conventions**: Booleans MUST be prefixed with `is` or `has` (e.g., `isLoading`, `isDraftFocusGlowing`, `hasSelection`).
- **Constants Management**: Avoid "magic numbers". Use `frontend/src/constants/limits.ts` and `backend/src/config/limits.js` for all configuration values (timeouts, sizes, counts).
- **Icon Sizing**: Use `ICON_SIZE_SM` (14), `ICON_SIZE_MD` (16), or `ICON_SIZE_LG` (20) from `limits.ts`.
- For Office tool changes, ensure host-specific behavior remains correct (Word vs Excel vs PowerPoint vs Outlook).

Current host/tool landscape (keep in mind for tool/agent changes):

- Word tools: 41 tools (includes `proposeRevision` for format-preserving edits, `eval_wordjs` via SES sandbox)
- Excel tools: 45 tools (high-count set, including `eval_officejs` and OpenExcel ports)
- PowerPoint tools: 16 tools (includes `proposeShapeTextRevision` for diff reporting, slides, shapes, notes, and `eval_powerpointjs`)
- Outlook tools: 14 tools (mail compose/read helpers and `eval_outlookjs`)
- General tools: 6 tools (`executeBash` VFS, `calculateMath`, `getCurrentDate`, file operations)
- **Total**: 129 tools across all hosts

**Agent Stability Features** (implemented):

- **Skills System**: Office.js best practices auto-injected into agent prompts via `frontend/src/skills/` (5 markdown skill documents: common.skill.md + host-specific for Word/Excel/PowerPoint/Outlook)
- **Code Validator**: Pre-execution validation for all `eval_*` tools via `officeCodeValidator.ts` (catches missing load/sync, wrong namespaces, infinite loops)
- **Diffing Integration**: Format-preserving text editing via `wordDiffUtils.ts` (wraps `office-word-diff` local library — Token Map → Sentence Diff → Block Replace cascade). `proposeRevision` is the only Word tool that preserves formatting on unchanged text; agent is instructed to always call `getSelectedTextWithFormatting` first, and is explicitly forbidden from using `eval_wordjs` with `insertText(..., 'Replace')` which destroys formatting. `office-word-diff` is Word-only (PowerPoint/Outlook lack the Range API and Track Changes).
- **Vision Image Handling (Registry Pattern)**: Large Base64 image data MUST BE decoupled from the agent prompt. Use `powerpointImageRegistry` (Map) in `powerpointTools.ts` to store raw Base64. Agent tools (like `insertImageOnSlide`) MUST accept a `filename` parameter and retrieve the data from the registry.
- **Session Persistence**: Uploaded files and images MUST be persisted in `sessionUploadedFiles` and `sessionUploadedImages` refs in `useAgentLoop.ts`. Their context (extracted text for files, filename list for images) MUST be re-injected into every message to prevent amnesia.
- **Log Sanitization**: All payloads containing potential Base64 images MUST pass through `sanitizePayloadForLogs` in `backend.ts` before being logged to prevent disk saturation and terminal crashes.
- **Token Manager**: Images in multi-part content MUST have a fixed token cost (default 1000) in `tokenManager.ts` to prevent the manager from truncating history due to Base64 length.
- **Sandbox Enhancement**: Host filtering in `sandbox.ts` prevents cross-namespace API access (e.g., Word API blocked in Excel context)

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
- **Node.js Version Constraint**: Both `backend/package.json` and `frontend/package.json` specify `"engines": { "node": ">=20.19.0 || >=22.0.0" }`. When updating Dockerfiles or CI configuration, maintain Node.js 22 compatibility.
- **Docker Image Constraint**: All Dockerfiles MUST use Debian-based images (`node:22-slim`, `nginx:stable`). DO NOT use Alpine variants due to hardware incompatibility with older Intel Celeron processors (musl libc + AVX issues). See §2 and DESIGN_REVIEW.md §C0c.

## 7) Documentation Guidelines

- **Language**: All documentation files (`.md`) must be written in **English**, not French. This includes `README.md`, `DESIGN_REVIEW.md`, `UX_REVIEW.md`, `SKILLS_AUDIT.md`, `agents.md`, and any future documentation files.
- When updating `README.md`:
  - Prefer **incremental updates** to existing sections.
  - Only modify outdated statements.
  - Keep deployment and environment variable sections accurate.
  - Reflect implemented capabilities without speculative roadmap edits.
  - **Keep the model tiers table in sync** with `backend/src/config/models.js` and `backend/.env.example`.

## 8) Product Requirements Document (PRD.md) Guidelines

**CRITICAL: Always update PRD.md when adding or modifying features.**

- **When to update**: Every time you add, modify, or remove a user-facing feature, update `PRD.md` to reflect the change.
- **What to include**:
  - Feature descriptions from a **user perspective**
  - Business logic and acceptance criteria
  - User personas and use cases
  - Constraints and limitations
- **What NOT to include**:
  - Technical implementation details
  - Code snippets or code examples
  - Library names, framework names, or specific technologies
  - File paths, function names, or class names
  - Architecture diagrams or technical workflows
- **Keep it high-level**: The PRD is a product document, not a technical specification. Focus on **what** the product does and **why**, not **how** it's implemented.
- **Single source of truth**: The PRD should be the authoritative reference for product features. When implementing new features, read the PRD first to understand the requirements.
- **Language**: PRD.md must be written in **English**.

## 9) PowerPoint Agent

- **Persona**: Expert in visual communication and public speaking.
- **Style**: Concise, punchy, structured (bullet points), slide-oriented.
- **System prompt basis**: "You are an expert in PowerPoint presentations. Your goal is to help the user create clear and impactful slides. You favor bullet points, short sentences, and direct style. You can also write speaker notes that are, in contrast, conversational and engaging."
- **API layer**: Uses the Office Common API (`Office.context.document.getSelectedDataAsync` / `setSelectedDataAsync`) with `CoercionType.Text`. Unlike Word or Excel, PowerPoint has no host-specific `run()` context. Text interaction is limited to the active text selection within a shape.
- **Quick actions**: Bullets, Speaker Notes, Impact (Punchify), Shrink, Visual (draft mode).

## 10) Known Issues to Watch For

**ALWAYS consult [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) before making architectural changes.**

This document is the single source of truth for:

- Issue inventory (CRITICAL/HIGH/MEDIUM/LOW severity)
- Fix status tracking (✅ FIXED / open)
- Architectural gaps and technical debt
- Recommended remediation strategies

Do not duplicate issue tracking here — refer to the authoritative document.

## 11) Validation Checklist Before Commit

- Run at least one relevant build/check command (frontend and/or backend depending on change).
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- Keep commit message clear and scoped.
- If changing templates, manifest generation, ports, or host URLs: regenerate manifests and verify both outputs are updated as expected.
- If changing model parameters in `config/models.js`: test both streaming (quick actions) AND sync (chat) paths.
- If changing tool definitions: verify the tool count stays under `MAX_TOOLS` (default 128).
- **If modifying Dockerfiles**: Ensure only Debian-based images are used (`node:22-slim`, `nginx:stable`). Never introduce Alpine variants (`node:*-alpine`, `nginx:alpine`). See §2 for hardware compatibility requirements.

## 12) Commit/PR Quality Bar

- Commit title should describe user-facing impact.
- PR summary should include:
  - what changed,
  - why it changed,
  - how it was validated,
  - any compatibility notes.

## 13) Pull Request Workflow

When the user asks to create a Pull Request, **always follow this exact sequence**:

1. **Update `PRD.md`**: If the changes include new or modified features, update the PRD to reflect these changes from a **user perspective** (no technical implementation details).
2. **Update `README.md`**: Add the newly developed features directly to the README. Be concise, list the features clearly, and avoid detailing the underlying code implementation.
3. **Update `CHANGELOG.md`**: Add an entry for the new version/changes. Use a standard changelog format. Clearly outline the user-facing changes (features, fixes, improvements) without deep technical implementation details. Include the link to the PR if known, or mention the PR branch/name.
4. **Stage & Commit**:
   - Use `git add .` to stage changes.
   - Use `git commit -m "feat: your concise commit message"`
   - **Important**: If chaining commands, use PowerShell syntax `;` (e.g., `git add .; git commit -m "..."`), NOT `&&`.
5. **Push to Remote**: `git push origin <branch_name>`
6. **Draft the PR**:
   - Create a temporary markdown file (e.g., `.github/pr_body.md`) containing the detailed PR body outlining the changes. This avoids multi-line string escaping errors in PowerShell.
   - Run the GitHub CLI command: `gh pr create --title "feat: Your PR Title" --body-file .github/pr_body.md`
7. **Notify the User**: Confirm the PR creation and provide the generated link.

## 14) Strict Agent Instructions

1. **Restricted Scope:** You must STRICTLY remain within the current working directory (kickoffice).
2. **No Outside Access:** You are strictly forbidden from reading, listing, or modifying files outside of this directory (for example, absolutely no reading of `~/.claude/settings.json` or any other file in `~/`).
3. **Focus:** Ignore your own system configuration settings when you review code. Focus solely on the source code and files present in this repository.

## 15) Vibe Coding Rules (Error Prevention)

To prevent common errors during vibe coding and automated command execution, **STRICTLY ADHERE** to the following rules:

1. **PowerShell Command Chaining:**
   - **NEVER** use `&&` to chain commands in PowerShell. It is not a valid statement separator.
   - **ALWAYS** use `;` to chain commands (e.g., `npm run build ; npm run check`).
2. **NPM Scripts:**
   - **DO NOT** run `npm run check` or `npm run type-check` or `npm run typecheck` unless you have explicitly verified that the script exists in `package.json`.
   - If you need to verify types and no script is available, run `npx tsc --noEmit` directly.
