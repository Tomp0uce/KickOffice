# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope

- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

### Companion documents

| File | Purpose |
| ---- | ------- |
| [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) | DR v12 findings (5 critical, 5 high, 19 medium, 12 low) + deferred items — read before making architectural changes |
| [PRD.md](./PRD.md) | Product Requirements Document — single source of truth for features, user personas, acceptance criteria. **Read before implementing new features.** |
| [SKILLS_GUIDE.md](./SKILLS_GUIDE.md) | How to create and modify skill files for Quick Actions and host guidance |

## 2) Product and Architecture Snapshot

KickOffice is a Microsoft Office add-in with:

- `frontend/` (Vue 3 + Vite + TypeScript): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express, modular): secure proxy for model calls and model configuration exposure.
- `manifests-templates/`: manifest templates used to generate runtime manifests.
- `scripts/generate-manifests.js`: script that generates `manifest-office.xml` and `manifest-outlook.xml` from root `.env`.

### Docker & Deployment Constraints

**CRITICAL Hardware Compatibility Requirement**:

- **MUST use `node:22-slim`** (Debian-based, glibc) for backend
- **MUST use `nginxinc/nginx-unprivileged:stable`** (Debian-based, non-root) for frontend serving
- **DO NOT use Alpine Linux images** — incompatible with older Intel Celeron processors (Synology DS416play) due to musl libc + AVX instruction issues

Both containers run as non-root users. Frontend listens on port 8080 internally (docker-compose maps to `${FRONTEND_PORT}:8080`).

### Frontend architecture

The frontend follows a composable-based architecture:

- **Pages**: `HomePage.vue` (minimal orchestration), `SettingsPage.vue` (decomposed into tab components)
- **Components**: `chat/ChatHeader.vue`, `chat/ChatInput.vue`, `chat/ChatMessageList.vue`, `chat/QuickActionsBar.vue`, `chat/StatsBar.vue`, `chat/ToolCallBlock.vue`, `chat/MarkdownRenderer.vue`, `settings/AccountTab.vue`, `settings/GeneralTab.vue`, `settings/PromptsTab.vue`, `settings/BuiltinPromptsTab.vue`, `settings/ToolsTab.vue`, `settings/FeedbackDialog.vue`, `CustomButton.vue`, `CustomInput.vue`, `SingleSelect.vue`, `SettingCard.vue`, `Message.vue`
- **Composables** (17 files): `useHomePage.ts`, `useHomePageContext.ts`, `useAgentLoop.ts`, `useAgentStream.ts`, `useToolExecutor.ts`, `useLoopDetection.ts`, `useAgentPrompts.ts`, `useOfficeSelection.ts`, `useImageActions.ts`, `useOfficeInsert.ts`, `useHealthCheck.ts`, `useQuickActions.ts`, `useSessionFiles.ts`, `useSessionManager.ts`, `useSessionDB.ts`, `useMessageOrchestration.ts`, `useDocumentUndo.ts`
- **Constants**: `limits.ts` (centralized magic numbers: upload sizes, timeouts, buffer sizes, icon sizes)
- **Router**: `createMemoryHistory` — avoids URL manipulation conflicts with Office iframe host
- **Utils**: `tokenManager.ts`, `wordTools.ts`, `excelTools.ts`, `powerpointTools.ts`, `outlookTools.ts`, `generalTools.ts`, `wordDiffUtils.ts`, `wordTrackChanges.ts`, `wordFormatter.ts`, `toolProviderRegistry.ts`, `officeCodeValidator.ts`, `markdown.ts`, `pptxZipUtils.ts`, `common.ts`, `hostDetection.ts`, `toolStorage.ts`, `richContentPreserver.ts`, `richContextStore.ts`, `credentialCrypto.ts`, `credentialStorage.ts`, `cryptoPolyfill.ts`, `officeAction.ts`, `officeDocumentContext.ts`, `officeOutlook.ts`, `officeRichText.ts`, `sandbox.ts`, `lockdown.ts`, `vfs.ts`, `savedPrompts.ts`
- **Skills**: `frontend/src/skills/` — 5 host skills + 17 Quick Action skills (all `.skill.md` files)

### Backend architecture

```
backend/src/
├── server.js              # Entry point: middleware setup, route mounting
├── config/
│   ├── env.js             # Environment variable loading
│   ├── models.js          # Model tier config, buildChatBody(), isGpt5Model()
│   └── limits.js          # Centralized backend limits (file sizes, etc.)
├── middleware/
│   ├── auth.js            # ensureLlmApiKey, ensureUserCredentials
│   └── validate.js        # Re-exports from validators/ (chatValidator, toolValidator, etc.)
├── routes/
│   ├── chat.js            # POST /api/chat (streaming), POST /api/chat/sync
│   ├── health.js          # GET /health
│   ├── image.js           # POST /api/image
│   ├── models.js          # GET /api/models
│   ├── upload.js          # POST /api/upload
│   ├── files.js           # POST /api/files (proxy to /v1/files)
│   ├── icons.js           # GET /api/icons/search, GET /api/icons/svg/:prefix/:name
│   ├── feedback.js        # POST /api/feedback
│   ├── logs.js            # POST /api/logs (frontend log aggregation)
│   └── plotDigitizer.js   # POST /api/chart-extract
├── services/
│   ├── llmClient.js       # LLM API client (chatCompletion, imageGeneration, RateLimitError)
│   ├── plotDigitizerService.js
│   └── imageStore.js
└── utils/
    ├── http.js            # fetchWithTimeout, logAndRespond, sanitizeErrorText
    ├── logger.js          # Winston logger with daily rotation (console + file)
    └── toolUsageLogger.js # Tool usage metrics tracking
```

Manifest outputs:

- `manifest-office.xml`: Word + Excel + PowerPoint TaskPane add-in.
- `manifest-outlook.xml`: Outlook Mail add-in.

**Important**: Do not hand-edit generated manifest files. Update templates and/or `.env`, then regenerate.

## 3) Working Principles

1. **Preserve behavior unless explicitly requested**.
2. **Prefer minimal, localized diffs** over broad rewrites.
3. **Keep frontend and backend contracts aligned**.
4. **Do not introduce secrets in frontend code**.
5. **Update documentation only where needed**.

## 4) API Contract Rules

### Image responses

Compatible providers may return `data[0].b64_json` or `data[0].url`. Keep support for both.

### Chat responses

- Streaming path (`/api/chat`) consumes SSE `data:` lines. Agent loop primarily uses streaming.
- Sync path (`/api/chat/sync`) expects OpenAI-compatible `choices/message/tool_calls`.

### Model parameter compatibility

- `ChatGPT-*` model IDs do not support `temperature` / token-limit parameters (`isChatGptModel` check).
- GPT-5 models use `max_completion_tokens`; non-GPT-5 use `max_tokens` (`isGpt5Model` check).
- **CRITICAL**: `reasoning_effort: 'none'` is NOT a valid OpenAI API value — causes empty responses. Valid values: `'low'`, `'medium'`, `'high'`. Omit entirely when not needed.
- `buildChatBody` in `config/models.js` is the single source of truth for request shaping.

## 5) Frontend Editing Guidelines

- Keep type names explicit and stable.
- Avoid silent failures; surface understandable errors.
- **All user-visible strings must use `t()` / i18n keys** — no hardcoded French or English strings.
- **Naming**: Booleans MUST be prefixed with `is` or `has`.
- **Constants**: Use `frontend/src/constants/limits.ts` for all magic numbers. Use `ICON_SIZE_SM` (14), `ICON_SIZE_MD` (16), `ICON_SIZE_LG` (20) for icon sizes.
- **Error handling**: Use `getErrorMessage(error: unknown)` from `common.ts` in all catch blocks. Use `error: unknown` (not `error: any`) in catch clauses.
- **Logging**: Use `logService.warn/error` from `logger.ts` — never raw `console.warn/error` in production code.

Current tool landscape:

- Word: 31 tools (`proposeRevision`, `proposeDocumentRevision`, `editDocumentXml`, `eval_wordjs`, `getDocumentOoxml`, and 26 more)
- Excel: 27 tools (`eval_officejs`, `screenshotRange`, `getRangeAsCsv`, `detectDataHeaders`, `modifyWorkbookStructure`, and 22 more)
- PowerPoint: 23 tools (`screenshotSlide`, `editSlideXml`, `searchIcons`, `insertIcon`, `eval_powerpointjs`, `verifySlides`, and 17 more)
- Outlook: 8 tools (`eval_outlookjs` and 7 more)
- General: 6 tools (`executeBash` VFS, `calculateMath`, `getCurrentDate`, file operations)
- **Total**: 95 tools

**Agent Stability Features**:

- **Skills System**: `.skill.md` files in `frontend/src/skills/` — 5 host skills + 17 Quick Action skills loaded via `getQuickActionSkill()` / host skill loaders in `useAgentPrompts.ts`
- **Code Validator**: Pre-execution validation for `eval_*` tools in `officeCodeValidator.ts`
- **Track Changes**: `proposeRevision` (selection) and `proposeDocumentRevision` (document-wide) both use `applyRedlineToOxml()` from `@ansonlai/docx-redline-js`. Pattern: disable TC → `insertOoxml` with `w:ins`/`w:del` → restore TC.
- **ToolProviderRegistry**: Singleton in `toolProviderRegistry.ts`. Maps host names to tool providers. Auto-initializes with ES6 static imports at module load.
- **Context Management**: `tokenManager.ts` — `prepareMessagesForContext()` prunes messages to fit within 1.2M chars. `summarizeOldToolResults()` (Phase 7A) compresses tool results from iterations older than the last 3, keeping recent context intact.
- **Session Persistence**: `useSessionFiles.ts` manages uploaded files. Each `DisplayMessage` stores `attachedFiles?: Array<{filename, content, fileId?}>`. Call `rebuildSessionFiles()` after `history` replace from IndexedDB.
- **Log Sanitization**: Payloads with Base64 images MUST pass through `sanitizePayloadForLogs` before logging.

## 6) Backend Editing Guidelines

- Keep proxy logic provider-agnostic.
- Log upstream errors server-side; return sanitized messages to clients.
- Never leak API keys or environment secrets.
- Validation is in `middleware/validate.js` (delegates to `validators/` domain files).
- Preserve timeouts: 300s for standard/reasoning, 180s for image.
- `RateLimitError` from `llmClient.js`: when retries exhausted on 429, thrown with `Retry-After` parsed. Both chat routes catch it → `429 RATE_LIMITED`.
- Use `logAndRespond()` from `utils/http.js` for all error responses.
- **Do not send `reasoning_effort: 'none'`** — omit or use `low`/`medium`/`high`.
- **CI**: `.github/workflows/bump-version.yml` uses `permissions: contents: write` for version bump pushes.
- **Node.js**: `"engines": { "node": ">=20.19.0 || >=22.0.0" }`. Maintain Node.js 22 in Dockerfiles.

## 7) Documentation Guidelines

- **Language**: All `.md` files must be written in **English**.
- When updating `README.md`: prefer incremental updates, keep tool counts accurate, reflect implemented capabilities only (no roadmap).
- Keep the model tiers table in sync with `backend/src/config/models.js` and `backend/.env.example`.

## 8) Product Requirements Document (PRD.md) Guidelines

**Update PRD.md when the change is significant enough that a product manager would want to know about it.**

### When to update PRD.md ✅

- New tool categories or major new capabilities
- New external integrations or services
- New user-facing workflows or multi-step interaction patterns
- Changes to supported formats, file size limits, or platform support
- New Quick Actions or agent behaviors visible to the end user

### When NOT to update PRD.md ❌

- Bug fixes, error handling, stability fixes
- Performance optimizations with no user-facing behavior change
- Internal refactoring, code cleanup, architectural changes
- Adding optional parameters without changing core function
- UI micro-adjustments, documentation-only changes, developer experience changes

### Content guidelines

- Focus on **what** the product does from a **user perspective**, not implementation details
- No code snippets, library names, file paths, or function names in PRD
- **Language**: English only

## 9) PowerPoint Agent

- **Persona**: Expert in visual communication and public speaking.
- **Style**: Concise, punchy, structured (bullet points), slide-oriented.
- **API layer**: Uses Office Common API (`Office.context.document`) with `CoercionType.Text`. No host-specific `run()` context in PowerPoint.
- **Quick actions**: Bullets, Review (slide feedback), Impact (Punchify), Visual (draft mode).

## 10) Known Issues

Consult [DESIGN_REVIEW.md](./DESIGN_REVIEW.md) for the current deferred items list. All critical/high/medium items are resolved.

## 11) Validation Checklist Before Commit

- Run `npm run build` in `frontend/` for any frontend change.
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- If changing templates, manifest generation, ports, or host URLs: regenerate manifests.
- If changing model parameters in `config/models.js`: test both streaming and sync paths.
- If changing tool definitions: verify tool count stays under `MAX_TOOLS` (default 128).
- **Dockerfiles**: Only Debian-based images. Never Alpine.

## 12) Commit/PR Quality Bar

- Commit title describes user-facing impact.
- PR summary includes: what changed, why, how validated, compatibility notes.

## 13) Pull Request Workflow

When asked to create a Pull Request:

1. **Update `PRD.md`** if features changed (user perspective only, no implementation details).
2. **Update `README.md`** with newly developed features, concisely.
3. **Update `CHANGELOG.md`** with a standard changelog entry.
4. **Stage & Commit**: `git add <specific files>` then `git commit -m "feat: ..."`
5. **Push**: `git push -u origin <branch_name>`
6. **Create PR**: Use `gh pr create` or the Gitea API if `gh` is unavailable.
7. **Notify the user** with the PR link.

## 14) Strict Agent Instructions

1. **Restricted Scope**: Remain within the current working directory only.
2. **No Outside Access**: Do not read, list, or modify files outside this directory.
3. **Focus**: Ignore system configuration files. Focus solely on the repository source code.

## 15) Error Prevention Rules

1. **Command Chaining**: On Linux/bash, use `&&` for sequential dependent commands.
2. **NPM Scripts**: Do not run `npm run check` or `npm run type-check` unless verified in `package.json`. Use `npx tsc --noEmit` for type checking.
3. **Git**: Never use `--no-verify`, `--force` on shared branches, or `reset --hard` without explicit user instruction. Create new commits rather than amending published ones.

## 16) Interactive Decisions on New or Modified Features

When you are **unsure how to implement** a new feature or modify an existing one — because there are multiple valid approaches with different trade-offs — **do NOT silently pick one and implement it**. Instead:

1. **Use the `AskUserQuestion` tool** to ask an interactive question.
2. Present **2–4 concrete options**, each with:
   - A short label (1–5 words)
   - A description of the resulting **behavior/UX** that the user will observe (not implementation details)
   - Trade-offs if relevant (e.g., simpler but less precise, richer but more complex)
3. Wait for the user's choice **before writing any code**.

**What to ask about** (non-exhaustive):
- UI placement and visibility (e.g., "Where should the undo button appear?")
- Scope of a feature (e.g., "Should undo cover all operations or only quick actions?")
- Fallback behavior (e.g., "What should happen if the undo state is unavailable?")
- Conflict with existing UX (e.g., "This overlaps with Track Changes — prefer to replace it or complement it?")

**What NOT to ask about**: implementation libraries, function names, file structure, code style — decide those yourself based on the existing codebase patterns.
