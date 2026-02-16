# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope
- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

## 2) Product and Architecture Snapshot
KickOffice is a Microsoft Office add-in with:
- `frontend/` (Vue 3 + Vite): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express): secure proxy for model calls and model configuration exposure.
- `manifests-templates/`: manifest templates used to generate runtime manifests.
- `scripts/generate-manifests.js`: script that generates `manifest-office.xml` and `manifest-outlook.xml` from root `.env`.

Manifest outputs:
- `manifest-office.xml`: Word + Excel + PowerPoint TaskPane add-in.
- `manifest-outlook.xml`: Outlook Mail add-in.

Important manifest rule:
- Do not hand-edit generated manifest files when URL/host values change. Update templates and/or `.env`, then regenerate.

Backend API surface:
- `POST /api/chat` (streaming)
- `POST /api/chat/sync` (agent tool loop)
- `POST /api/image`
- `GET /api/models`
- `GET /health`

Operational backend behavior to preserve:
- IP rate limiting is enabled on `/api/chat*` and `/api/image`.
- Server logs include request logging (`morgan`) and sanitized API error responses.
- Upstream provider errors are logged server-side; clients receive generic/safe error messages.

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
- Streaming path consumes SSE-like `data:` lines.
- Sync path expects OpenAI-compatible `choices/message/tool_calls` structures.

### Model parameter compatibility
- Keep compatibility with OpenAI-compatible providers that may differ on parameter support.
- `ChatGPT-*` model IDs do not support `temperature` / token-limit parameters in current backend logic.
- GPT-5 models use `max_completion_tokens` while non-GPT-5 chat models use `max_tokens`.
- If model plumbing is changed, update backend validation + request shaping + frontend expectations together.

Any contract change should update both backend and frontend in the same change set.

## 5) Frontend Editing Guidelines
- Keep type names explicit and stable.
- Avoid silent failures in user-facing flows; surface understandable errors.
- Preserve i18n behavior and existing translation keys when possible.
- For Office tool changes, ensure host-specific behavior remains correct (Word vs Excel vs PowerPoint vs Outlook).
- Host capability gaps must be handled explicitly in UX (clear message instead of degraded/broken fallback behavior).

Current host/tool landscape (keep in mind for tool/agent changes):
- Word tools: high-count set (with general tools) used in agent mode.
- Excel tools: high-count set (with general tools) used in agent mode.
- PowerPoint tools: focused set for slide editing and speaker notes.
- Outlook tools: mail compose/read helpers.

## 6) Backend Editing Guidelines
- Keep proxy logic provider-agnostic for OpenAI-compatible endpoints.
- Log upstream provider errors on the server, but return sanitized client-facing messages (no raw upstream `details`).
- Do not leak API keys or environment secrets in logs or responses.
- Keep input validation strict for: `messages`, `tools`, `temperature`, `maxTokens`, image prompt/size/quality/count.
- Preserve timeout behavior per endpoint/tier (`/api/chat`, `/api/chat/sync`, `/api/image`).

## 7) Documentation Guidelines
- **Language**: All documentation files (`.md`) must be written in **English**, not French. This includes `README.md`, `DESIGN_REVIEW.md`, `SKILLS_AUDIT.md`, `agents.md`, and any future documentation files.
- When updating `README.md`:
  - Prefer **incremental updates** to existing sections.
  - Only modify outdated statements.
  - Keep deployment and environment variable sections accurate.
  - Reflect implemented capabilities without speculative roadmap edits.

## 8) PowerPoint Agent
- **Persona**: Expert in visual communication and public speaking.
- **Style**: Concise, punchy, structured (bullet points), slide-oriented.
- **System prompt basis**: "Tu es un expert en présentations PowerPoint. Ton but est d'aider l'utilisateur à créer des diapositives claires et percutantes. Tu privilégies les listes à puces, les phrases courtes et le style direct. Tu peux aussi rédiger des notes pour l'orateur qui sont, à l'inverse, conversationnelles et engageantes."
- **API layer**: Uses the Office Common API (`Office.context.document.getSelectedDataAsync` / `setSelectedDataAsync`) with `CoercionType.Text`. Unlike Word or Excel, PowerPoint has no host-specific `run()` context. Text interaction is limited to the active text selection within a shape.
- **Quick actions**: Bullets, Speaker Notes, Impact (Punchify), Shrink, Visual (draft mode).

## 9) Validation Checklist Before Commit
- Run at least one relevant build/check command (frontend and/or backend depending on change).
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- Keep commit message clear and scoped.
- If changing templates, manifest generation, ports, or host URLs: regenerate manifests and verify both outputs are updated as expected.

## 10) Commit/PR Quality Bar
- Commit title should describe user-facing impact.
- PR summary should include:
  - what changed,
  - why it changed,
  - how it was validated,
  - any compatibility notes.
