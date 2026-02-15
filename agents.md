# AGENTS Guide for KickOffice

This document provides operational guidance for AI coding agents working in this repository.

## 1) Scope
- This guide applies to the whole repository unless a more specific guide exists in a nested directory.
- Follow system/developer/user instructions first when conflicts occur.

## 2) Product and Architecture Snapshot
KickOffice is a Microsoft Office add-in with:
- `frontend/` (Vue 3 + Vite): task pane UI, chat/agent/image UX, Office.js tool execution.
- `backend/` (Express): secure proxy for model calls and model configuration exposure.
- `manifest.xml`: Office add-in manifest for Word and Excel hosts.

Backend API surface:
- `POST /api/chat` (streaming)
- `POST /api/chat/sync` (agent tool loop)
- `POST /api/image`
- `GET /api/models`
- `GET /health`

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

Any contract change should update both backend and frontend in the same change set.

## 5) Frontend Editing Guidelines
- Keep type names explicit and stable.
- Avoid silent failures in user-facing flows; surface understandable errors.
- Preserve i18n behavior and existing translation keys when possible.
- For Office tool changes, ensure host-specific behavior remains correct (Word vs Excel).

## 6) Backend Editing Guidelines
- Keep proxy logic provider-agnostic for OpenAI-compatible endpoints.
- Maintain clear error forwarding (`status` + `details`) for easier debugging.
- Do not leak API keys or environment secrets in logs or responses.

## 7) Documentation Guidelines
When updating `README.md`:
- Prefer **incremental updates** to existing sections.
- Only modify outdated statements.
- Keep deployment and environment variable sections accurate.
- Reflect implemented capabilities without speculative roadmap edits.

## 8) Validation Checklist Before Commit
- Run at least one relevant build/check command (frontend and/or backend depending on change).
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- Keep commit message clear and scoped.

## 9) Commit/PR Quality Bar
- Commit title should describe user-facing impact.
- PR summary should include:
  - what changed,
  - why it changed,
  - how it was validated,
  - any compatibility notes.
