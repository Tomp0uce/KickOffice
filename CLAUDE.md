# CLAUDE.md — KickOffice

> **KickOffice** — Microsoft Office add-in (Word, Excel, PowerPoint, Outlook) powered by LLM agents.
> This file is read on every turn. Keep it concise and high-signal. Target: < 200 lines.

---

## Language Rule

**All generated content must be in English** — documentation, comments, commit messages, variable names, error messages, PR descriptions, and test labels.
Exception: user-created skill files (`.skill.md`) which the user writes in their language of choice.

---

## Project Overview

KickOffice is a Microsoft Office add-in that provides an AI-powered chat interface inside the Office task pane. It executes tools directly on the active document (Word, Excel, PowerPoint) or email (Outlook) via Office.js APIs.

**Stack:** Vue 3 + TypeScript + Tailwind CSS v4 + Vite (frontend) · Node.js/Express SSE proxy (backend) · LiteLLM (production) / OpenAI (testing) · Vitest (unit) + Playwright (e2e) · Docker Compose on Synology DS416play.

**Architecture:** Composable-based frontend (`useAgentLoop` orchestrates tool execution, `useQuickActions` handles one-click actions), ToolProviderRegistry maps hosts to tool definitions, Express backend proxies LLM calls with rate limiting and auth.

---

## Workflow Orchestration

### 1. Plan Mode Default

- Enter plan mode for ANY non-trivial task (3+ steps or architectural decisions).
- If something goes sideways, STOP and re-plan immediately — don't keep pushing.
- Use plan mode for verification steps, not just building.
- Write detailed specs upfront to reduce ambiguity.

### 2. Subagent Strategy

- Use subagents liberally to keep the main context window clean.
- Offload research, exploration, and parallel analysis to subagents.
- For complex problems, throw more compute at it via subagents.
- One task per subagent for focused execution.

### 3. Self-Improvement Loop

- After ANY correction from the user: update `tasks/lessons.md` with the pattern.
- Write rules for yourself that prevent the same mistake.
- Ruthlessly iterate on these lessons until mistake rate drops.
- Review lessons at session start for relevant project context.

### 4. Verification Before Done

- Never mark a task complete without proving it works.
- Run `npm run build` in `frontend/` for any frontend change.
- Ask yourself: "Would a staff engineer approve this?"
- Run tests, check logs, demonstrate correctness.

### 5. Demand Elegance (Balanced)

- For non-trivial changes: pause and ask "is there a more elegant way?"
- If a fix feels hacky: "Knowing everything I know now, implement the elegant solution."
- Skip this for simple, obvious fixes — don't over-engineer.
- Challenge your own work before presenting it.

### 6. Autonomous Bug Fixing

- When given a bug report: just fix it. Don't ask for hand-holding.
- Read logs, errors, failing tests — then resolve them.
- Zero context switching required from the user.
- Go fix failing CI tests without being told how.

---

## Task Management

1. **Plan First**: Write plan to `tasks/todo.md` with checkable items.
2. **Verify Plan**: Check in before starting implementation.
3. **Track Progress**: Mark items `[x]` complete as you go — do not batch.
4. **Explain Changes**: High-level summary at each step.
5. **Document Results**: Add review section to `tasks/todo.md`.
6. **Capture Lessons**: Update `tasks/lessons.md` after corrections.

---

## Core Principles

- **Simplicity First**: Make every change as simple as possible. Impact minimal code.
- **No Laziness**: Find root causes. No temporary fixes. Senior developer standards.
- **Minimal Impact**: Changes should only touch what's necessary. Avoid introducing bugs.
- **No Sycophancy**: Disagree when warranted. Flag bad ideas directly. Don't just agree.

---

## Decision Checkpoints

Ask **only** when there is genuine ambiguity. Two triggers:
1. **Multiple valid approaches** with meaningfully different trade-offs.
2. **A smarter approach exists** than what was requested.

**Do NOT ask** when there is only one sensible implementation — just do it.

Format: `AskUserQuestion` tool, 1–2 questions, 2–4 options. Recommendation labelled "(Recommandé)".

---

## Code Standards

### Frontend

- **All user-visible strings**: use `t()` / i18n keys — no hardcoded strings.
- **Booleans**: MUST be prefixed with `is` or `has`.
- **Constants**: Use `frontend/src/constants/limits.ts` for all magic numbers.
- **Error handling**: Use `getErrorMessage(error: unknown)` from `common.ts`. Use `error: unknown` (not `error: any`).
- **Logging**: Use `logService.warn/error` from `logger.ts` — never raw `console.warn/error`.
- **Router**: `createMemoryHistory` — avoids URL manipulation conflicts with Office iframe host.
- **LSP**: `typescript-language-server` is active. Use `goToDefinition`, `findReferences`, `hover` before editing. Do not guess type shapes.

### Backend

- Keep proxy logic provider-agnostic.
- Log upstream errors server-side; return sanitized messages to clients.
- Never leak API keys or environment secrets.
- Use `logAndRespond()` from `utils/http.js` for all error responses.
- Preserve timeouts: 300s for standard/reasoning, 180s for image.
- `buildChatBody` in `config/models.js` is the single source of truth for request shaping.
- **CRITICAL**: `reasoning_effort: 'none'` is NOT valid — causes empty responses. See `tasks/lessons.md` L01.

### Test Conventions

- Test files live in `__tests__/` subdirectories next to the module under test.
- Use `vi.mock()` before imports for modules that touch Office.js, DOM, or env vars.
- Mirror the source directory structure: `composables/__tests__/`, `utils/__tests__/`.

---

## Domain Knowledge

- **Hosts**: Word (34 tools), Excel (28), PowerPoint (24), Outlook (9), General (6) — 101 total.
- **Skills**: `.skill.md` files in `frontend/src/skills/` — host skills + Quick Action skills.
- **Manifests**: Generated from `manifests-templates/` + root `.env` via `scripts/generate-manifests.js`. Never hand-edit generated files.
- **Office.js constraints**: PowerPoint has no host-specific `run()` context — uses `Office.context.document` with `CoercionType.Text`. Word eval tools use SES `Compartment` sandbox.
- **Context management**: `tokenManager.ts` prunes messages to fit 1.2M chars. `summarizeOldToolResults()` compresses tool results older than 3 iterations.

---

## Docker & Deployment

**CRITICAL — Synology DS416play (Intel Celeron) compatibility:**
- **MUST use `node:22-slim`** (Debian-based) for backend.
- **MUST use `nginxinc/nginx-unprivileged:stable`** (Debian-based) for frontend.
- **DO NOT use Alpine Linux images** — musl libc + AVX instruction issues.

Both containers run as non-root. Frontend listens on port 8080 internally (`FRONTEND_PORT:8080`).
Node.js engine: `>=20.19.0 || >=22.0.0`. Maintain Node.js 22 in Dockerfiles.

---

## Security Rules

- Never introduce command injection, XSS, SQL injection, or OWASP Top 10 vulnerabilities.
- If insecure code is written, fix it immediately.
- Base64 image payloads MUST pass through `sanitizePayloadForLogs` before logging.
- `eval_*` tools are sandboxed via `officeCodeValidator.ts` + SES `Compartment`. Code is validated before `new Function()` execution.

---

## Git & Collaboration

- Branch naming: `feat/<slug>` or `fix/<slug>`. Claude-initiated: `claude/<slug>-<id>`.
- Commit messages: imperative mood, describe the *why*. In English.
- Confirm before: force-push, branch deletion, PR creation, any action visible to others.

### PR Workflow

Before committing, update docs **in this order**:
1. `PRD.md` — if user-facing features changed (product perspective, no code details)
2. `README.md` — newly developed features; keep tool counts accurate
3. `CHANGELOG.md` — standard entry

---

## Validation Checklist Before Commit

- Run `npm run build` in `frontend/` for any frontend change.
- Verify touched UI flows if applicable.
- Ensure changed docs match actual code behavior.
- If changing templates/manifests/ports: regenerate manifests.
- If changing model parameters in `config/models.js`: test both streaming and sync paths.
- If changing tool definitions: verify count stays under `MAX_TOOLS` (128).
- **Dockerfiles**: Only Debian-based images. Never Alpine.

---

## Documentation Hierarchy

| Layer | Location | Purpose |
|---|---|---|
| Global context | `~/.claude/CLAUDE.md` | Consistent behavior across all projects |
| Project context | `CLAUDE.md` (this file) | KickOffice-specific rules and structure |
| Product spec | `PRD.md` | Full product requirements — source of truth for features |
| Design health | `DESIGN_REVIEW.md` | Open architectural items, scores, fix plan |
| Skills guide | `SKILLS_GUIDE.md` | How to create and modify skill files |
| Task tracking | `tasks/todo.md` | Active work items with checkboxes |
| Lessons learned | `tasks/lessons.md` | Mistake patterns and prevention rules |

---

*Last updated: 2026-03-28*
