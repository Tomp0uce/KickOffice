# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added

- **Frontend Logging System**: Secure collection of frontend logs (warn/error) sent to the backend (`POST /api/logs`) and stored in server-side files.
- **Persistent Log Storage**: Local logs are now persisted in IndexedDB (`useSessionDB.ts`) to survive taskpane reloads.
- **Request Tracing**: Added `X-Request-Id` header to all API requests, captured in both frontend and backend logs for full-stack debugging.
- **Header Caching**: Implemented a Promise-based cache for global headers in `backend.ts` with explicit invalidation on credential changes, reducing storage access overhead.
- **Centralized Constants**: Created `limits.ts` (frontend) and `limits.js` (backend) to eliminate magic numbers for timeouts, buffer sizes, and limits.
- **Modular Settings**: Refactored `SettingsPage.vue` into 5 independent tab components for better maintainability and performance.
- **`useHomePage` Composable**: Extracted business logic, session management, and navigation from `HomePage.vue`.
- **Skeleton Loaders**: Added animated skeleton states for Settings page loading.
- **Accessibility Enhancements**: Added `aria-label` to all interactive elements, implemented focus traps in dialogues, and unified `:focus-visible` styles.
- **Keyboard Shortcut Hint**: Added a visible hint for `Shift + Enter` (new line) in the chat input.

### Fixed

- **Taskpane Width**: Increased default `RequestedWidth` to 450px in manifest templates for better readability of indicators and settings.
- **Word Insertion Logic**: Consolidated duplicated insertion logic between `wordApi.ts` and `WordFormatter.ts`.
- **Line Ending Normalization**: Centralized newline normalization into a single `common.ts` helper.
- **Markdown Config**: Deduplicated `MarkdownIt` initialization into a reusable helper.
- **Dead Code Removal**: Deleted 4 legacy Python scripts, unused imports in `wordTools.ts`, and redundant dependencies.
- **Boolean Naming**: Standardized boolean variables with `is*` and `has*` prefixes across the codebase.
- **Feedback Writing**: Improved reliability of the feedback submission route with atomic file writes.

### [1.0.112] - 2026-03-08

### Added

- **Message Timestamps**: Chat messages now display the creation time (HH:MM format) for better conversation context tracking

---

## [Previous Unreleased v4]

### Added

- **Complete Documentation Overhaul**:
  - **README.md**: Full rewrite with accurate project overview, architecture diagram, model tiers, agent stability system, tool summary (129 tools), and proper credits for inspiration projects (word-GPT-Plus, excel-ai-assistant, office-word-diff, Redink)
  - **CLAUDE.md**: Streamlined agent guide with architecture quick reference, working principles, critical rules, and commit/PR workflow
  - **DESIGN_REVIEW.md v4**: Complete fresh audit with 43 issues (10 CRITICAL, 9 HIGH, 12 MEDIUM, 8 LOW, 4 DEAD CODE), including new Docker build issues and dead code analysis

### Fixed

- **Security Hardening (10 Critical Issues from DESIGN_REVIEW v4)**:
  - **CSRF Protection**: Added explicit origin validation to CSRF middleware. POST/PUT/PATCH/DELETE requests without X-User-Key header now require valid origin from allowlist
  - **Credential Encryption**: Replaced XOR obfuscation with Web Crypto API (AES-GCM 256-bit encryption) for credentials stored in localStorage. SessionStorage credentials remain unencrypted (session-only)
  - **Rate Limiting**: Added IP-based rate limiter to `/api/upload` endpoint (10 uploads/min) to prevent memory exhaustion DoS attacks
  - **Stream Abort Handling**: Fixed hanging requests and resource leaks in streaming chat endpoint. Now properly calls `reader.cancel()` on client disconnect, adds 30s read timeout, and handles write errors after disconnection
  - **Safe JSON Stringify**: Added `safeStringify()` function with depth validation (max 10 levels) and circular reference detection to prevent DoS via deeply nested tool arguments
  - **Agent Iteration Limit**: Enforced explicit iteration count check in agent loop. User-configured `agentMaxIterations` setting is now respected instead of only timeout-based enforcement
  - **Quick Actions Loading State**: Quick actions now check `loading.value` and `abortController` before execution to prevent duplicate requests and history corruption
  - **Docker npm ci Compatibility**: Changed `npm ci` to `npm install` in frontend Dockerfile for better compatibility with local file dependencies on Synology NAS
- **Docker Build Failure**: Fixed critical Docker build issues preventing deployment on Synology NAS:
  - Changed frontend build context from `./frontend` to root (`.`) to include `office-word-diff` dependency
  - Restructured `frontend/Dockerfile` to properly copy local `office-word-diff` library
  - Added `office-word-diff` to `package-lock.json` for `npm ci` compatibility
- **Synology DS416play Compatibility**: Alpine Linux images cause "Illegal instruction" errors on Celeron CPUs (musl libc uses AVX instructions not supported by older processors):
  - Changed `node:22-alpine` to `node:22-slim` (Debian/glibc)
  - Changed `nginx:alpine` to `nginx:stable` (Debian-based)

---

## [Previous Unreleased]

### Added

- **Agent Stability System (Three Pillars)**:
  - **Skills System (Pillar 2)**: Office.js best practices automatically injected into agent prompts. Five skill documents (common + Word/Excel/PowerPoint/Outlook-specific) teach the model THE PROXY PATTERN, 5 critical rules (always load, always sync, use try/catch, check empty selections, prefer dedicated tools), and host-specific patterns. Reduces common Office.js errors (missing load/sync, wrong namespaces, undefined properties) through defensive prompting.
  - **Code Validator (Pillar 3)**: Pre-execution validation for all `eval_*` tools via `officeCodeValidator.ts`. Blocks execution if code is missing `context.sync()`, missing `.load()` before property reads, uses wrong namespace (e.g., Word API in Excel), contains infinite loops, or uses dangerous operations (`eval()`, `new Function()`). Provides validation feedback to the model for self-correction. Warnings for missing try/catch, excessive sync calls, incorrect Excel array formats, and large hardcoded ranges.
  - **Diffing Integration (Pillar 1)**: Format-preserving text editing tools that apply surgical changes while keeping formatting intact. Word's `proposeRevision` tool computes word-level diffs and applies only insertions/deletions, preserving bold/italic/colors/fonts on unchanged text, with optional Track Changes integration. PowerPoint's `proposeShapeTextRevision` tool provides diff statistics with full text replacement (API limitation). Uses `office-word-diff` library with cascading fallback strategies (token → sentence → block).
  - **Sandbox Enhancement**: Host filtering in `sandbox.ts` prevents cross-namespace API access, blocking Word/Excel/PowerPoint APIs in wrong contexts.
  - **Tool Count**: Increased from 127 to 129 tools (Word: 40→41 with `proposeRevision`, PowerPoint: 15→16 with `proposeShapeTextRevision`).
- **Secure Credential Persistence**: Added "Remember credentials" toggle in Settings > Account. Credentials are now stored with XOR obfuscation + Base64 encoding in localStorage when enabled, falling back to sessionStorage when disabled. Automatic migration from legacy sessionStorage format.
- **Smart Scroll Behavior**: Improved chat scroll UX with context-aware scrolling:
  - Scrolls to bottom when sending a user message
  - Scrolls to the top of the assistant message when receiving a response
  - Scrolls to bottom of history on initial page load
- **Credential Error Detection**: API now detects 401 credential errors and displays user-friendly localized messages instead of raw technical errors.
- **Secure Dynamic Code Execution**: Integrated `ses` (Secure ECMAScript) sandbox to allow the AI to safely execute custom JavaScript code via new escape-hatch tools (`eval_officejs`, `eval_wordjs`, `eval_powerpointjs`, `eval_outlookjs`) across all Office applications.
- **Agent File Processing**: Added a new `/api/upload` backend endpoint to process file uploads (PDF, DOCX, XLSX, CSV). Uploaded files are now attached to the AI prompt (`<attachments>`) and can be read by the AI using the `read` tool, enabling multi-modal data analysis.
- **Extended Excel Tools (OpenExcel Port)**: Ported advanced Excel manipulation tools from the OpenExcel project: `findData` (regex/case search), `duplicateWorksheet`, `hideUnhideRowColumn`, `getAllObjects` (charts/pivot tables), and `modifyObject` (delete charts/pivots).
- **Prompt Optimizations**: Enhanced prompt system in `useAgentPrompts.ts` to instruct the AI to leverage batch operations (`batchProcessRange`) instead of iterative cell updates, drastically improving performance for large data tasks. Improved overwrite protection documentation in the tool schemas.
- **Tool Execution Success Message**: When agent tools execute successfully without text response (e.g., proofreading with comments), displays "Actions completed successfully" instead of an error.

### Fixed

- **Azure/LiteLLM Compatibility**: Fixed message payload compatibility with Azure OpenAI and LiteLLM proxies by removing empty `tool_calls` arrays from assistant messages. These providers reject messages with `tool_calls: []`, causing agent failures when no tools are invoked.
- **Word Proofreading Empty Response**: Fixed issue where proofreading in Word displayed "empty response" error after successfully adding comments via the `addComment` tool.
- **Outlook Reply Language Detection**: Outlook reply quick action now properly detects the language of the original email and responds in the same language, overriding the configured reply language setting.

### Changed

- **DESIGN_REVIEW.md v2**: Complete fresh audit with 28 new issues identified (3 CRITICAL, 5 HIGH, 10 MEDIUM, 7 LOW, 3 BUILD). Previous v1 audit (38 issues, all resolved) preserved as reference.
- **README.md updated**: Fixed model default values to match actual code (`gpt-5.1`, `gpt-image-1`). Updated backend API surface description. Added undocumented frontend env vars (`VITE_REQUEST_TIMEOUT_MS`, `VITE_VERBOSE_LOGGING`). Refreshed Known Open Issues section with v2 findings. Cleaned up Not Yet Implemented list.
- **agents.md updated**: Fixed PowerPoint tool count (14, not 8). Updated backend architecture section to include `services/llmClient.js`. Corrected streaming endpoint documentation (now supports tools). Refreshed known issues reference to DESIGN_REVIEW.md v2.
- **CHANGELOG.md updated**: Added documentation update entries.

---

## [Previous]

### Added

- **Auto-scroll to beginning of AI response**: Automatically scrolls the chat window so the beginning of the AI response remains in view when generating long messages.
- **Excel, Outlook, PowerPoint Tooltips**: Added localized definitions for quick action tooltips.
- **PowerPoint Visual Quick Action automation**: Automatically triggers image model selection and generates images instantly when 5+ words are selected.
- **Extended PowerPoint Agent Skills**: Added `deleteSlide`, `getShapes`, `deleteShape`, `setShapeFill`, `moveResizeShape`, and `getAllSlidesOverview`.
- **PowerPoint Bullets Rendering**: Fixed an issue where AI-generated bullet points were inserted as raw text. They are now inserted as properly formatted HTML lists to enforce native slide bullets.
- **Translate Quick Action Unification**: Standardized the "Translate" action across Word, Outlook, and PowerPoint to automatically detect the source language (FR/EN) and translate to the other, sharing the same prompt, icon, tooltip, and menu position.
- **E2E Test Infrastructure**: Added Playwright test setup with navigation and settings tests (`npm run test:e2e`).
- **LLM API Service Abstraction**: Centralized LLM API calls in `services/llmClient.js` with timeout configuration.
- **Rate limiting for info endpoints**: Added rate limiting to `/health` and `/api/models` endpoints.
- **HSTS security headers**: Enabled strict transport security in production mode.

### Changed

- **DESIGN_REVIEW.md rewritten in English**: Full code audit with 38 identified issues organized by severity (3 CRITICAL, 6 HIGH, 16 MEDIUM, 10 LOW, 3 BUILD). All issues now resolved.
- **Node.js version constraint**: Updated to `>=20.19.0 || >=22.0.0` to align with LTS versions.
- **Vite build optimization**: Added `manualChunks` configuration to split vendor dependencies and reduce chunk sizes.
- **useAgentLoop refactored**: Grouped 34 options into logical sub-interfaces (`AgentLoopRefs`, `AgentLoopModels`, `AgentLoopHost`, `AgentLoopSettings`, `AgentLoopActions`, `AgentLoopHelpers`).

### Fixed

- **Tool state synchronization**: The interface "Settings" now correctly restricts the tools dynamically passed to the agent.
- **Streaming in Agent Loop**: Replaced synchronous `chatSync` calls with `chatStream` for immediate feedback during basic tool queries.
- **Context Management**: Implemented an intelligent pruning strategy ensuring message context doesn't exceed LLM token limits while keeping function and function call pairs together.
- **Chat History Persistence**: Historize conversations locally via `localStorage` (segmented by Office Host) to avoid data loss on taskpane closure.
- **UI UX translations/texts**: Extracted hardcoded "Thought process" string to `fr.json` & `en.json`.
- **Developer syntax**: Replaced confusing `${text}` / `${language}` syntax with standardized `[TEXT]` / `[LANGUAGE]` indicators in the built-in prompts editor in settings.
- **Image Generation Prompts**: Refined prompts to prevent text rejection and improve the corporate presentation aesthetic.
- **Image Clipboard Fallback**: Fixed an issue where users couldn't copy AI-generated images in environments lacking standard clipboard support (Office Webviews).
- **GitHub Action Permissions**: Grant write permissions to the bumps-version GitHub action workflow.

### Security

- **XSS Protection**: Added strict `ALLOWED_TAGS` and `ALLOWED_ATTR` allowlists to DOMPurify configuration.
- **Credential Sanitization**: User credentials (`X-User-Key`, `X-User-Email`) are now redacted from error logs.
- **Session Storage for Credentials**: Migrated LiteLLM credentials from localStorage to sessionStorage.
- **Credential Validation**: Added email format validation and minimum key length (8 chars) check.
- **Request Timeout**: Added configurable request timeout middleware (default 10 minutes).
- **Stream Backpressure**: Added proper backpressure handling for SSE streams.
- **Input Validation**: Enhanced message validation (role, content, max 200 messages).
- **Startup Validation**: API key is now validated at server startup (fatal in production).
- **Agent Loop Race Condition**: Fixed abort handling to prevent state corruption during tool execution.
- **Tool Preference Migration**: Tool preferences are now preserved when tool definitions change.
