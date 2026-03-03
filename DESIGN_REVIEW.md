# DESIGN_REVIEW.md — Code Audit v4

**Date**: 2026-03-03
**Version**: 4.0
**Scope**: Full codebase review (frontend, backend, office-word-diff)

---

## Summary

| Severity | Count | Description |
|----------|-------|-------------|
| **CRITICAL** | 10 | Security vulnerabilities, data loss, crashes, **build failures** |
| **HIGH** | 9 | Major bugs, broken features |
| **MEDIUM** | 12 | Minor bugs, inconsistencies |
| **LOW** | 8 | Code quality, style issues |
| **DEAD CODE** | 4 | Unused files/artifacts |
| **Total** | **43** | |

---

## CRITICAL Issues

### C0a. Docker Build Failure — office-word-diff Not in package-lock.json
**File**: `frontend/package.json`, `frontend/package-lock.json`
**Status**: IMMEDIATE FIX REQUIRED

The `office-word-diff` dependency is defined as `"file:../office-word-diff"` in package.json but is **missing from package-lock.json**. This causes `npm ci` to fail because it requires exact lockfile match.

**Error**: `npm ci` fails during Docker build
**Impact**: Docker build completely broken

**Fix**: Run `cd frontend && npm install` to regenerate package-lock.json with office-word-diff included.

---

### C0b. Docker Build Context Missing office-word-diff
**File**: `frontend/Dockerfile`, `docker-compose.yml`
**Status**: IMMEDIATE FIX REQUIRED

The frontend Dockerfile context is `./frontend`, but `office-word-diff` is at `../office-word-diff` (outside build context). Even with a correct package-lock.json, npm cannot resolve the local file dependency.

**Impact**: Docker build fails on any Synology NAS or CI environment

**Fix Options**:
1. **Recommended**: Change docker-compose to use root context and update Dockerfile path
2. **Alternative**: Publish office-word-diff to npm or bundle it into frontend
3. **Workaround**: Copy office-word-diff into frontend before build

---

### C0c. Synology DS416play Compatibility Issues
**File**: `docker-compose.yml`, Dockerfiles
**Status**: FIXED

The Synology DS416play has an Intel Celeron processor that is **NOT compatible with Alpine Linux**. Alpine uses musl libc which executes instructions (AVX) that the Celeron doesn't support, causing "Illegal instruction (core dumped)" errors.

**Requirements**:
- **MUST use `node:22-slim`** (Debian-based, glibc) — NOT Alpine
- **MUST use `nginx:stable`** (Debian-based) — NOT nginx:alpine
- Pre-build images on a more powerful machine if build is too slow
- Consider `--max-old-space-size=512` for memory-constrained builds

**DO NOT USE on Synology DS416play**:
- `node:*-alpine` — causes "Illegal instruction" on Celeron CPUs
- `nginx:alpine` — same issue with musl libc
- Any musl-based images

---

### C1. CSRF Token with SameSite=None
**File**: `backend/src/server.js` (Lines 113-139)
**Status**: Open

CSRF token set with `sameSite: 'none'`, allowing cross-site requests to include the token. The skip condition for `X-User-Key` header creates a bypass where Office add-in requests skip CSRF checks entirely.

**Impact**: CSRF attacks possible in certain scenarios.

**Fix**: Use `sameSite: 'strict'` or implement proper origin validation.

---

### C2. Unvalidated File Upload Memory Exhaustion
**File**: `backend/src/routes/upload.js` (Lines 9-52)
**Status**: Open

- Multer configured with 10MB limit per request, but no rate limiting on upload endpoint
- PDF parser buffers entire file in memory
- XLSX library loads entire workbook without sheet size limits
- Large concurrent uploads can cause OOM

**Impact**: DoS vulnerability via memory exhaustion.

**Fix**: Add rate limiting to `/api/upload`, implement streaming for large files.

---

### C3. Plaintext Credentials in localStorage
**File**: `frontend/src/utils/credentialStorage.ts` (Lines 3-50)
**Status**: Open

User credentials stored with XOR obfuscation (easily reversible) in localStorage. Accessible to any XSS payload in Office context.

**Impact**: Credential theft via XSS.

**Fix**: Use Web Crypto API for encryption or avoid persistent credential storage.

---

### C4. Unsafe JSON.stringify for Tool Signature
**File**: `frontend/src/composables/useAgentLoop.ts` (Line 159)
**Status**: Open

`JSON.stringify(toolArgs)` on untrusted LLM-provided arguments. Vulnerable to prototype pollution and can cause DoS with deeply nested structures.

**Impact**: Agent loop detection bypass, DoS via nested objects.

**Fix**: Validate object structure and depth before stringification.

---

### C5. Unhandled Abort Signal in Stream Processing
**File**: `backend/src/routes/chat.js` (Lines 80-110)
**Status**: Open

Reader loop continues draining upstream response after client abort. `res.write()` after disconnection throws unhandled error.

**Impact**: Hanging requests, zombie connections, resource leak.

**Fix**: Implement proper abort signal handling with read timeout.

---

### C6. Agent Max Iterations Silently Capped
**File**: `frontend/src/composables/useAgentLoop.ts` (Line 349)
**Status**: Open

User can set `agentMaxIterations` up to 100, but the actual enforcement is only via timeout. Agent can exceed configured max iterations.

**Impact**: User setting is misleading and ignored.

**Fix**: Enforce iteration count explicitly in the agent loop.

---

### C7. Invalid `reasoning_effort=none` in .env.example
**File**: `backend/.env.example`
**Status**: Open

The value `'none'` is not a valid OpenAI API value for `reasoning_effort`. Causes empty responses when used with tools.

**Impact**: Breaks tool use with reasoning models.

**Fix**: Remove the line or use valid values: `low`, `medium`, `high`.

---

### C8. Quick Actions Bypass Loading/Abort State
**File**: `frontend/src/pages/HomePage.vue`
**Status**: Open

Quick actions can be triggered while another request is in progress, bypassing the loading state and abort handling.

**Impact**: History corruption, duplicate requests.

**Fix**: Disable quick actions while loading, use shared abort controller.

---

## HIGH Issues

### H1. Missing Validation for Empty Model Response
**File**: `backend/src/routes/chat.js` (Lines 180-220)
**Status**: Open

Response validation for `/api/chat/sync` doesn't validate empty `message.content`. Frontend creates assistant message with empty string.

**Impact**: Silent message loss, agent retries indefinitely.

**Fix**: Validate non-empty content before returning success.

---

### H2. IndexedDB Quota Exceeded Silent Failure
**File**: `frontend/src/composables/useSessionDB.ts` (Lines 119-134)
**Status**: Open

No try-catch around `idbPut()`. If IndexedDB quota is exceeded, operation fails silently and session messages are lost.

**Impact**: Silent data loss, chat history corruption.

**Fix**: Add exception handling with user notification and fallback.

---

### H3. Unvalidated Tool Definition Injection
**File**: `backend/src/middleware/validate.js` (Lines 52-85)
**Status**: Open

Tool validation accepts any schema without recursive depth limit. Deeply nested parameter schemas (1000+ levels) can cause parser DoS.

**Impact**: Tool-based DoS via malformed definitions.

**Fix**: Add depth limit to schema validation.

---

### H4. Race Condition in Session Switching
**File**: `frontend/src/composables/useSessionManager.ts` (Lines 30-38)
**Status**: Open

No mutex on session switching. Rapid session switches cause concurrent saves to interleave, losing data.

**Impact**: Session data loss on rapid switching.

**Fix**: Implement mutex/queue for session operations.

---

### H5. Memory Leak in SSE Stream Processing
**File**: `frontend/src/api/backend.ts` (Lines 224-280)
**Status**: Open

Buffer grows unbounded during large responses. No limit on buffer size—malicious server can OOM frontend.

**Impact**: Frontend memory exhaustion, browser tab crash.

**Fix**: Add buffer size limits, implement backpressure.

---

### H6. Chat History Unbounded Growth
**File**: `frontend/src/composables/useSessionDB.ts`
**Status**: Open

Chat history stored in IndexedDB without size limits. Long sessions accumulate megabytes of data.

**Impact**: Slow performance, storage exhaustion.

**Fix**: Implement history pruning or pagination.

---

### H7. Token Manager Uses Character Count
**File**: `frontend/src/utils/tokenManager.ts`
**Status**: Open

Token budget calculated using character count instead of actual token count. Inaccurate for non-ASCII text.

**Impact**: Context corruption, premature truncation or overflow.

**Fix**: Use tiktoken or similar for accurate token counting.

---

### H8. No Confirmation Dialog for New Chat
**File**: `frontend/src/pages/HomePage.vue`
**Status**: Open

"New Chat" button immediately clears history without confirmation. Easy to lose conversation accidentally.

**Impact**: Accidental data loss.

**Fix**: Add confirmation dialog or undo functionality.

---

### H9. Incomplete PDF Parser Error Handling
**File**: `backend/src/routes/upload.js` (Line 45)
**Status**: Open

PDFParse may return partial text for corrupted PDFs. No validation of extraction success.

**Impact**: Silent PDF corruption, incomplete analysis.

**Fix**: Validate extraction success, detect truncation.

---

## MEDIUM Issues

### M1. Inconsistent Error Handling in Token Manager
**File**: `frontend/src/utils/tokenManager.ts` (Lines 35-75)
**Status**: Open

Message truncation logic truncates system prompt first without warning. No user notification of truncation.

**Impact**: Context corruption, lost conversation history.

---

### M2. Unvalidated LocalStorage Language Preference
**File**: `frontend/src/composables/useAgentLoop.ts` (Lines 551, 645, 933)
**Status**: Open

Three separate hardcoded language comparisons instead of centralized i18n usage.

**Impact**: Inconsistent language selection.

---

### M3. Type Assertions Without Runtime Validation
**File**: `frontend/src/composables/useAgentLoop.ts` (Lines 119, 138, 347, 672)
**Status**: Open

Multiple `as Record<string, any>` casts without runtime type checks.

**Impact**: Runtime type errors, unexpected failures.

---

### M4. Accessibility (ARIA) Incomplete
**File**: Various frontend components
**Status**: Open

Missing ARIA attributes on interactive elements, insufficient keyboard navigation.

**Impact**: Poor accessibility for screen reader users.

---

### M5. No Unit Test Infrastructure
**File**: `frontend/`
**Status**: Open

Only E2E tests via Playwright. No vitest/jest configuration for unit testing.

**Impact**: Harder to catch regressions, slower feedback loop.

---

### M6. No ESLint/Prettier Configuration
**File**: Project root
**Status**: Open

No linting or formatting configuration. Inconsistent code style.

**Impact**: Code quality drift, harder reviews.

---

### M7. Missing Request Cancellation on Route Change
**File**: `frontend/src/pages/HomePage.vue`
**Status**: Open

In-flight requests not cancelled when navigating away from page.

**Impact**: Wasted resources, potential state corruption.

---

### M8. Excel Formula Language Detection Incomplete
**File**: `frontend/src/utils/excelTools.ts`
**Status**: Open

Formula language detection only handles en/fr. Other locales may have issues.

**Impact**: Incorrect formula syntax for non-EN/FR users.

---

### M9. PowerPoint Bullet Rendering Fragile
**File**: `frontend/src/utils/powerpointTools.ts`
**Status**: Open

HTML list insertion depends on Office version support. May fail silently on older versions.

**Impact**: Raw text instead of bullets on older Office.

---

### M10. Image Generation Error Messages Generic
**File**: `frontend/src/composables/useImageActions.ts`
**Status**: Open

All image generation errors show same generic message. No specific error codes.

**Impact**: Hard to diagnose image generation failures.

---

### M11. Outlook Compose Mode Detection Fragile
**File**: `frontend/src/utils/outlookTools.ts`
**Status**: Open

Compose vs read mode detection relies on checking multiple properties. May fail on edge cases.

**Impact**: Wrong tools used in wrong context.

---

### M12. Settings Page Performance
**File**: `frontend/src/pages/SettingsPage.vue` (1157 lines)
**Status**: Open

Large single component with all settings. May cause rendering performance issues.

**Impact**: Slow settings page, especially on mobile.

---

## LOW Issues

### L1. Sensitive Header Redaction Regex Inefficiency
**File**: `backend/src/utils/http.js` (Lines 3-17)
**Status**: Open

Pre-compiled regexes with global flag, manual `lastIndex` reset. Pattern may over-redact.

**Impact**: Over-redaction in logs, lost debugging info.

---

### L2. Inconsistent Logging in Router Handlers
**File**: `backend/src/routes/image.js` (Line 34)
**Status**: Open

Image router doesn't have verbose logging like chat router.

**Impact**: Harder to debug image generation errors.

---

### L3. Magic Numbers Without Constants
**File**: `backend/src/routes/upload.js` (Lines 18, 99)
**Status**: Open

`10 * 1024 * 1024` and `MAX_CHARS = 100000` as magic numbers.

**Impact**: Maintenance burden, unclear rationale.

---

### L4. Implicit Type Coercion in Token Calculations
**File**: `frontend/src/utils/tokenManager.ts` (Line 24)
**Status**: Open

Number arithmetic without explicit NaN checks.

**Impact**: Subtle logic bugs if message length miscalculated.

---

### L5. Unused Variables in Model Configuration
**File**: `backend/src/config/models.js` (Lines 56-64)
**Status**: Open

`temperatureEffort` variable defined but never used.

**Impact**: Code clutter, cognitive load.

---

### L6. Console.log Statements in Production Code
**File**: Various
**Status**: Open

Debug console.log statements left in some utility files.

**Impact**: Noisy browser console.

---

### L7. Duplicate Error Message Definitions
**File**: `frontend/src/i18n/locales/`
**Status**: Open

Some error messages defined in both locale files with slight variations.

**Impact**: Inconsistent error messages.

---

### L8. Backend Health Check Minimal
**File**: `backend/src/routes/health.js`
**Status**: Open

Health check only returns `{ status: 'ok' }`. No dependency checks.

**Impact**: May report healthy when LLM API is down.

---

## DEAD CODE

### D1. Old PR Body File
**File**: `.github/pr_body_12_14.md`
**Status**: Can be deleted

Old PR template artifact from previous PR submission.

---

### D2. Analysis Documents
**Files**: `agents.md`, `gemini.md`, `AGENT_MODE_ANALYSIS.md`
**Status**: Review needed

Analysis/reference documents that may be outdated. Verify if still needed.

---

### D3. LiteLLM Local Proxy Directory
**File**: `litellm-local-proxy/`
**Status**: Review needed

Directory exists but not actively used in current deployment model.

---

### D4. Legacy wordApi.ts Exports
**File**: `frontend/src/api/wordApi.ts`
**Status**: Partial

File marked as legacy in comments but still has active exports used by `useOfficeInsert.ts`.

---

## Deferred Items (From Previous Audits)

### IC2 — Containers Run as Root
**Files**: `frontend/Dockerfile`, `backend/Dockerfile`

Containers run as root (Node.js default). Adding `USER node` is best practice but retained for deployment simplicity.

### IH2 — Private IP in Build Arg
**File**: `frontend/Dockerfile`

Private IP baked into default `VITE_BACKEND_URL`. Users must override at build time.

### IH3 — DuckDNS Domain in Example
**File**: `.env.example`

External DuckDNS domain as example value. Users replace with their own.

### UM10 — PowerPoint HTML Reconstruction
**File**: `frontend/src/utils/powerpointTools.ts`

Character-by-character HTML reconstruction—high complexity, low ROI. Deferred.

---

## Recommendations by Priority

### Immediate (Critical)
1. Fix `sameSite: 'none'` on CSRF cookie OR implement origin validation
2. Add rate limiting to `/api/upload` endpoint
3. Encrypt credentials properly or remove persistent storage
4. Fix `reasoning_effort=none` in `.env.example`
5. Implement abort signal handling in streaming

### Short-term (High)
1. Add IndexedDB quota exception handling
2. Validate model responses for non-empty content
3. Implement mutex for session switching
4. Add buffer size limits to SSE processing
5. Enforce agent max iterations explicitly

### Medium-term (Medium/Low)
1. Add unit test infrastructure (vitest)
2. Configure ESLint + Prettier
3. Improve accessibility (ARIA)
4. Add confirmation for destructive actions
5. Clean up dead code and artifacts

---

## Verification Commands

```bash
# Frontend build check
cd frontend && npm run build

# Backend start check
cd backend && npm start

# E2E tests
cd frontend && npm run test:e2e

# Check for TypeScript errors
cd frontend && npx tsc --noEmit
```

---

## Changelog

| Version | Date | Changes |
|---------|------|---------|
| v4.0 | 2026-03-03 | Complete fresh audit, added dead code analysis |
| v3.0 | 2026-02-28 | 162 issues identified, 131 resolved |
| v2.0 | 2026-02-22 | 28 new issues after major refactor |
| v1.0 | 2026-02-15 | Initial audit, 38 issues (all resolved) |
