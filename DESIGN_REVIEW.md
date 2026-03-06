# DESIGN_REVIEW.md — Code Audit v4

**Date**: 2026-03-03
**Version**: 4.0
**Scope**: Full codebase review (frontend, backend, office-word-diff)

---

## Summary

| Severity      | Count  | Fixed  | Remaining | Description                                                      |
| ------------- | ------ | ------ | --------- | ---------------------------------------------------------------- |
| **CRITICAL**  | 10     | ✅ 10  | 0         | Security vulnerabilities, data loss, crashes, **build failures** |
| **HIGH**      | 13     | ✅ 13  | 0         | Major bugs, broken features                                      |
| **MEDIUM**    | 15     | ✅ 15  | 0         | Minor bugs, inconsistencies                                      |
| **LOW**       | 8      | ✅ 8   | 0         | Code quality, style issues                                       |
| **DEAD CODE** | 4      | ✅ 4   | 0         | Unused files/artifacts                                           |
| **Total**     | **50** | **50** | **0**    |                                                                  |

### ✅ Critical Issues Resolution (2026-03-03)

All 10 critical security issues have been addressed:

- ✅ **C0a**: Docker npm ci compatibility - Changed to `npm install`
- ✅ **C0b**: Docker build context - Root context with proper office-word-diff copy
- ✅ **C1**: CSRF protection - Added explicit origin validation
- ✅ **C2**: Upload rate limiting - 10 uploads/min per IP
- ✅ **C3**: Credential encryption - AES-GCM 256-bit with Web Crypto API
- ✅ **C4**: Safe JSON stringify - Depth validation + circular detection
- ✅ **C5**: Stream abort handling - reader.cancel() + 30s timeout
- ✅ **C6**: Agent max iterations - Explicit iteration count enforcement
- ✅ **C7**: Invalid reasoning_effort - Removed from .env.example
- ✅ **C8**: Quick actions loading - Prevent execution during active requests

---

## CRITICAL Issues

### C0a. Docker Build Failure — office-word-diff Not in package-lock.json

**File**: `frontend/package.json`, `frontend/package-lock.json`
**Status**: ✅ **FIXED**
---

### C0b. Docker Build Context Missing office-word-diff

**File**: `frontend/Dockerfile`, `docker-compose.yml`
**Status**: ✅ **FIXED**
---

### C0c. Synology DS416play Compatibility Issues

**File**: `docker-compose.yml`, Dockerfiles
**Status**: ✅ **FIXED**
---

### C1. CSRF Token with SameSite=None

**File**: `backend/src/server.js` (Lines 113-139)
**Status**: ✅ **FIXED**
---

### C2. Unvalidated File Upload Memory Exhaustion

**File**: `backend/src/routes/upload.js` (Lines 9-52)
**Status**: ✅ **FIXED**
---

### C3. Plaintext Credentials in localStorage

**File**: `frontend/src/utils/credentialStorage.ts` (Lines 3-50)
**Status**: ✅ **FIXED**
---

### C4. Unsafe JSON.stringify for Tool Signature

**File**: `frontend/src/composables/useAgentLoop.ts` (Line 159)
**Status**: ✅ **FIXED**
---

### C5. Unhandled Abort Signal in Stream Processing

**File**: `backend/src/routes/chat.js` (Lines 80-110)
**Status**: ✅ **FIXED**
---

### C6. Agent Max Iterations Silently Capped

**File**: `frontend/src/composables/useAgentLoop.ts` (Line 349)
**Status**: ✅ **FIXED**
---

### C7. Invalid `reasoning_effort=none` in .env.example

**File**: `backend/.env.example`
**Status**: ✅ **FIXED**
---

### C8. Quick Actions Bypass Loading/Abort State

**File**: `frontend/src/pages/HomePage.vue`
**Status**: ✅ **FIXED**
---

## HIGH Issues

### H1. Missing Validation for Empty Model Response

**File**: `backend/src/routes/chat.js` (Lines 180-220)
**Status**: ✅ **FIXED**
---

### H2. IndexedDB Quota Exceeded Silent Failure

**File**: `frontend/src/composables/useSessionDB.ts` (Lines 119-134)
**Status**: ✅ **FIXED**
---

### H3. Unvalidated Tool Definition Injection

**File**: `backend/src/middleware/validate.js` (Lines 52-85)
**Status**: ✅ **FIXED**
---

### H4. Race Condition in Session Switching

**File**: `frontend/src/composables/useSessionManager.ts` (Lines 30-38)
**Status**: ✅ **FIXED**
---

### H5. Memory Leak in SSE Stream Processing

**File**: `frontend/src/api/backend.ts` (Lines 224-280)
**Status**: ✅ **FIXED**
---

### H6. Chat History Unbounded Growth

**File**: `frontend/src/composables/useSessionDB.ts`
**Status**: ✅ **FIXED**
---

### H7. Token Manager Uses Character Count

**File**: `frontend/src/utils/tokenManager.ts`
**Status**: ✅ **FIXED**
---

### H8. No Confirmation Dialog for New Chat

**File**: `frontend/src/pages/HomePage.vue`
**Status**: ✅ **FIXED**
---

### H9. Incomplete PDF Parser Error Handling

**File**: `backend/src/routes/upload.js` (Line 45)
**Status**: ✅ **FIXED**
---

### H10. Missing Specific Error Message for Invalid Credentials

**File**: `frontend/src/api/backend.ts` / `frontend/src/composables/useAgentLoop.ts`
**Status**: ✅ **FIXED**
---

### H11. Agent State Hanging Between Tools and After Task Completion

**File**: `frontend/src/composables/useAgentLoop.ts`
**Status**: ✅ **FIXED**
---

### H12. Custom Prompts Not Accessible in Dropdown

**File**: `frontend/src/pages/HomePage.vue` / `frontend/src/components/`
**Status**: ✅ **FIXED**
---

### H13. Max Token Limit Too Low

**File**: `frontend/src/utils/tokenManager.ts` / `backend/src/config/`
**Status**: ✅ **FIXED**
---

## MEDIUM Issues

### M1. Inconsistent Error Handling in Token Manager

**File**: `frontend/src/utils/tokenManager.ts` (Lines 35-75)
**Status**: ✅ **FIXED**
---

### M2. Unvalidated LocalStorage Language Preference

**File**: `frontend/src/composables/useAgentLoop.ts` (Lines 551, 645, 933)
**Status**: ✅ **FIXED**
---

### M3. Type Assertions Without Runtime Validation

**File**: `frontend/src/composables/useAgentLoop.ts` (Lines 119, 138, 347, 672)
**Status**: ✅ **FIXED**
---

### M4. Accessibility (ARIA) Incomplete

**File**: Various frontend components
**Status**: ✅ **FIXED**
---

### M5. No Unit Test Infrastructure

**File**: `frontend/`
**Status**: ✅ **FIXED**
---

### M6. No ESLint/Prettier Configuration

**File**: Project root
**Status**: ✅ **FIXED**
---

### M7. Missing Request Cancellation on Route Change

**File**: `frontend/src/pages/HomePage.vue`
**Status**: ✅ **FIXED**
---

### M9. PowerPoint Bullet Rendering Fragile

**File**: `frontend/src/utils/powerpointTools.ts`
**Status**: ✅ **FIXED**
---

### M10. Image Generation Error Messages Generic

**File**: `frontend/src/composables/useImageActions.ts`
**Status**: ✅ **FIXED**
---

### M11. Outlook Compose Mode Detection Fragile

**File**: `frontend/src/utils/outlookTools.ts`
**Status**: ✅ **FIXED**
---

### M12. Settings Page Performance

**File**: `frontend/src/pages/SettingsPage.vue` (1157 lines)
**Status**: ✅ **FIXED**
---

### M13. Redundant Dual Status and Missing Status States

**File**: `frontend/src/components/chat/`, `frontend/src/components/StatusBar.vue`
**Status**: ✅ **FIXED**
---

### M14. Missing Context Usage Indicator

**File**: `frontend/src/components/StatusBar.vue`
**Status**: ✅ **FIXED**
---

### M15. Word Proofreading Too Intrusive

**File**: `frontend/src/prompts/` / `frontend/src/utils/wordTools.ts`
**Status**: ✅ **FIXED**
---

### M16. Uploaded File Content Populates Chat Input Directly

**File**: `frontend/src/components/chat/` / `frontend/src/pages/HomePage.vue`
**Status**: ✅ **FIXED**
---

## LOW Issues

### L1. Sensitive Header Redaction Regex Inefficiency

**File**: `backend/src/utils/http.js` (Lines 3-17)
**Status**: ✅ **FIXED**
---

### L2. Inconsistent Logging in Router Handlers

**File**: `backend/src/routes/image.js` (Line 34)
**Status**: ✅ **FIXED**
---

### L3. Magic Numbers Without Constants

**File**: `backend/src/routes/upload.js` (Lines 18, 99)
**Status**: ✅ **FIXED**
---

### L4. Implicit Type Coercion in Token Calculations

**File**: `frontend/src/utils/tokenManager.ts` (Line 24)
**Status**: ✅ **FIXED**
---

### L5. Unused Variables in Model Configuration

**File**: `backend/src/config/models.js` (Lines 56-64)
**Status**: ✅ **FIXED**
---

### L6. Console.log Statements in Production Code

**File**: Various
**Status**: ✅ **FIXED**
---

### L7. Duplicate Error Message Definitions

**File**: `frontend/src/i18n/locales/`
**Status**: ✅ **FIXED**
---

### L8. Backend Health Check Minimal

**File**: `backend/src/routes/health.js`
**Status**: ✅ **FIXED**
---

## DEAD CODE

### D1. Old PR Body File

**File**: `.github/pr_body_12_14.md`
**Status**: ✅ **FIXED**
---

### D2. Analysis Documents

**Files**: `AGENT_MODE_ANALYSIS.md`
**Status**: ✅ **FIXED**
---

### D3. LiteLLM Local Proxy Directory

**File**: `litellm-local-proxy/`
**Status**: ✅ **FIXED**
---

### D4. Legacy wordApi.ts Exports

**File**: `frontend/src/api/wordApi.ts`
**Status**: ✅ **FIXED**
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

| Version | Date       | Changes                                        |
| ------- | ---------- | ---------------------------------------------- |
| v4.0    | 2026-03-03 | Complete fresh audit, added dead code analysis |
| v3.0    | 2026-02-28 | 162 issues identified, 131 resolved            |
| v2.0    | 2026-02-22 | 28 new issues after major refactor             |
| v1.0    | 2026-02-15 | Initial audit, 38 issues (all resolved)        |
