## Summary

This PR fixes three critical issues affecting production deployments and user experience:

1. **Credential decryption errors** causing console spam when encrypted data becomes corrupted
2. **Rate limiting failures** when running behind Synology/nginx reverse proxies
3. **Missing timestamps** in chat messages for better conversation context

## Changes

### 🔧 Backend Fixes

**Reverse Proxy Compatibility** (`backend/src/server.js`)
- Enabled Express `trust proxy` setting to allow `express-rate-limit` to correctly identify client IPs via `X-Forwarded-For` header
- Fixes `ERR_ERL_UNEXPECTED_X_FORWARDED_FOR` error when running behind Synology NAS or nginx reverse proxies

### 🎨 Frontend Improvements

**Credential Storage Error Handling** (`frontend/src/utils/credentialStorage.ts`)
- Enhanced `decryptValue()` to accept optional `key` parameter for identifying corrupted data
- Automatically clears corrupted encrypted credentials from localStorage when decryption fails
- Prevents repeated `OperationError` console spam on subsequent page loads

**Message Timestamps** (`frontend/src/types/chat.ts`, `frontend/src/composables/useImageActions.ts`, `frontend/src/components/chat/ChatMessageList.vue`)
- Added `timestamp` field to `DisplayMessage` interface
- Automatically capture message creation time in `createDisplayMessage()`
- Display time in HH:MM format below each message with subtle styling (10px, 60% opacity)

### 📝 Documentation

- Updated `README.md` with new features: reverse proxy support and message timestamps
- Updated `CHANGELOG.md` with detailed description of all fixes

## Testing

**Reverse Proxy Fix**:
- ✅ Backend starts without `ERR_ERL_UNEXPECTED_X_FORWARDED_FOR` error
- ✅ Rate limiting works correctly with `X-Forwarded-For` header

**Credential Storage**:
- ✅ Corrupted encrypted data is automatically cleaned from localStorage
- ✅ No repeated decryption errors in console on page reload
- ✅ New credentials are stored and retrieved correctly

**Message Timestamps**:
- ✅ All new messages display creation time
- ✅ Time format is concise (HH:MM)
- ✅ Styling is subtle and doesn't interfere with message content
- ✅ Timestamp only appears when available (backward compatible)

## Validation

- [x] Backend builds successfully
- [x] Frontend builds successfully
- [x] No TypeScript errors
- [x] Changes aligned with CLAUDE.md guidelines
- [x] Documentation updated

## Related Issues

Fixes console errors reported in Synology Docker deployments and improves chat UX with message timestamps.

---

🤖 Generated with [Claude Code](https://claude.com/claude-code)
