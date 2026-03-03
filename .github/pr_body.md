## Summary

Fix Docker build compatibility for Synology DS416play NAS with Intel Celeron processor.

## Problem

Alpine Linux images (`node:*-alpine`, `nginx:alpine`) cause **"Illegal instruction (core dumped)"** errors on Synology DS416play because:
- Alpine uses **musl libc** which executes AVX instructions
- Intel Celeron processors don't support these instructions
- This causes immediate crashes when running Node.js or nginx

Additionally, `office-word-diff` was missing from `package-lock.json`, causing `npm ci` to fail.

## Solution

### Docker Images
- `node:22-alpine` → `node:22-slim` (Debian/glibc)
- `nginx:alpine` → `nginx:stable` (Debian-based)

### Package Lock
- Added `office-word-diff` to `package-lock.json` for `npm ci` compatibility

### Documentation
- Updated DESIGN_REVIEW.md with correct Synology compatibility requirements
- Updated CHANGELOG.md with fix details

## Test plan

- [ ] Run `docker compose build --no-cache`
- [ ] Run `docker compose up -d`
- [ ] Verify all containers start without "Illegal instruction" errors
- [ ] Test backend health: `curl http://localhost:3003/health`
- [ ] Test frontend loads at `http://localhost:3002`

🤖 Generated with [Claude Code](https://claude.ai/code)
