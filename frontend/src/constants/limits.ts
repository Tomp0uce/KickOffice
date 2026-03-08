/**
 * Application-wide numeric constants.
 * Centralises all magic numbers to avoid duplication and ease future changes.
 */

// ─── File upload ─────────────────────────────────────────────────────────────

/** Maximum allowed size for a user-attached file (10 MB). */
export const MAX_UPLOAD_BYTES = 10 * 1024 * 1024

// ─── UI sizing ───────────────────────────────────────────────────────────────

/** Maximum height of the auto-growing chat textarea (px). */
export const TEXTAREA_MAX_HEIGHT_PX = 120

/** Small icon size (Lucide :size prop). */
export const ICON_SIZE_SM = 12

/** Medium icon size (Lucide :size prop). */
export const ICON_SIZE_MD = 16

/** Large icon size (Lucide :size prop). */
export const ICON_SIZE_LG = 20

// ─── Backend / polling ───────────────────────────────────────────────────────

/** Interval between backend health-check polls (ms). */
export const HEALTH_CHECK_INTERVAL_MS = 30_000

// ─── Logging ─────────────────────────────────────────────────────────────────

/** Maximum number of log entries kept in the in-memory ring buffer. */
export const LOG_RING_BUFFER_SIZE = 500
