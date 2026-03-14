/**
 * Application-wide numeric constants.
 * Centralises all magic numbers to avoid duplication and ease future changes.
 */

// ─── File upload ─────────────────────────────────────────────────────────────

/** Maximum allowed size for a user-attached file (10 MB). */
export const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;

// ─── UI sizing ───────────────────────────────────────────────────────────────

/** Maximum height of the auto-growing chat textarea (px). */
export const TEXTAREA_MAX_HEIGHT_PX = 120;

/** Small icon size (Lucide :size prop). */
export const ICON_SIZE_SM = 12;

/** Medium icon size (Lucide :size prop). */
export const ICON_SIZE_MD = 16;

/** Large icon size (Lucide :size prop). */
export const ICON_SIZE_LG = 20;

// ─── Backend / polling ───────────────────────────────────────────────────────

/** Interval between backend health-check polls (ms). */
export const HEALTH_CHECK_INTERVAL_MS = 30_000;

// ─── Logging ─────────────────────────────────────────────────────────────────

/** Maximum number of log entries kept in the in-memory ring buffer. */
export const LOG_RING_BUFFER_SIZE = 500;

// ─── Word tool limits ────────────────────────────────────────────────────────
// QUAL-M1: Centralized Word-specific limits

/** Maximum length for search text in Word tools. */
export const WORD_SEARCH_TEXT_MAX_LENGTH = 255;

/** Font size threshold for H1 heading detection. */
export const WORD_HEADING_1_FONT_SIZE = 20;

/** Font size threshold for H2 heading detection. */
export const WORD_HEADING_2_FONT_SIZE = 15;

/** Font size threshold for H3 heading detection. */
export const WORD_HEADING_3_FONT_SIZE = 12.5;

/** Code truncation length for validation errors. */
export const WORD_CODE_TRUNCATE_SHORT = 200;

/** Code truncation length for execution errors. */
export const WORD_CODE_TRUNCATE_LONG = 300;

// ─── Outlook tool limits ─────────────────────────────────────────────────────
// QUAL-M1: Centralized Outlook-specific limits

/** Timeout for Outlook API actions (ms). */
export const OUTLOOK_ACTION_TIMEOUT_MS = 20_000;

// ─── Office action retry ─────────────────────────────────────────────────────
// QUAL-M1: Centralized Office.js retry configuration

/** First retry delay for Office actions (ms). */
export const OFFICE_RETRY_BACKOFF_DELAY_1 = 1000;

/** Second retry delay for Office actions (ms). */
export const OFFICE_RETRY_BACKOFF_DELAY_2 = 2000;

// ─── Backend file limits ─────────────────────────────────────────────────────
// QUAL-M1: Backend-side file size limits (must match backend config)

/** Maximum file size for backend file uploads (50 MB). */
export const BACKEND_MAX_FILE_SIZE = 50 * 1024 * 1024;
