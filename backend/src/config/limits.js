/**
 * QUAL-M1: Centralized limits and magic numbers for backend
 */

/**
 * File upload limits
 */
const FILE_LIMITS = {
  /** Maximum file size for uploads (50 MB) */
  MAX_FILE_SIZE: 50 * 1024 * 1024,
}

module.exports = {
  FILE_LIMITS,
}
