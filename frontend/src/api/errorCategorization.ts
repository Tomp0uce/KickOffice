// ────────────────────────────────────────────────────────────────────────────
// Error categorisation — exposes structured info for user-facing messages
// ────────────────────────────────────────────────────────────────────────────
export type ErrorType = 'timeout' | 'network' | 'rate_limit' | 'auth' | 'server' | 'unknown';

export interface CategorizedError {
  type: ErrorType;
  /** i18n key to use in the UI */
  i18nKey: string;
  /** Optional raw message from the upstream provider (e.g. LiteLLM BadRequestError detail) */
  rawDetail?: string;
}

/** Maps backend error codes to i18n keys. Falls back to message inspection if no code present. */
const ERROR_CODE_MAP: Record<string, CategorizedError> = {
  VALIDATION_ERROR: { type: 'unknown', i18nKey: 'failedToResponse' },
  AUTH_REQUIRED: { type: 'auth', i18nKey: 'credentialsRequired' },
  RATE_LIMITED: { type: 'rate_limit', i18nKey: 'errorRateLimit' },
  LLM_BAD_REQUEST: { type: 'unknown', i18nKey: 'errorLlmBadRequest' },
  LLM_UPSTREAM_ERROR: { type: 'server', i18nKey: 'errorServer' },
  LLM_EMPTY_RESPONSE: { type: 'server', i18nKey: 'errorServer' },
  LLM_INVALID_JSON: { type: 'server', i18nKey: 'errorServer' },
  LLM_NO_CHOICES: { type: 'server', i18nKey: 'errorServer' },
  LLM_CONTENT_FILTERED: { type: 'server', i18nKey: 'errorServer' },
  LLM_TIMEOUT: { type: 'timeout', i18nKey: 'errorTimeout' },
  IMAGE_TIMEOUT: { type: 'timeout', i18nKey: 'errorTimeout' },
  INTERNAL_ERROR: { type: 'server', i18nKey: 'errorServer' },
  PDF_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
  DOCX_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
  NO_FILE_UPLOADED: { type: 'unknown', i18nKey: 'failedToResponse' },
  UNSUPPORTED_FILE_TYPE: { type: 'unknown', i18nKey: 'failedToResponse' },
  FILE_EMPTY: { type: 'unknown', i18nKey: 'failedToResponse' },
  CHART_IMAGE_NOT_FOUND: { type: 'unknown', i18nKey: 'failedToResponse' },
  CHART_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
};

export function categorizeError(error: unknown): CategorizedError {
  if (error instanceof DOMException && error.name === 'AbortError') {
    return { type: 'unknown', i18nKey: 'generationStop' };
  }
  if (error instanceof DOMException && error.name === 'TimeoutError') {
    return { type: 'timeout', i18nKey: 'errorTimeout' };
  }
  if (error instanceof TypeError) {
    return { type: 'network', i18nKey: 'errorNetwork' };
  }

  // Try structured error code first (from backend ErrorCodes registry)
  if (error instanceof Error && 'code' in error) {
    const errorWithCode = error as Error & { code?: string; detail?: string };
    const mapped = ERROR_CODE_MAP[errorWithCode.code ?? ''];
    if (mapped) {
      const rawDetail = errorWithCode.detail;
      return rawDetail ? { ...mapped, rawDetail } : mapped;
    }
  }

  // Fallback: inspect error message string
  const msg = (error instanceof Error ? error.message : String(error)).toLowerCase();
  if (
    msg.includes('401') ||
    msg.includes('403') ||
    msg.includes('credentials') ||
    msg.includes('x-user-key') ||
    msg.includes('x-user-email')
  ) {
    return { type: 'auth', i18nKey: 'credentialsRequired' };
  }
  if (msg.includes('429') || msg.includes('rate limit') || msg.includes('too many')) {
    return { type: 'rate_limit', i18nKey: 'errorRateLimit' };
  }
  if (
    msg.includes('500') ||
    msg.includes('502') ||
    msg.includes('503') ||
    msg.includes('internal server')
  ) {
    return { type: 'server', i18nKey: 'errorServer' };
  }
  if (msg.includes('timeout') || msg.includes('timed out')) {
    return { type: 'timeout', i18nKey: 'errorTimeout' };
  }
  return { type: 'unknown', i18nKey: 'failedToResponse' };
}
