import { getUserKey, getUserEmail } from '@/utils/credentialStorage';
import { logService } from '@/utils/logger';

// Timeouts by model tier — reasoning models need more time (up to 6 min LLM + overhead)
const BASE_TIMEOUT_MS = Number(import.meta.env.VITE_REQUEST_TIMEOUT_MS) || 180_000;
const TIMEOUT_BY_TIER: Record<string, number> = {
  reasoning: 600_000, // 10 min — GPT-5.2 up to 65k output tokens
  standard: 300_000, // 5 min — GPT-5.2 up to 32k output tokens
  fast: 120_000,
};

function getTimeoutForTier(modelTier?: string): number {
  if (modelTier && TIMEOUT_BY_TIER[modelTier]) return TIMEOUT_BY_TIER[modelTier];
  return BASE_TIMEOUT_MS;
}

const RETRY_DELAYS_MS = [1_500, 4_000] as const;

function wait(ms: number): Promise<void> {
  return new Promise(resolve => {
    setTimeout(resolve, ms);
  });
}

function isRetryableError(error: unknown): boolean {
  return (
    error instanceof TypeError || (error instanceof DOMException && error.name === 'TimeoutError')
  );
}

function createTimeoutSignal(
  timeoutMs: number,
  externalSignal?: AbortSignal,
): { signal: AbortSignal; cleanup: () => void } {
  const timeoutController = new AbortController();

  const timeoutId = setTimeout(() => {
    timeoutController.abort(new DOMException('Request timed out', 'TimeoutError'));
  }, timeoutMs);

  const abortFromExternal = () => {
    timeoutController.abort(externalSignal?.reason);
  };

  if (externalSignal) {
    if (externalSignal.aborted) {
      abortFromExternal();
    } else {
      externalSignal.addEventListener('abort', abortFromExternal, { once: true });
    }
  }

  return {
    signal: timeoutController.signal,
    cleanup: () => {
      clearTimeout(timeoutId);
      externalSignal?.removeEventListener('abort', abortFromExternal);
    },
  };
}

export async function fetchWithTimeoutAndRetry(
  url: string,
  init: RequestInit = {},
  modelTier?: string,
): Promise<Response> {
  let attempt = 0;
  const timeoutMs = getTimeoutForTier(modelTier);

  while (true) {
    const { signal, cleanup } = createTimeoutSignal(timeoutMs, init.signal ?? undefined);

    try {
      return await fetch(url, {
        ...init,
        credentials: 'include',
        signal,
      });
    } catch (error) {
      if (init.signal?.aborted) {
        throw error;
      }

      const isPost = init.method?.toUpperCase() === 'POST';
      // Allow 1 retry on POST for timeout/network errors (transient failures)
      const maxPostRetries = 1;
      const shouldRetry =
        attempt < RETRY_DELAYS_MS.length &&
        isRetryableError(error) &&
        (!isPost || attempt < maxPostRetries);
      if (!shouldRetry) {
        logService.error(`Network request failed: ${url}`, error);
        throw error;
      }

      logService.warn(`Network retry ${attempt + 1}/${RETRY_DELAYS_MS.length} for ${url}`, error);
      await wait(RETRY_DELAYS_MS[attempt]);
      attempt += 1;
    } finally {
      cleanup();
    }
  }
}

function getCsrfToken(): string {
  const match = document.cookie.match(/(?:^| )csrf_token=([^;]+)/);
  if (match) return match[1];
  return '';
}

// ERR-L1: Generate a per-request UUID for frontend↔backend log correlation
export function generateRequestId(): string {
  return typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

// ─── QC-L2: Header cache ─────────────────────────────────────────────────────
// getGlobalHeaders() is called on every network request. The calls to
// getUserKey(), getUserEmail() and logService.getContext() touch async storage
// on each invocation. We cache the resolved headers and expose
// invalidateHeaderCache() so callers can bust the cache after credential changes.
let _headerCache: Promise<Record<string, string>> | null = null;

/** Bust the cached credentials headers (call after saving new credentials). */
export function invalidateHeaderCache(): void {
  _headerCache = null;
}

async function buildGlobalHeaders(): Promise<Record<string, string>> {
  const userKey = await getUserKey();
  const userEmail = await getUserEmail();
  const ctx = await logService.getContext();

  const headers: Record<string, string> = {};
  if (userKey) headers['X-User-Key'] = userKey;
  if (userEmail) headers['X-User-Email'] = userEmail;
  if (ctx.host) headers['X-Office-Host'] = ctx.host;
  if (ctx.sessionId) headers['X-Session-Id'] = ctx.sessionId;

  return headers;
}

export async function getGlobalHeaders(): Promise<Record<string, string>> {
  if (!_headerCache) {
    _headerCache = buildGlobalHeaders();
  }
  const cached = await _headerCache;
  // CSRF token reads from document.cookie — always fresh (not cached)
  const csrf = getCsrfToken();
  if (csrf) return { ...cached, 'x-csrf-token': csrf };
  return cached;
}
