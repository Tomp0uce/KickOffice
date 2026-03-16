import { logService } from './logger';
import { OFFICE_RETRY_BACKOFF_DELAY_1, OFFICE_RETRY_BACKOFF_DELAY_2 } from '@/constants/limits';

const DEFAULT_OFFICE_ACTION_TIMEOUT_MS = 10_000;
const OFFICE_BUSY_TIMEOUT_MESSAGE =
  'Office app is busy. Please exit cell editing or close dialogs.';

// Helper to delay execution
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export async function executeOfficeAction<T>(
  action: () => Promise<T>,
  actionName: string = 'unknown_action',
  timeoutMs: number = DEFAULT_OFFICE_ACTION_TIMEOUT_MS,
  abortSignal?: AbortSignal,
): Promise<T> {
  if (abortSignal?.aborted) {
    throw new Error('Operation aborted by user');
  }

  const maxRetries = 2;
  const backoffDelays = [OFFICE_RETRY_BACKOFF_DELAY_1, OFFICE_RETRY_BACKOFF_DELAY_2];

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    if (abortSignal?.aborted) {
      throw new Error('Operation aborted by user');
    }

    let timeoutId: ReturnType<typeof setTimeout> | undefined;
    // ERR-C4: Register abort listener OUTSIDE the Promise constructor so it is always
    // assigned before Promise.race and always removed in finally — no race, no leak.
    let abortListener: (() => void) | undefined;
    let rejectTimeoutPromise: ((err: Error) => void) | undefined;

    const timeoutPromise = new Promise<never>((_, reject) => {
      rejectTimeoutPromise = reject;
      timeoutId = setTimeout(() => {
        logService.warn(`[OfficeAction] Timeout executing ${actionName} (attempt ${attempt + 1})`, {
          actionName,
        });
        reject(new Error(OFFICE_BUSY_TIMEOUT_MESSAGE));
      }, timeoutMs);
    });

    if (abortSignal) {
      abortListener = () => {
        logService.warn(`[OfficeAction] Aborted ${actionName}`, { actionName });
        rejectTimeoutPromise?.(new Error('Operation aborted by user'));
      };
      abortSignal.addEventListener('abort', abortListener);
    }

    try {
      const result = await Promise.race([action(), timeoutPromise]);
      return result;
    } catch (err) {
      // Don't retry if aborted
      if (err instanceof Error && err.message === 'Operation aborted by user') {
        throw err;
      }

      const isGeneralException = err instanceof Error && err.message.includes('GeneralException');
      const isBusy =
        err instanceof Error &&
        (err.message === OFFICE_BUSY_TIMEOUT_MESSAGE || err.message.includes('busy'));

      const shouldRetry = (isGeneralException || isBusy) && attempt < maxRetries;

      if (!shouldRetry) {
        if (
          err instanceof Error &&
          err.message !== OFFICE_BUSY_TIMEOUT_MESSAGE &&
          err.message !== 'Operation aborted by user'
        ) {
          logService.error(`[OfficeAction] Error executing ${actionName}`, err, { actionName });
        }
        throw err;
      }

      logService.warn(
        `[OfficeAction] Retrying ${actionName} after error: ${err instanceof Error ? err.message : String(err)}. Attempt ${attempt + 1} of ${maxRetries}`,
        { actionName },
      );
      await delay(backoffDelays[attempt]);
    } finally {
      if (timeoutId) clearTimeout(timeoutId);
      // ERR-C4: abortListener is always defined when abortSignal is provided — safe to remove
      if (abortSignal && abortListener) {
        abortSignal.removeEventListener('abort', abortListener);
      }
    }
  }

  throw new Error('Unreachable');
}
