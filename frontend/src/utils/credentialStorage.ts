/**
 * Credential storage utility.
 * Storage layer: selects between localStorage (encrypted) and sessionStorage (plaintext).
 * Encryption layer: delegated to credentialCrypto.ts.
 */

import { logCryptoStatus } from './cryptoPolyfill';
import { encryptValue, decryptValue, clearEncryptionKeys } from './credentialCrypto';
import { logService } from './logger';

const STORAGE_PREFIX = 'ko_cred_';

let _cryptoStatusLogged = false;
function ensureCryptoStatusLogged(): void {
  if (!_cryptoStatusLogged) {
    _cryptoStatusLogged = true;
    logCryptoStatus();
  }
}

/**
 * Safe localStorage setter with quota handling.
 */
function safeSetItem(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    if (e instanceof DOMException && e.name === 'QuotaExceededError') {
      logService.warn(
        '[CredentialStorage] localStorage quota exceeded — credential not saved:',
        key,
      );
    } else {
      throw e;
    }
  }
}

/**
 * ARCH-M3: One-time credential migration at app startup
 * Migrates credentials between localStorage and sessionStorage based on "remember" preference.
 * This runs once at startup instead of on every credential read.
 */
export async function migrateCredentialsOnStartup(): Promise<void> {
  ensureCryptoStatusLogged();
  const remember = getRememberCredentials();
  const credentialKeys = ['litellmUserKey', 'litellmUserEmail'];

  for (const key of credentialKeys) {
    if (remember) {
      // Should use localStorage - migrate from sessionStorage if needed
      const sessionValue = sessionStorage.getItem(key);
      const localEncrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`);

      if (sessionValue && !localEncrypted) {
        try {
          const encrypted = await encryptValue(sessionValue, true);
          safeSetItem(`${STORAGE_PREFIX}${key}`, encrypted);
          sessionStorage.removeItem(key);
          logService.info(
            `[CredentialStorage] Migrated "${key}" from sessionStorage to localStorage`,
          );
        } catch (err) {
          logService.error(
            `[CredentialStorage] Failed to migrate "${key}" to localStorage`,
            err instanceof Error ? err : new Error(String(err)),
          );
        }
      }
    } else {
      // Should use sessionStorage - migrate from localStorage if needed
      const localEncrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`);
      const sessionValue = sessionStorage.getItem(key);

      if (localEncrypted && !sessionValue) {
        try {
          const decrypted = await decryptValue(localEncrypted, false, key);
          if (decrypted) {
            sessionStorage.setItem(key, decrypted);
            localStorage.removeItem(`${STORAGE_PREFIX}${key}`);
            logService.info(
              `[CredentialStorage] Migrated "${key}" from localStorage to sessionStorage`,
            );
          }
        } catch (err) {
          logService.error(
            `[CredentialStorage] Failed to migrate "${key}" to sessionStorage`,
            err instanceof Error ? err : new Error(String(err)),
          );
        }
      }
    }
  }
}

/**
 * ARCH-M3: Simplified credential retrieval without inline migration
 * Get credential from the appropriate storage (with decryption for localStorage).
 * Migration is now handled by migrateCredentialsOnStartup() called at app init.
 */
async function getCredential(key: string, remember: boolean): Promise<string> {
  if (remember) {
    const encrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`);
    if (encrypted) {
      const decrypted = await decryptValue(encrypted, true, key);
      if (decrypted) return decrypted;
    }
    return '';
  }

  // sessionStorage path
  const sessionValue = sessionStorage.getItem(key);
  return sessionValue || '';
}

/**
 * Set credential in the appropriate storage (encrypted localStorage or plaintext sessionStorage).
 */
async function setCredential(key: string, value: string, remember: boolean): Promise<void> {
  if (remember) {
    if (value) {
      const encrypted = await encryptValue(value, true);
      safeSetItem(`${STORAGE_PREFIX}${key}`, encrypted);
    } else {
      localStorage.removeItem(`${STORAGE_PREFIX}${key}`);
    }
    sessionStorage.removeItem(key);
  } else {
    if (value) {
      sessionStorage.setItem(key, value);
    } else {
      sessionStorage.removeItem(key);
    }
    localStorage.removeItem(`${STORAGE_PREFIX}${key}`);
  }
}

/**
 * Check if "remember credentials" is enabled.
 * Defaults to true for Office Add-ins (sessionStorage is wiped on restart).
 */
export function getRememberCredentials(): boolean {
  const value = localStorage.getItem('rememberCredentials');
  return value === null ? true : value === 'true';
}

/**
 * Set "remember credentials" preference and migrate data accordingly.
 */
export async function setRememberCredentials(value: boolean): Promise<void> {
  const wasRemembering = getRememberCredentials();
  safeSetItem('rememberCredentials', value ? 'true' : 'false');

  if (value && !wasRemembering) {
    // Migrate from sessionStorage to encrypted localStorage
    const key = sessionStorage.getItem('litellmUserKey');
    const email = sessionStorage.getItem('litellmUserEmail');

    if (key) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, await encryptValue(key, true));
      sessionStorage.removeItem('litellmUserKey');
    }
    if (email) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, await encryptValue(email, true));
      sessionStorage.removeItem('litellmUserEmail');
    }
  } else if (!value && wasRemembering) {
    // Migrate from encrypted localStorage to sessionStorage
    const key = await getCredential('litellmUserKey', true);
    const email = await getCredential('litellmUserEmail', true);

    if (key) sessionStorage.setItem('litellmUserKey', key);
    if (email) sessionStorage.setItem('litellmUserEmail', email);

    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`);
    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`);
  }
}

export async function getUserKey(): Promise<string> {
  return getCredential('litellmUserKey', getRememberCredentials());
}

export async function setUserKey(value: string): Promise<void> {
  await setCredential('litellmUserKey', value, getRememberCredentials());
}

export async function getUserEmail(): Promise<string> {
  return getCredential('litellmUserEmail', getRememberCredentials());
}

export async function setUserEmail(value: string): Promise<void> {
  await setCredential('litellmUserEmail', value, getRememberCredentials());
}

export function clearCredentials(): void {
  sessionStorage.removeItem('litellmUserKey');
  sessionStorage.removeItem('litellmUserEmail');
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`);
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`);
  clearEncryptionKeys();
}

/**
 * Migrate credentials from old plaintext format (backward compatibility).
 */
export async function migrateFromPlaintext(): Promise<void> {
  const remember = getRememberCredentials();
  if (!remember) return;

  const isPlaintext = (value: string | null): boolean => {
    if (!value) return false;
    try {
      atob(value);
      return value.length < 20;
    } catch {
      return true;
    }
  };

  const key = localStorage.getItem(`${STORAGE_PREFIX}litellmUserKey`);
  const email = localStorage.getItem(`${STORAGE_PREFIX}litellmUserEmail`);

  if (key && isPlaintext(key)) {
    safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, await encryptValue(key, true));
  }
  if (email && isPlaintext(email)) {
    safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, await encryptValue(email, true));
  }
}

/**
 * Check if credentials are configured.
 */
export async function areCredentialsConfigured(): Promise<boolean> {
  try {
    const key = await getUserKey();
    const email = await getUserEmail();
    const hasCredentials = key.length > 0 && email.length > 0;
    if (!hasCredentials) {
      logService.warn('[CredentialStorage] Credentials not configured - key or email missing');
    }
    return hasCredentials;
  } catch (error) {
    logService.error(
      '[CredentialStorage] Error checking credentials',
      error instanceof Error ? error : new Error(String(error)),
    );
    return false;
  }
}
