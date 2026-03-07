/**
 * Credential storage utility.
 * Storage layer: selects between localStorage (encrypted) and sessionStorage (plaintext).
 * Encryption layer: delegated to credentialCrypto.ts.
 */

import { logCryptoStatus } from './cryptoPolyfill'
import { encryptValue, decryptValue, clearEncryptionKeys } from './credentialCrypto'

const STORAGE_PREFIX = 'ko_cred_'

// Log crypto availability on module load
logCryptoStatus()

/**
 * Safe localStorage setter with quota handling.
 */
function safeSetItem(key: string, value: string): void {
  try {
    localStorage.setItem(key, value)
  } catch (e) {
    if (e instanceof DOMException && e.name === 'QuotaExceededError') {
      console.warn('[CredentialStorage] localStorage quota exceeded — credential not saved:', key)
    } else {
      throw e
    }
  }
}

/**
 * Get credential from the appropriate storage (with decryption for localStorage).
 * Checks both storages as fallback to survive migrations.
 */
async function getCredential(key: string, remember: boolean): Promise<string> {
  if (remember) {
    const encrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`)
    if (encrypted) {
      const decrypted = await decryptValue(encrypted, true, key)
      if (decrypted) return decrypted
    }

    // Fallback: migrate from sessionStorage
    const sessionValue = sessionStorage.getItem(key)
    if (sessionValue) {
      console.info(`[CredentialStorage] Migrating "${key}" from sessionStorage to localStorage`)
      await setCredential(key, sessionValue, true)
      return sessionValue
    }

    return ''
  }

  // sessionStorage path
  const sessionValue = sessionStorage.getItem(key)
  if (sessionValue) return sessionValue

  // Fallback: migrate from encrypted localStorage
  const encrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`)
  if (encrypted) {
    const decrypted = await decryptValue(encrypted, false, key)
    if (decrypted) {
      console.info(`[CredentialStorage] Migrating "${key}" from localStorage to sessionStorage`)
      sessionStorage.setItem(key, decrypted)
      return decrypted
    }
  }

  return ''
}

/**
 * Set credential in the appropriate storage (encrypted localStorage or plaintext sessionStorage).
 */
async function setCredential(key: string, value: string, remember: boolean): Promise<void> {
  if (remember) {
    if (value) {
      const encrypted = await encryptValue(value, true)
      safeSetItem(`${STORAGE_PREFIX}${key}`, encrypted)
    } else {
      localStorage.removeItem(`${STORAGE_PREFIX}${key}`)
    }
    sessionStorage.removeItem(key)
  } else {
    if (value) {
      sessionStorage.setItem(key, value)
    } else {
      sessionStorage.removeItem(key)
    }
    localStorage.removeItem(`${STORAGE_PREFIX}${key}`)
  }
}

/**
 * Check if "remember credentials" is enabled.
 * Defaults to true for Office Add-ins (sessionStorage is wiped on restart).
 */
export function getRememberCredentials(): boolean {
  const value = localStorage.getItem('rememberCredentials')
  return value === null ? true : value === 'true'
}

/**
 * Set "remember credentials" preference and migrate data accordingly.
 */
export async function setRememberCredentials(value: boolean): Promise<void> {
  const wasRemembering = getRememberCredentials()
  safeSetItem('rememberCredentials', value ? 'true' : 'false')

  if (value && !wasRemembering) {
    // Migrate from sessionStorage to encrypted localStorage
    const key = sessionStorage.getItem('litellmUserKey')
    const email = sessionStorage.getItem('litellmUserEmail')

    if (key) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, await encryptValue(key, true))
      sessionStorage.removeItem('litellmUserKey')
    }
    if (email) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, await encryptValue(email, true))
      sessionStorage.removeItem('litellmUserEmail')
    }
  } else if (!value && wasRemembering) {
    // Migrate from encrypted localStorage to sessionStorage
    const key = await getCredential('litellmUserKey', true)
    const email = await getCredential('litellmUserEmail', true)

    if (key) sessionStorage.setItem('litellmUserKey', key)
    if (email) sessionStorage.setItem('litellmUserEmail', email)

    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`)
    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`)
  }
}

export async function getUserKey(): Promise<string> {
  return getCredential('litellmUserKey', getRememberCredentials())
}

export async function setUserKey(value: string): Promise<void> {
  await setCredential('litellmUserKey', value, getRememberCredentials())
}

export async function getUserEmail(): Promise<string> {
  return getCredential('litellmUserEmail', getRememberCredentials())
}

export async function setUserEmail(value: string): Promise<void> {
  await setCredential('litellmUserEmail', value, getRememberCredentials())
}

export function clearCredentials(): void {
  sessionStorage.removeItem('litellmUserKey')
  sessionStorage.removeItem('litellmUserEmail')
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`)
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`)
  clearEncryptionKeys()
}

/**
 * Migrate credentials from old plaintext format (backward compatibility).
 */
export async function migrateFromPlaintext(): Promise<void> {
  const remember = getRememberCredentials()
  if (!remember) return

  const isPlaintext = (value: string | null): boolean => {
    if (!value) return false
    try {
      atob(value)
      return value.length < 20
    } catch {
      return true
    }
  }

  const key = localStorage.getItem(`${STORAGE_PREFIX}litellmUserKey`)
  const email = localStorage.getItem(`${STORAGE_PREFIX}litellmUserEmail`)

  if (key && isPlaintext(key)) {
    safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, await encryptValue(key, true))
  }
  if (email && isPlaintext(email)) {
    safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, await encryptValue(email, true))
  }
}

/**
 * Check if credentials are configured.
 */
export async function areCredentialsConfigured(): Promise<boolean> {
  try {
    const key = await getUserKey()
    const email = await getUserEmail()
    const hasCredentials = key.length > 0 && email.length > 0
    if (!hasCredentials) {
      console.warn('[CredentialStorage] Credentials not configured - key or email missing')
    }
    return hasCredentials
  } catch (error) {
    console.error('[CredentialStorage] Error checking credentials:', error)
    return false
  }
}
