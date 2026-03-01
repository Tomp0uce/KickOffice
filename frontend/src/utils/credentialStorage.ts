/**
 * Credential storage utility
 * Uses sessionStorage by default for better security.
 * Can use localStorage if "remember credentials" is explicitly enabled.
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

const STORAGE_PREFIX = 'ko_cred_'

/**
 * Get credential from the appropriate storage
 */
function getCredential(key: string, remember: boolean): string {
  if (remember) {
    return localStorage.getItem(`${STORAGE_PREFIX}${key}`) || ''
  }
  return sessionStorage.getItem(key) || ''
}

/**
 * Set credential in the appropriate storage
 */
function setCredential(key: string, value: string, remember: boolean): void {
  if (remember) {
    if (value) {
      safeSetItem(`${STORAGE_PREFIX}${key}`, value)
    } else {
      localStorage.removeItem(`${STORAGE_PREFIX}${key}`)
    }
    // Clear from sessionStorage if exists
    sessionStorage.removeItem(key)
  } else {
    if (value) {
      sessionStorage.setItem(key, value)
    } else {
      sessionStorage.removeItem(key)
    }
    // Clear from localStorage if exists
    localStorage.removeItem(`${STORAGE_PREFIX}${key}`)
  }
}

/**
 * Check if "remember credentials" is enabled
 */
export function getRememberCredentials(): boolean {
  return localStorage.getItem('rememberCredentials') === 'true'
}

/**
 * Set "remember credentials" preference and migrate data accordingly
 */
export function setRememberCredentials(value: boolean): void {
  const wasRemembering = getRememberCredentials()
  safeSetItem('rememberCredentials', value ? 'true' : 'false')

  if (value && !wasRemembering) {
    // Migrate from sessionStorage to localStorage
    const key = sessionStorage.getItem('litellmUserKey')
    const email = sessionStorage.getItem('litellmUserEmail')
    if (key) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, key)
      sessionStorage.removeItem('litellmUserKey')
    }
    if (email) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, email)
      sessionStorage.removeItem('litellmUserEmail')
    }
  } else if (!value && wasRemembering) {
    // Migrate from localStorage to sessionStorage
    const key = getCredential('litellmUserKey', true)
    const email = getCredential('litellmUserEmail', true)
    if (key) {
      sessionStorage.setItem('litellmUserKey', key)
    }
    if (email) {
      sessionStorage.setItem('litellmUserEmail', email)
    }
    // Clear from localStorage
    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`)
    localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`)
  }
}

export function getUserKey(): string {
  return getCredential('litellmUserKey', getRememberCredentials())
}

export function setUserKey(value: string): void {
  setCredential('litellmUserKey', value, getRememberCredentials())
}

export function getUserEmail(): string {
  return getCredential('litellmUserEmail', getRememberCredentials())
}

export function setUserEmail(value: string): void {
  setCredential('litellmUserEmail', value, getRememberCredentials())
}

export function clearCredentials(): void {
  // Clear from both storages
  sessionStorage.removeItem('litellmUserKey')
  sessionStorage.removeItem('litellmUserEmail')
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`)
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`)
}

/**
 * Migrate credentials from old sessionStorage format (for backward compatibility)
 */
export function migrateFromSessionStorage(): void {
  const remember = getRememberCredentials()
  if (remember) {
    // Check if we have credentials in sessionStorage that should be in localStorage
    const sessionKey = sessionStorage.getItem('litellmUserKey')
    const sessionEmail = sessionStorage.getItem('litellmUserEmail')
    if (sessionKey && !localStorage.getItem(`${STORAGE_PREFIX}litellmUserKey`)) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, sessionKey)
      sessionStorage.removeItem('litellmUserKey')
    }
    if (sessionEmail && !localStorage.getItem(`${STORAGE_PREFIX}litellmUserEmail`)) {
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, sessionEmail)
      sessionStorage.removeItem('litellmUserEmail')
    }
  }
}
