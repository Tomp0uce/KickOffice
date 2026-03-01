/**
 * Secure credential storage utility
 * Uses localStorage with obfuscation for persistent storage
 * Falls back to sessionStorage for session-only storage
 */

const OBFUSCATION_KEY = 'K1ck0ff1c3' // Simple obfuscation key

/**
 * Simple XOR-based obfuscation (not encryption, just makes credentials non-plaintext)
 */
function obfuscate(text: string): string {
  if (!text) return ''
  let result = ''
  for (let i = 0; i < text.length; i++) {
    result += String.fromCharCode(text.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length))
  }
  return btoa(result) // Base64 encode the result
}

/**
 * Reverse the obfuscation
 */
function deobfuscate(encoded: string): string {
  if (!encoded) return ''
  try {
    const decoded = atob(encoded)
    let result = ''
    for (let i = 0; i < decoded.length; i++) {
      result += String.fromCharCode(decoded.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length))
    }
    return result
  } catch {
    return ''
  }
}

const STORAGE_PREFIX = 'ko_cred_'

/**
 * Get credential from the appropriate storage
 */
function getCredential(key: string, remember: boolean): string {
  if (remember) {
    const stored = localStorage.getItem(`${STORAGE_PREFIX}${key}`)
    return stored ? deobfuscate(stored) : ''
  }
  return sessionStorage.getItem(key) || ''
}

/**
 * Set credential in the appropriate storage
 */
function setCredential(key: string, value: string, remember: boolean): void {
  if (remember) {
    if (value) {
      localStorage.setItem(`${STORAGE_PREFIX}${key}`, obfuscate(value))
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
  localStorage.setItem('rememberCredentials', value ? 'true' : 'false')

  if (value && !wasRemembering) {
    // Migrate from sessionStorage to localStorage
    const key = sessionStorage.getItem('litellmUserKey')
    const email = sessionStorage.getItem('litellmUserEmail')
    if (key) {
      localStorage.setItem(`${STORAGE_PREFIX}litellmUserKey`, obfuscate(key))
      sessionStorage.removeItem('litellmUserKey')
    }
    if (email) {
      localStorage.setItem(`${STORAGE_PREFIX}litellmUserEmail`, obfuscate(email))
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
      localStorage.setItem(`${STORAGE_PREFIX}litellmUserKey`, obfuscate(sessionKey))
      sessionStorage.removeItem('litellmUserKey')
    }
    if (sessionEmail && !localStorage.getItem(`${STORAGE_PREFIX}litellmUserEmail`)) {
      localStorage.setItem(`${STORAGE_PREFIX}litellmUserEmail`, obfuscate(sessionEmail))
      sessionStorage.removeItem('litellmUserEmail')
    }
  }
}
