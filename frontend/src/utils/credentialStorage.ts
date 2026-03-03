/**
 * Credential storage utility with Web Crypto API encryption
 * Uses sessionStorage by default for better security.
 * Can use localStorage if "remember credentials" is explicitly enabled.
 */

const STORAGE_PREFIX = 'ko_cred_'
const ENCRYPTION_KEY_NAME = 'ko_encryption_key'

/**
 * Get or create an encryption key using Web Crypto API
 */
async function getEncryptionKey(): Promise<CryptoKey> {
  let keyData = sessionStorage.getItem(ENCRYPTION_KEY_NAME)

  if (!keyData) {
    // Generate a new random key
    const key = await crypto.subtle.generateKey(
      { name: 'AES-GCM', length: 256 },
      true, // extractable
      ['encrypt', 'decrypt']
    )

    // Export and store the key
    const exported = await crypto.subtle.exportKey('jwk', key)
    keyData = JSON.stringify(exported)
    sessionStorage.setItem(ENCRYPTION_KEY_NAME, keyData)
    return key
  }

  // Import existing key
  const keyJwk = JSON.parse(keyData)
  return crypto.subtle.importKey(
    'jwk',
    keyJwk,
    { name: 'AES-GCM', length: 256 },
    true,
    ['encrypt', 'decrypt']
  )
}

/**
 * Encrypt a string using Web Crypto API
 */
async function encryptValue(plaintext: string): Promise<string> {
  if (!plaintext) return ''

  try {
    const key = await getEncryptionKey()
    const iv = crypto.getRandomValues(new Uint8Array(12)) // 12 bytes for AES-GCM
    const encoded = new TextEncoder().encode(plaintext)

    const ciphertext = await crypto.subtle.encrypt(
      { name: 'AES-GCM', iv },
      key,
      encoded
    )

    // Combine IV and ciphertext
    const combined = new Uint8Array(iv.length + ciphertext.byteLength)
    combined.set(iv, 0)
    combined.set(new Uint8Array(ciphertext), iv.length)

    // Convert to base64
    return btoa(String.fromCharCode(...combined))
  } catch (error) {
    console.error('[CredentialStorage] Encryption failed:', error)
    // Fallback to plaintext if encryption fails (shouldn't happen)
    return plaintext
  }
}

/**
 * Decrypt a string using Web Crypto API
 */
async function decryptValue(encrypted: string): Promise<string> {
  if (!encrypted) return ''

  try {
    const key = await getEncryptionKey()

    // Decode from base64
    const combined = Uint8Array.from(atob(encrypted), c => c.charCodeAt(0))

    // Extract IV and ciphertext
    const iv = combined.slice(0, 12)
    const ciphertext = combined.slice(12)

    const decrypted = await crypto.subtle.decrypt(
      { name: 'AES-GCM', iv },
      key,
      ciphertext
    )

    return new TextDecoder().decode(decrypted)
  } catch (error) {
    console.error('[CredentialStorage] Decryption failed:', error)
    // Return empty string if decryption fails
    return ''
  }
}

/**
 * Safe localStorage setter with quota handling
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
 * Get credential from the appropriate storage (with decryption for localStorage)
 */
async function getCredential(key: string, remember: boolean): Promise<string> {
  if (remember) {
    const encrypted = localStorage.getItem(`${STORAGE_PREFIX}${key}`)
    if (!encrypted) return ''
    return decryptValue(encrypted)
  }
  return sessionStorage.getItem(key) || ''
}

/**
 * Set credential in the appropriate storage (with encryption for localStorage)
 */
async function setCredential(key: string, value: string, remember: boolean): Promise<void> {
  if (remember) {
    if (value) {
      const encrypted = await encryptValue(value)
      safeSetItem(`${STORAGE_PREFIX}${key}`, encrypted)
    } else {
      localStorage.removeItem(`${STORAGE_PREFIX}${key}`)
    }
    // Clear from sessionStorage if exists
    sessionStorage.removeItem(key)
  } else {
    if (value) {
      // sessionStorage doesn't need encryption (session-only)
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
export async function setRememberCredentials(value: boolean): Promise<void> {
  const wasRemembering = getRememberCredentials()
  safeSetItem('rememberCredentials', value ? 'true' : 'false')

  if (value && !wasRemembering) {
    // Migrate from sessionStorage to localStorage (with encryption)
    const key = sessionStorage.getItem('litellmUserKey')
    const email = sessionStorage.getItem('litellmUserEmail')
    if (key) {
      const encrypted = await encryptValue(key)
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, encrypted)
      sessionStorage.removeItem('litellmUserKey')
    }
    if (email) {
      const encrypted = await encryptValue(email)
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, encrypted)
      sessionStorage.removeItem('litellmUserEmail')
    }
  } else if (!value && wasRemembering) {
    // Migrate from localStorage to sessionStorage (with decryption)
    const key = await getCredential('litellmUserKey', true)
    const email = await getCredential('litellmUserEmail', true)
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
  // Clear from both storages
  sessionStorage.removeItem('litellmUserKey')
  sessionStorage.removeItem('litellmUserEmail')
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserKey`)
  localStorage.removeItem(`${STORAGE_PREFIX}litellmUserEmail`)
  // Clear encryption key
  sessionStorage.removeItem(ENCRYPTION_KEY_NAME)
}

/**
 * Migrate credentials from old plaintext format (for backward compatibility)
 */
export async function migrateFromPlaintext(): Promise<void> {
  const remember = getRememberCredentials()
  if (remember) {
    // Check if we have plaintext credentials that should be encrypted
    const key = localStorage.getItem(`${STORAGE_PREFIX}litellmUserKey`)
    const email = localStorage.getItem(`${STORAGE_PREFIX}litellmUserEmail`)

    // Check if they look like plaintext (not base64 encrypted)
    const isPlaintext = (value: string | null) => {
      if (!value) return false
      try {
        // If it's valid base64 and long enough to be encrypted, assume it's encrypted
        atob(value)
        return value.length < 20 // Encrypted values should be longer
      } catch {
        return true // Not base64, so plaintext
      }
    }

    if (key && isPlaintext(key)) {
      const encrypted = await encryptValue(key)
      safeSetItem(`${STORAGE_PREFIX}litellmUserKey`, encrypted)
    }
    if (email && isPlaintext(email)) {
      const encrypted = await encryptValue(email)
      safeSetItem(`${STORAGE_PREFIX}litellmUserEmail`, encrypted)
    }
  }
}
