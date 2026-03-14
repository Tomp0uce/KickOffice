/**
 * Web Crypto API utilities for credential encryption/decryption.
 * Separated from credentialStorage.ts to keep encryption and storage concerns isolated.
 */

import { isCryptoAvailable } from './cryptoPolyfill'
import { logService } from './logger'

const ENCRYPTION_KEY_NAME = 'ko_encryption_key'

/**
 * Get or create an AES-GCM encryption key.
 * Stored in localStorage (remember=true) or sessionStorage (remember=false).
 * Checks both storages to survive migrations.
 */
export async function getEncryptionKey(remember: boolean): Promise<CryptoKey | null> {
  if (!isCryptoAvailable()) return null

  const primaryStorage = remember ? localStorage : sessionStorage
  const fallbackStorage = remember ? sessionStorage : localStorage

  let keyData = primaryStorage.getItem(ENCRYPTION_KEY_NAME)

  if (!keyData) {
    keyData = fallbackStorage.getItem(ENCRYPTION_KEY_NAME)
    if (keyData) {
      primaryStorage.setItem(ENCRYPTION_KEY_NAME, keyData)
      logService.info('[CredentialCrypto] Migrated encryption key to', remember ? 'localStorage' : 'sessionStorage')
    }
  }

  if (!keyData) {
    const key = await crypto.subtle.generateKey(
      { name: 'AES-GCM', length: 256 },
      true,
      ['encrypt', 'decrypt']
    )
    const exported = await crypto.subtle.exportKey('jwk', key)
    primaryStorage.setItem(ENCRYPTION_KEY_NAME, JSON.stringify(exported))
    return key
  }

  return crypto.subtle.importKey(
    'jwk',
    JSON.parse(keyData),
    { name: 'AES-GCM', length: 256 },
    true,
    ['encrypt', 'decrypt']
  )
}

/**
 * Encrypt a string. Falls back to plaintext if Web Crypto is unavailable.
 */
export async function encryptValue(plaintext: string, remember: boolean): Promise<string> {
  if (!plaintext) return ''

  try {
    const key = await getEncryptionKey(remember)
    if (!key) {
      logService.warn('[CredentialCrypto] Crypto not available, storing plaintext')
      return plaintext
    }

    const iv = crypto.getRandomValues(new Uint8Array(12))
    const encoded = new TextEncoder().encode(plaintext)
    const ciphertext = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, encoded)

    const combined = new Uint8Array(iv.length + ciphertext.byteLength)
    combined.set(iv, 0)
    combined.set(new Uint8Array(ciphertext), iv.length)
    return btoa(String.fromCharCode(...combined))
  } catch (error) {
    logService.error('[CredentialCrypto] Encryption failed', error instanceof Error ? error : new Error(String(error)))
    return plaintext
  }
}

/**
 * Decrypt a string. Returns empty string if crypto is unavailable or decryption fails.
 * @param credentialKey - Storage key name, used for cleanup on failure
 */
export async function decryptValue(encrypted: string, remember: boolean, credentialKey?: string): Promise<string> {
  if (!encrypted) return ''

  try {
    const cryptoKey = await getEncryptionKey(remember)

    if (!cryptoKey) {
      logService.warn('[CredentialCrypto] Crypto not available, treating as plaintext')
      try {
        atob(encrypted)
        // Looks like encrypted data but we can't decrypt it — clear it
        logService.error('[CredentialCrypto] Found encrypted data but crypto is not available')
        if (credentialKey) {
          logService.warn(`[CredentialCrypto] Clearing unusable encrypted data for key: ${credentialKey}`)
          localStorage.removeItem(`ko_cred_${credentialKey}`)
          sessionStorage.removeItem(`ko_cred_${credentialKey}`)
        }
        return ''
      } catch {
        return encrypted // Not base64 — treat as plaintext
      }
    }

    const combined = Uint8Array.from(atob(encrypted), c => c.charCodeAt(0))
    const iv = combined.slice(0, 12)
    const ciphertext = combined.slice(12)
    const decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, cryptoKey, ciphertext)
    return new TextDecoder().decode(decrypted)
  } catch (error) {
    logService.error('[CredentialCrypto] Decryption failed', error instanceof Error ? error : new Error(String(error)))
    if (credentialKey) {
      logService.warn(`[CredentialCrypto] Clearing corrupted data for key: ${credentialKey}`)
      localStorage.removeItem(`ko_cred_${credentialKey}`)
      sessionStorage.removeItem(`ko_cred_${credentialKey}`)
      sessionStorage.removeItem(credentialKey)
    }
    return ''
  }
}

/**
 * Clear encryption keys from both storages.
 */
export function clearEncryptionKeys(): void {
  localStorage.removeItem(ENCRYPTION_KEY_NAME)
  sessionStorage.removeItem(ENCRYPTION_KEY_NAME)
}
