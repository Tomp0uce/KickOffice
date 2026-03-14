/**
 * Crypto API polyfills and fallbacks
 * Provides fallback implementations when Web Crypto API is not available
 */

import { logService } from './logger'

/**
 * Generate a UUID (v4)
 * Falls back to a simple implementation if crypto.randomUUID is not available
 */
export function randomUUID(): string {
  // Try native implementation first
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID()
  }

  // Fallback implementation (RFC 4122 version 4)
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0
    const v = c === 'x' ? r : (r & 0x3) | 0x8
    return v.toString(16)
  })
}

/**
 * Check if Web Crypto API is available
 * crypto.subtle is only available in secure contexts (HTTPS)
 */
export function isCryptoAvailable(): boolean {
  return typeof crypto !== 'undefined' && typeof crypto.subtle !== 'undefined'
}

/**
 * Get a warning message if crypto is not available
 */
export function getCryptoWarning(): string {
  if (!isCryptoAvailable()) {
    return 'Web Crypto API not available. Credentials will be stored unencrypted. Please ensure the app is running over HTTPS.'
  }
  return ''
}

/**
 * Log crypto availability status
 */
export function logCryptoStatus(): void {
  if (isCryptoAvailable()) {
    logService.info('[CryptoPolyfill] Web Crypto API available - using encrypted storage')
  } else {
    logService.warn('[CryptoPolyfill] Web Crypto API NOT available - using unencrypted storage')
    logService.warn('[CryptoPolyfill]', getCryptoWarning())
  }
}
