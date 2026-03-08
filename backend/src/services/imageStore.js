import crypto from 'crypto'

const store = new Map()
const TTL_MS = 30 * 60 * 1000 // 30 minutes

/**
 * Store an image buffer and return a unique ID.
 * Old entries are automatically purged on each call.
 */
export function storeImage(buffer, mimeType) {
  // Purge expired entries
  const now = Date.now()
  for (const [key, val] of store.entries()) {
    if (now - val.createdAt > TTL_MS) store.delete(key)
  }

  const id = crypto.randomUUID()
  store.set(id, { buffer, mimeType, createdAt: now })
  return id
}

/**
 * Retrieve an image buffer by ID. Returns null if not found or expired.
 */
export function getImage(id) {
  const entry = store.get(id)
  if (!entry) return null
  if (Date.now() - entry.createdAt > TTL_MS) {
    store.delete(id)
    return null
  }
  return entry
}
