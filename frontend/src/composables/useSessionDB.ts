/**
 * Lightweight IndexedDB helper for KickOffice session management.
 * Adapted from OpenExcel's db.ts (open-excel-main).
 */
import type { DisplayMessage } from '@/types/chat'

export interface ChatSession {
  id: string
  hostType: string
  name: string
  messages: DisplayMessage[]
  createdAt: number
  updatedAt: number
}

const DB_NAME = 'KickOfficeDB'
const DB_VERSION = 1
const SESSIONS_STORE = 'sessions'

let dbInstance: IDBDatabase | null = null

function openDB(): Promise<IDBDatabase> {
  if (dbInstance) return Promise.resolve(dbInstance)

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION)

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result
      if (!db.objectStoreNames.contains(SESSIONS_STORE)) {
        const store = db.createObjectStore(SESSIONS_STORE, { keyPath: 'id' })
        store.createIndex('hostType', 'hostType', { unique: false })
        store.createIndex('updatedAt', 'updatedAt', { unique: false })
      }
    }

    request.onsuccess = (event) => {
      dbInstance = (event.target as IDBOpenDBRequest).result
      resolve(dbInstance)
    }

    request.onerror = (event) => {
      reject((event.target as IDBOpenDBRequest).error)
    }
  })
}

function idbGet<T>(store: IDBObjectStore, key: string): Promise<T | undefined> {
  return new Promise((resolve, reject) => {
    const req = store.get(key)
    req.onsuccess = () => resolve(req.result as T | undefined)
    req.onerror = () => reject(req.error)
  })
}

function idbPut(store: IDBObjectStore, value: unknown): Promise<void> {
  return new Promise((resolve, reject) => {
    const req = store.put(value)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

function idbAdd(store: IDBObjectStore, value: unknown): Promise<void> {
  return new Promise((resolve, reject) => {
    const req = store.add(value)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

function idbDelete(store: IDBObjectStore, key: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const req = store.delete(key)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

function idbGetAllByIndex<T>(store: IDBObjectStore, indexName: string, value: string): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const index = store.index(indexName)
    const req = index.getAll(value)
    req.onsuccess = () => resolve(req.result as T[])
    req.onerror = () => reject(req.error)
  })
}

function deriveSessionName(messages: DisplayMessage[]): string | null {
  const firstUser = messages.find(m => m.role === 'user')
  if (!firstUser) return null
  const text = firstUser.content.trim()
  if (!text) return null
  return text.length > 40 ? `${text.slice(0, 37)}...` : text
}

export async function listSessions(hostType: string): Promise<ChatSession[]> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readonly')
  const store = tx.objectStore(SESSIONS_STORE)
  const sessions = await idbGetAllByIndex<ChatSession>(store, 'hostType', hostType)
  sessions.sort((a, b) => b.updatedAt - a.updatedAt)
  return sessions
}

export async function createSession(hostType: string, name?: string): Promise<ChatSession> {
  const db = await openDB()
  const now = Date.now()
  const session: ChatSession = {
    id: crypto.randomUUID(),
    hostType,
    name: name ?? 'New Chat',
    messages: [],
    createdAt: now,
    updatedAt: now,
  }
  const tx = db.transaction(SESSIONS_STORE, 'readwrite')
  await idbAdd(tx.objectStore(SESSIONS_STORE), session)
  return session
}

export async function getSession(sessionId: string): Promise<ChatSession | undefined> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readonly')
  return idbGet<ChatSession>(tx.objectStore(SESSIONS_STORE), sessionId)
}

export async function saveSession(sessionId: string, messages: DisplayMessage[]): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readwrite')
  const store = tx.objectStore(SESSIONS_STORE)
  const session = await idbGet<ChatSession>(store, sessionId)
  if (!session) return
  let name = session.name
  if (name === 'New Chat') {
    const derived = deriveSessionName(messages)
    if (derived) name = derived
  }
  await idbPut(store, { ...session, messages, name, updatedAt: Date.now() })
}

export async function deleteSession(sessionId: string): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readwrite')
  await idbDelete(tx.objectStore(SESSIONS_STORE), sessionId)
}

export function getSessionMessageCount(session: ChatSession): number {
  return session.messages.filter(m => m.role === 'user' || m.role === 'assistant').length
}
