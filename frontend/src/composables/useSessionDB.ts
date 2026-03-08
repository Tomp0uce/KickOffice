/**
 * Lightweight IndexedDB helper for KickOffice session management.
 * Adapted from OpenExcel's db.ts (open-excel-main).
 * Includes per-session VFS (Virtual File System) persistence.
 */
import type { DisplayMessage } from '@/types/chat'
import type { LogEntry } from '@/utils/logger'
import { snapshotVfs, restoreVfs } from '@/utils/vfs'
import { randomUUID } from '@/utils/cryptoPolyfill'
import { message as messageUtil } from '@/utils/message'
import { i18n } from '@/i18n'

export interface VfsFile {
  path: string
  data: Uint8Array
}

export interface ChatSession {
  id: string
  hostType: string
  name: string
  messages: DisplayMessage[]
  vfsFiles?: VfsFile[]
  createdAt: number
  updatedAt: number
}

const DB_NAME = 'KickOfficeDB'
const DB_VERSION = 2
const SESSIONS_STORE = 'sessions'
const LOGS_STORE = 'logs'

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
      if (!db.objectStoreNames.contains(LOGS_STORE)) {
        const logsStore = db.createObjectStore(LOGS_STORE, { keyPath: 'id', autoIncrement: true })
        logsStore.createIndex('sessionId', 'sessionId', { unique: false })
        logsStore.createIndex('timestamp', 'timestamp', { unique: false })
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

function formatSessionDate(timestamp: number): string {
  const d = new Date(timestamp)
  const dd = String(d.getDate()).padStart(2, '0')
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  const yy = String(d.getFullYear()).slice(-2)
  const hh = String(d.getHours()).padStart(2, '0')
  const min = String(d.getMinutes()).padStart(2, '0')
  return `${dd}/${mm}/${yy} ${hh}:${min}`
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
    id: randomUUID(),
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
  const session = await idbGet<ChatSession>(tx.objectStore(SESSIONS_STORE), sessionId)
  if (session) {
    // Restore VFS state for this session
    try {
      await restoreVfs(session.vfsFiles || [])
    } catch (e) {
      console.warn('[SessionDB] Failed to restore VFS for session', sessionId, e)
    }
  }
  return session
}

export async function saveSession(sessionId: string, messages: DisplayMessage[]): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readwrite')
  const store = tx.objectStore(SESSIONS_STORE)
  const session = await idbGet<ChatSession>(store, sessionId)
  if (!session) return
  // Deep-clone to strip Vue reactive proxies — IDB's structured clone algorithm cannot serialize them
  let plainMessages: DisplayMessage[] = JSON.parse(JSON.stringify(messages))
  
  // H6 Fix: Prune chat history to prevent unbounded growth max (200)
  const MAX_HISTORY = 200
  if (plainMessages.length > MAX_HISTORY) {
    plainMessages = plainMessages.slice(plainMessages.length - MAX_HISTORY)
  }

  const newName = (session.name === 'New Chat') ? formatSessionDate(session.createdAt) : session.name
  // Snapshot VFS state for this session
  let vfsFiles: VfsFile[] = []
  try {
    vfsFiles = await snapshotVfs()
  } catch (e) {
    console.warn('[SessionDB] Failed to snapshot VFS', e)
  }

  // H2 Fix: Catch IndexedDB QuotaExceeded errors
  try {
    await idbPut(store, { ...session, messages: plainMessages, vfsFiles, name: newName, updatedAt: Date.now() })
  } catch (putErr) {
    console.error('[SessionDB] Failed to save session:', putErr)
    messageUtil.error((i18n.global.t as any)('failedToSaveChatData'))
  }
}

export async function deleteSession(sessionId: string): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(SESSIONS_STORE, 'readwrite')
  await idbDelete(tx.objectStore(SESSIONS_STORE), sessionId)
}

export function getSessionMessageCount(session: ChatSession): number {
  return session.messages.filter(m => m.role === 'user' || m.role === 'assistant').length
}

export async function appendLogEntry(entry: LogEntry): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(LOGS_STORE, 'readwrite')
  await idbAdd(tx.objectStore(LOGS_STORE), entry)
}

export async function getLogsForSession(sessionId: string): Promise<LogEntry[]> {
  const db = await openDB()
  const tx = db.transaction(LOGS_STORE, 'readonly')
  return idbGetAllByIndex<LogEntry>(tx.objectStore(LOGS_STORE), 'sessionId', sessionId)
}

export async function pruneOldLogs(maxEntries: number = 1000): Promise<void> {
  const db = await openDB()
  const tx = db.transaction(LOGS_STORE, 'readwrite')
  const store = tx.objectStore(LOGS_STORE)
  const allEntries = await new Promise<Array<LogEntry & { id: number }>>((resolve, reject) => {
    const index = store.index('timestamp')
    const req = index.getAll()
    req.onsuccess = () => resolve(req.result as Array<LogEntry & { id: number }>)
    req.onerror = () => reject(req.error)
  })
  if (allEntries.length <= maxEntries) return
  const toDelete = allEntries.slice(0, allEntries.length - maxEntries)
  for (const entry of toDelete) {
    await new Promise<void>((resolve, reject) => {
      const req = store.delete(entry.id)
      req.onsuccess = () => resolve()
      req.onerror = () => reject(req.error)
    })
  }
}
