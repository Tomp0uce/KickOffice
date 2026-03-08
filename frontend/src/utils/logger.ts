import { detectOfficeHost } from './hostDetection'
import { getUserEmail } from './credentialStorage'
import { appendLogEntry } from '@/composables/useSessionDB'

export interface LogEntry {
  timestamp: string
  level: 'error' | 'warn' | 'info' | 'debug'
  source: 'frontend'
  traffic: 'user' | 'llm' | 'auto' | 'system'
  sessionId?: string
  userId?: string
  host?: string
  message: string
  data?: Record<string, unknown>
  error?: {
    name: string
    message: string
    stack?: string
  }
}

class LogService {
  private buffer: Map<string, LogEntry[]> = new Map()
  private readonly MAX_ENTRIES = 500
  private _currentSessionId: string = 'default'

  public readonly originalConsole = {
    log: console.log,
    info: console.info,
    warn: console.warn,
    error: console.error,
    debug: console.debug
  }

  setCurrentSessionId(id: string) {
    this._currentSessionId = id
  }

  async getContext() {
    return {
      host: detectOfficeHost(),
      sessionId: this._currentSessionId,
      userId: (await getUserEmail()) || 'anonymous'
    }
  }

  private async addEntry(
    level: LogEntry['level'],
    message: string,
    traffic: LogEntry['traffic'] = 'system',
    data?: any,
    errorObj?: any
  ) {
    const ctx = await this.getContext()
    const entry: LogEntry = {
      timestamp: new Date().toISOString(),
      level,
      source: 'frontend',
      traffic,
      sessionId: ctx.sessionId,
      userId: ctx.userId,
      host: ctx.host,
      message,
    }

    if (data) entry.data = data
    if (errorObj instanceof Error) {
      entry.error = {
        name: errorObj.name,
        message: errorObj.message,
        stack: errorObj.stack
      }
    } else if (errorObj) {
      entry.data = { ...entry.data, rawError: errorObj }
    }

    const sessionId = ctx.sessionId
    if (!this.buffer.has(sessionId)) {
      this.buffer.set(sessionId, [])
    }
    const sessionLogs = this.buffer.get(sessionId)!
    sessionLogs.push(entry)

    if (sessionLogs.length > this.MAX_ENTRIES) {
      sessionLogs.shift()
    }

    appendLogEntry(entry).catch(() => {})
  }

  error(message: string, error?: any, data?: any) {
    // Avoid infinite loop if we use console.error in original mode
    this.originalConsole.error(`[KO] ${message}`, error || '', data || '')
    this.addEntry('error', message, 'system', data, error)
  }

  warn(message: string, data?: any) {
    this.originalConsole.warn(`[KO] ${message}`, data || '')
    this.addEntry('warn', message, 'system', data)
  }

  info(message: string, traffic: LogEntry['traffic'] = 'system', data?: any) {
    this.originalConsole.info(`[KO] ${message}`, data || '')
    this.addEntry('info', message, traffic, data)
  }

  debug(message: string, data?: any) {
    this.originalConsole.debug(`[KO] ${message}`, data || '')
    this.addEntry('debug', message, 'system', data)
  }

  getSessionLogs(sessionId: string): LogEntry[] {
    return this.buffer.get(sessionId) || []
  }

  clearSessionLogs(sessionId: string) {
    this.buffer.delete(sessionId)
  }
}

export const logService = new LogService()
