import { detectOfficeHost } from './hostDetection';
import { LOG_RING_BUFFER_SIZE } from '@/constants/limits';

// ERR-M3: Flush interval in ms (30 s) and max entries per batch (matches backend MAX_ENTRIES)
const LOG_FLUSH_INTERVAL_MS = 30_000;
const LOG_FLUSH_MAX_ENTRIES = 200;

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface LogEntry {
  timestamp: string;
  level: LogLevel;
  source: 'frontend';
  traffic: 'user' | 'llm' | 'auto' | 'system';
  sessionId?: string;
  userId?: string;
  host?: string;
  message: string;
  data?: Record<string, unknown>;
  error?: {
    name: string;
    message: string;
    stack?: string;
  };
}

const LOG_LEVEL_MAP: Record<LogLevel, number> = {
  debug: 0,
  info: 1,
  warn: 2,
  error: 3,
};

class LogService {
  private buffer: Map<string, LogEntry[]> = new Map();
  private readonly MAX_ENTRIES = LOG_RING_BUFFER_SIZE;
  private _currentSessionId: string = 'default';
  private _logLevel: LogLevel = import.meta.env.PROD ? 'warn' : 'debug';
  // ERR-M3: pending entries waiting for the next flush to /api/logs
  private _pendingFlush: LogEntry[] = [];
  private _flushTimer: ReturnType<typeof setInterval> | null = null;

  public readonly originalConsole = {
    log: console.log,
    info: console.info,
    warn: console.warn,
    error: console.error,
    debug: console.debug,
  };

  setCurrentSessionId(id: string) {
    this._currentSessionId = id;
  }

  setLogLevel(level: LogLevel) {
    this._logLevel = level;
  }

  // ERR-M3: Start the periodic flush timer (called once at app boot)
  startFlushTimer() {
    if (this._flushTimer !== null) return;
    this._flushTimer = setInterval(() => {
      this._flushToBackend();
    }, LOG_FLUSH_INTERVAL_MS);
  }

  private async _flushToBackend() {
    if (this._pendingFlush.length === 0) return;
    const batch = this._pendingFlush.splice(0, LOG_FLUSH_MAX_ENTRIES);
    try {
      // Lazy import to avoid circular dependency (backend.ts -> logger.ts)
      const { submitLogs } = await import('@/api/backend');
      await submitLogs(batch);
    } catch {
      // Silent: log forwarding failures must not affect the UI
    }
  }

  async getContext() {
    return {
      host: detectOfficeHost(),
      sessionId: this._currentSessionId,
      userId: await import('./credentialStorage').then(m => m.getUserEmail()).catch(() => 'anonymous'),
    };
  }

  private async addEntry(
    level: LogEntry['level'],
    message: string,
    traffic: LogEntry['traffic'] = 'system',
    data?: any,
    errorObj?: any,
  ) {
    if (LOG_LEVEL_MAP[level] < LOG_LEVEL_MAP[this._logLevel]) {
      return;
    }

    const ctx = await this.getContext();
    const entry: LogEntry = {
      timestamp: new Date().toISOString(),
      level,
      source: 'frontend',
      traffic,
      sessionId: ctx.sessionId,
      userId: ctx.userId,
      host: ctx.host,
      message,
    };

    if (data) entry.data = data;
    if (errorObj instanceof Error) {
      entry.error = {
        name: errorObj.name,
        message: errorObj.message,
        stack: errorObj.stack,
      };
    } else if (errorObj) {
      entry.data = { ...entry.data, rawError: errorObj };
    }

    const sessionId = ctx.sessionId;
    if (!this.buffer.has(sessionId)) {
      this.buffer.set(sessionId, []);
    }
    const sessionLogs = this.buffer.get(sessionId)!;
    sessionLogs.push(entry);

    if (sessionLogs.length > this.MAX_ENTRIES) {
      sessionLogs.shift();
    }

    import('@/composables/useSessionDB').then(m => m.appendLogEntry(entry)).catch(() => {});

    // ERR-M3: Queue for backend forwarding (warn + error only, to match backend ALLOWED_LEVELS)
    if (level === 'warn' || level === 'error') {
      this._pendingFlush.push(entry);
      // Flush immediately on error so critical events are not lost on the next 30 s tick
      if (level === 'error') {
        this._flushToBackend();
      }
    }
  }

  error(message: string, error?: any, data?: any) {
    // Avoid infinite loop if we use console.error in original mode
    this.originalConsole.error(`[KO] ${message}`, error || '', data || '');
    this.addEntry('error', message, 'system', data, error);
  }

  warn(message: string, data?: any) {
    this.originalConsole.warn(`[KO] ${message}`, data || '');
    this.addEntry('warn', message, 'system', data);
  }

  info(message: string, traffic: LogEntry['traffic'] = 'system', data?: any) {
    this.originalConsole.info(`[KO] ${message}`, data || '');
    this.addEntry('info', message, traffic, data);
  }

  debug(message: string, data?: any) {
    this.originalConsole.debug(`[KO] ${message}`, data || '');
    this.addEntry('debug', message, 'system', data);
  }

  getSessionLogs(sessionId: string): LogEntry[] {
    return this.buffer.get(sessionId) || [];
  }

  clearSessionLogs(sessionId: string) {
    this.buffer.delete(sessionId);
  }
}

export const logService = new LogService();
