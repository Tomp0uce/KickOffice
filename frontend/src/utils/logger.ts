import { detectOfficeHost } from './hostDetection';
import { LOG_RING_BUFFER_SIZE } from '@/constants/limits';

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
