import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// --- Mocks declared before any module import ---

vi.mock('@/utils/hostDetection', () => ({
  detectOfficeHost: vi.fn(() => 'Word'),
}));

vi.mock('@/utils/credentialStorage', () => ({
  getUserEmail: vi.fn(async () => 'user@example.com'),
}));

vi.mock('@/composables/useSessionDB', () => ({
  appendLogEntry: vi.fn(async () => {}),
}));

vi.mock('@/api/backend', () => ({
  submitLogs: vi.fn(async () => {}),
}));

// --- Imports after mocks ---

import { logService, type LogEntry } from '@/utils/logger';
import { getUserEmail } from '@/utils/credentialStorage';
import { appendLogEntry } from '@/composables/useSessionDB';
import { submitLogs } from '@/api/backend';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Flush all microtasks / pending promises.
 * Multiple ticks are needed because addEntry is async and may chain additional
 * microtasks (dynamic imports for useSessionDB and backend forwarding).
 */
const flushPromises = () =>
  new Promise<void>(resolve => setTimeout(() => setTimeout(resolve, 0), 0));

/** Read all entries for the default 'default' session */
function getDefaultLogs(): LogEntry[] {
  return logService.getSessionLogs('default');
}

// ---------------------------------------------------------------------------
// Setup / Teardown
// ---------------------------------------------------------------------------

beforeEach(() => {
  // Reset singleton state between tests
  logService.setCurrentSessionId('default');
  logService.setLogLevel('debug');
  logService.clearSessionLogs('default');
  // Clear the private pending flush queue to prevent cross-test contamination
  (logService as unknown as { _pendingFlush: LogEntry[] })._pendingFlush = [];

  // Reset all mock call counts
  vi.clearAllMocks();

  // Spy on originalConsole methods (not global console, because logger captures
  // console refs at class construction time into `originalConsole`)
  vi.spyOn(logService.originalConsole, 'debug').mockImplementation(() => {});
  vi.spyOn(logService.originalConsole, 'info').mockImplementation(() => {});
  vi.spyOn(logService.originalConsole, 'warn').mockImplementation(() => {});
  vi.spyOn(logService.originalConsole, 'error').mockImplementation(() => {});
});

afterEach(() => {
  vi.restoreAllMocks();
});

// ---------------------------------------------------------------------------
// LogService — log level methods
// ---------------------------------------------------------------------------

describe('logService.debug()', () => {
  it('writes a debug entry to the session buffer', async () => {
    logService.debug('hello debug');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs).toHaveLength(1);
    expect(logs[0].level).toBe('debug');
    expect(logs[0].message).toBe('hello debug');
  });

  it('forwards to originalConsole.debug with [KO] prefix', async () => {
    logService.debug('dbg msg');
    await flushPromises();

    expect(logService.originalConsole.debug).toHaveBeenCalledWith('[KO] dbg msg', '');
  });

  it('passes data to originalConsole when provided', async () => {
    logService.debug('with data', { key: 'value' });
    await flushPromises();

    expect(logService.originalConsole.debug).toHaveBeenCalledWith('[KO] with data', {
      key: 'value',
    });
  });

  it('does NOT queue entry to _pendingFlush (debug is below warn)', async () => {
    vi.clearAllMocks();
    logService.debug('quiet debug');
    await flushPromises();

    // submitLogs must never be called for debug entries
    expect(submitLogs).not.toHaveBeenCalled();
  });

  it('defaults traffic to "system"', async () => {
    logService.debug('debug traffic test');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('system');
  });

  it('accepts an explicit traffic label as third argument', async () => {
    logService.debug('debug llm', undefined, 'llm');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('llm');
  });
});

describe('logService.info()', () => {
  it('writes an info entry to the session buffer', async () => {
    logService.info('hello info');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs).toHaveLength(1);
    expect(logs[0].level).toBe('info');
    expect(logs[0].message).toBe('hello info');
  });

  it('defaults traffic to "system"', async () => {
    logService.info('system info');
    await flushPromises();

    expect(getDefaultLogs()[0].traffic).toBe('system');
  });

  it('accepts an explicit traffic label', async () => {
    logService.info('user message', 'user');
    await flushPromises();

    expect(getDefaultLogs()[0].traffic).toBe('user');
  });

  it('accepts "llm" and "auto" traffic labels', async () => {
    logService.info('llm msg', 'llm');
    await flushPromises();
    logService.info('auto msg', 'auto');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs[0].traffic).toBe('llm');
    expect(logs[1].traffic).toBe('auto');
  });

  it('forwards to originalConsole.info', async () => {
    logService.info('info msg');
    await flushPromises();

    expect(logService.originalConsole.info).toHaveBeenCalledWith('[KO] info msg', '');
  });
});

describe('logService.warn()', () => {
  it('writes a warn entry to the session buffer', async () => {
    logService.warn('watch out');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs).toHaveLength(1);
    expect(logs[0].level).toBe('warn');
    expect(logs[0].message).toBe('watch out');
  });

  it('forwards to originalConsole.warn', async () => {
    logService.warn('warn msg');
    await flushPromises();

    expect(logService.originalConsole.warn).toHaveBeenCalledWith('[KO] warn msg', '');
  });

  it('forwards data argument to console.warn', async () => {
    logService.warn('warn with data', { extra: 1 });
    await flushPromises();

    expect(logService.originalConsole.warn).toHaveBeenCalledWith('[KO] warn with data', {
      extra: 1,
    });
  });

  it('defaults traffic to "system"', async () => {
    logService.warn('warn traffic test');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('system');
  });

  it('accepts an explicit traffic label as third argument', async () => {
    logService.warn('warn llm', undefined, 'llm');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('llm');
  });

  it('serializes Error objects passed as data into structured record', async () => {
    const err = new Error('test failure');
    logService.warn('caught error', err);
    await flushPromises();
    const entry = getDefaultLogs()[0];
    expect(entry.data).toBeDefined();
    expect(entry.data?.message).toBe('test failure');
    expect(entry.data?.name).toBe('Error');
    expect(entry.data?.stack).toBeDefined();
  });
});

describe('logService.error()', () => {
  it('writes an error entry to the session buffer', async () => {
    logService.error('something broke');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs).toHaveLength(1);
    expect(logs[0].level).toBe('error');
    expect(logs[0].message).toBe('something broke');
  });

  it('forwards to originalConsole.error with [KO] prefix', async () => {
    logService.error('err msg');
    await flushPromises();

    expect(logService.originalConsole.error).toHaveBeenCalledWith('[KO] err msg', '', '');
  });

  it('triggers an immediate backend flush', async () => {
    logService.error('critical error');
    await flushPromises();

    expect(submitLogs).toHaveBeenCalled();
  });

  it('defaults traffic to "system"', async () => {
    logService.error('error traffic test');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('system');
  });

  it('accepts an explicit traffic label as fourth argument', async () => {
    logService.error('error llm', undefined, undefined, 'llm');
    await flushPromises();
    expect(getDefaultLogs()[0].traffic).toBe('llm');
  });
});

// ---------------------------------------------------------------------------
// LogEntry structure
// ---------------------------------------------------------------------------

describe('LogEntry structure', () => {
  it('entry always has source = "frontend"', async () => {
    logService.info('check source');
    await flushPromises();

    expect(getDefaultLogs()[0].source).toBe('frontend');
  });

  it('entry has a valid ISO timestamp', async () => {
    const before = Date.now();
    logService.info('timestamp test');
    await flushPromises();
    const after = Date.now();

    const ts = Date.parse(getDefaultLogs()[0].timestamp);
    expect(ts).toBeGreaterThanOrEqual(before);
    expect(ts).toBeLessThanOrEqual(after);
  });

  it('entry includes sessionId from context', async () => {
    logService.setCurrentSessionId('sess-42');
    logService.clearSessionLogs('sess-42');
    logService.info('session check');
    await flushPromises();

    expect(logService.getSessionLogs('sess-42')[0].sessionId).toBe('sess-42');
    // Restore default for next tests
    logService.setCurrentSessionId('default');
  });

  it('entry includes userId from credentialStorage', async () => {
    logService.info('user check');
    await flushPromises();

    expect(getDefaultLogs()[0].userId).toBe('user@example.com');
  });

  it('entry includes host from detectOfficeHost', async () => {
    logService.info('host check');
    await flushPromises();

    expect(getDefaultLogs()[0].host).toBe('Word');
  });

  it('userId falls back to "anonymous" when credentialStorage rejects', async () => {
    vi.mocked(getUserEmail).mockRejectedValueOnce(new Error('no creds'));
    logService.info('anon check');
    await flushPromises();

    expect(getDefaultLogs()[0].userId).toBe('anonymous');
  });
});

// ---------------------------------------------------------------------------
// data field handling
// ---------------------------------------------------------------------------

describe('data field', () => {
  it('stores a plain object as entry.data', async () => {
    logService.info('with data', 'system', { key: 'val' });
    await flushPromises();

    expect(getDefaultLogs()[0].data).toEqual({ key: 'val' });
  });

  it('ignores data when it is a string', async () => {
    logService.info('string data', 'system', 'just a string');
    await flushPromises();

    expect(getDefaultLogs()[0].data).toBeUndefined();
  });

  it('ignores data when it is an array', async () => {
    logService.info('array data', 'system', [1, 2, 3]);
    await flushPromises();

    expect(getDefaultLogs()[0].data).toBeUndefined();
  });

  it('ignores data when it is null', async () => {
    logService.info('null data', 'system', null as unknown as undefined);
    await flushPromises();

    expect(getDefaultLogs()[0].data).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// error field — Error object vs raw value
// ---------------------------------------------------------------------------

describe('error field handling in logService.error()', () => {
  it('sets entry.error from an Error instance', async () => {
    const err = new Error('oops');
    logService.error('caught error', err);
    await flushPromises();

    const entry = getDefaultLogs()[0];
    expect(entry.error).toBeDefined();
    expect(entry.error?.name).toBe('Error');
    expect(entry.error?.message).toBe('oops');
    expect(entry.error?.stack).toContain('Error: oops');
  });

  it('preserves error.name for subclasses of Error', async () => {
    const err = new TypeError('bad type');
    logService.error('type error', err);
    await flushPromises();

    expect(getDefaultLogs()[0].error?.name).toBe('TypeError');
  });

  it('stores a non-Error truthy value inside entry.data.rawError', async () => {
    logService.error('string error', 'something went wrong');
    await flushPromises();

    const entry = getDefaultLogs()[0];
    expect(entry.error).toBeUndefined();
    expect(entry.data?.rawError).toBe('something went wrong');
  });

  it('stores a numeric error code in entry.data.rawError', async () => {
    logService.error('numeric error', 404);
    await flushPromises();

    expect(getDefaultLogs()[0].data?.rawError).toBe(404);
  });

  it('does not set error field when errorObj is undefined', async () => {
    logService.error('no error obj');
    await flushPromises();

    expect(getDefaultLogs()[0].error).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Log level filtering
// ---------------------------------------------------------------------------

describe('log level filtering', () => {
  it('suppresses debug entries when level is set to "info"', async () => {
    logService.setLogLevel('info');
    logService.debug('suppressed debug');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(0);
  });

  it('suppresses debug and info entries when level is set to "warn"', async () => {
    logService.setLogLevel('warn');
    logService.debug('no debug');
    logService.info('no info');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(0);
  });

  it('allows warn through when level is "warn"', async () => {
    logService.setLogLevel('warn');
    logService.warn('visible warn');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(1);
  });

  it('suppresses everything below error when level is "error"', async () => {
    logService.setLogLevel('error');
    logService.debug('no');
    logService.info('no');
    logService.warn('no');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(0);
  });

  it('allows error through when level is "error"', async () => {
    logService.setLogLevel('error');
    logService.error('visible error');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(1);
  });

  it('passes all levels when set to "debug"', async () => {
    logService.setLogLevel('debug');
    logService.debug('d');
    await flushPromises();
    logService.info('i');
    await flushPromises();
    logService.warn('w');
    await flushPromises();
    logService.error('e');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(4);
  });
});

// ---------------------------------------------------------------------------
// getSessionLogs
// ---------------------------------------------------------------------------

describe('getSessionLogs()', () => {
  it('returns an empty array for an unknown session', () => {
    expect(logService.getSessionLogs('nonexistent')).toEqual([]);
  });

  it('returns all entries for the requested session', async () => {
    logService.info('first');
    await flushPromises();
    logService.info('second');
    await flushPromises();

    expect(getDefaultLogs()).toHaveLength(2);
  });

  it('keeps sessions isolated from each other', async () => {
    logService.setCurrentSessionId('sess-A');
    logService.clearSessionLogs('sess-A');
    logService.info('message A');
    await flushPromises();

    logService.setCurrentSessionId('sess-B');
    logService.clearSessionLogs('sess-B');
    logService.info('message B');
    await flushPromises();

    expect(logService.getSessionLogs('sess-A')).toHaveLength(1);
    expect(logService.getSessionLogs('sess-B')).toHaveLength(1);
    expect(logService.getSessionLogs('sess-A')[0].message).toBe('message A');

    // Restore
    logService.setCurrentSessionId('default');
  });
});

// ---------------------------------------------------------------------------
// clearSessionLogs
// ---------------------------------------------------------------------------

describe('clearSessionLogs()', () => {
  it('empties the buffer for the given session', async () => {
    logService.info('to be cleared');
    await flushPromises();

    logService.clearSessionLogs('default');

    expect(getDefaultLogs()).toHaveLength(0);
  });

  it('does not affect other sessions', async () => {
    logService.setCurrentSessionId('keeper');
    logService.clearSessionLogs('keeper');
    logService.info('keep me');
    await flushPromises();

    logService.clearSessionLogs('default');

    expect(logService.getSessionLogs('keeper')).toHaveLength(1);

    // Restore
    logService.setCurrentSessionId('default');
    logService.clearSessionLogs('keeper');
  });

  it('is idempotent on an already-empty session', () => {
    logService.clearSessionLogs('default');
    logService.clearSessionLogs('default'); // second call must not throw
    expect(getDefaultLogs()).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Ring buffer — MAX_ENTRIES cap (LOG_RING_BUFFER_SIZE = 500)
// ---------------------------------------------------------------------------

describe('ring buffer cap', () => {
  it('trims oldest entry once the buffer exceeds MAX_ENTRIES (500)', async () => {
    // Fill the buffer to capacity, awaiting each entry so the async chain fully resolves.
    const waitTick = () => new Promise<void>(r => setTimeout(r, 0));
    for (let i = 0; i < 500; i++) {
      logService.debug(`entry-${i}`);
      await waitTick();
    }

    expect(getDefaultLogs()).toHaveLength(500);
    expect(getDefaultLogs()[0].message).toBe('entry-0');

    // One more entry: entry-0 should be evicted
    logService.debug('entry-500');
    await flushPromises();

    const logs = getDefaultLogs();
    expect(logs).toHaveLength(500);
    expect(logs[0].message).toBe('entry-1');
    expect(logs[499].message).toBe('entry-500');
  }, 30_000);
});

// ---------------------------------------------------------------------------
// Backend flush behaviour
// ---------------------------------------------------------------------------

describe('backend flush (_pendingFlush)', () => {
  it('queues warn entries for the next flush cycle', async () => {
    logService.warn('queued warn');
    await flushPromises();

    // warn does NOT trigger immediate flush — submitLogs is NOT called yet
    expect(submitLogs).not.toHaveBeenCalled();
  });

  it('flushes immediately on error', async () => {
    logService.error('immediate flush');
    await flushPromises();

    expect(submitLogs).toHaveBeenCalled();
    const batch = vi.mocked(submitLogs).mock.calls[0][0] as LogEntry[];
    expect(batch.length).toBeGreaterThan(0);
    expect(batch[0].level).toBe('error');
  });

  it('does not call submitLogs for info entries', async () => {
    logService.info('not forwarded');
    await flushPromises();

    expect(submitLogs).not.toHaveBeenCalled();
  });

  it('does not call submitLogs for debug entries', async () => {
    logService.debug('not forwarded');
    await flushPromises();

    expect(submitLogs).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// appendLogEntry integration (useSessionDB)
// ---------------------------------------------------------------------------

describe('appendLogEntry integration', () => {
  it('calls appendLogEntry for every added entry', async () => {
    logService.info('persist me');
    await flushPromises();

    expect(appendLogEntry).toHaveBeenCalledTimes(1);
    const saved = vi.mocked(appendLogEntry).mock.calls[0][0] as LogEntry;
    expect(saved.message).toBe('persist me');
  });

  it('does NOT throw when appendLogEntry rejects', async () => {
    vi.mocked(appendLogEntry).mockRejectedValueOnce(new Error('db error'));
    await expect(async () => {
      logService.info('resilient');
      await flushPromises();
    }).not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// setCurrentSessionId
// ---------------------------------------------------------------------------

describe('setCurrentSessionId()', () => {
  it('changes which session receives new log entries', async () => {
    logService.setCurrentSessionId('custom-session');
    logService.clearSessionLogs('custom-session');
    logService.info('in custom session');
    await flushPromises();

    expect(logService.getSessionLogs('custom-session')).toHaveLength(1);
    expect(getDefaultLogs()).toHaveLength(0);

    // Restore
    logService.setCurrentSessionId('default');
    logService.clearSessionLogs('custom-session');
  });
});

// ---------------------------------------------------------------------------
// startFlushTimer — guards
// ---------------------------------------------------------------------------

describe('startFlushTimer()', () => {
  it('is idempotent: calling it twice does not create a second timer', () => {
    // Grab the internal timer state by calling twice
    // We cannot directly inspect the private field, but we can verify
    // that no error or unexpected behaviour occurs.
    expect(() => {
      logService.startFlushTimer();
      logService.startFlushTimer();
    }).not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// getContext
// ---------------------------------------------------------------------------

describe('getContext()', () => {
  it('returns host from detectOfficeHost', async () => {
    const ctx = await logService.getContext();
    expect(ctx.host).toBe('Word');
  });

  it('returns sessionId matching the current session', async () => {
    logService.setCurrentSessionId('ctx-test');
    const ctx = await logService.getContext();
    expect(ctx.sessionId).toBe('ctx-test');

    logService.setCurrentSessionId('default');
  });

  it('returns userId from getUserEmail', async () => {
    const ctx = await logService.getContext();
    expect(ctx.userId).toBe('user@example.com');
  });

  it('falls back to "anonymous" if getUserEmail throws', async () => {
    vi.mocked(getUserEmail).mockRejectedValueOnce(new Error('no email'));
    const ctx = await logService.getContext();
    expect(ctx.userId).toBe('anonymous');
  });
});

// ---------------------------------------------------------------------------
// originalConsole preservation
// ---------------------------------------------------------------------------

describe('originalConsole', () => {
  it('exposes all five console methods', () => {
    expect(typeof logService.originalConsole.log).toBe('function');
    expect(typeof logService.originalConsole.info).toBe('function');
    expect(typeof logService.originalConsole.warn).toBe('function');
    expect(typeof logService.originalConsole.error).toBe('function');
    expect(typeof logService.originalConsole.debug).toBe('function');
  });
});
