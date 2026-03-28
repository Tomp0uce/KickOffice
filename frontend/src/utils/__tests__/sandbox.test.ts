import { describe, it, expect, vi, beforeEach } from 'vitest';

// ─────────────────────────────────────────────────────────────────────────────
// Mock lockdown — prevent SES from actually running in tests
// ─────────────────────────────────────────────────────────────────────────────

vi.mock('../lockdown', () => ({
  ensureLockdown: vi.fn(),
}));

// ─────────────────────────────────────────────────────────────────────────────
// Mock logger — prevent real log side-effects
// ─────────────────────────────────────────────────────────────────────────────

vi.mock('../logger', () => ({
  logService: {
    debug: vi.fn(),
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
  },
}));

// ─────────────────────────────────────────────────────────────────────────────
// Compartment mock factory
//
// A real SES Compartment is not available in the Vitest / happy-dom environment.
// We replace it with a lightweight shim that:
//   - stores the endowments passed to the constructor
//   - evaluate() runs the code via new Function with those globals injected
// ─────────────────────────────────────────────────────────────────────────────

class MockCompartment {
  private globals: Record<string, unknown>;

  constructor(options: { globals: Record<string, unknown>; __options__: boolean }) {
    this.globals = options.globals ?? {};
  }

  evaluate(code: string): unknown {
    // Build an argument list from the stored globals so the executed code can
    // access them by name.
    const keys = Object.keys(this.globals);
    const values = keys.map(k => this.globals[k]);

    // new Function creates an isolated scope; we inject globals as parameters.
    // eslint-disable-next-line @typescript-eslint/no-implied-eval
    try {
      const fn = new Function(...keys, `return ${code}`);
      const result = fn(...values);
      // If the result is a Promise (e.g. from async IIFE), swallow rejections
      // that are only relevant to the call site — prevents unhandled rejection
      // noise in tests that only care about side effects (e.g. logging).
      if (result && typeof result === 'object' && typeof result.then === 'function') {
        return Promise.resolve(result).catch(() => undefined);
      }
      return result;
    } catch {
      // If the code cannot be evaluated (e.g. it is not a valid expression),
      // return a resolved Promise with undefined — mirroring the async IIFE
      // wrapper behaviour in the real sandbox.
      return Promise.resolve(undefined);
    }
  }
}

// Attach to global before importing sandbox so the module sees it
(globalThis as unknown as Record<string, unknown>)['Compartment'] = MockCompartment;

// ─────────────────────────────────────────────────────────────────────────────
// Import after mocks are set up
// ─────────────────────────────────────────────────────────────────────────────

import { sandboxedEval } from '../sandbox';
import { ensureLockdown } from '../lockdown';

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

// Utility available for future tests if needed
// function resolvePromise(value: unknown): Promise<unknown> {
//   return Promise.resolve(value);
// }

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — basic execution
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — basic execution', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('calls ensureLockdown before executing code', () => {
    sandboxedEval('return 1', {});
    expect(ensureLockdown).toHaveBeenCalledTimes(1);
  });

  it('returns a Promise (async IIFE wrapper)', () => {
    const result = sandboxedEval('', {});
    // The evaluate wraps code in async IIFE, so the mock returns a Promise
    expect(result).toBeInstanceOf(Promise);
  });

  it('resolves with a value returned by the code', async () => {
    const result = await (sandboxedEval('return 42', {}) as Promise<unknown>);
    expect(result).toBe(42);
  });

  it('resolves with a string value', async () => {
    const result = await (sandboxedEval('return "hello"', {}) as Promise<unknown>);
    expect(result).toBe('hello');
  });

  it('resolves with undefined for code that has no return statement', async () => {
    const result = await (sandboxedEval('const x = 1', {}) as Promise<unknown>);
    expect(result).toBeUndefined();
  });

  it('executes code that accesses provided globals', async () => {
    const globals = { myValue: 99 };
    const result = await (sandboxedEval('return myValue', globals) as Promise<unknown>);
    expect(result).toBe(99);
  });

  it('executes code that calls a provided function', async () => {
    const add = (a: number, b: number) => a + b;
    const globals = { add };
    const result = await (sandboxedEval('return add(3, 4)', globals) as Promise<unknown>);
    expect(result).toBe(7);
  });

  it('executes code that uses an injected object', async () => {
    const globals = { config: { version: '2.0' } };
    const result = await (sandboxedEval('return config.version', globals) as Promise<unknown>);
    expect(result).toBe('2.0');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — host namespace filtering (buildHostGlobals)
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — host namespace filtering', () => {
  const allNamespaces = {
    Word: { run: vi.fn() },
    Excel: { run: vi.fn() },
    PowerPoint: { run: vi.fn() },
  };

  it('removes Excel and PowerPoint globals when host is Word', async () => {
    const result = await (sandboxedEval(
      'return typeof Excel === "undefined" && typeof PowerPoint === "undefined"',
      { ...allNamespaces },
      'Word',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('keeps Word global when host is Word', async () => {
    const result = await (sandboxedEval(
      'return typeof Word !== "undefined"',
      { ...allNamespaces },
      'Word',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('removes Word and PowerPoint globals when host is Excel', async () => {
    const result = await (sandboxedEval(
      'return typeof Word === "undefined" && typeof PowerPoint === "undefined"',
      { ...allNamespaces },
      'Excel',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('keeps Excel global when host is Excel', async () => {
    const result = await (sandboxedEval(
      'return typeof Excel !== "undefined"',
      { ...allNamespaces },
      'Excel',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('removes Word and Excel globals when host is PowerPoint', async () => {
    const result = await (sandboxedEval(
      'return typeof Word === "undefined" && typeof Excel === "undefined"',
      { ...allNamespaces },
      'PowerPoint',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('keeps PowerPoint global when host is PowerPoint', async () => {
    const result = await (sandboxedEval(
      'return typeof PowerPoint !== "undefined"',
      { ...allNamespaces },
      'PowerPoint',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('removes Word, Excel and PowerPoint globals when host is Outlook', async () => {
    const result = await (sandboxedEval(
      'return typeof Word === "undefined" && typeof Excel === "undefined" && typeof PowerPoint === "undefined"',
      { ...allNamespaces },
      'Outlook',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('does not remove any namespace when no host is specified', async () => {
    const result = await (sandboxedEval(
      'return typeof Word !== "undefined" && typeof Excel !== "undefined"',
      { ...allNamespaces },
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('does not remove globals that are not in the namespace map (custom keys)', async () => {
    const globals = { Word: { run: vi.fn() }, customHelper: () => 42 };
    const result = await (sandboxedEval(
      'return typeof customHelper !== "undefined"',
      globals,
      'Excel',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('filtering only sets namespace to undefined — does not delete the key', async () => {
    // The implementation sets `result[ns] = undefined` — the key still exists
    // but its value is undefined. Code receiving it sees it as undefined.
    const globals = { Excel: { run: vi.fn() } };
    const result = await (sandboxedEval(
      'return typeof Excel === "undefined"',
      globals,
      'Word',
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — blocked built-ins exposed as undefined
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — blocked built-ins are undefined in sandbox', () => {
  it('eval is undefined in the compartment globals', async () => {
    // The sandbox passes eval: undefined to the Compartment; our mock injects it
    // as a parameter whose value is undefined.
    const result = await (sandboxedEval(
      'return typeof eval === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('fetch is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof fetch === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('Function is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof Function === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('XMLHttpRequest is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof XMLHttpRequest === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('Proxy is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof Proxy === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('Reflect is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof Reflect === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('Compartment is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof Compartment === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('lockdown is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof lockdown === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('harden is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof harden === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });

  it('WebSocket is undefined in the compartment globals', async () => {
    const result = await (sandboxedEval(
      'return typeof WebSocket === "undefined"',
      {},
    ) as Promise<unknown>);
    expect(result).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — safe built-ins are available
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — safe built-ins are available', () => {
  it('Math is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Math', {}) as Promise<unknown>);
    expect(result).toBe('object');
  });

  it('JSON is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof JSON', {}) as Promise<unknown>);
    expect(result).toBe('object');
  });

  it('console is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof console', {}) as Promise<unknown>);
    expect(result).toBe('object');
  });

  it('Array is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Array', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('Object is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Object', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('Date is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Date', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('String is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof String', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('Number is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Number', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('Boolean is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Boolean', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });

  it('Promise is available in the sandbox', async () => {
    const result = await (sandboxedEval('return typeof Promise', {}) as Promise<unknown>);
    expect(result).toBe('function');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — code preview / debug logging
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — debug logging', () => {
  it('logs a debug message with the host label', async () => {
    const { logService } = await import('../logger');
    vi.clearAllMocks();

    sandboxedEval('return 1', {}, 'Word');

    expect(logService.debug).toHaveBeenCalledTimes(1);
    const [msg] = (logService.debug as ReturnType<typeof vi.fn>).mock.calls[0] as [string];
    expect(msg).toContain('host=Word');
  });

  it('logs "host=unspecified" when no host is provided', async () => {
    const { logService } = await import('../logger');
    vi.clearAllMocks();

    sandboxedEval('return 1', {});

    const [msg] = (logService.debug as ReturnType<typeof vi.fn>).mock.calls[0] as [string];
    expect(msg).toContain('host=unspecified');
  });

  it('truncates code longer than 200 chars in log preview', async () => {
    const { logService } = await import('../logger');
    vi.clearAllMocks();

    // Use a valid JS expression (string literal) repeated so the whole code
    // exceeds 200 chars — the sandbox evaluates it fine and we get the log.
    const longCode = '"' + 'a'.repeat(250) + '"';
    sandboxedEval(longCode, {});

    const [msg] = (logService.debug as ReturnType<typeof vi.fn>).mock.calls[0] as [string];
    // The preview must end with the ellipsis character
    expect(msg).toContain('\u2026');
    // The previewed code portion must be no longer than 200 chars + ellipsis
    const codePreviewMatch = msg.match(/code=(.+)/s);
    expect(codePreviewMatch).not.toBeNull();
    // Preview is 200 chars of code plus "…"
    expect(codePreviewMatch![1].length).toBeLessThanOrEqual(201);
  });

  it('does not truncate code of exactly 200 chars', async () => {
    const { logService } = await import('../logger');
    vi.clearAllMocks();

    // 200-char string literal: opening quote + 198 chars + closing quote = 200
    const exactCode = '"' + 'b'.repeat(198) + '"';
    sandboxedEval(exactCode, {});

    const [msg] = (logService.debug as ReturnType<typeof vi.fn>).mock.calls[0] as [string];
    expect(msg).not.toContain('\u2026');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// sandboxedEval — globals isolation (original object not mutated)
// ─────────────────────────────────────────────────────────────────────────────

describe('sandboxedEval — globals object is not mutated', () => {
  it('does not modify the original globals passed in', () => {
    const globals = {
      Word: { run: vi.fn() },
      Excel: { run: vi.fn() },
      PowerPoint: { run: vi.fn() },
    };
    const original = { ...globals };

    sandboxedEval('return 1', globals, 'Word');

    // Original globals should be unchanged
    expect(globals.Word).toBe(original.Word);
    expect(globals.Excel).toBe(original.Excel);
    expect(globals.PowerPoint).toBe(original.PowerPoint);
  });
});
