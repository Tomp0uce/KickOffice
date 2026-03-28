import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * The module caches state in module-level variables (detectedHost, officeReady).
 * We must re-import the module on each test to get a clean slate.
 */
async function freshImport() {
  vi.resetModules();
  return import('../hostDetection');
}

function setOfficeHost(host: string | undefined) {
  if (host === undefined) {
    // Remove Office entirely
    delete (window as unknown as Record<string, unknown>)['Office'];
  } else {
    (window as unknown as Record<string, unknown>)['Office'] = {
      context: { host },
    };
  }
}

function setOfficeWithMailbox(hasMailbox: boolean) {
  (window as unknown as Record<string, unknown>)['Office'] = {
    context: {
      host: undefined,
      mailbox: hasMailbox ? {} : undefined,
    },
  };
}

function setGlobalNamespace(name: string, value: unknown) {
  (window as unknown as Record<string, unknown>)[name] = value;
}

// Utility available for future tests if needed
// function removeGlobalNamespace(name: string) {
//   delete (window as unknown as Record<string, unknown>)[name];
// }

// ─────────────────────────────────────────────────────────────────────────────
// Cleanup after each test
// ─────────────────────────────────────────────────────────────────────────────

afterEach(() => {
  delete (window as unknown as Record<string, unknown>)['Office'];
  delete (window as unknown as Record<string, unknown>)['Word'];
  delete (window as unknown as Record<string, unknown>)['Excel'];
  delete (window as unknown as Record<string, unknown>)['PowerPoint'];
  vi.resetModules();
});

// ─────────────────────────────────────────────────────────────────────────────
// detectOfficeHost — Office.context.host detection
// ─────────────────────────────────────────────────────────────────────────────

describe('detectOfficeHost — Office.context.host', () => {
  it('returns Word when host is "Word"', async () => {
    setOfficeHost('Word');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Word');
  });

  it('returns Word when host is "Document"', async () => {
    setOfficeHost('Document');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Word');
  });

  it('returns Excel when host is "Excel"', async () => {
    setOfficeHost('Excel');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Excel');
  });

  it('returns Excel when host is "Workbook"', async () => {
    setOfficeHost('Workbook');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Excel');
  });

  it('returns PowerPoint when host is "PowerPoint"', async () => {
    setOfficeHost('PowerPoint');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('PowerPoint');
  });

  it('returns PowerPoint when host is "Presentation"', async () => {
    setOfficeHost('Presentation');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('PowerPoint');
  });

  it('returns Outlook when host is "Outlook"', async () => {
    setOfficeHost('Outlook');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Outlook');
  });

  it('returns Outlook when host is "Mailbox"', async () => {
    setOfficeHost('Mailbox');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Outlook');
  });

  it('returns Unknown when host is an unrecognised string', async () => {
    setOfficeHost('SomeOtherHost');
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });

  it('returns Unknown when Office is not present on window', async () => {
    setOfficeHost(undefined);
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// detectOfficeHost — fallback global namespace detection
// ─────────────────────────────────────────────────────────────────────────────

describe('detectOfficeHost — global namespace fallback', () => {
  beforeEach(() => {
    // Ensure Office.context.host is absent so fallback path runs
    setOfficeHost(undefined);
  });

  it('returns Word when global Word namespace is defined', async () => {
    setGlobalNamespace('Word', {});
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Word');
  });

  it('returns Excel when global Excel namespace is defined', async () => {
    setGlobalNamespace('Excel', {});
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Excel');
  });

  it('returns PowerPoint when global PowerPoint namespace is defined', async () => {
    setGlobalNamespace('PowerPoint', {});
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('PowerPoint');
  });

  it('returns Outlook when Office.context.mailbox is defined', async () => {
    setOfficeWithMailbox(true);
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Outlook');
  });

  it('returns Unknown when no global namespaces are present', async () => {
    // All cleaned up in afterEach; nothing extra to set
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// detectOfficeHost — partial / malformed Office object edge cases
// ─────────────────────────────────────────────────────────────────────────────

describe('detectOfficeHost — edge cases', () => {
  it('returns Unknown when Office exists but context is missing', async () => {
    (window as unknown as Record<string, unknown>)['Office'] = {};
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });

  it('returns Unknown when Office.context exists but host is missing', async () => {
    (window as unknown as Record<string, unknown>)['Office'] = { context: {} };
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });

  it('returns Unknown when Office.context.host is null', async () => {
    (window as unknown as Record<string, unknown>)['Office'] = {
      context: { host: null },
    };
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });

  it('returns Unknown when Office is set to a non-object primitive', async () => {
    (window as unknown as Record<string, unknown>)['Office'] = 42;
    const { detectOfficeHost } = await freshImport();
    expect(detectOfficeHost()).toBe('Unknown');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// detectOfficeHost — caching after markOfficeReady
// ─────────────────────────────────────────────────────────────────────────────

describe('detectOfficeHost — caching with markOfficeReady', () => {
  it('caches the host after markOfficeReady is called', async () => {
    setOfficeHost('Word');
    const { detectOfficeHost, markOfficeReady } = await freshImport();

    markOfficeReady();
    // First call: detects Word and stores in cache
    expect(detectOfficeHost()).toBe('Word');

    // Now change Office.context.host to something else
    setOfficeHost('Excel');

    // Cached value should still be Word
    expect(detectOfficeHost()).toBe('Word');
  });

  it('re-detects on first call after markOfficeReady (cache reset)', async () => {
    setOfficeHost('Excel');
    const { detectOfficeHost, markOfficeReady } = await freshImport();

    // First call before markOfficeReady
    expect(detectOfficeHost()).toBe('Excel');

    // markOfficeReady resets detectedHost to Unknown
    setOfficeHost('PowerPoint');
    markOfficeReady();

    // Should re-detect with the current Office context
    expect(detectOfficeHost()).toBe('PowerPoint');
  });

  it('does not cache result before markOfficeReady', async () => {
    setOfficeHost('Word');
    const { detectOfficeHost } = await freshImport();

    expect(detectOfficeHost()).toBe('Word');

    // Change Office context — no cache in place, so it should re-detect
    setOfficeHost('Excel');
    expect(detectOfficeHost()).toBe('Excel');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// isWord / isExcel / isPowerPoint / isOutlook
// ─────────────────────────────────────────────────────────────────────────────

describe('isWord', () => {
  it('returns true when host is Word', async () => {
    setOfficeHost('Word');
    const { isWord } = await freshImport();
    expect(isWord()).toBe(true);
  });

  it('returns false when host is Excel', async () => {
    setOfficeHost('Excel');
    const { isWord } = await freshImport();
    expect(isWord()).toBe(false);
  });
});

describe('isExcel', () => {
  it('returns true when host is Excel', async () => {
    setOfficeHost('Excel');
    const { isExcel } = await freshImport();
    expect(isExcel()).toBe(true);
  });

  it('returns false when host is PowerPoint', async () => {
    setOfficeHost('PowerPoint');
    const { isExcel } = await freshImport();
    expect(isExcel()).toBe(false);
  });
});

describe('isPowerPoint', () => {
  it('returns true when host is PowerPoint', async () => {
    setOfficeHost('PowerPoint');
    const { isPowerPoint } = await freshImport();
    expect(isPowerPoint()).toBe(true);
  });

  it('returns false when host is Word', async () => {
    setOfficeHost('Word');
    const { isPowerPoint } = await freshImport();
    expect(isPowerPoint()).toBe(false);
  });
});

describe('isOutlook', () => {
  it('returns true when host is Outlook', async () => {
    setOfficeHost('Outlook');
    const { isOutlook } = await freshImport();
    expect(isOutlook()).toBe(true);
  });

  it('returns false when no host is set', async () => {
    setOfficeHost(undefined);
    const { isOutlook } = await freshImport();
    expect(isOutlook()).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// forHost
// ─────────────────────────────────────────────────────────────────────────────

describe('forHost', () => {
  it('returns word option when host is Word', async () => {
    setOfficeHost('Word');
    const { forHost } = await freshImport();
    expect(forHost({ word: 'w', excel: 'e', default: 'd' })).toBe('w');
  });

  it('returns excel option when host is Excel', async () => {
    setOfficeHost('Excel');
    const { forHost } = await freshImport();
    expect(forHost({ word: 'w', excel: 'e', default: 'd' })).toBe('e');
  });

  it('returns powerpoint option when host is PowerPoint', async () => {
    setOfficeHost('PowerPoint');
    const { forHost } = await freshImport();
    expect(forHost({ powerpoint: 'pp', default: 'd' })).toBe('pp');
  });

  it('returns outlook option when host is Outlook', async () => {
    setOfficeHost('Outlook');
    const { forHost } = await freshImport();
    expect(forHost({ outlook: 'ol', default: 'd' })).toBe('ol');
  });

  it('falls back to default when the specific option is absent', async () => {
    setOfficeHost('Word');
    const { forHost } = await freshImport();
    expect(forHost({ excel: 'e', default: 'fallback' })).toBe('fallback');
  });

  it('returns default when host is Unknown', async () => {
    setOfficeHost(undefined);
    const { forHost } = await freshImport();
    expect(forHost({ word: 'w', excel: 'e', default: 'unknown-default' })).toBe('unknown-default');
  });

  it('returns undefined when no matching option and no default', async () => {
    setOfficeHost('Excel');
    const { forHost } = await freshImport();
    expect(forHost({ word: 'w' })).toBeUndefined();
  });

  it('returns undefined when host is Unknown and no default provided', async () => {
    setOfficeHost(undefined);
    const { forHost } = await freshImport();
    expect(forHost({ word: 'w' })).toBeUndefined();
  });

  it('handles an empty options object', async () => {
    setOfficeHost('Word');
    const { forHost } = await freshImport();
    expect(forHost({})).toBeUndefined();
  });

  it('returns numeric values correctly', async () => {
    setOfficeHost('Excel');
    const { forHost } = await freshImport();
    expect(forHost({ excel: 42, default: 0 })).toBe(42);
  });

  it('handles boolean values', async () => {
    setOfficeHost('Word');
    const { forHost } = await freshImport();
    expect(forHost({ word: false, default: true })).toBe(false);
  });
});
