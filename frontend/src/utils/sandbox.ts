import { ensureLockdown } from './lockdown'

/* global Compartment */

export type SandboxHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook'

/**
 * Execute code in a sandboxed environment with SES.
 *
 * @param code - JavaScript code to execute
 * @param globals - Global variables to expose (context, Word/Excel/etc)
 * @param host - Optional host to restrict available namespaces
 * @returns Promise with execution result
 */
export function sandboxedEval(
  code: string,
  globals: Record<string, any>,
  host?: SandboxHost
): unknown {
  ensureLockdown()

  // Build filtered globals based on host
  const filteredGlobals = buildHostGlobals(globals, host)

  // @ts-ignore - Compartment is provided by SES
  const compartment = new Compartment({
    globals: {
      ...filteredGlobals,
      // Safe built-ins
      console,
      Math,
      Date,
      JSON,
      Array,
      Object,
      String,
      Number,
      Boolean,
      Promise,
      // Explicitly blocked
      Function: undefined,
      Reflect: undefined,
      Proxy: undefined,
      Compartment: undefined,
      harden: undefined,
      lockdown: undefined,
      eval: undefined,
      // Blocked browser APIs that could cause issues
      fetch: undefined,
      XMLHttpRequest: undefined,
      WebSocket: undefined,
    },
    __options__: true,
  })

  // Audit trail: log host and truncated code for debugging
  const preview = code.length > 200 ? `${code.slice(0, 200)}…` : code
  console.info(`[sandbox] host=${host ?? 'unspecified'} code=${preview}`)

  // Wrap in async IIFE and execute
  return compartment.evaluate(`(async () => { ${code} })()`)
}

/**
 * Build globals object filtered by host.
 * Prevents cross-host API access (e.g., using Word API in Excel).
 */
function buildHostGlobals(
  globals: Record<string, any>,
  host?: SandboxHost
): Record<string, any> {
  const result = { ...globals }

  // If no host specified, allow all (backwards compatibility)
  if (!host) {
    return result
  }

  // Remove namespaces not matching the current host
  const namespaceMap: Record<SandboxHost, string[]> = {
    Word: ['Excel', 'PowerPoint'],      // Remove these from Word
    Excel: ['Word', 'PowerPoint'],       // Remove these from Excel
    PowerPoint: ['Word', 'Excel'],       // Remove these from PowerPoint
    Outlook: ['Word', 'Excel', 'PowerPoint'],  // Remove all from Outlook
  }

  const toRemove = namespaceMap[host] || []
  for (const ns of toRemove) {
    if (ns in result) {
      result[ns] = undefined
    }
  }

  return result
}

/**
 * Create a safe error message from an execution error.
 * Strips sensitive information while keeping useful details.
 */
function sanitizeExecutionError(error: any): string {
  const message = error?.message || String(error)

  // Remove stack traces that might expose internal paths
  const sanitized = message
    .replace(/at\s+.*:\d+:\d+/g, '')  // Remove stack trace lines
    .replace(/\n\s*\n/g, '\n')         // Remove empty lines
    .trim()

  return sanitized || 'Unknown error occurred during code execution'
}
