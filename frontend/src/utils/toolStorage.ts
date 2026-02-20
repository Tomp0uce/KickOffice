const ENABLED_TOOLS_STORAGE_KEY = 'enabledTools'
const ENABLED_TOOLS_STORAGE_VERSION = 1

interface EnabledToolsStorageState {
  version: number
  signature: string
  enabledToolNames: string[]
}

export function buildToolSignature(allToolNames: string[]): string {
  return allToolNames.slice().sort().join('|')
}

export function persistEnabledTools(allToolNames: string[], enabledToolNames: Set<string>): void {
  const payload: EnabledToolsStorageState = {
    version: ENABLED_TOOLS_STORAGE_VERSION,
    signature: buildToolSignature(allToolNames),
    enabledToolNames: allToolNames.filter(name => enabledToolNames.has(name)),
  }
  localStorage.setItem(ENABLED_TOOLS_STORAGE_KEY, JSON.stringify(payload))
}

export function getEnabledToolNamesFromStorage(allToolNames: string[]): Set<string> {
  const fallback = new Set(allToolNames)
  const allToolNameSet = new Set(allToolNames)
  const toolSignature = buildToolSignature(allToolNames)

  try {
    const stored = localStorage.getItem(ENABLED_TOOLS_STORAGE_KEY)
    if (!stored) {
      persistEnabledTools(allToolNames, fallback)
      return fallback
    }

    const parsed = JSON.parse(stored)

    // Handle legacy raw array format from old SettingsPage versions
    if (Array.isArray(parsed)) {
      const legacyEnabledTools = new Set<string>(parsed.filter((name): name is string => typeof name === 'string' && allToolNameSet.has(name)))
      const enabledTools = legacyEnabledTools.size > 0 ? legacyEnabledTools : fallback
      persistEnabledTools(allToolNames, enabledTools)
      return enabledTools
    }

    const isValidState = parsed
      && typeof parsed === 'object'
      && Array.isArray(parsed.enabledToolNames)
      && typeof parsed.version === 'number'
      && typeof parsed.signature === 'string'

    if (!isValidState || parsed.version !== ENABLED_TOOLS_STORAGE_VERSION || parsed.signature !== toolSignature) {
      persistEnabledTools(allToolNames, fallback)
      return fallback
    }

    const enabledTools = new Set<string>(
      parsed.enabledToolNames.filter((name: unknown): name is string => typeof name === 'string' && allToolNameSet.has(name))
    )
    const result = enabledTools.size > 0 ? enabledTools : fallback
    
    // Self-heal if there are missing/extra tool definitions loaded now vs when saved
    if (result.size !== parsed.enabledToolNames.length) {
      persistEnabledTools(allToolNames, result)
    }

    return result
  } catch {
    persistEnabledTools(allToolNames, fallback)
    return fallback
  }
}
