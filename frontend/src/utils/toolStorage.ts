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

/**
 * Migrates tool preferences when tool definitions change.
 * - Preserves enabled state for tools that still exist
 * - New tools are enabled by default
 * - Removed tools are silently dropped
 * Returns the migrated set of enabled tool names.
 */
function migrateToolPreferences(
  storedEnabledNames: string[],
  allToolNames: string[]
): Set<string> {
  const allToolNameSet = new Set(allToolNames)
  const storedEnabledSet = new Set(storedEnabledNames)

  // Start with tools that were enabled AND still exist
  const migratedEnabled = new Set<string>()
  for (const name of storedEnabledNames) {
    if (allToolNameSet.has(name)) {
      migratedEnabled.add(name)
    }
  }

  // New tools (not in stored set) are enabled by default
  for (const name of allToolNames) {
    if (!storedEnabledSet.has(name) && !storedEnabledNames.includes(name)) {
      // This is a new tool - check if it was explicitly disabled or just new
      // Since we don't track disabled tools, we enable new tools by default
      migratedEnabled.add(name)
    }
  }

  // Log migration info if there were changes
  const addedTools = allToolNames.filter(n => !storedEnabledSet.has(n))
  const removedTools = storedEnabledNames.filter(n => !allToolNameSet.has(n))
  if (addedTools.length > 0 || removedTools.length > 0) {
    console.info('[ToolStorage] Migrated tool preferences', {
      addedTools: addedTools.length,
      removedTools: removedTools.length,
      preservedEnabled: migratedEnabled.size,
    })
  }

  return migratedEnabled
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

    if (!isValidState) {
      persistEnabledTools(allToolNames, fallback)
      return fallback
    }

    // If signature changed, migrate preferences instead of resetting
    if (parsed.signature !== toolSignature) {
      const migratedTools = migrateToolPreferences(parsed.enabledToolNames, allToolNames)
      persistEnabledTools(allToolNames, migratedTools)
      return migratedTools
    }

    // Version mismatch with same signature — migrate as well
    if (parsed.version !== ENABLED_TOOLS_STORAGE_VERSION) {
      const migratedTools = migrateToolPreferences(parsed.enabledToolNames, allToolNames)
      persistEnabledTools(allToolNames, migratedTools)
      return migratedTools
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
