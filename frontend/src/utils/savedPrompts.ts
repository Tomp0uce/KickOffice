export interface SavedPrompt {
  id: string
  name: string
  systemPrompt: string
  userPrompt: string
}

/**
 * Validates that an object conforms to the SavedPrompt interface.
 */
function isValidSavedPrompt(item: unknown): item is SavedPrompt {
  if (!item || typeof item !== 'object') return false
  const obj = item as Record<string, unknown>
  return (
    typeof obj.id === 'string' &&
    typeof obj.name === 'string' &&
    typeof obj.systemPrompt === 'string' &&
    typeof obj.userPrompt === 'string'
  )
}

export function loadSavedPromptsFromStorage(fallback: SavedPrompt[] = []): SavedPrompt[] {
  const stored = localStorage.getItem('savedPrompts')
  if (!stored) return fallback

  try {
    const parsed = JSON.parse(stored)
    if (!Array.isArray(parsed)) {
      console.warn('[SavedPrompts] Invalid storage format: expected array')
      return fallback
    }

    // Filter and validate each item
    const validPrompts = parsed.filter((item, index) => {
      if (isValidSavedPrompt(item)) {
        return true
      }
      console.warn(`[SavedPrompts] Invalid item at index ${index}, skipping`)
      return false
    })

    return validPrompts
  } catch (err) {
    console.warn('[SavedPrompts] Failed to parse stored prompts:', err)
    return fallback
  }
}
