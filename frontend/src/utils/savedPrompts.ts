export interface SavedPrompt {
  id: string
  name: string
  systemPrompt: string
  userPrompt: string
}

export function loadSavedPromptsFromStorage(fallback: SavedPrompt[] = []): SavedPrompt[] {
  const stored = localStorage.getItem('savedPrompts')
  if (!stored) return fallback

  try {
    const parsed = JSON.parse(stored)
    if (Array.isArray(parsed)) {
      return parsed as SavedPrompt[]
    }
  } catch {
    // Ignore invalid payload and fallback
  }

  return fallback
}
