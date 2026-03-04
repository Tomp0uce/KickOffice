/**
 * Global store for rich content context
 * Stores the last extracted rich content context for use in tool execution
 */

import type { RichContentContext } from './richContentPreserver'
import { ref } from 'vue'

// Global reactive store for the last rich context
const lastRichContext = ref<RichContentContext | null>(null)

export function setLastRichContext(context: RichContentContext | null): void {
  lastRichContext.value = context
}

export function getLastRichContext(): RichContentContext | null {
  return lastRichContext.value
}

export function clearLastRichContext(): void {
  lastRichContext.value = null
}
