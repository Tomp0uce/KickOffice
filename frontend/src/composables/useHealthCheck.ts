import { ref, onMounted, onUnmounted, type Ref } from 'vue'
import { healthCheck, fetchModels } from '@/api/backend'
import { logService } from '@/utils/logger'
import type { ModelTier, ModelInfo } from '@/types'
import { HEALTH_CHECK_INTERVAL_MS } from '@/constants/limits'

export function useHealthCheck(
  availableModels: Ref<Record<string, ModelInfo>>,
  selectedModelTier: Ref<ModelTier>
) {
  const backendOnline = ref(false)
  let intervalId: number | null = null

  async function checkBackend() {
    if (document.visibilityState === 'hidden') return
    backendOnline.value = await healthCheck()
    if (!backendOnline.value) return
    try {
      availableModels.value = await fetchModels()
      if (!availableModels.value[selectedModelTier.value]) {
        const [firstTier] = Object.keys(availableModels.value)
        if (firstTier) selectedModelTier.value = firstTier as ModelTier
      }
    } catch {
      logService.error('Failed to fetch models')
    }
  }

  function startPolling() {
    checkBackend()
    intervalId = window.setInterval(checkBackend, HEALTH_CHECK_INTERVAL_MS)
  }

  function stopPolling() {
    if (intervalId !== null) window.clearInterval(intervalId)
  }

  onMounted(startPolling)
  onUnmounted(stopPolling)

  return { backendOnline, checkBackend }
}
