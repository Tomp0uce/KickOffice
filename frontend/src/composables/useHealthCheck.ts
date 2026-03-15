import { ref, onMounted, onUnmounted, type Ref } from 'vue';
import { healthCheck, fetchModels } from '@/api/backend';
import { logService } from '@/utils/logger';
import type { ModelTier, ModelInfo } from '@/types';
import { HEALTH_CHECK_INTERVAL_MS } from '@/constants/limits';

export function useHealthCheck(
  availableModels: Ref<Record<string, ModelInfo>>,
  selectedModelTier: Ref<ModelTier>,
) {
  const backendOnline = ref(false);
  let intervalId: number | null = null;

  async function runCheck() {
    backendOnline.value = await healthCheck();
    if (!backendOnline.value) return;
    try {
      availableModels.value = await fetchModels();
      if (!availableModels.value[selectedModelTier.value]) {
        const [firstTier] = Object.keys(availableModels.value);
        if (firstTier) selectedModelTier.value = firstTier as ModelTier;
      }
    } catch {
      logService.error('Failed to fetch models');
    }
  }

  // Interval polls are skipped when the document is hidden to avoid
  // unnecessary requests. The initial check and visibility-change handler
  // always run regardless so Office add-ins (where visibilityState can
  // remain 'hidden' inside a WebView2 iframe) still get an accurate status.
  async function checkBackend() {
    if (document.visibilityState === 'hidden') return;
    await runCheck();
  }

  function onVisibilityChange() {
    if (document.visibilityState !== 'hidden') runCheck();
  }

  function startPolling() {
    runCheck(); // Initial check always runs, ignoring visibility state
    intervalId = window.setInterval(checkBackend, HEALTH_CHECK_INTERVAL_MS);
  }

  function stopPolling() {
    if (intervalId !== null) window.clearInterval(intervalId);
  }

  onMounted(() => {
    startPolling();
    document.addEventListener('visibilitychange', onVisibilityChange);
  });

  onUnmounted(() => {
    stopPolling();
    document.removeEventListener('visibilitychange', onVisibilityChange);
  });

  return { backendOnline, checkBackend };
}
