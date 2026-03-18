/**
 * useSkillCreator.ts
 *
 * Calls POST /api/skill-creator to generate a skill from a natural language description.
 * The LLM system prompt and tool knowledge are embedded in the backend — not exposed here.
 */

import { ref } from 'vue';
import { logService } from '@/utils/logger';
import { fetchWithTimeoutAndRetry, getGlobalHeaders, generateRequestId } from '@/api/httpClient';
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser';

const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || '';

export interface SkillCreatorResult {
  name: string;
  description: string;
  host: SkillHost;
  executionMode: SkillExecutionMode;
  icon: string;
  skillContent: string;
}

export function useSkillCreator() {
  const generating = ref(false);
  const error = ref<string | null>(null);

  async function generateSkill(
    description: string,
    host: SkillHost,
  ): Promise<SkillCreatorResult | null> {
    generating.value = true;
    error.value = null;

    try {
      const globalHeaders = await getGlobalHeaders();
      const res = await fetchWithTimeoutAndRetry(
        `${BACKEND_URL}/api/skill-creator`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            ...globalHeaders,
            'X-Request-Id': generateRequestId(),
          },
          body: JSON.stringify({ description, host }),
        },
        'standard', // use standard tier timeout
      );

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
        throw new Error((err as { error?: string }).error || `HTTP ${res.status}`);
      }

      return (await res.json()) as SkillCreatorResult;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown error';
      logService.error('[SkillCreator] Generation failed', err);
      error.value = msg;
      return null;
    } finally {
      generating.value = false;
    }
  }

  return { generating, error, generateSkill };
}
