<template>
  <div
    v-if="hasStats || modelName"
    class="flex items-center justify-between border-t border-border-secondary bg-bg-secondary px-2 py-1 text-[10px] text-secondary"
    style="font-family: monospace"
  >
    <!-- Token counts -->
    <div class="flex items-center gap-2">
      <span v-if="hasStats" :title="`Input tokens: ${sessionStats.inputTokens}`">
        ↑{{ formatTokens(sessionStats.inputTokens) }}
      </span>
      <span v-if="hasStats" :title="`Output tokens: ${sessionStats.outputTokens}`">
        ↓{{ formatTokens(sessionStats.outputTokens) }}
      </span>
      <div
        v-if="contextWindowTokens > 0 && sessionStats.inputTokens > 0"
        class="ml-1 flex items-center gap-1 w-20"
        :title="`Context usage: ${sessionStats.inputTokens} / ${contextWindowTokens} tokens`"
      >
        <div class="h-1.5 flex-1 bg-border rounded-full overflow-hidden">
          <div class="h-full bg-accent transition-all" :style="{ width: contextPct + '%' }"></div>
        </div>
        <span class="text-[9px]">{{ contextPct }}%</span>
      </div>
    </div>

    <!-- Model name -->
    <div v-if="modelName" class="truncate ml-2 text-secondary/80">
      {{ modelName }}
    </div>
  </div>
</template>

<script lang="ts" setup>
import { computed } from 'vue'

interface TokenStats {
  inputTokens: number
  outputTokens: number
  totalTokens: number
}

const props = defineProps<{
  sessionStats: TokenStats
  modelName?: string
  contextWindowTokens?: number
  currentAction?: string
  loading?: boolean
}>()

const hasStats = computed(
  () => props.sessionStats.inputTokens > 0 || props.sessionStats.outputTokens > 0,
)

const contextPct = computed(() => {
  if (!props.contextWindowTokens || props.contextWindowTokens === 0) return '0'
  return ((props.sessionStats.inputTokens / props.contextWindowTokens) * 100).toFixed(1)
})

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}k`
  return n.toString()
}
</script>
