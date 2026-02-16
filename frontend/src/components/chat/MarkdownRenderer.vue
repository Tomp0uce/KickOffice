<template>
  <div class="markdown-renderer" v-html="renderedHtml" />
</template>

<script lang="ts" setup>
import { computed } from 'vue'

import { renderSanitizedMarkdown } from '@/utils/markdown'

const props = defineProps<{
  content: string
}>()

const renderedHtml = computed(() => renderSanitizedMarkdown(props.content))
</script>

<style scoped>
.markdown-renderer {
  overflow-wrap: anywhere;
}

.markdown-renderer :deep(p) {
  margin: 0 0 0.5rem;
}

.markdown-renderer :deep(p:last-child) {
  margin-bottom: 0;
}

.markdown-renderer :deep(ul),
.markdown-renderer :deep(ol) {
  margin: 0.25rem 0 0.5rem;
  padding-left: 1.25rem;
}

.markdown-renderer :deep(pre) {
  margin: 0.25rem 0;
  overflow-x: auto;
  border-radius: 0.25rem;
  padding: 0.5rem;
  background-color: color-mix(in srgb, var(--color-bg-secondary) 75%, black);
}

.markdown-renderer :deep(code) {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace;
}
</style>
