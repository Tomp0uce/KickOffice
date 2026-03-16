<template>
  <div class="markdown-renderer" v-html="renderedHtml" />
</template>

<script lang="ts" setup>
import { computed } from 'vue';

import { renderSanitizedMarkdown } from '@/utils/officeRichText';

const props = defineProps<{
  content: string;
}>();

const renderedHtml = computed(() => renderSanitizedMarkdown(props.content));
</script>

<style scoped>
.markdown-renderer {
  overflow-wrap: anywhere;
  min-width: 0;
  max-width: 100%;
}

.markdown-renderer :deep(p) {
  margin: 0 0 0.5rem;
}

.markdown-renderer :deep(p:last-child) {
  margin-bottom: 0;
}

.markdown-renderer :deep(ul) {
  margin: 0.25rem 0 0.5rem;
  padding-left: 1.5rem;
  list-style: disc outside none !important;
  overflow: visible;
}

.markdown-renderer :deep(ol) {
  margin: 0.25rem 0 0.5rem;
  padding-left: 1.5rem;
  list-style: decimal outside none !important;
  overflow: visible;
}

.markdown-renderer :deep(li) {
  display: list-item;
  list-style: inherit;
}

.markdown-renderer :deep(pre) {
  margin: 0.25rem 0;
  overflow-x: auto;
  max-width: 100%;
  border-radius: 0.25rem;
  padding: 0.5rem;
  background-color: color-mix(in srgb, var(--color-bg-secondary) 75%, black);
}

.markdown-renderer :deep(code) {
  font-family:
    ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New',
    monospace;
}
</style>
