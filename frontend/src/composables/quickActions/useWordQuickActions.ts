/**
 * useWordQuickActions — Quick action definitions for the Word host.
 * Extracted from HomePage.vue as part of UX-H1 / QUAL-H2.
 */
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';
import { BookOpen, CheckCheck, FileCheck, Globe, Sparkle } from 'lucide-vue-next';
import type { QuickAction } from '@/types/chat';

export function useWordQuickActions() {
  const { t } = useI18n();

  const wordQuickActions = computed<QuickAction[]>(() => [
    {
      key: 'word-proofread',
      label: t('proofread'),
      icon: CheckCheck,
      executeWithAgent: true,
      tooltipKey: 'proofread_tooltip',
    },
    {
      key: 'word-translate',
      label: t('translate'),
      icon: Globe,
      executeWithAgent: true,
      tooltipKey: 'translate_tooltip',
    },
    {
      key: 'word-review',
      label: t('wordReview', 'Review'),
      icon: BookOpen,
      executeWithAgent: true,
      tooltipKey: 'wordReview_tooltip',
    },
    {
      key: 'polish',
      label: t('polish'),
      icon: Sparkle,
      tooltipKey: 'polish_tooltip',
    },
    {
      key: 'summary',
      label: t('summary'),
      icon: FileCheck,
      tooltipKey: 'summary_tooltip',
    },
  ]);

  return { wordQuickActions };
}
