/**
 * usePowerPointQuickActions — Quick action definitions for the PowerPoint host.
 * Extracted from HomePage.vue as part of UX-H1 / QUAL-H2.
 */
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';
import { CheckCheck, Globe, Image, ScanSearch, Zap } from 'lucide-vue-next';
import type { PowerPointQuickAction } from '@/types/chat';

export function usePowerPointQuickActions() {
  const { t } = useI18n();

  const powerPointQuickActions = computed<PowerPointQuickAction[]>(() => [
    {
      key: 'ppt-proofread',
      label: t('proofread'),
      icon: CheckCheck,
      mode: 'immediate',
      executeWithAgent: true,
      tooltipKey: 'ppt_proofread_tooltip',
    },
    {
      key: 'ppt-translate',
      label: t('translate'),
      icon: Globe,
      mode: 'immediate',
      executeWithAgent: true,
      tooltipKey: 'translate_tooltip',
    },
    {
      // PPT-H2: replaced speakerNotes with review — no text selection required
      key: 'review',
      label: t('pptReview'),
      icon: ScanSearch,
      mode: 'immediate',
      tooltipKey: 'pptReview_tooltip',
    },
    {
      key: 'punchify',
      label: t('pptPunchify'),
      icon: Zap,
      mode: 'immediate',
      tooltipKey: 'pptPunchify_tooltip',
      executeWithAgent: true,
    },
    {
      key: 'visual',
      label: t('pptVisual'),
      icon: Image,
      mode: 'immediate',
      tooltipKey: 'pptVisual_tooltip',
    },
  ]);

  return { powerPointQuickActions };
}
