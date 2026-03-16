/**
 * useOutlookQuickActions — Quick action definitions for the Outlook host.
 * Extracted from HomePage.vue as part of UX-H1 / QUAL-H2.
 */
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';
import { CheckCheck, Globe, ListTodo, Mail, NotebookPen } from 'lucide-vue-next';
import type { OutlookQuickAction } from '@/types/chat';

export function useOutlookQuickActions() {
  const { t } = useI18n();

  const outlookQuickActions = computed<OutlookQuickAction[]>(() => [
    {
      key: 'proofread',
      label: t('outlookProofread'),
      icon: CheckCheck,
      tooltipKey: 'outlookProofread_tooltip',
    },
    {
      key: 'translate',
      label: t('translate'),
      icon: Globe,
      tooltipKey: 'translate_tooltip',
    },
    {
      key: 'reply',
      label: t('outlookReply'),
      icon: Mail,
      mode: 'smart-reply',
      prefix: t('outlookReplyPrePrompt'),
      tooltipKey: 'outlookReply_tooltip',
    },
    {
      key: 'extract',
      label: t('outlookExtract'),
      icon: ListTodo,
      tooltipKey: 'outlookExtract_tooltip',
    },
    {
      key: 'mom',
      label: t('outlookMoM', 'MoM'),
      icon: NotebookPen,
      mode: 'mom',
      prefix: t('outlookMoMPrefix', 'Génère moi un compte rendu de réunion pour ces notes de réunion : '),
      tooltipKey: 'outlookMoM_tooltip',
    },
  ]);

  return { outlookQuickActions };
}
