/**
 * useExcelQuickActions — Quick action definitions for the Excel host.
 * Extracted from HomePage.vue as part of UX-H1 / QUAL-H2.
 */
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';
import {
  BookOpen,
  ChartBarBig,
  FunctionSquare,
  Grid3X3,
  Table,
  TrendingUp,
} from 'lucide-vue-next';
import type { ExcelQuickAction } from '@/types/chat';

export function useExcelQuickActions() {
  const { t } = useI18n();

  const excelQuickActions = computed<ExcelQuickAction[]>(() => [
    {
      key: 'ingest',
      label: t('excelIngest', 'Smart Ingestion'),
      icon: Table,
      mode: 'immediate',
      executeWithAgent: true,
      tooltipKey: 'excelIngest_tooltip',
    },
    {
      key: 'digitizeChart',
      label: t('excelDigitizeChart', 'Digitize Chart'),
      icon: ChartBarBig,
      mode: 'immediate',
      executeWithAgent: true,
      imageUpload: true,
      tooltipKey: 'excelDigitizeChart_tooltip',
    },
    {
      key: 'explain',
      label: t('excelExplain', 'Explain Formula'),
      icon: BookOpen,
      mode: 'immediate',
      executeWithAgent: true,
      systemPrompt:
        'You are an Excel expert. Explain the selected formula or data in simple terms: what it does, how it works, and any edge cases to be aware of.',
      tooltipKey: 'excelExplain_tooltip',
    },
    {
      key: 'formulaGenerator',
      label: t('excelFormulaGenerator', 'Formula Generator'),
      icon: FunctionSquare,
      mode: 'draft',
      prefix: t('excelFormulaGeneratorPrefix', 'Help me build a formula'),
      tooltipKey: 'excelFormulaGenerator_tooltip',
    },
    {
      key: 'dataTrend',
      label: t('excelDataTrend', 'Data Trend'),
      icon: TrendingUp,
      mode: 'immediate',
      executeWithAgent: true,
      systemPrompt:
        'You are a data analyst. Analyze the trends in the selected data: identify patterns, outliers, growth rates, and provide a concise summary with actionable insights.',
      tooltipKey: 'excelDataTrend_tooltip',
    },
    {
      key: 'pixelArt',
      label: t('excelPixelArt', 'Pixel Art'),
      icon: Grid3X3,
      mode: 'immediate',
      executeWithAgent: true,
      imageUpload: true,
      tooltipKey: 'excelPixelArt_tooltip',
    },
  ]);

  return { excelQuickActions };
}
