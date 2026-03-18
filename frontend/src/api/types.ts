import type { ModelTier } from '@/types';

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string | any[];
  tool_calls?: Array<{
    id: string;
    type: 'function';
    function: {
      name: string;
      arguments: string;
    };
  }>;
}

export interface ToolChatMessage {
  role: 'tool';
  tool_call_id: string;
  content: string;
}

export type ChatRequestMessage = ChatMessage | ToolChatMessage;

export interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
}

export interface ChatStreamOptions {
  messages: ChatRequestMessage[];
  modelTier: ModelTier;
  tools?: ApiToolDefinition[];
  onStream: (text: string) => void;
  onToolCallDelta?: (toolCallDeltas: any[]) => void;
  onFinishReason?: (finishReason: string | null) => void;
  onUsage?: (usage: TokenUsage) => void;
  abortSignal?: AbortSignal;
}

export interface ApiToolDefinition {
  type: 'function';
  function: {
    name: string;
    description?: string;
    parameters: Record<string, any>;
    strict?: boolean;
  };
}

export interface ImageGenerateOptions {
  prompt: string;
  size?: string;
  quality?: string;
  abortSignal?: AbortSignal;
}

export interface PlotAreaBox {
  /** Left edge of the chart's plot area. Value in [0,1] = fraction of image width; value > 1 = raw pixels. */
  xMinPx: number;
  /** Right edge of the chart's plot area. */
  xMaxPx: number;
  /** Top edge of the chart's plot area (smaller pixel value = higher on screen). */
  yMinPx: number;
  /** Bottom edge of the chart's plot area (larger pixel value = lower on screen, where X axis sits). */
  yMaxPx: number;
}

export interface ChartExtractParams {
  imageId: string;
  xAxisRange: [number, number];
  yAxisRange: [number, number];
  targetColor: string;
  plotAreaBox: PlotAreaBox;
  chartType?: string;
  colorTolerance?: number;
  numPoints?: number;
}

export interface ChartExtractResult {
  points: Array<{ x: number; y: number }>;
  pixelsMatched: number;
  imageSize: { width: number; height: number };
  plotBounds?: { pxMin: number; pxMax: number; pyMin: number; pyMax: number };
  warning?: string;
}

export interface FeedbackSystemContext {
  host: string;
  appVersion: string;
  modelTier: string;
  userAgent: string;
}
