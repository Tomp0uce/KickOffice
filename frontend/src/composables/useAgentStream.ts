import type { Ref } from 'vue';
import { chatStream, type ChatRequestMessage, type TokenUsage } from '@/api/backend';
import type { ModelTier } from '@/types';
import type { DisplayMessage } from '@/types/chat';

interface ToolCallFunction {
  name: string;
  arguments: string;
}

export interface ToolCall {
  id: string;
  type: 'function';
  function: ToolCallFunction;
}

export interface AssistantMessage {
  role: 'assistant';
  content: string;
  tool_calls: ToolCall[];
}

interface StreamResponseChoice {
  message: AssistantMessage;
  finish_reason?: string | null;
}

export interface StreamResponse {
  choices: StreamResponseChoice[];
}

export function useAgentStream() {
  async function executeStream(options: {
    messages: ChatRequestMessage[];
    modelTier: ModelTier;
    tools?: any[];
    abortSignal?: AbortSignal;
    currentAction?: Ref<string>;
    currentAssistantMessage?: DisplayMessage;
    scrollToBottom?: () => Promise<void>;
    accumulateUsage?: (usage: TokenUsage) => void;
  }) {
    let response: StreamResponse = {
      choices: [{ message: { role: 'assistant', content: '', tool_calls: [] } }],
    };
    let truncatedByLength = false;

    await chatStream({
      messages: options.messages,
      modelTier: options.modelTier,
      tools: options.tools,
      abortSignal: options.abortSignal,
      onStream: text => {
        if (options.currentAction) options.currentAction.value = '';
        if (options.currentAssistantMessage) options.currentAssistantMessage.content = text;
        response.choices[0].message.content = text;
        // Auto-follow if user is near bottom; respects isAutoScrollEnabled flag inside scrollToBottom.
        options.scrollToBottom?.();
      },
      onToolCallDelta: toolCallDeltas => {
        if (options.currentAction) options.currentAction.value = '';
        for (const delta of toolCallDeltas) {
          const idx = delta.index;
          if (!response.choices[0].message.tool_calls[idx]) {
            response.choices[0].message.tool_calls[idx] = {
              id: delta.id,
              type: 'function',
              function: { name: delta.function.name || '', arguments: '' },
            };
          }
          if (delta.function?.arguments) {
            response.choices[0].message.tool_calls[idx].function.arguments +=
              delta.function.arguments;
          }
        }
      },
      onFinishReason: finishReason => {
        if (finishReason === 'length') truncatedByLength = true;
      },
      onUsage: options.accumulateUsage,
    });

    if (response.choices[0].message.tool_calls) {
      response.choices[0].message.tool_calls =
        response.choices[0].message.tool_calls.filter(Boolean);
    }

    return { response, truncatedByLength };
  }

  return { executeStream };
}
