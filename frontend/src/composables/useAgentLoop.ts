import type { ModelTier, ModelInfo, ToolCategory } from '@/types';
import type { UndoSnapshot } from '@/composables/useDocumentUndo';
import { ref, type Ref, type ComputedRef } from 'vue';

import {
  type ChatMessage,
  type ChatRequestMessage,
  type TokenUsage,
  uploadFile,
  uploadFileToPlatform,
  categorizeError,
  chatStream,
  generateImage,
} from '@/api/backend';
import {
  type ExcelFormulaLanguage,
  GLOBAL_STYLE_INSTRUCTIONS,
  getOutlookBuiltInPrompt,
} from '@/utils/constant';
import { getDisplayLanguage } from '@/utils/common';
import { getGeneralToolDefinitions } from '@/utils/generalTools';
import { message as messageUtil } from '@/utils/message';
import { powerpointImageRegistry } from '@/utils/powerpointTools';
import { prepareMessagesForContext, estimateContextUsagePercent } from '@/utils/tokenManager';
import { getEnabledToolNamesFromStorage } from '@/utils/toolStorage';
import { getToolsForHost } from '@/utils/toolProviderRegistry';
import { getPreservationInstruction, extractTextFromHtml } from '@/utils/richContentPreserver';
import { useAgentPrompts } from '@/composables/useAgentPrompts';
import { useOfficeSelection } from '@/composables/useOfficeSelection';
import {
  setLastRichContext,
  clearLastRichContext,
  getLastRichContext,
} from '@/utils/richContextStore';
import { areCredentialsConfigured } from '@/utils/credentialStorage';
import { logService } from '@/utils/logger';
import { writeFile as vfsWriteFile } from '@/utils/vfs';

import type {
  DisplayMessage,
  ExcelQuickAction,
  PowerPointQuickAction,
  OutlookQuickAction,
  QuickAction,
} from '@/types/chat';

import { useAgentStream, type StreamResponse } from './useAgentStream';
import { executeAgentToolCall } from './useToolExecutor';
import { useLoopDetection } from './useLoopDetection';
import { useSessionFiles } from './useSessionFiles';
import { useQuickActions } from './useQuickActions';
import { useMessageOrchestration } from './useMessageOrchestration';
interface AgentLoopRefs {
  history: Ref<DisplayMessage[]>;
  userInput: Ref<string>;
  loading: Ref<boolean>;
  imageLoading: Ref<boolean>;
  backendOnline: Ref<boolean>;
  abortController: Ref<AbortController | null>;
  inputTextarea: Ref<HTMLTextAreaElement | undefined>;
  isDraftFocusGlowing: Ref<boolean>;
}

interface AgentLoopModels {
  availableModels: Ref<Record<string, ModelInfo>>;
  selectedModelTier: Ref<ModelTier>;
  selectedModelInfo: Ref<ModelInfo | undefined>;
  firstChatModelTier: Ref<ModelTier>;
}

interface AgentLoopHost {
  isOutlook: boolean;
  isPowerPoint: boolean;
  isExcel: boolean;
  isWord: boolean;
}

interface AgentLoopSettings {
  agentMaxIterations: Ref<number>;
  excelFormulaLanguage: Ref<ExcelFormulaLanguage>; // TOOL-M4: extended from 'en' | 'fr'
  userGender: Ref<string>;
  userFirstName: Ref<string>;
  userLastName: Ref<string>;
}

interface AgentLoopActions {
  quickActions: ComputedRef<QuickAction[] | undefined>;
  outlookQuickActions?: Ref<OutlookQuickAction[]>;
  excelQuickActions: Ref<ExcelQuickAction[]>;
  powerPointQuickActions: Ref<PowerPointQuickAction[]>;
}

interface AgentLoopHelpers {
  createDisplayMessage: (
    role: DisplayMessage['role'],
    content: string,
    imageSrc?: string,
  ) => DisplayMessage;
  adjustTextareaHeight: () => void;
  scrollToBottom: () => Promise<void>;
  scrollToMessageTop?: () => Promise<void>;
  scrollToVeryBottom?: () => Promise<void>;
  /**
   * Capture full document state (body OOXML for Word) before the agent runs.
   * Preferred over captureBeforeInsert for the agent loop because agent tools that replace
   * whole paragraphs cannot destroy a body-OOXML snapshot (unlike the CC-based approach).
   */
  captureDocumentState?: () => Promise<Partial<UndoSnapshot> | null>;
  /** Capture selection state before a quick action so undo is available afterwards */
  captureBeforeInsert?: () => Promise<Partial<UndoSnapshot> | null>;
  /** Save the snapshot captured by captureDocumentState/captureBeforeInsert and mark undo available */
  saveSnapshot?: (partial: Partial<UndoSnapshot>, tag?: string) => void;
}

interface UseAgentLoopOptions {
  t: (key: string) => string;
  refs: AgentLoopRefs;
  models: AgentLoopModels;
  host: AgentLoopHost;
  settings: AgentLoopSettings;
  actions: AgentLoopActions;
  helpers: AgentLoopHelpers;
}

export interface SessionStats {
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  /** Input tokens for the most recent API call only — used for context-window bar */
  lastCallInputTokens: number;
}

export function useAgentLoop(options: UseAgentLoopOptions) {
  const { t, refs, models, host, settings, actions, helpers } = options;

  const { executeStream } = useAgentStream();
  const { addSignatureAndCheckLoop, clearSignatures } = useLoopDetection(5, 2);

  // ARCH-H1: Extract session files management
  const { addSessionFile, rebuildSessionFiles, getSessionFilesForChat } = useSessionFiles({
    history: refs.history,
  });
  const sessionUploadedImages = ref<
    { filename: string; dataUri: string; imageId?: string; fileId?: string }[]
  >([]);

  // ARCH-H1: Extract message orchestration
  const hostIsWord = !host.isOutlook && !host.isPowerPoint && !host.isExcel;
  const { prepareMessages } = useMessageOrchestration({
    history: refs.history,
    hostIsOutlook: host.isOutlook,
    hostIsPowerPoint: host.isPowerPoint,
    hostIsExcel: host.isExcel,
    hostIsWord,
  });

  // Destructure refs
  const {
    history,
    userInput,
    loading,
    imageLoading,
    backendOnline,
    abortController,
    inputTextarea,
    isDraftFocusGlowing,
  } = refs;

  // Destructure models
  const { availableModels, selectedModelTier, selectedModelInfo, firstChatModelTier } = models;

  // Destructure host flags (aliased to match existing code)
  const { isOutlook: hostIsOutlook, isPowerPoint: hostIsPowerPoint, isExcel: hostIsExcel } = host;

  // Destructure settings
  const { agentMaxIterations, excelFormulaLanguage, userGender, userFirstName, userLastName } =
    settings;

  // Destructure actions
  const { quickActions, outlookQuickActions, excelQuickActions, powerPointQuickActions } = actions;

  // Destructure helpers
  const {
    createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop = scrollToBottom, // fallback to scrollToBottom if not provided
    captureDocumentState,
    captureBeforeInsert,
    saveSnapshot,
  } = helpers;

  const currentAction = ref('');
  const pendingMoM = ref(false);

  const sessionStats = ref({
    inputTokens: 0,
    outputTokens: 0,
    totalTokens: 0,
    lastCallInputTokens: 0,
  });

  function resetSessionStats() {
    sessionStats.value = {
      inputTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      lastCallInputTokens: 0,
    };
  }

  function accumulateUsage(usage: TokenUsage) {
    sessionStats.value.inputTokens += usage.promptTokens;
    sessionStats.value.outputTokens += usage.completionTokens;
    sessionStats.value.totalTokens += usage.totalTokens;
    // Track the last individual call for the context-window bar (cumulative inflates past 100%)
    sessionStats.value.lastCallInputTokens = usage.promptTokens;
  }

  const getActionLabelForCategory = (category?: ToolCategory) => {
    switch (category) {
      case 'read':
        return t('agentActionReading');
      case 'format':
        return t('agentActionFormatting');
      case 'write':
      default:
        return t('agentActionRunning');
    }
  };

  const { agentPrompt } = useAgentPrompts({
    t,
    userGender,
    userFirstName,
    userLastName,
    excelFormulaLanguage,
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
  });

  const { getOfficeSelection, getOfficeSelectionAsHtml } = useOfficeSelection({
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
  });

  const resolveChatModelTier = (): ModelTier =>
    selectedModelInfo.value?.type === 'image' ? firstChatModelTier.value : selectedModelTier.value;

  async function runAgentLoop(messages: ChatMessage[], modelTier: ModelTier) {
    // ARCH-M1: Use ToolProviderRegistry instead of direct imports
    const appToolDefs = getToolsForHost({
      isOutlook: hostIsOutlook,
      isPowerPoint: hostIsPowerPoint,
      isExcel: hostIsExcel,
    });

    const generalToolDefs = getGeneralToolDefinitions();
    const allToolDefs = [...generalToolDefs, ...appToolDefs];
    const enabledToolNames = getEnabledToolNamesFromStorage(allToolDefs.map(def => def.name));
    const enabledToolDefs = allToolDefs.filter(def => enabledToolNames.has(def.name));
    const tools = enabledToolDefs.map(def => ({
      type: 'function' as const,
      function: {
        name: def.name,
        description: def.description,
        parameters: def.inputSchema as Record<string, unknown>,
      },
    }));

    // Add preservation instruction to system prompt if we have rich content
    const richContext = getLastRichContext();
    if (richContext?.hasRichContent && messages[0]?.role === 'system') {
      messages[0].content += getPreservationInstruction(richContext);
    }

    let iteration = 0;
    const maxIter = Number(agentMaxIterations.value) || 10;
    const startTime = Date.now();
    const timeoutMs = maxIter * 60 * 1000; // up to 1 minute per iteration allowed
    let currentMessages: ChatRequestMessage[] = [...messages];
    // Sliding window loop detection (P6) uses useLoopDetection composable
    clearSignatures();
    let toolsWereExecuted = false; // Track if any tools were successfully executed
    currentAction.value = t('agentAnalyzing');
    history.value.push(createDisplayMessage('assistant', ''));
    // Scroll to bottom so the thinking banner ("Agent Action Indicator") stays visible.
    // We do NOT use scrollToMessageTop here — the empty assistant bubble is hidden (v-show=false)
    // and its getBoundingClientRect() returns zeros, which would scroll to the top of the chat.
    await scrollToBottom();
    const currentAssistantMessage = history.value[history.value.length - 1];
    let abortedByUser = false;
    while (Date.now() - startTime < timeoutMs) {
      if (abortController.value?.signal.aborted) {
        abortedByUser = true;
        break;
      }

      iteration++;

      // Enforce max iterations limit
      if (iteration > maxIter) {
        currentAssistantMessage.content += `\n\n⚠️ ${t('agentMaxIterationsReached')}`;
        break;
      }

      // H11: Show "agentAnalyzing" initially, or "agentWaitingForLLM" if tools were just executed and we are generating a response
      const llmWaitLabel = iteration === 1 ? t('agentAnalyzing') : t('agentWaitingForLLM');
      currentAction.value = llmWaitLabel;
      const llmWaitStart = Date.now();

      const currentSystemPrompt =
        messages[0]?.role === 'system'
          ? typeof messages[0].content === 'string'
            ? messages[0].content
            : ''
          : '';
      const contextPct = estimateContextUsagePercent(currentMessages, currentSystemPrompt);

      const llmWaitTimer = setInterval(() => {
        // Stop the timer once currentAction was cleared externally (streaming started)
        if (!currentAction.value) {
          clearInterval(llmWaitTimer);
          return;
        }
        const elapsed = Math.round((Date.now() - llmWaitStart) / 1000);
        const ctxSuffix = contextPct >= 50 ? ` · ctx ${contextPct}%` : '';
        currentAction.value = `${llmWaitLabel} (${elapsed}s${ctxSuffix})`;
      }, 1000);
      const contextSafeMessages = prepareMessagesForContext(currentMessages, currentSystemPrompt);
      logService.info('llm_request', 'llm', {
        model: modelTier,
        messageCount: contextSafeMessages.length,
        contextPct,
      });

      let response: StreamResponse;
      let truncatedByLength = false;

      try {
        const streamResult = await executeStream({
          messages: contextSafeMessages,
          modelTier,
          tools,
          abortSignal: abortController.value?.signal || undefined,
          currentAction,
          currentAssistantMessage,
          scrollToBottom,
          accumulateUsage,
        });

        clearInterval(llmWaitTimer);
        response = streamResult.response;
        truncatedByLength = streamResult.truncatedByLength;
        logService.info('llm_response_complete', 'llm', {
          tokensUsed: sessionStats.value.totalTokens,
        });
      } catch (err: unknown) {
        clearInterval(llmWaitTimer);
        if (
          (err instanceof Error && err.name === 'AbortError') ||
          abortController.value?.signal.aborted
        ) {
          abortedByUser = true;
          break;
        }
        logService.error('[AgentLoop] chatStream failed', err, {
          host: hostIsOutlook
            ? 'outlook'
            : hostIsPowerPoint
              ? 'powerpoint'
              : hostIsExcel
                ? 'excel'
                : 'word',
          modelTier,
          iteration,
          messageCount: currentMessages.length,
          traffic: 'system',
        });
        const errInfo = categorizeError(err);
        if (errInfo.type === 'auth') {
          currentAssistantMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`;
        } else {
          currentAssistantMessage.content = errInfo.rawDetail
            ? `${t(errInfo.i18nKey)}\n\n> ${errInfo.rawDetail}`
            : t(errInfo.i18nKey);
        }
        // ERR-L2: Mark the message as stream-interrupted so the UI can show a prominent Retry affordance
        const isStreamInterrupted =
          err instanceof Error && err.message.includes('stream_interrupted');
        if (isStreamInterrupted) {
          currentAssistantMessage.streamError = true;
        }
        currentAction.value = '';
        break;
      }

      // Handle finish_reason: "length" — model was cut off mid-response (P7)
      if (truncatedByLength) {
        currentAction.value = '';
        if (!currentAssistantMessage.content?.trim()) {
          currentAssistantMessage.content = t('errorTruncated');
        } else {
          // Append warning to existing content
          currentAssistantMessage.content += `\n\n${t('errorTruncated')}`;
        }
        break;
      }

      const choice = response.choices?.[0];
      if (!choice) break;
      const assistantMsg = choice.message;
      const assistantMsgForHistory: ChatRequestMessage = {
        role: 'assistant',
        content: assistantMsg.content || '',
      };
      // Only include tool_calls if non-empty (Azure/LiteLLM rejects empty arrays)
      if (assistantMsg.tool_calls?.length) {
        assistantMsgForHistory.tool_calls = assistantMsg.tool_calls;
      }
      currentMessages.push(assistantMsgForHistory);
      if (assistantMsg.content) currentAssistantMessage.content = assistantMsg.content;
      if (!assistantMsg.tool_calls?.length) {
        currentAction.value = '';
        break;
      }
      // Collect all tool results before adding to messages (atomic update)
      const toolResults: { tool_call_id: string; content: string }[] = [];
      let toolLoopAborted = false;

      for (const toolCall of assistantMsg.tool_calls) {
        // Check abort before each tool execution
        if (abortController.value?.signal.aborted) {
          toolLoopAborted = true;
          break;
        }

        const toolResult = await executeAgentToolCall(
          toolCall,
          enabledToolDefs,
          currentAssistantMessage,
          currentAction,
          getActionLabelForCategory,
          scrollToBottom,
        );
        const sig = toolResult.signature;

        // Sliding window loop detection (P6) — same signature repeated
        if (sig && addSignatureAndCheckLoop(sig)) {
          toolResults.push({
            tool_call_id: toolCall.id,
            content:
              'Error: You have called this exact tool with the same arguments multiple times in a row. This is a loop. Stop repeating and try a different approach.',
          });
          continue;
        }

        if (toolResult.success) toolsWereExecuted = true;
        if (toolResult.screenshotBase64) {
          const mimeType = toolResult.screenshotMimeType || 'image/png';
          const dataUri = `data:${mimeType};base64,${toolResult.screenshotBase64}`;
          const filename = `screenshot_${Date.now()}.png`;
          sessionUploadedImages.value.push({ filename, dataUri });
        }
        toolResults.push({ tool_call_id: toolResult.tool_call_id, content: toolResult.content });
      }

      // If aborted mid-tool-loop, rollback partial state by removing incomplete assistant message
      if (toolLoopAborted) {
        // Remove the last assistant message with tool_calls since we didn't complete all tools
        const lastMsgIdx = currentMessages.length - 1;
        if (lastMsgIdx >= 0 && currentMessages[lastMsgIdx].role === 'assistant') {
          currentMessages.pop();
        }
        abortedByUser = true;
        break;
      }

      // Atomically add all tool results now that loop completed successfully
      for (const toolResult of toolResults) {
        currentMessages.push({
          role: 'tool',
          tool_call_id: toolResult.tool_call_id,
          content: toolResult.content,
        });
      }

      // H11: Switch status from tool execution to waiting for LLM response
      currentAction.value = t('agentWaitingForLLM');
    }

    // P8: Persist full tool call sequence in history so subsequent turns have context
    const initialMsgCount = messages.length;
    const newMessages = currentMessages.slice(initialMsgCount);
    if (newMessages.length > 0) {
      currentAssistantMessage.rawMessages = newMessages;
    }

    if (abortedByUser) {
      currentAction.value = '';
      history.value.push(createDisplayMessage('system', t('agentStoppedByUser')));
      return;
    }

    const assistantContent = currentAssistantMessage?.content?.trim() || '';
    if (!assistantContent) {
      // If tools were executed successfully but no text response, that's OK (e.g., proofreading with comments)
      if (toolsWereExecuted) {
        currentAssistantMessage.content = t('toolsExecutedSuccessfully');
      } else {
        currentAssistantMessage.content = t('noModelResponse');
      }
    }

    if (Date.now() - startTime >= timeoutMs) messageUtil.warning(t('recursionLimitExceeded'));
    currentAction.value = '';

    // Scroll to the top of the completed assistant response so the user sees it from the start.
    // The assistant message now has content, so scrollToMessageTop finds it correctly.
    await scrollToMessageTop();
  }

  // ARCH-M6: Shared one-shot stream helper for SmartReply / MoM
  async function streamOneShot(messages: ChatMessage[], logTag: string): Promise<void> {
    history.value.push(createDisplayMessage('assistant', ''));
    try {
      await chatStream({
        messages,
        modelTier: resolveChatModelTier(),
        onStream: (text: string) => {
          const idx = history.value.length - 1;
          history.value[idx] = { ...history.value[idx], role: 'assistant', content: text };
        },
        onUsage: accumulateUsage,
        abortSignal: abortController.value?.signal,
      });
      const lastIdx = history.value.length - 1;
      const lastMessage = history.value[lastIdx];
      if (!lastMessage?.content?.trim()) {
        history.value[lastIdx] = { ...lastMessage, content: t('noModelResponse') };
      }
    } catch (err: unknown) {
      if (err instanceof Error && err.name === 'AbortError') return;
      logService.error(`[AgentLoop] ${logTag} chatStream failed`, err);
      const errIdx = history.value.length - 1;
      const errInfo = categorizeError(err);
      if (errInfo.type === 'auth') {
        history.value[errIdx] = {
          ...history.value[errIdx],
          content: `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`,
        };
      } else {
        const errContent = errInfo.rawDetail
          ? `${t(errInfo.i18nKey)}\n\n> ${errInfo.rawDetail}`
          : t(errInfo.i18nKey);
        history.value[errIdx] = { ...history.value[errIdx], content: errContent };
      }
    }
    await scrollToMessageTop();
  }

  async function handleMoM(userMessage: string) {
    pendingMoM.value = false;

    const lang = getDisplayLanguage();

    const momPrompt = getOutlookBuiltInPrompt()['mom'];
    const systemMsg = momPrompt.system(lang) + `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`;
    const userMsg = momPrompt.user(userMessage, lang);

    await streamOneShot(
      [
        { role: 'system', content: systemMsg },
        { role: 'user', content: userMsg },
      ],
      'MoM',
    );
  }

  async function fetchSelectionWithTimeout() {
    let textTimeoutId: ReturnType<typeof setTimeout> | null = null;
    let htmlTimeoutId: ReturnType<typeof setTimeout> | null = null;
    let localSelectedText = '';
    try {
      const timeoutPromise = new Promise<string>((_, reject) => {
        textTimeoutId = setTimeout(() => reject(new Error('getOfficeSelection timeout')), 3000);
      }).catch(() => '') as Promise<string>;

      if (!hostIsExcel) {
        // F1: Extract formatted HTML natively and convert to markdown to preserve styling (Word, PPT, Outlook)
        const htmlPromise = new Promise<string>((_, reject) => {
          htmlTimeoutId = setTimeout(
            () => reject(new Error('getOfficeSelectionAsHtml timeout')),
            3000,
          );
        }).catch(() => '') as Promise<string>;

        try {
          const htmlContent = await Promise.race([
            getOfficeSelectionAsHtml({ includeOutlookSelectedText: true }),
            htmlPromise,
          ]);
          if (htmlContent) {
            const richContext = extractTextFromHtml(htmlContent);
            localSelectedText = richContext.cleanText || localSelectedText;
            // Store rich context globally so tools can access it (especially for Outlook image preservation)
            if (richContext.hasRichContent) {
              setLastRichContext(richContext);
            }
          } else {
            localSelectedText = await Promise.race([
              getOfficeSelection({ includeOutlookSelectedText: true }),
              timeoutPromise,
            ]);
          }
        } catch {
          localSelectedText = await Promise.race([
            getOfficeSelection({ includeOutlookSelectedText: true }),
            timeoutPromise,
          ]);
        }
      } else {
        localSelectedText = await Promise.race([
          getOfficeSelection({ includeOutlookSelectedText: true }),
          timeoutPromise,
        ]);
      }
    } catch (error) {
      logService.warn('[AgentLoop] Failed to fetch selection before sending message', error);
    } finally {
      if (textTimeoutId) clearTimeout(textTimeoutId);
      if (htmlTimeoutId) clearTimeout(htmlTimeoutId);
    }
    return localSelectedText;
  }

  async function processChat(
    userMessage: string,
    visionImages?: Array<{ filename: string; dataUri: string; imageId?: string }>,
    selectionContext?: string,
    uploadedFiles?: Array<{ filename: string; content: string; fileId?: string }>,
  ) {
    const modelConfig = availableModels.value[selectedModelTier.value];
    if (modelConfig?.type === 'image') {
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')));
      await scrollToMessageTop(); // Scroll to top of assistant message
      imageLoading.value = true;
      try {
        // Include document selection context in the image prompt when available so the
        // generated image matches the document content the user is working on.
        let imagePrompt = userMessage;
        if (selectionContext) {
          const truncatedContext =
            selectionContext.length > 2000
              ? selectionContext.slice(0, 2000) + '…'
              : selectionContext;
          imagePrompt = `${userMessage}\n\nDocument context: ${truncatedContext}`;
        }
        const imageSrc = await generateImage({ prompt: imagePrompt });
        const message = history.value[history.value.length - 1];
        message.role = 'assistant';
        message.content = '';
        message.imageSrc = imageSrc;
      } catch (err: unknown) {
        logService.error('[AgentLoop] image generation failed', err);
        const message = history.value[history.value.length - 1];
        const errInfo = categorizeError(err);
        const baseMsg = t(errInfo.i18nKey);
        const detail = err instanceof Error ? err.message : String(err);
        message.role = 'assistant';
        message.content = `${baseMsg}\n\n${detail}`;
        message.imageSrc = undefined;
      } finally {
        imageLoading.value = false;
      }
      await scrollToBottom(); // Final scroll after image loads
      return;
    }

    const lang = getDisplayLanguage();
    const systemPrompt = agentPrompt(lang);
    const modelTier = resolveChatModelTier();

    // ARCH-H1: Use prepareMessages from useMessageOrchestration
    let messages = await prepareMessages(systemPrompt, uploadedFiles);

    // Additional context injections (selection, vision images)
    try {
      if (selectionContext) {
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user');
        if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
          const selectionLabel = hostIsOutlook
            ? 'Selected text'
            : hostIsPowerPoint
              ? 'Selected slide text'
              : hostIsExcel
                ? 'Selected cells'
                : 'Selected text';
          const sanitizedSelection = selectionContext.replace(
            new RegExp('</?document_content>', 'g'),
            '',
          );
          messages[lastUserIdx].content +=
            `\n\nHere is the current context from the user's document (${selectionLabel}). IMPORTANT: First evaluate if this context is relevant to the user's query. If it is not relevant, ignore it completely and answer the query normally.\n\n<document_content>\n${sanitizedSelection}\n</document_content>`;
        }
      }
    } catch (ctxErr) {
      logService.warn('[AgentLoop] Failed to fetch document context', ctxErr);
    }

    // Inject vision images as multipart content into the last user message
    // Point 2 Fix: Use ALL session images for vision injection (Session Persistence)
    if ((visionImages && visionImages.length > 0) || sessionUploadedImages.value.length > 0) {
      const lastUserIdx = messages.map(m => m.role).lastIndexOf('user');
      if (lastUserIdx !== -1) {
        let textContent = messages[lastUserIdx].content || userMessage;
        const imageContextLines: string[] = [];
        for (const img of sessionUploadedImages.value) {
          const idTag = img.imageId ? ` (imageId: ${img.imageId})` : '';
          imageContextLines.push(`- [${img.filename}]${idTag}`);
        }
        if (imageContextLines.length > 0) {
          textContent += `\n\n<uploaded_images>\nThe following images are available in session memory:\n${imageContextLines.join('\n')}\nTo embed an image in a slide, use insertImageOnSlide with the filename. To extract chart data into Excel, use extract_chart_data with the imageId.\n</uploaded_images>`;
        }

        const parts: { type: string; text?: string; image_url?: { url: string } }[] = [
          { type: 'text', text: String(textContent) },
        ];
        for (const img of sessionUploadedImages.value) {
          // Use provider fileId when available (avoids re-sending base64 bytes each iteration)
          const imageUrl = img.fileId ?? img.dataUri;
          parts.push({ type: 'image_url', image_url: { url: imageUrl } });
        }
        (messages[lastUserIdx] as { role: string; content: unknown }).content = parts;
      }
    }

    return await runAgentLoop(messages, modelTier);
  }

  // ARCH-H1: Extract Quick Actions management
  const { applyQuickAction } = useQuickActions({
    t,
    refs: {
      history,
      userInput,
      loading,
      imageLoading,
      backendOnline,
      abortController,
      inputTextarea,
      isDraftFocusGlowing,
    },
    models: { availableModels },
    hosts: { hostIsOutlook, hostIsPowerPoint, hostIsExcel },
    quickActions,
    outlookQuickActions,
    excelQuickActions,
    powerPointQuickActions,
    helpers: {
      createDisplayMessage,
      adjustTextareaHeight,
      scrollToBottom,
      scrollToMessageTop,
      getOfficeSelection,
      getOfficeSelectionAsHtml,
      runAgentLoop,
      resolveChatModelTier,
    },
    pendingMoM,
    captureBeforeInsert,
    saveSnapshot,
  });

  async function sendMessage(payload?: string, files?: File[]) {
    // Clear any previous rich context at the start of a new request
    clearLastRichContext();

    let textToSend = '';

    if (payload) {
      textToSend = payload;
    } else if (userInput.value && typeof userInput.value === 'string') {
      textToSend = userInput.value;
    }

    textToSend = textToSend?.trim() || '';

    if (!textToSend && (!files || files.length === 0)) {
      if (availableModels.value[selectedModelTier.value]?.type !== 'image') {
        return;
      }
    }

    if (loading.value) {
      return;
    }

    loading.value = true;

    if (!backendOnline.value) {
      loading.value = false;
      return messageUtil.error(t('backendOffline'));
    }

    // BUGFIX: Validate credentials are configured before sending request
    const hasCredentials = await areCredentialsConfigured();
    if (!hasCredentials) {
      loading.value = false;
      messageUtil.error(t('credentialsRequired'));
      return;
    }

    if (userInput.value.trim() === textToSend) {
      userInput.value = '';
      adjustTextareaHeight();
    }

    const userMessage = textToSend;

    let isImageFromSelection = false;
    let selectedText = '';

    // For direct image generation from selection
    if (!userMessage && availableModels.value[selectedModelTier.value]?.type === 'image') {
      try {
        selectedText = await getOfficeSelection();
      } catch (err) {
        logService.warn('[AgentLoop] Failed to fetch selection for image generation', err);
      }
      let wordCount = selectedText
        .trim()
        .split(/\s+/)
        .filter(w => w.length > 0).length;

      if (wordCount < 5 && hostIsPowerPoint) {
        try {
          const { executeOfficeAction } = await import('@/utils/officeAction');
          selectedText = await executeOfficeAction(() => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const PPT = (window as unknown as Record<string, any>)['PowerPoint'];
            if (!PPT?.run) return Promise.resolve('');
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return PPT.run(async (context: any) => {
              let activeSlideIndex = 0;
              try {
                if (typeof context.presentation.getSelectedSlides === 'function') {
                  const selectedSlides = context.presentation.getSelectedSlides();
                  selectedSlides.load('items/id');
                  await context.sync();
                  if (selectedSlides.items.length > 0) {
                    const slides = context.presentation.slides;
                    slides.load('items/id');
                    await context.sync();
                    const selectedId = selectedSlides.items[0].id;
                    const idx = slides.items.findIndex((s: { id: string }) => s.id === selectedId);
                    if (idx !== -1) activeSlideIndex = idx;
                  }
                }
              } catch (e) {}

              const slides = context.presentation.slides;
              slides.load('items');
              await context.sync();
              if (activeSlideIndex >= slides.items.length) return '';
              const slide = slides.items[activeSlideIndex];

              const shapes = slide.shapes;
              shapes.load('items');
              await context.sync();

              for (const shape of shapes.items) {
                try {
                  shape.textFrame.textRange.load('text');
                } catch {}
              }
              await context.sync();

              const texts = [];
              for (const shape of shapes.items) {
                try {
                  texts.push((shape.textFrame.textRange.text || '').trim());
                } catch {}
              }
              return texts.filter(Boolean).join('\n');
            });
          });
          wordCount = selectedText
            .trim()
            .split(/\s+/)
            .filter(w => w.length > 0).length;
        } catch (e) {
          logService.warn('[AgentLoop] Fallback to PowerPoint slide content failed', e);
        }
      }

      if (wordCount < 5) {
        loading.value = false;
        return messageUtil.error(t('fileExtractError'));
      }
      isImageFromSelection = true;
    }

    abortController.value = new AbortController();

    // If it's pure selection image, we show the selection as the user message bubble
    const displayMessageText = isImageFromSelection ? selectedText : userMessage;
    history.value.push(createDisplayMessage('user', displayMessageText));
    const userMsgIdx = history.value.length - 1;
    await scrollToMessageTop(); // Scroll to top of user message just sent

    try {
      // MoM interception: when user sends notes after clicking "MoM" quick action
      if (pendingMoM.value && hostIsOutlook) {
        await handleMoM(userMessage);
        return;
      }

      // GEN-L3: Always fetch selected text as Phantom Context (if not already purely image generation)
      if (!isImageFromSelection) {
        selectedText = await fetchSelectionWithTimeout();
      }

      let fullMessage = displayMessageText;

      if (files && files.length > 0) {
        currentAction.value = t('agentUploadingFiles') || 'Extraction des fichiers...';
        try {
          const newTextFiles: Array<{ filename: string; content: string; fileId?: string }> = [];
          for (const file of files) {
            const result = await uploadFile(file);
            if (result.imageBase64) {
              // Try to upload image to /v1/files for provider-side caching (best-effort)
              let imageFileId: string | undefined;
              try {
                const platformResult = await uploadFileToPlatform(file, 'vision');
                if (platformResult.fileId) imageFileId = platformResult.fileId;
              } catch {
                // Provider doesn't support /v1/files for images — inline base64 fallback
                logService.warn(
                  '[AgentLoop] /v1/files upload failed for image — using inline base64',
                  { filename: file.name },
                );
                messageUtil.warning(
                  t('warningFileFallbackInline') ||
                    `Image "${file.name}" sent inline (provider does not support /v1/files)`,
                );
              }

              // Step 4: Store in session images (with imageId for chart extraction)
              sessionUploadedImages.value.push({
                filename: result.filename,
                dataUri: result.imageBase64,
                imageId: result.imageId,
                fileId: imageFileId,
              });

              // Write image binary to VFS so tools like imageToSheet can access it
              try {
                const base64Data = result.imageBase64.replace(/^data:[^;]+;base64,/, '');
                const binaryStr = atob(base64Data);
                const bytes = new Uint8Array(binaryStr.length);
                for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
                vfsWriteFile(`/home/user/uploads/${result.filename}`, bytes).catch(err => {
                  logService.warn('[AgentLoop] VFS write failed for uploaded image', err);
                  // ERR-C3: Surface VFS failure so agent context gap is visible to user
                  messageUtil.warning(
                    t('warningVfsWriteFailed') ||
                      `File context for "${result.filename}" could not be saved — the agent may lack access on the next turn.`,
                  );
                });
              } catch {
                /* best-effort */
              }

              // Point 3 Fix: Store in PPT registry for tool access (by filename AND imageId)
              if (hostIsPowerPoint) {
                const rawBase64 = result.imageBase64.replace(/^data:[^;]+;base64,/, '');
                powerpointImageRegistry.set(result.filename, rawBase64);
                if (result.imageId) powerpointImageRegistry.set(result.imageId, rawBase64);
              }
              // Show a preview thumbnail in the user message bubble
              history.value[userMsgIdx].imageSrc = result.imageBase64;
            } else {
              // Store extracted text in persistent session memory
              const entry: { filename: string; content: string; fileId?: string } = {
                filename: result.filename,
                content: result.extractedText,
              };
              // Try to upload to LLM provider for file_id referencing (best-effort)
              try {
                const platformResult = await uploadFileToPlatform(file);
                if (platformResult.fileId) {
                  entry.fileId = platformResult.fileId;
                }
              } catch {
                // Provider doesn't support /v1/files or network error — fall back to inline content
                logService.warn(
                  '[AgentLoop] /v1/files upload failed — using inline content fallback',
                  { filename: file.name },
                );
                messageUtil.warning(
                  t('warningFileFallbackInline') ||
                    `File "${file.name}" sent inline (provider does not support /v1/files)`,
                );
              }
              addSessionFile(entry);
              // Write to VFS so the agent can access content via vfsReadFile on future turns
              vfsWriteFile(`/home/user/uploads/${result.filename}`, result.extractedText).catch(
                err => {
                  logService.warn('[AgentLoop] VFS write failed for uploaded file', err);
                  // ERR-C3: Surface VFS failure so agent context gap is visible to user
                  messageUtil.warning(
                    t('warningVfsWriteFailed') ||
                      `File context for "${result.filename}" could not be saved — the agent may lack access on the next turn.`,
                  );
                },
              );
              newTextFiles.push({
                filename: result.filename,
                content: result.extractedText,
                fileId: entry.fileId,
              });
            }
          }
          // Persist file info on the user message for session restore
          if (newTextFiles.length > 0) {
            history.value[userMsgIdx].attachedFiles = newTextFiles;
          }
        } catch (uploadObjErr: unknown) {
          logService.error('[AgentLoop] File upload failed', uploadObjErr);
          return messageUtil.error(t('somethingWentWrong'));
        }
      }

      // Step 4: Pass session uploaded files to processChat (inline or file_id reference)
      const uploadedFilesForChat = getSessionFilesForChat();

      // Capture document state before agent modifies the document (enables undo).
      // captureDocumentState snapshots full body OOXML (Word) so para-level replacements
      // by agent tools cannot destroy the undo anchor (unlike the CC approach).
      let undoSnapshot: Partial<UndoSnapshot> | null = null;
      try {
        const captureFn = captureDocumentState ?? captureBeforeInsert;
        if (captureFn) undoSnapshot = await captureFn();
      } catch (snapErr) {
        logService.warn('[AgentLoop] sendMessage: captureDocumentState failed', snapErr);
      }

      // Only append context to standard text chats, not pure image generations
      // selectedText is passed separately to processChat so it never pollutes the UI history
      if (isImageFromSelection) {
        if (hostIsPowerPoint) {
          fullMessage = t('pptVisualPrefix') + '\n' + selectedText;
        } else {
          // Truncate selected text to avoid overwhelming the image model with too much content
          const truncatedText =
            selectedText.length > 2000 ? selectedText.slice(0, 2000) + '…' : selectedText;
          fullMessage = t('imageGenerationPrompt').replace('{text}', truncatedText);
        }
        await processChat(fullMessage.trim(), undefined, undefined, uploadedFilesForChat);
      } else {
        // Pass selectedText as selectionContext: injected into LLM payload only, not shown in UI
        await processChat(
          fullMessage.trim(),
          undefined,
          selectedText || undefined,
          uploadedFilesForChat,
        );
      }

      // Save snapshot so undo button becomes available after agent completes
      if (undoSnapshot && saveSnapshot) {
        try {
          saveSnapshot(undoSnapshot);
        } catch (saveErr) {
          logService.warn('[AgentLoop] sendMessage: saveSnapshot failed', saveErr);
        }
      }
    } catch (error: unknown) {
      if (!(error instanceof Error) || error.name !== 'AbortError') {
        logService.error('[AgentLoop] sendMessage failed', error);
        const errInfo = categorizeError(error);
        messageUtil.error(t(errInfo.i18nKey));
      }
    } finally {
      currentAction.value = '';
      loading.value = false;
      abortController.value = null;
    }
  }

  // rebuildSessionFiles now provided by useSessionFiles composable

  return {
    sendMessage,
    applyQuickAction,
    runAgentLoop,
    getOfficeSelection,
    currentAction,
    sessionStats,
    resetSessionStats,
    rebuildSessionFiles,
  };
}
