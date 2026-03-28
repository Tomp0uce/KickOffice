/**
 * useQuickActions.ts
 *
 * Manages Quick Action execution for all Office hosts.
 * Handles:
 * - Visual image generation (PowerPoint)
 * - Review slide feedback (PowerPoint)
 * - Smart reply / draft modes (Outlook, Excel)
 * - Text transformation actions (all hosts)
 *
 * Extracted from useAgentLoop.ts as part of ARCH-H1 refactoring.
 */

import { type Ref, type ComputedRef, nextTick } from 'vue';
import type { ModelTier, ModelInfo } from '@/types';
import type { UndoSnapshot } from '@/composables/useDocumentUndo';
import type {
  DisplayMessage,
  ExcelQuickAction,
  PowerPointQuickAction,
  OutlookQuickAction,
  QuickAction,
} from '@/types/chat';
import type { ChatMessage } from '@/api/backend';

import {
  chatStream,
  generateImage,
  categorizeError,
  uploadFile,
  uploadFileToPlatform,
} from '@/api/backend';
import { getDisplayLanguage } from '@/utils/common';
import {
  GLOBAL_STYLE_INSTRUCTIONS,
  getBuiltInPrompt,
  getExcelBuiltInPrompt,
  getOutlookBuiltInPrompt,
  getPowerPointBuiltInPrompt,
  outlookBuiltInPrompt,
  powerPointBuiltInPrompt,
  builtInPrompt,
  excelBuiltInPrompt,
} from '@/utils/constant';
import { message as messageUtil } from '@/utils/message';
import { powerpointToolDefinitions } from '@/utils/powerpointTools';
import {
  extractTextFromHtml,
  reassembleWithFragments,
  getPreservationInstruction,
  type RichContentContext,
} from '@/utils/richContentPreserver';
import { applyInheritedStyles, renderOfficeCommonApiHtml } from '@/utils/markdown';
import { areCredentialsConfigured } from '@/utils/credentialStorage';
import { logService } from '@/utils/logger';
import { getQuickActionSkill } from '@/skills';
import { writeFile as vfsWriteFile } from '@/utils/vfs';

// DRY-H1: Grouped sub-interfaces to reduce ambient coupling
export interface QuickActionRefs {
  history: Ref<DisplayMessage[]>;
  userInput: Ref<string>;
  loading: Ref<boolean>;
  imageLoading: Ref<boolean>;
  backendOnline: Ref<boolean>;
  abortController: Ref<AbortController | null>;
  inputTextarea: Ref<HTMLTextAreaElement | undefined>;
  isDraftFocusGlowing: Ref<boolean>;
}

export interface QuickActionModels {
  availableModels: Ref<Record<string, ModelInfo>>;
}

export interface QuickActionHosts {
  hostIsOutlook: boolean;
  hostIsPowerPoint: boolean;
  hostIsExcel: boolean;
}

export interface QuickActionHelpers {
  createDisplayMessage: (
    role: DisplayMessage['role'],
    content: string,
    imageSrc?: string,
  ) => DisplayMessage;
  adjustTextareaHeight: () => void;
  scrollToBottom: () => Promise<void>;
  scrollToMessageTop?: () => Promise<void>;
  getOfficeSelection: (opts?: {
    includeOutlookSelectedText?: boolean;
    actionKey?: string;
  }) => Promise<string>;
  getOfficeSelectionAsHtml: (opts?: {
    includeOutlookSelectedText?: boolean;
    actionKey?: string;
  }) => Promise<string>;
  runAgentLoop: (messages: ChatMessage[], modelTier: ModelTier) => Promise<void>;
  resolveChatModelTier: () => ModelTier;
}

export interface UseQuickActionsOptions {
  t: (key: string) => string;
  refs: QuickActionRefs;
  models: QuickActionModels;
  hosts: QuickActionHosts;
  quickActions: ComputedRef<QuickAction[] | undefined>;
  outlookQuickActions?: Ref<OutlookQuickAction[]>;
  excelQuickActions: Ref<ExcelQuickAction[]>;
  powerPointQuickActions: Ref<PowerPointQuickAction[]>;
  helpers: QuickActionHelpers;
  pendingMoM?: Ref<boolean>;
  captureBeforeInsert?: () => Promise<Partial<UndoSnapshot> | null>;
  saveSnapshot?: (partial: Partial<UndoSnapshot>, tag?: string) => void;
}

export function useQuickActions(options: UseQuickActionsOptions) {
  const {
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
  } = options;

  async function applyQuickAction(actionKey: string) {
    if (!backendOnline.value) return messageUtil.error(t('backendOffline'));

    // BUGFIX: Validate credentials are configured before sending request
    const hasCredentials = await areCredentialsConfigured();
    if (!hasCredentials) {
      messageUtil.error(t('credentialsRequired'));
      return;
    }

    // Prevent quick actions from running while another request is in progress
    if (loading.value || abortController.value) {
      return messageUtil.warning(
        t('requestInProgress') ||
          'A request is already in progress. Please wait or stop the current request.',
      );
    }

    const selectedQuickAction = hostIsExcel
      ? excelQuickActions.value.find((a: ExcelQuickAction) => a.key === actionKey)
      : hostIsPowerPoint
        ? powerPointQuickActions.value.find((a: PowerPointQuickAction) => a.key === actionKey)
        : hostIsOutlook && outlookQuickActions?.value
          ? outlookQuickActions.value.find((a: OutlookQuickAction) => a.key === actionKey)
          : quickActions.value?.find((a: QuickAction) => a.key === actionKey);

    const selectedExcelQuickAction = hostIsExcel
      ? (selectedQuickAction as ExcelQuickAction | undefined)
      : undefined;
    const selectedPowerPointQuickAction = hostIsPowerPoint
      ? (selectedQuickAction as PowerPointQuickAction | undefined)
      : undefined;
    const selectedOutlookQuickAction = hostIsOutlook
      ? (selectedQuickAction as OutlookQuickAction | undefined)
      : undefined;

    if (actionKey === 'visual' && hostIsPowerPoint) {
      const imageModelTier = (Object.entries(availableModels.value) as [string, ModelInfo][]).find(
        ([_, info]) => info.type === 'image',
      )?.[0] as ModelTier | undefined;
      if (!imageModelTier) {
        return messageUtil.error(t('imageError') || 'No image model configured.');
      }

      // PPT-IMG: Determine context source:
      // 1. If selected text has >= 10 words → illustrate the selection
      // 2. Otherwise → illustrate the full slide
      // In both cases we try to capture a screenshot so the vision LLM can match
      // the slide's visual style/colours/tone even when illustrating selected text.
      const MIN_SELECTION_WORDS = 10;
      let imageContext = '';
      let imageContextMode: 'selection' | 'slide' = 'slide';

      // Step A: Try to get selected text
      try {
        const selectedText = (await getOfficeSelection({ actionKey })) || '';
        const selWordCount = selectedText.trim().split(/\s+/).filter(Boolean).length;
        if (selWordCount >= MIN_SELECTION_WORDS) {
          imageContext = selectedText.trim();
          imageContextMode = 'selection';
        }
      } catch {
        /* no selection — fall through to slide screenshot */
      }

      // Step B: Always try to screenshot the current slide.
      // - selection mode: screenshot provides visual style context alongside the text
      // - slide mode: screenshot is the primary source, passed directly via vision
      let slideScreenshotUri: string | null = null;
      let currentSlideNum = 1;
      try {
        const sn = await powerpointToolDefinitions.getCurrentSlideIndex.execute({});
        currentSlideNum = parseInt(sn, 10) || 1;
      } catch {}
      try {
        const screenshotJson = await powerpointToolDefinitions.screenshotSlide.execute({
          slideNumber: currentSlideNum,
        });
        const screenshot = JSON.parse(screenshotJson);
        if (screenshot.base64 && !screenshot.error) {
          slideScreenshotUri = `data:image/png;base64,${screenshot.base64}`;
        }
      } catch (err) {
        logService.warn('[AgentLoop] PPT-IMG: slide screenshot failed', err);
      }

      // Last resort when slide mode and no screenshot: fall back to raw slide text
      if (imageContextMode === 'slide' && !slideScreenshotUri) {
        try {
          imageContext = await powerpointToolDefinitions.getSlideContent.execute({
            slideNumber: currentSlideNum,
          });
        } catch {}
        if (!imageContext) imageContext = await getOfficeSelection({ actionKey });
      }

      // Build user-facing label indicating which mode was used
      const lang = getDisplayLanguage();
      const modeLabel =
        imageContextMode === 'selection'
          ? lang === 'English'
            ? '📝 Illustrating selected text'
            : '📝 Illustration du texte sélectionné'
          : lang === 'English'
            ? '🖼️ Illustrating the full slide'
            : '🖼️ Illustration de la slide complète';

      // Step 1: call LLM to generate a proper image generation prompt
      // - If we have a screenshot: pass image directly (vision) — one LLM call, most accurate
      // - Otherwise: pass text content through the visual prompt template
      const visualPrompt = getPowerPointBuiltInPrompt().visual;
      const systemMsg = visualPrompt.system(lang);

      // Build user message: vision (screenshot) or text (selection/fallback)
      const visualRequirements = `Requirements for the image generation prompt:
- The image must visually represent the SPECIFIC topic, concept, or data from this slide — not a generic illustration.
- Choose the most appropriate visual style: photo-realistic scene, flat vector illustration, isometric diagram, infographic, conceptual metaphor, data visualization, etc.
- If the concept benefits from labels or short text in the image, explicitly request it.
- Describe composition: foreground, background, key focal elements.
- Specify color palette, mood, and lighting that match the slide's tone (professional, energetic, calm, technical).
- Wide landscape format (16:9), high resolution, suitable for professional presentation slides.
- No generic filler images (e.g., no random handshakes or abstract blobs unless directly relevant).

Constraints:
1. Respond in ${lang}.
2. OUTPUT ONLY the image prompt, ready to be sent directly to an image generation API. No explanation, no preamble.`;

      // Build vision user message when we have a screenshot.
      // - selection mode: include both the screenshot (style/tone) and the selected text (what to illustrate)
      // - slide mode: screenshot alone is the source
      const buildVisionUserMessage = () => {
        const imagePart = { type: 'image_url', image_url: { url: slideScreenshotUri! } };
        if (imageContextMode === 'selection') {
          return {
            role: 'user' as const,
            content: [
              imagePart,
              {
                type: 'text',
                text: `Task: Write a detailed prompt for an image generation model to illustrate the selected text below.

PRIORITY RULES (strictly follow this weighting):
1. [HIGHEST WEIGHT] The image MUST illustrate the selected text — this is the sole subject of the image. Every visual element must directly serve this text.
2. [MEDIUM WEIGHT] Borrow the slide's visual style from the screenshot: colour palette, graphic style (flat, realistic, technical, etc.), mood and tone.
3. [IGNORE] All other text visible on the slide that is NOT the selected text must be completely ignored — do not let it influence the subject or composition of the image.

Selected text to illustrate:
"${imageContext}"

${visualRequirements}`,
              },
            ],
          };
        }
        return {
          role: 'user' as const,
          content: [
            imagePart,
            {
              type: 'text',
              text: `Task: Based on this presentation slide image, write a detailed prompt for an image generation model that will produce a visual directly illustrating this slide's content.\n\n${visualRequirements}`,
            },
          ],
        };
      };

      const promptMessages: ChatMessage[] = slideScreenshotUri
        ? [{ role: 'system', content: systemMsg }, buildVisionUserMessage()]
        : [
            { role: 'system', content: systemMsg },
            { role: 'user', content: visualPrompt.user(imageContext || '', lang) },
          ];

      const actionLabel = selectedQuickAction?.label || t(actionKey);
      history.value.push(
        createDisplayMessage(
          'user',
          `[${actionLabel}] ${modeLabel}\n${(imageContext || '').substring(0, 100)}...`,
        ),
      );
      history.value.push(createDisplayMessage('assistant', t('imageGenerating')));
      await scrollToMessageTop?.();

      loading.value = true;
      abortController.value = new AbortController();
      try {
        let imagePrompt = '';
        await chatStream({
          messages: promptMessages,
          modelTier: resolveChatModelTier(),
          onStream: async (text: string) => {
            imagePrompt = text;
          },
          abortSignal: abortController.value?.signal,
        });

        if (!imagePrompt.trim()) {
          history.value[history.value.length - 1].content = t('somethingWentWrong');
          return;
        }

        // Step 2: use the generated description to produce the image
        history.value[history.value.length - 1].content = t('imageGenerating');
        imageLoading.value = true;
        const imageSrc = await generateImage({
          prompt: imagePrompt.trim(),
          abortSignal: abortController.value?.signal,
        });
        const message = history.value[history.value.length - 1];
        message.role = 'assistant';
        message.content = '';
        message.imageSrc = imageSrc;
        await scrollToBottom();
      } catch (err: unknown) {
        if (!(err instanceof Error) || err.name !== 'AbortError') {
          logService.error('[AgentLoop] visual quick action failed', err);
          const errInfo = categorizeError(err);
          history.value[history.value.length - 1].content = t(errInfo.i18nKey);
        }
      } finally {
        imageLoading.value = false;
        loading.value = false;
        abortController.value = null;
      }
      return;
    }

    // PPT-H2: "review" — screenshots current slide + gathers overview, then runs agent loop
    // Does NOT require selected text (bypasses the selectedText guard below)
    if (actionKey === 'review' && hostIsPowerPoint) {
      const lang = getDisplayLanguage();
      const actionLabel = selectedQuickAction?.label || t(actionKey);
      history.value.push(createDisplayMessage('user', `[${actionLabel}]`));

      loading.value = true;
      abortController.value = new AbortController();
      try {
        const systemMsg = `You are an expert presentation coach reviewing a PowerPoint presentation. Respond in ${lang}.
Instructions:
1. Call \`getCurrentSlideIndex\` to find the current slide number.
2. Call \`screenshotSlide\` with that slide number to see the visual layout.
3. Call \`getAllSlidesOverview\` to understand the full presentation context.
4. Based on the screenshot and the presentation overview, provide 3-5 specific, actionable improvement suggestions for THIS slide only.
Review areas: content clarity, visual balance (too much/too little text), message impact, consistency with the rest of the presentation.
Format your response as numbered suggestions. Be concrete and direct. Do NOT suggest changes to other slides.`;
        await runAgentLoop(
          [
            { role: 'system', content: systemMsg },
            {
              role: 'user',
              content: 'Review the current slide and provide improvement suggestions.',
            },
          ],
          resolveChatModelTier(),
        );
      } catch (err: unknown) {
        if (!(err instanceof Error) || err.name !== 'AbortError') {
          logService.error('[AgentLoop] review quick action failed', err);
          const errInfo = categorizeError(err);
          const last = history.value[history.value.length - 1];
          if (last?.role === 'assistant') last.content = t(errInfo.i18nKey);
        }
      } finally {
        loading.value = false;
        abortController.value = null;
      }
      return;
    }

    if (selectedOutlookQuickAction?.mode === 'smart-reply') {
      userInput.value = selectedOutlookQuickAction.prefix || '';
      adjustTextareaHeight();
      isDraftFocusGlowing.value = true;
      setTimeout(() => {
        isDraftFocusGlowing.value = false;
      }, 1500);
      await nextTick();
      const el = inputTextarea.value;
      if (el) {
        el.focus();
        const len = userInput.value.length;
        el.setSelectionRange(len, len);
      }
      return;
    }

    if (selectedOutlookQuickAction?.mode === 'mom') {
      if (pendingMoM) pendingMoM.value = true;
      userInput.value = selectedOutlookQuickAction.prefix || '';
      adjustTextareaHeight();
      isDraftFocusGlowing.value = true;
      setTimeout(() => {
        isDraftFocusGlowing.value = false;
      }, 1500);
      await nextTick();
      const el = inputTextarea.value;
      if (el) {
        el.focus();
        const len = userInput.value.length;
        el.setSelectionRange(len, len);
      }
      return;
    }

    if (selectedOutlookQuickAction?.mode === 'draft') {
      userInput.value = selectedOutlookQuickAction.prefix || '';
      adjustTextareaHeight();
      isDraftFocusGlowing.value = true;
      setTimeout(() => {
        isDraftFocusGlowing.value = false;
      }, 1500);
      await nextTick();
      const el = inputTextarea.value;
      if (el) {
        el.focus();
        const len = userInput.value.length;
        el.setSelectionRange(len, len);
      }
      return;
    }

    // IMAGE-UPLOAD actions: open file picker, upload image, then run agent with chart digitizer skill
    if (hostIsExcel && selectedExcelQuickAction?.imageUpload) {
      const lang = getDisplayLanguage();
      return new Promise<void>(resolve => {
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = 'image/png,image/jpeg,image/jpg,image/gif,image/webp';

        fileInput.onchange = async () => {
          const file = fileInput.files?.[0];
          if (!file) {
            resolve();
            return;
          }

          const actionLabel = selectedExcelQuickAction.label || t(actionKey);
          history.value.push(createDisplayMessage('user', `[${actionLabel}] ${file.name}`));

          if (loading.value) {
            resolve();
            return;
          }
          loading.value = true;
          abortController.value = new AbortController();

          try {
            // Upload image to backend to get base64 + imageId
            const result = await uploadFile(file);
            if (!result.imageBase64) {
              messageUtil.error(t('errorImageUploadFailed') || 'Image upload failed.');
              return;
            }

            // Optionally upload to provider for caching
            let fileId: string | undefined;
            try {
              const platformResult = await uploadFileToPlatform(file, 'vision');
              if (platformResult.fileId) fileId = platformResult.fileId;
            } catch {
              /* inline fallback */
            }

            // Write image to VFS so tools like imageToSheet can access it via filePath
            if (result.imageBase64) {
              try {
                const base64Data = result.imageBase64.replace(/^data:[^;]+;base64,/, '');
                const binaryStr = atob(base64Data);
                const bytes = new Uint8Array(binaryStr.length);
                for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
                await vfsWriteFile(`/home/user/uploads/${result.filename}`, bytes);
              } catch (vfsErr) {
                logService.warn('[QuickActions] VFS write failed for image', vfsErr);
              }
            }

            const imageIdTag = result.imageId ? ` (imageId: ${result.imageId})` : '';
            const vfsPath = `/home/user/uploads/${result.filename}`;
            const imageContext =
              `<uploaded_images>\nThe following images are available in session memory:\n` +
              `- [${result.filename}]${imageIdTag}\n` +
              (actionKey === 'pixelArt'
                ? `The image has been saved to VFS at: ${vfsPath}\nUse the imageToSheet tool with filePath="${vfsPath}" to generate pixel art in Excel.\n`
                : `To extract chart data into Excel, use extract_chart_data with the imageId.\n`) +
              `</uploaded_images>`;

            const userText =
              `[UI language: ${lang}]\n\n` +
              (actionKey === 'pixelArt'
                ? `Convert this image into pixel art in Excel using the imageToSheet tool.\n\n`
                : `Digitize this chart image: extract the data and write it to Excel, then create a matching chart.\n\n`) +
              imageContext;

            // Build multipart user message (text + image)
            // NOTE: result.imageBase64 is already a full data URI (data:mime;base64,...)
            // returned by the backend — do NOT wrap it again with data:${file.type};base64,...
            const imageUrl = fileId ?? result.imageBase64;
            const userMessage: ChatMessage = {
              role: 'user',
              content: [
                { type: 'text', text: userText },
                { type: 'image_url', image_url: { url: imageUrl } },
              ],
            };

            const skillContent = getQuickActionSkill(actionKey);
            const systemMsg = skillContent
              ? skillContent + `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`
              : `You are an expert chart data extractor. ${GLOBAL_STYLE_INSTRUCTIONS}`;

            // Capture undo state before the agent modifies the document
            let undoSnapshot: Partial<UndoSnapshot> | null = null;
            try {
              if (captureBeforeInsert) undoSnapshot = await captureBeforeInsert();
            } catch {
              /* best-effort */
            }

            await runAgentLoop(
              [{ role: 'system', content: systemMsg }, userMessage],
              resolveChatModelTier(),
            );

            // Mark undo as available now that the agent has finished modifying the document
            if (undoSnapshot && saveSnapshot) {
              try {
                saveSnapshot(undoSnapshot);
              } catch {
                /* best-effort */
              }
            }
          } catch (err) {
            logService.error('[QuickActions] digitizeChart error', err);
            messageUtil.error(t('errorGeneric') || 'An error occurred.');
          } finally {
            loading.value = false;
            abortController.value = null;
            resolve();
          }
        };

        fileInput.oncancel = () => resolve();
        fileInput.click();
      });
    }

    if (selectedExcelQuickAction?.mode === 'draft') {
      userInput.value = selectedExcelQuickAction.prefix || '';
      adjustTextareaHeight();
      isDraftFocusGlowing.value = true;
      setTimeout(() => {
        isDraftFocusGlowing.value = false;
      }, 1000);
      await nextTick();
      const el = inputTextarea.value;
      if (el) {
        el.focus();
        const len = userInput.value.length;
        el.setSelectionRange(len, len);
      }
      return;
    }

    if (loading.value) return;
    loading.value = true;
    abortController.value = new AbortController();

    try {
      const selectedText = await getOfficeSelection({
        includeOutlookSelectedText: true,
        actionKey,
      });
      if (!selectedText) {
        messageUtil.error(
          t(
            hostIsOutlook
              ? 'selectEmailPrompt'
              : hostIsPowerPoint
                ? 'selectSlideTextPrompt'
                : hostIsExcel
                  ? 'selectCellsPrompt'
                  : 'selectTextPrompt',
          ),
        );
        return;
      }

      // F1: Try to get HTML selection for rich content preservation (Word, Outlook)
      let richContext: RichContentContext | null = null;
      const isTextModifyingAction = !selectedQuickAction?.executeWithAgent && !hostIsExcel;
      if (isTextModifyingAction) {
        try {
          const htmlContent = await getOfficeSelectionAsHtml({
            includeOutlookSelectedText: true,
            actionKey,
          });
          if (htmlContent) {
            richContext = extractTextFromHtml(htmlContent);
          }
        } catch (err) {
          logService.warn(
            '[AgentLoop] Failed to get HTML selection for rich content preservation',
            err,
          );
        }
      }

      // Use Markdown text if HTML was parsed successfully, otherwise fallback to plain text selection.
      // Also fall back when cleanText is empty (e.g. HTML coercion unavailable in some Outlook modes).
      const rawTextForLlm = richContext?.cleanText || selectedText;
      const textForLlm =
        '\n<document_content>\n' +
        rawTextForLlm.replace(new RegExp('</?document_content>', 'g'), '') +
        '\n<' +
        '/document_content>\n';

      let action:
        | { system: (lang: string) => string; user: (text: string, lang: string) => string }
        | undefined;
      let systemMsg = '';
      let userMsg = '';

      // Resolve UI language once — used for skill injection and fallback prompts
      const lang = getDisplayLanguage();

      // SKILL-L1: Try to load skill file first (priority 1)
      const skillContent = getQuickActionSkill(actionKey);
      if (skillContent) {
        systemMsg = skillContent;
        // Inject UI language so skills can respond in the correct language.
        // Translation DIRECTION is determined by the skill from the text content itself.
        userMsg = `[UI language: ${lang}]\n\n${textForLlm}`;
      } else {
        // Priority 2: systemPrompt from Quick Action definition
        if (hostIsOutlook) {
          action =
            getOutlookBuiltInPrompt()[actionKey as keyof typeof outlookBuiltInPrompt] ||
            getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt];
        } else if (hostIsPowerPoint) {
          if (selectedPowerPointQuickAction?.systemPrompt) {
            systemMsg = selectedPowerPointQuickAction.systemPrompt;
            userMsg = selectedText || t('applyToCurrentSlide') || 'Apply to the current slide.';
          } else {
            action =
              getPowerPointBuiltInPrompt()[actionKey as keyof typeof powerPointBuiltInPrompt] ||
              getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt];
          }
        } else if (hostIsExcel) {
          if (
            selectedExcelQuickAction?.mode === 'immediate' &&
            selectedExcelQuickAction.systemPrompt
          ) {
            systemMsg = selectedExcelQuickAction.systemPrompt;
            userMsg = `Selection:\n${selectedText}`;
          } else action = getExcelBuiltInPrompt()[actionKey as keyof typeof excelBuiltInPrompt];
        } else action = getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt];

        // Priority 3: Fallback to constant.ts prompts
        if (!systemMsg || !userMsg) {
          if (!action) action = getBuiltInPrompt()[actionKey as keyof typeof builtInPrompt];
          if (!action) return;
          systemMsg = action.system(lang);
          userMsg = action.user(textForLlm, lang);
        }
      }

      // LANG-UNIVERSAL: Inject UI language tag into ALL quick-action user messages.
      // Skill-based actions already have it (from SKILL-L1 above); this covers
      // built-in prompts and hardcoded systemPrompt paths so every host (Word,
      // Excel, PowerPoint, Outlook) respects the interface language.
      if (userMsg && !userMsg.startsWith('[UI language:')) {
        userMsg = `[UI language: ${lang}]\n\n${userMsg}`;
      }

      // Enforce global formatting constraints on all Quick Actions
      systemMsg += `\n\n${GLOBAL_STYLE_INSTRUCTIONS}`;

      // F1: Add preservation instruction if rich content was detected
      if (richContext?.hasRichContent) {
        systemMsg += getPreservationInstruction(richContext);
      }

      const actionLabel = selectedQuickAction?.label || t(actionKey);
      history.value.push(
        createDisplayMessage('user', `[${actionLabel}] ${selectedText.substring(0, 100)}...`),
      );

      if (selectedQuickAction?.executeWithAgent) {
        // Capture undo state before the agent modifies the document
        let undoSnapshotQA: Partial<UndoSnapshot> | null = null;
        try {
          if (captureBeforeInsert) undoSnapshotQA = await captureBeforeInsert();
        } catch {
          /* best-effort */
        }

        await runAgentLoop(
          [
            { role: 'system', content: systemMsg },
            { role: 'user', content: userMsg },
          ],
          resolveChatModelTier(),
        );

        // Mark undo as available now that the agent has finished
        if (undoSnapshotQA && saveSnapshot) {
          try {
            saveSnapshot(undoSnapshotQA);
          } catch {
            /* best-effort */
          }
        }
      } else {
        history.value.push(createDisplayMessage('assistant', ''));
        await scrollToMessageTop?.(); // Scroll to show start of assistant response
        try {
          await chatStream({
            messages: [
              { role: 'system', content: systemMsg },
              { role: 'user', content: userMsg },
            ],
            modelTier: resolveChatModelTier(),
            onStream: async (text: string) => {
              const message = history.value[history.value.length - 1];
              message.role = 'assistant';
              message.content = text;
              await scrollToBottom();
            },
            abortSignal: abortController.value?.signal,
          });
          // Check for empty response
          const lastMessage = history.value[history.value.length - 1];
          if (!lastMessage?.content?.trim()) {
            lastMessage.content = t('noModelResponse');
          }

          // F1: Reassemble rich content with preserved fragments and inject native styles
          if (lastMessage?.content) {
            let finalHtml = '';
            if (richContext?.hasRichContent) {
              finalHtml = reassembleWithFragments(lastMessage.content, richContext);
            }
            if (richContext?.extractedStyles) {
              if (!finalHtml) finalHtml = renderOfficeCommonApiHtml(lastMessage.content);
              finalHtml = applyInheritedStyles(finalHtml, richContext.extractedStyles);
            }
            if (finalHtml) {
              lastMessage.richHtml = finalHtml;
            }
          }
        } catch (err: unknown) {
          if (err instanceof Error && err.name === 'AbortError') return;
          logService.error('[AgentLoop] Quick action chatStream failed', err);
          const lastMessage = history.value[history.value.length - 1];
          const errInfo = categorizeError(err);
          if (errInfo.type === 'auth') {
            lastMessage.content = `⚠️ ${t('credentialsRequiredTitle')}\n\n${t('credentialsRequired')}`;
            messageUtil.warning(t('credentialsRequired'));
          } else {
            lastMessage.content = t(errInfo.i18nKey);
            messageUtil.error(t(errInfo.i18nKey));
          }
        }
      }
    } finally {
      loading.value = false;
      abortController.value = null;
    }
  }

  return {
    applyQuickAction,
  };
}
