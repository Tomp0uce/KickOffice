/**
 * useUserSkillExecution.ts
 *
 * Executes user-created skills with the same pipeline as built-in quick actions.
 * - immediate → chatStream (text response in chat)
 * - agent     → runAgentLoop (full Office tool access)
 * - draft     → pre-fills the textarea + focus glow
 */

import { nextTick } from 'vue';
import type { Ref } from 'vue';
import type { UserSkill } from '@/types/userSkill';
import type { ChatMessage } from '@/api/backend';
import { chatStream, categorizeError } from '@/api/backend';
import { GLOBAL_STYLE_INSTRUCTIONS } from '@/utils/constant';
import { message as messageUtil } from '@/utils/message';
import { logService } from '@/utils/logger';
import type { ModelTier } from '@/types';
import type { DisplayMessage } from '@/types/chat';

export interface UseUserSkillExecutionOptions {
  t: (key: string) => string;
  history: Ref<DisplayMessage[]>;
  userInput: Ref<string>;
  loading: Ref<boolean>;
  abortController: Ref<AbortController | null>;
  inputTextarea: Ref<HTMLTextAreaElement | undefined>;
  isDraftFocusGlowing: Ref<boolean>;
  getOfficeSelection: (opts?: Record<string, unknown>) => Promise<string>;
  runAgentLoop: (messages: ChatMessage[], modelTier: ModelTier) => Promise<void>;
  resolveChatModelTier: () => ModelTier;
  createDisplayMessage: (role: DisplayMessage['role'], content: string) => DisplayMessage;
  adjustTextareaHeight: () => void;
  scrollToBottom: () => Promise<void>;
  scrollToMessageTop?: () => Promise<void>;
}

export function useUserSkillExecution(options: UseUserSkillExecutionOptions) {
  const {
    t,
    history,
    userInput,
    loading,
    abortController,
    inputTextarea,
    isDraftFocusGlowing,
    getOfficeSelection,
    runAgentLoop,
    resolveChatModelTier,
    createDisplayMessage,
    adjustTextareaHeight,
    scrollToBottom,
    scrollToMessageTop,
  } = options;

  async function executeUserSkill(skill: UserSkill): Promise<void> {
    if (loading.value || abortController.value) {
      messageUtil.warning(
        t('requestInProgress') || 'A request is already in progress. Please wait.',
      );
      return;
    }

    const lang = localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français';

    // ── Draft mode: pre-fill textarea and focus ──────────────────────────────
    if (skill.executionMode === 'draft') {
      userInput.value = '';
      adjustTextareaHeight();
      isDraftFocusGlowing.value = true;
      setTimeout(() => {
        isDraftFocusGlowing.value = false;
      }, 1500);
      await nextTick();
      const el = inputTextarea.value;
      if (el) {
        el.focus();
        el.setSelectionRange(0, 0);
      }
      return;
    }

    // ── Get selected text (required for immediate; optional for agent) ───────
    let selectedText = '';
    try {
      selectedText = await getOfficeSelection();
    } catch {
      // Agent skills can proceed without selection — they call their own tools
    }

    if (!selectedText && skill.executionMode !== 'agent') {
      messageUtil.error(t('selectTextPrompt') || 'Please select some text first.');
      return;
    }

    // ── Build messages ────────────────────────────────────────────────────────
    const systemMsg = `${skill.skillContent}\n\n${GLOBAL_STYLE_INSTRUCTIONS}`;

    const userMsg = selectedText
      ? `[UI language: ${lang}]\n\n<document_content>\n${selectedText}\n</document_content>`
      : `[UI language: ${lang}]`;

    const label = skill.name;
    const preview = selectedText ? `${selectedText.substring(0, 100)}${selectedText.length > 100 ? '…' : ''}` : '';
    history.value.push(createDisplayMessage('user', `[${label}]${preview ? ` ${preview}` : ''}`));

    const messages: ChatMessage[] = [
      { role: 'system', content: systemMsg },
      { role: 'user', content: userMsg },
    ];

    // ── Agent mode ────────────────────────────────────────────────────────────
    if (skill.executionMode === 'agent') {
      loading.value = true;
      abortController.value = new AbortController();
      try {
        await runAgentLoop(messages, resolveChatModelTier());
      } catch (err: unknown) {
        if (!(err instanceof Error) || err.name !== 'AbortError') {
          logService.error('[UserSkillExecution] agent skill failed', err);
          const last = history.value[history.value.length - 1];
          if (last?.role === 'assistant') {
            const errInfo = categorizeError(err);
            last.content = t(errInfo.i18nKey);
          }
        }
      } finally {
        loading.value = false;
        abortController.value = null;
      }
      return;
    }

    // ── Immediate mode: chatStream ────────────────────────────────────────────
    history.value.push(createDisplayMessage('assistant', ''));
    await scrollToMessageTop?.();

    loading.value = true;
    abortController.value = new AbortController();
    try {
      await chatStream({
        messages,
        modelTier: resolveChatModelTier(),
        onStream: async (text: string) => {
          const last = history.value[history.value.length - 1];
          last.role = 'assistant';
          last.content = text;
          await scrollToBottom();
        },
        abortSignal: abortController.value.signal,
      });
      const last = history.value[history.value.length - 1];
      if (!last?.content?.trim()) {
        last.content = t('noModelResponse');
      }
    } catch (err: unknown) {
      if (!(err instanceof Error) || err.name !== 'AbortError') {
        logService.error('[UserSkillExecution] chatStream failed', err);
        const last = history.value[history.value.length - 1];
        if (last) {
          const errInfo = categorizeError(err);
          last.content = t(errInfo.i18nKey);
        }
      }
    } finally {
      loading.value = false;
      abortController.value = null;
    }
  }

  return { executeUserSkill };
}
