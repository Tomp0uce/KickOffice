/**
 * useMessageOrchestration.ts
 *
 * Orchestrates message construction for the LLM:
 * - Builds chat messages from history
 * - Injects document context (Excel, PowerPoint, Word, Outlook)
 * - Injects uploaded files (text files, platform file IDs)
 * - Injects rich content preservation instructions
 *
 * Extracted from useAgentLoop.ts as part of ARCH-H1 refactoring.
 */

import { type Ref } from 'vue';
import type { DisplayMessage } from '@/types/chat';
import type { ChatMessage, ChatRequestMessage } from '@/api/backend';
import {
  getExcelDocumentContext,
  getPowerPointDocumentContext,
  getOutlookDocumentContext,
  getWordDocumentContext,
} from '@/utils/officeDocumentContext';
import { getPreservationInstruction } from '@/utils/richContentPreserver';
import { getLastRichContext } from '@/utils/richContextStore';
import { logService } from '@/utils/logger';
import type { SessionFile } from './useSessionFiles';

export interface UseMessageOrchestrationOptions {
  history: Ref<DisplayMessage[]>;
  hostIsOutlook: boolean;
  hostIsPowerPoint: boolean;
  hostIsExcel: boolean;
  hostIsWord: boolean;
}

export function useMessageOrchestration(options: UseMessageOrchestrationOptions) {
  const { history, hostIsOutlook, hostIsPowerPoint, hostIsExcel, hostIsWord } = options;

  /**
   * Build base chat messages from history.
   * Converts DisplayMessage[] to ChatMessage[].
   */
  function buildChatMessages(systemPrompt: string): ChatMessage[] {
    const msgs: ChatRequestMessage[] = [{ role: 'system', content: systemPrompt }];
    for (const m of history.value) {
      let contentToKeep = m.content;
      // If the assistant message only had tool calls and no content, ensure it's not totally empty
      if (
        m.role === 'assistant' &&
        !contentToKeep?.trim() &&
        m.rawMessages &&
        m.rawMessages.length > 0
      ) {
        contentToKeep = `[Tools executed internally]`;
      }
      msgs.push({ role: m.role, content: contentToKeep || '' });
    }
    return msgs as ChatMessage[];
  }

  /**
   * Inject document context into messages.
   * Fetches document metadata (sheets, slides, email info) and appends to last user message.
   *
   * @param messages - Messages array to modify (modified in-place)
   * @returns The modified messages array
   */
  async function injectDocumentContext(messages: ChatMessage[]): Promise<ChatMessage[]> {
    try {
      let docContextJson = '';
      if (hostIsExcel) docContextJson = await getExcelDocumentContext();
      else if (hostIsPowerPoint) docContextJson = await getPowerPointDocumentContext();
      else if (hostIsOutlook) docContextJson = await getOutlookDocumentContext();
      else if (hostIsWord) docContextJson = await getWordDocumentContext();

      if (docContextJson) {
        const lastUserIdx = messages.map(m => m.role).lastIndexOf('user');
        if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
          messages[lastUserIdx].content += `\n\n<doc_context>\n${docContextJson}\n</doc_context>`;
        }
      }
    } catch (err) {
      // Document context is optional — continue without it if it fails
      logService.warn('[MessageOrchestration] Failed to inject document context', err);
    }

    return messages;
  }

  /**
   * Inject uploaded files into messages.
   * Supports both platform file IDs (Claude API /v1/files) and inline content fallback.
   *
   * @param messages - Messages array to modify (modified in-place)
   * @param uploadedFiles - Files to inject
   * @param injectedContext - Legacy string-based file content (deprecated, kept for compatibility)
   * @returns The modified messages array
   */
  function injectUploadedFiles(
    messages: ChatMessage[],
    uploadedFiles?: SessionFile[],
    injectedContext?: string,
  ): ChatMessage[] {
    // Legacy string-based injection (kept for any call sites that still pass it)
    if (injectedContext) {
      const lastUserIdx = messages.map(m => m.role).lastIndexOf('user');
      if (lastUserIdx !== -1 && typeof messages[lastUserIdx].content === 'string') {
        messages[lastUserIdx].content +=
          `\n\n<attached_files>\n${injectedContext}\n</attached_files>`;
      }
    }

    // Modern file injection with platform file IDs
    if (uploadedFiles && uploadedFiles.length > 0) {
      const lastUserIdx = messages.map(m => m.role).lastIndexOf('user');
      if (lastUserIdx !== -1) {
        const hasFileRefs = uploadedFiles.some(f => f.fileId);
        if (hasFileRefs && typeof messages[lastUserIdx].content === 'string') {
          // Convert to content array: text + file references + inline fallback for files without fileId
          const parts: any[] = [{ type: 'text', text: messages[lastUserIdx].content as string }];
          for (const f of uploadedFiles) {
            if (f.fileId) {
              parts.push({ type: 'file', file: { file_id: f.fileId } });
            } else {
              parts.push({
                type: 'text',
                text: `\n\n[Contenu du fichier "${f.filename}"]:\n${f.content}\n[Fin du fichier]`,
              });
            }
          }
          messages[lastUserIdx].content = parts;
        } else if (typeof messages[lastUserIdx].content === 'string') {
          // All inline fallback
          const inlineText = uploadedFiles
            .map(f => `\n\n[Contenu du fichier "${f.filename}"]:\n${f.content}\n[Fin du fichier]`)
            .join('');
          messages[lastUserIdx].content += `\n\n<attached_files>${inlineText}\n</attached_files>`;
        }
      }
    }

    return messages;
  }

  /**
   * Inject rich content preservation instructions into system message.
   * Used for Word/Outlook to preserve embedded images and formatting.
   *
   * @param messages - Messages array to modify (modified in-place)
   * @returns The modified messages array
   */
  function injectRichContentInstructions(messages: ChatMessage[]): ChatMessage[] {
    const richContext = getLastRichContext();
    if (richContext?.hasRichContent && messages[0]?.role === 'system') {
      messages[0].content += getPreservationInstruction(richContext);
    }
    return messages;
  }

  /**
   * Build and prepare messages for sending to LLM.
   * Combines all injection steps:
   * 1. Build base messages from history
   * 2. Inject document context
   * 3. Inject uploaded files
   * 4. Inject rich content preservation instructions
   *
   * @param systemPrompt - System prompt to use
   * @param uploadedFiles - Optional uploaded files to include
   * @param injectedContext - Optional legacy file content
   * @returns Messages ready for LLM
   */
  async function prepareMessages(
    systemPrompt: string,
    uploadedFiles?: SessionFile[],
    injectedContext?: string,
  ): Promise<ChatMessage[]> {
    let messages = buildChatMessages(systemPrompt);
    messages = await injectDocumentContext(messages);
    messages = injectUploadedFiles(messages, uploadedFiles, injectedContext);
    messages = injectRichContentInstructions(messages);
    return messages;
  }

  return {
    buildChatMessages,
    injectDocumentContext,
    injectUploadedFiles,
    injectRichContentInstructions,
    prepareMessages,
  };
}
