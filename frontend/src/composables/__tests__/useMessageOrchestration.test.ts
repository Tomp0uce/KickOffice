import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ref } from 'vue';

// ── Mocks must be declared before the module import ───────────────────────────

vi.mock('@/utils/officeDocumentContext', () => ({
  getExcelDocumentContext: vi.fn().mockResolvedValue(''),
  getPowerPointDocumentContext: vi.fn().mockResolvedValue(''),
  getOutlookDocumentContext: vi.fn().mockResolvedValue(''),
  getWordDocumentContext: vi.fn().mockResolvedValue(''),
}));

vi.mock('@/utils/richContentPreserver', () => ({
  getPreservationInstruction: vi.fn().mockReturnValue(' [preserve-rich-content]'),
}));

vi.mock('@/utils/richContextStore', () => ({
  getLastRichContext: vi.fn().mockReturnValue(null),
}));

vi.mock('@/utils/logger', () => ({
  logService: { warn: vi.fn(), error: vi.fn(), info: vi.fn() },
}));

vi.mock('@/api/backend', () => ({}));

import { useMessageOrchestration } from '../useMessageOrchestration';
import type { DisplayMessage } from '@/types/chat';
import {
  getExcelDocumentContext,
  getPowerPointDocumentContext,
  getOutlookDocumentContext,
  getWordDocumentContext,
} from '@/utils/officeDocumentContext';
import { getLastRichContext } from '@/utils/richContextStore';
import { getPreservationInstruction } from '@/utils/richContentPreserver';

function makeOrchestration(
  messages: Partial<DisplayMessage>[] = [],
  host: { excel?: boolean; powerpoint?: boolean; outlook?: boolean; word?: boolean } = {},
) {
  return useMessageOrchestration({
    history: ref(messages as DisplayMessage[]),
    hostIsExcel: host.excel ?? false,
    hostIsPowerPoint: host.powerpoint ?? false,
    hostIsOutlook: host.outlook ?? false,
    hostIsWord: host.word ?? false,
  });
}

describe('useMessageOrchestration', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getExcelDocumentContext).mockResolvedValue('');
    vi.mocked(getPowerPointDocumentContext).mockResolvedValue('');
    vi.mocked(getOutlookDocumentContext).mockResolvedValue('');
    vi.mocked(getWordDocumentContext).mockResolvedValue('');
    vi.mocked(getLastRichContext).mockReturnValue(null);
    vi.mocked(getPreservationInstruction).mockReturnValue(' [preserve-rich-content]');
  });

  // ─── buildChatMessages ────────────────────────────────────────────────────

  describe('buildChatMessages', () => {
    it('prepends a system message', () => {
      const { buildChatMessages } = makeOrchestration();
      const msgs = buildChatMessages('You are helpful.');
      expect(msgs[0]).toEqual({ role: 'system', content: 'You are helpful.' });
    });

    it('maps user and assistant messages from history', () => {
      const { buildChatMessages } = makeOrchestration([
        { id: '1', role: 'user', content: 'Hello' },
        { id: '2', role: 'assistant', content: 'Hi there' },
      ]);
      const msgs = buildChatMessages('sys');
      expect(msgs[1]).toEqual({ role: 'user', content: 'Hello' });
      expect(msgs[2]).toEqual({ role: 'assistant', content: 'Hi there' });
    });

    it('replaces empty assistant content with "[Tools executed internally]" when rawMessages exist', () => {
      const { buildChatMessages } = makeOrchestration([
        { id: '1', role: 'assistant', content: '', rawMessages: [{ type: 'tool_use' }] },
      ]);
      const msgs = buildChatMessages('sys');
      expect(msgs[1].content).toBe('[Tools executed internally]');
    });

    it('keeps non-empty assistant content as-is even when rawMessages exist', () => {
      const { buildChatMessages } = makeOrchestration([
        { id: '1', role: 'assistant', content: 'real text', rawMessages: [{}] },
      ]);
      const msgs = buildChatMessages('sys');
      expect(msgs[1].content).toBe('real text');
    });

    it('uses empty string for missing content', () => {
      const { buildChatMessages } = makeOrchestration([
        { id: '1', role: 'user', content: undefined as unknown as string },
      ]);
      const msgs = buildChatMessages('sys');
      expect(msgs[1].content).toBe('');
    });
  });

  // ─── injectDocumentContext ────────────────────────────────────────────────

  describe('injectDocumentContext', () => {
    it('calls getExcelDocumentContext when hostIsExcel', async () => {
      vi.mocked(getExcelDocumentContext).mockResolvedValue('{"sheets":[]}');
      const { buildChatMessages, injectDocumentContext } = makeOrchestration(
        [{ id: '1', role: 'user', content: 'analyze' }],
        { excel: true },
      );
      const msgs = buildChatMessages('sys');
      await injectDocumentContext(msgs);
      expect(getExcelDocumentContext).toHaveBeenCalledOnce();
      expect(msgs[1].content).toContain('<doc_context>');
    });

    it('does not call any document context function when no host flag is set', async () => {
      const { buildChatMessages, injectDocumentContext } = makeOrchestration([
        { id: '1', role: 'user', content: 'hello' },
      ]);
      const msgs = buildChatMessages('sys');
      await injectDocumentContext(msgs);
      expect(getExcelDocumentContext).not.toHaveBeenCalled();
      expect(getPowerPointDocumentContext).not.toHaveBeenCalled();
      expect(getOutlookDocumentContext).not.toHaveBeenCalled();
      expect(getWordDocumentContext).not.toHaveBeenCalled();
    });

    it('skips injection when docContext is empty string', async () => {
      vi.mocked(getExcelDocumentContext).mockResolvedValue('');
      const { buildChatMessages, injectDocumentContext } = makeOrchestration(
        [{ id: '1', role: 'user', content: 'hello' }],
        { excel: true },
      );
      const msgs = buildChatMessages('sys');
      const originalContent = msgs[1].content;
      await injectDocumentContext(msgs);
      expect(msgs[1].content).toBe(originalContent);
    });

    it('continues silently when doc context throws', async () => {
      vi.mocked(getWordDocumentContext).mockRejectedValue(new Error('Office not ready'));
      const { buildChatMessages, injectDocumentContext } = makeOrchestration(
        [{ id: '1', role: 'user', content: 'hello' }],
        { word: true },
      );
      const msgs = buildChatMessages('sys');
      await expect(injectDocumentContext(msgs)).resolves.toBeDefined();
    });
  });

  // ─── injectUploadedFiles ──────────────────────────────────────────────────

  describe('injectUploadedFiles', () => {
    it('returns messages unchanged when no files or context provided', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'hello' },
      ]);
      const msgs = buildChatMessages('sys');
      const original = msgs[1].content;
      injectUploadedFiles(msgs);
      expect(msgs[1].content).toBe(original);
    });

    it('appends legacy injectedContext to the last user message', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'analyze this' },
      ]);
      const msgs = buildChatMessages('sys');
      injectUploadedFiles(msgs, undefined, 'legacy file content');
      expect(msgs[1].content as string).toContain('<attached_files>');
      expect(msgs[1].content as string).toContain('legacy file content');
    });

    it('injects new inline file content and marks contentInjectedAt', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'process' },
      ]);
      const msgs = buildChatMessages('sys');
      const file = { filename: 'data.csv', content: 'a,b,c', contentInjectedAt: undefined };
      injectUploadedFiles(msgs, [file]);
      expect(msgs[1].content as string).toContain('data.csv');
      expect(msgs[1].content as string).toContain('a,b,c');
      expect(file.contentInjectedAt).toBeDefined();
    });

    it('converts message content to multipart array when file has a fileId', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'summarize' },
      ]);
      const msgs = buildChatMessages('sys');
      injectUploadedFiles(msgs, [
        { filename: 'doc.pdf', content: '', fileId: 'file_abc123' },
      ]);
      expect(Array.isArray(msgs[1].content)).toBe(true);
      const parts = msgs[1].content as any[];
      const filePart = parts.find((p: any) => p.type === 'file');
      expect(filePart?.file?.file_id).toBe('file_abc123');
    });

    it('appends a VFS reference note for already-injected files', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'continue' },
      ]);
      const msgs = buildChatMessages('sys');
      const seenFile = {
        filename: 'report.xlsx',
        content: 'big data',
        contentInjectedAt: Date.now() - 1000,
      };
      injectUploadedFiles(msgs, [seenFile]);
      expect(msgs[1].content as string).toContain('Previously uploaded files available in VFS');
      expect(msgs[1].content as string).toContain('"report.xlsx"');
    });

    it('does not re-inject content for already-seen files', () => {
      const { buildChatMessages, injectUploadedFiles } = makeOrchestration([
        { id: '1', role: 'user', content: 'hello' },
      ]);
      const msgs = buildChatMessages('sys');
      const seenFile = {
        filename: 'old.csv',
        content: 'secret data',
        contentInjectedAt: Date.now() - 5000,
      };
      injectUploadedFiles(msgs, [seenFile]);
      expect(msgs[1].content as string).not.toContain('secret data');
    });
  });

  // ─── injectRichContentInstructions ───────────────────────────────────────

  describe('injectRichContentInstructions', () => {
    it('leaves messages unchanged when there is no rich context', () => {
      vi.mocked(getLastRichContext).mockReturnValue(null);
      const { buildChatMessages, injectRichContentInstructions } = makeOrchestration();
      const msgs = buildChatMessages('Base system prompt.');
      const original = msgs[0].content;
      injectRichContentInstructions(msgs);
      expect(msgs[0].content).toBe(original);
    });

    it('leaves messages unchanged when rich context has no rich content', () => {
      vi.mocked(getLastRichContext).mockReturnValue({ hasRichContent: false } as any);
      const { buildChatMessages, injectRichContentInstructions } = makeOrchestration();
      const msgs = buildChatMessages('Base.');
      injectRichContentInstructions(msgs);
      expect(msgs[0].content).toBe('Base.');
    });

    it('appends preservation instruction to system message when hasRichContent is true', () => {
      vi.mocked(getLastRichContext).mockReturnValue({ hasRichContent: true } as any);
      const { buildChatMessages, injectRichContentInstructions } = makeOrchestration();
      const msgs = buildChatMessages('Base.');
      injectRichContentInstructions(msgs);
      expect(msgs[0].content).toBe('Base. [preserve-rich-content]');
      expect(getPreservationInstruction).toHaveBeenCalledOnce();
    });
  });
});
