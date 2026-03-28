import { describe, it, expect } from 'vitest';
import { ref } from 'vue';
import { useSessionFiles } from '../useSessionFiles';
import type { DisplayMessage } from '@/types/chat';
import type { SessionFile } from '../useSessionFiles';

function makeHistory(messages: Partial<DisplayMessage>[] = []) {
  return ref(messages as DisplayMessage[]);
}

describe('useSessionFiles', () => {
  describe('addSessionFile', () => {
    it('adds a file to the session', () => {
      const { sessionUploadedFiles, addSessionFile } = useSessionFiles({ history: makeHistory() });
      addSessionFile({ filename: 'report.xlsx', content: 'data' });
      expect(sessionUploadedFiles.value).toHaveLength(1);
      expect(sessionUploadedFiles.value[0].filename).toBe('report.xlsx');
    });

    it('deduplicates by filename — second add with same name is ignored', () => {
      const { sessionUploadedFiles, addSessionFile } = useSessionFiles({ history: makeHistory() });
      addSessionFile({ filename: 'doc.txt', content: 'v1' });
      addSessionFile({ filename: 'doc.txt', content: 'v2' });
      expect(sessionUploadedFiles.value).toHaveLength(1);
      expect(sessionUploadedFiles.value[0].content).toBe('v1');
    });

    it('allows adding files with different filenames', () => {
      const { sessionUploadedFiles, addSessionFile } = useSessionFiles({ history: makeHistory() });
      addSessionFile({ filename: 'a.txt', content: '' });
      addSessionFile({ filename: 'b.txt', content: '' });
      expect(sessionUploadedFiles.value).toHaveLength(2);
    });
  });

  describe('clearSessionFiles', () => {
    it('empties the file list', () => {
      const { sessionUploadedFiles, addSessionFile, clearSessionFiles } = useSessionFiles({
        history: makeHistory(),
      });
      addSessionFile({ filename: 'x.csv', content: '' });
      clearSessionFiles();
      expect(sessionUploadedFiles.value).toHaveLength(0);
    });
  });

  describe('getSessionFilesForChat', () => {
    it('returns undefined when no files are uploaded', () => {
      const { getSessionFilesForChat } = useSessionFiles({ history: makeHistory() });
      expect(getSessionFilesForChat()).toBeUndefined();
    });

    it('returns a copy of the files when files exist', () => {
      const { sessionUploadedFiles, addSessionFile, getSessionFilesForChat } = useSessionFiles({
        history: makeHistory(),
      });
      addSessionFile({ filename: 'data.csv', content: 'a,b' });
      const result = getSessionFilesForChat();
      expect(result).toHaveLength(1);
      expect(result![0].filename).toBe('data.csv');
      // Mutation of the returned copy must not affect the reactive store
      result!.push({ filename: 'injected.txt', content: '' });
      expect(sessionUploadedFiles.value).toHaveLength(1);
    });
  });

  describe('rebuildSessionFiles', () => {
    it('populates files from history messages that have attachedFiles', () => {
      const history = makeHistory([
        {
          id: '1',
          role: 'user',
          content: 'hello',
          attachedFiles: [{ filename: 'sheet.xlsx', content: 'abc' }],
        },
      ]);
      const { sessionUploadedFiles, rebuildSessionFiles } = useSessionFiles({ history });
      rebuildSessionFiles();
      expect(sessionUploadedFiles.value).toHaveLength(1);
      expect(sessionUploadedFiles.value[0].filename).toBe('sheet.xlsx');
    });

    it('deduplicates files across multiple messages', () => {
      const history = makeHistory([
        {
          id: '1',
          role: 'user',
          content: '',
          attachedFiles: [{ filename: 'f.txt', content: '1' }],
        },
        {
          id: '2',
          role: 'user',
          content: '',
          attachedFiles: [{ filename: 'f.txt', content: '2' }],
        },
      ]);
      const { sessionUploadedFiles, rebuildSessionFiles } = useSessionFiles({ history });
      rebuildSessionFiles();
      expect(sessionUploadedFiles.value).toHaveLength(1);
      expect(sessionUploadedFiles.value[0].content).toBe('1');
    });

    it('collects files from multiple different messages', () => {
      const history = makeHistory([
        { id: '1', role: 'user', content: '', attachedFiles: [{ filename: 'a.txt', content: '' }] },
        { id: '2', role: 'user', content: '', attachedFiles: [{ filename: 'b.txt', content: '' }] },
      ]);
      const { sessionUploadedFiles, rebuildSessionFiles } = useSessionFiles({ history });
      rebuildSessionFiles();
      expect(sessionUploadedFiles.value).toHaveLength(2);
    });

    it('clears previously added files before rebuilding', () => {
      const history = makeHistory([
        {
          id: '1',
          role: 'user',
          content: '',
          attachedFiles: [{ filename: 'new.txt', content: '' }],
        },
      ]);
      const { sessionUploadedFiles, addSessionFile, rebuildSessionFiles } = useSessionFiles({
        history,
      });
      addSessionFile({ filename: 'old.txt', content: '' });
      rebuildSessionFiles();
      expect(sessionUploadedFiles.value.map((f: SessionFile) => f.filename)).toEqual(['new.txt']);
    });

    it('handles messages without attachedFiles gracefully', () => {
      const history = makeHistory([
        { id: '1', role: 'user', content: 'just text' },
        { id: '2', role: 'assistant', content: 'reply' },
      ]);
      const { sessionUploadedFiles, rebuildSessionFiles } = useSessionFiles({ history });
      rebuildSessionFiles();
      expect(sessionUploadedFiles.value).toHaveLength(0);
    });
  });
});
