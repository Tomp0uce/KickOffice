import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ref } from 'vue';
import type { DisplayMessage } from '@/types/chat';

// Mock dependencies before import
vi.mock('@/composables/useSessionDB', () => ({
  listSessions: vi.fn(),
  createSession: vi.fn(),
  saveSession: vi.fn(),
  deleteSession: vi.fn(),
  getSessionMessageCount: vi.fn(),
}));

vi.mock('@/utils/powerpointTools', () => ({
  clearPowerpointImageRegistry: vi.fn(),
}));

vi.mock('@/utils/logger', () => ({
  logService: {
    setCurrentSessionId: vi.fn(),
    warn: vi.fn(),
    clearSessionLogs: vi.fn(),
  },
}));

import { useSessionManager } from '@/composables/useSessionManager';
import {
  listSessions,
  createSession,
  saveSession,
  deleteSession,
} from '@/composables/useSessionDB';
import { clearPowerpointImageRegistry } from '@/utils/powerpointTools';

const mockSession = (id: string, messages: DisplayMessage[] = []) => ({
  id,
  hostType: 'Word',
  name: `Session ${id}`,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  messages,
});

const userMsg: DisplayMessage = { id: '1', role: 'user', content: 'hello' };
const assistantMsg: DisplayMessage = { id: '2', role: 'assistant', content: 'world' };

beforeEach(() => {
  vi.clearAllMocks();
});

describe('useSessionManager', () => {
  // ─── init ──────────────────────────────────────────────────────────────────

  describe('init', () => {
    it('creates first session when no sessions exist', async () => {
      vi.mocked(listSessions).mockResolvedValue([]);
      vi.mocked(createSession).mockResolvedValue(mockSession('s1'));

      const history = ref<DisplayMessage[]>([]);
      const { init, currentSessionId } = useSessionManager('Word', history);
      await init();

      expect(createSession).toHaveBeenCalledWith('Word');
      expect(currentSessionId.value).toBe('s1');
      expect(history.value).toEqual([]);
    });

    it('loads latest session when sessions exist', async () => {
      const msgs = [userMsg, assistantMsg];
      vi.mocked(listSessions).mockResolvedValue([mockSession('s2', msgs), mockSession('s1')]);

      const history = ref<DisplayMessage[]>([]);
      const { init, currentSessionId } = useSessionManager('Word', history);
      await init();

      expect(currentSessionId.value).toBe('s2');
      expect(history.value).toEqual(msgs);
    });
  });

  // ─── newSession ────────────────────────────────────────────────────────────

  describe('newSession', () => {
    it('does nothing when current session is empty', async () => {
      vi.mocked(listSessions).mockResolvedValue([mockSession('s1')]);
      vi.mocked(createSession).mockResolvedValue(mockSession('s2'));

      const history = ref<DisplayMessage[]>([]);
      const { init, newSession } = useSessionManager('Word', history);
      await init();

      await newSession();
      // createSession only called once during init, not again
      expect(createSession).toHaveBeenCalledTimes(0); // init doesn't create since sessions exist
    });

    it('saves current and creates new session when history is non-empty', async () => {
      const msgs = [userMsg];
      vi.mocked(listSessions).mockResolvedValue([mockSession('s1', msgs)]);
      vi.mocked(createSession).mockResolvedValue(mockSession('s2'));

      const history = ref<DisplayMessage[]>(msgs);
      const { newSession, currentSessionId } = useSessionManager('Word', history);

      // Manually set session state as if init ran
      currentSessionId.value = 's1';

      await newSession();

      expect(saveSession).toHaveBeenCalledWith('s1', msgs);
      expect(createSession).toHaveBeenCalledWith('Word');
      expect(currentSessionId.value).toBe('s2');
      expect(history.value).toEqual([]);
      expect(clearPowerpointImageRegistry).toHaveBeenCalled();
    });
  });

  // ─── switchSession ─────────────────────────────────────────────────────────

  describe('switchSession', () => {
    it('switches to target session and saves current', async () => {
      const s1msgs = [userMsg];
      const s2msgs = [assistantMsg];
      vi.mocked(listSessions).mockResolvedValue([
        mockSession('s1', s1msgs),
        mockSession('s2', s2msgs),
      ]);

      const history = ref<DisplayMessage[]>(s1msgs);
      const { switchSession, currentSessionId } = useSessionManager('Word', history);
      currentSessionId.value = 's1';

      await switchSession('s2');

      expect(saveSession).toHaveBeenCalledWith('s1', s1msgs);
      expect(currentSessionId.value).toBe('s2');
      expect(history.value).toEqual(s2msgs);
    });

    it('does nothing when switching to current session', async () => {
      const history = ref<DisplayMessage[]>([]);
      const { switchSession, currentSessionId } = useSessionManager('Word', history);
      currentSessionId.value = 's1';

      await switchSession('s1');
      expect(saveSession).not.toHaveBeenCalled();
    });

    it('blocks when agent loop is running', async () => {
      const isAgentRunning = ref(true);
      const history = ref<DisplayMessage[]>([]);
      const { switchSession, currentSessionId } = useSessionManager(
        'Word',
        history,
        isAgentRunning,
      );
      currentSessionId.value = 's1';

      await switchSession('s2');
      expect(saveSession).not.toHaveBeenCalled();
      expect(currentSessionId.value).toBe('s1');
    });
  });

  // ─── deleteCurrentSession ──────────────────────────────────────────────────

  describe('deleteCurrentSession', () => {
    it('switches to another session and deletes the current one', async () => {
      const s2msgs = [assistantMsg];
      const s1 = mockSession('s1');
      const s2 = mockSession('s2', s2msgs);
      vi.mocked(listSessions).mockResolvedValue([s1, s2]);

      const history = ref<DisplayMessage[]>([]);
      const { deleteCurrentSession, currentSessionId, sessions } = useSessionManager(
        'Word',
        history,
      );
      currentSessionId.value = 's1';
      // Simulate sessions being loaded (as if init ran)
      sessions.value = [s1, s2];

      await deleteCurrentSession();

      expect(deleteSession).toHaveBeenCalledWith('s1');
      expect(currentSessionId.value).toBe('s2');
      expect(history.value).toEqual(s2msgs);
    });

    it('creates fresh session when deleting the only session', async () => {
      vi.mocked(listSessions).mockResolvedValue([mockSession('s1')]);
      vi.mocked(createSession).mockResolvedValue(mockSession('s2'));

      const history = ref<DisplayMessage[]>([userMsg]);
      const { deleteCurrentSession, currentSessionId, sessions } = useSessionManager(
        'Word',
        history,
      );
      currentSessionId.value = 's1';
      sessions.value = [mockSession('s1')];

      await deleteCurrentSession();

      expect(createSession).toHaveBeenCalledWith('Word');
      expect(currentSessionId.value).toBe('s2');
      expect(history.value).toEqual([]);
    });

    it('does nothing when no current session', async () => {
      const history = ref<DisplayMessage[]>([]);
      const { deleteCurrentSession, currentSessionId } = useSessionManager('Word', history);
      currentSessionId.value = null;

      await deleteCurrentSession();
      expect(deleteSession).not.toHaveBeenCalled();
    });
  });

  // ─── persistCurrentSession ─────────────────────────────────────────────────

  describe('persistCurrentSession', () => {
    it('saves current session and reloads sessions', async () => {
      vi.mocked(listSessions).mockResolvedValue([]);

      const history = ref<DisplayMessage[]>([userMsg]);
      const { persistCurrentSession, currentSessionId } = useSessionManager('Word', history);
      currentSessionId.value = 's1';

      await persistCurrentSession();

      expect(saveSession).toHaveBeenCalledWith('s1', [userMsg]);
      expect(listSessions).toHaveBeenCalled();
    });
  });
});
