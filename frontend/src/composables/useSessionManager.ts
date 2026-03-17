/**
 * Session management composable for KickOffice.
 * Adapté de l'implémentation OpenExcel (open-excel-main) — multi-session avec IndexedDB.
 */
import { ref, type Ref } from 'vue';
import type { DisplayMessage } from '@/types/chat';
import {
  type ChatSession,
  listSessions,
  createSession,
  saveSession,
  deleteSession,
  getSessionMessageCount,
} from '@/composables/useSessionDB';
import { logService } from '@/utils/logger';
import { clearPowerpointImageRegistry } from '@/utils/powerpointTools';

export type { ChatSession };
export { getSessionMessageCount };

export function useSessionManager(
  hostType: string,
  history: Ref<DisplayMessage[]>,
  isAgentRunning?: Ref<boolean>,
) {
  const sessions = ref<ChatSession[]>([]);
  const currentSessionId = ref<string | null>(null);
  const isSwitching = ref(false);

  const currentSession = (): ChatSession | undefined =>
    sessions.value.find((s: ChatSession) => s.id === currentSessionId.value);

  async function loadSessions() {
    sessions.value = await listSessions(hostType);
  }

  async function init() {
    await loadSessions();
    if (sessions.value.length === 0) {
      const session = await createSession(hostType);
      sessions.value = [session];
      currentSessionId.value = session.id;
      logService.setCurrentSessionId(session.id);
      history.value = [];
    } else {
      const latest = sessions.value[0];
      currentSessionId.value = latest.id;
      logService.setCurrentSessionId(latest.id);
      history.value = latest.messages ?? [];
    }
  }

  async function newSession() {
    // If the current session is already empty, do not create a new one.
    if (history.value.length === 0) {
      return;
    }

    // Save current session first
    if (currentSessionId.value) {
      await saveSession(currentSessionId.value, history.value);
    }
    const session = await createSession(hostType);
    await loadSessions();
    currentSessionId.value = session.id;
    logService.setCurrentSessionId(session.id);
    history.value = [];
    clearPowerpointImageRegistry(); // QUAL-M2: prevent memory leak across sessions
  }

  async function switchSession(sessionId: string) {
    if (isSwitching.value) return;
    if (sessionId === currentSessionId.value) return;
    // RACE-C1: Prevent history.value replacement while the agent loop is actively writing to it.
    // If the caller already guards (e.g. useHomePage.handleSwitchSession), this is a second
    // defensive layer that also protects any direct call paths.
    if (isAgentRunning?.value) {
      logService.warn('[SessionManager] switchSession blocked — agent loop is running', {
        targetSessionId: sessionId,
      });
      return;
    }
    isSwitching.value = true;
    try {
      // Save current session
      if (currentSessionId.value) {
        await saveSession(currentSessionId.value, history.value);
      }
      // Reload sessions to get latest names
      await loadSessions();
      const target = sessions.value.find((s: ChatSession) => s.id === sessionId);
      if (!target) return;
      currentSessionId.value = sessionId;
      logService.setCurrentSessionId(sessionId);
      history.value = target.messages ?? [];
      clearPowerpointImageRegistry(); // QUAL-M2: prevent memory leak across sessions
    } finally {
      isSwitching.value = false;
    }
  }

  async function persistCurrentSession() {
    if (currentSessionId.value) {
      await saveSession(currentSessionId.value, history.value);
      await loadSessions();
    }
  }

  async function deleteCurrentSession() {
    if (!currentSessionId.value) return;
    const idToDelete = currentSessionId.value;
    // Switch to another session first
    const others = sessions.value.filter((s: ChatSession) => s.id !== idToDelete);
    if (others.length > 0) {
      currentSessionId.value = others[0].id;
      logService.setCurrentSessionId(others[0].id);
      history.value = others[0].messages ?? [];
    } else {
      // Create a new one if none remain
      const fresh = await createSession(hostType);
      currentSessionId.value = fresh.id;
      logService.setCurrentSessionId(fresh.id);
      history.value = [];
    }
    await deleteSession(idToDelete);
    logService.clearSessionLogs(idToDelete);
    await loadSessions();
  }

  return {
    sessions,
    currentSessionId,
    currentSession,
    init,
    newSession,
    switchSession,
    persistCurrentSession,
    deleteCurrentSession,
  };
}
