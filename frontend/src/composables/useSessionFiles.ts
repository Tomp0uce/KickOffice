/**
 * useSessionFiles.ts
 *
 * Manages uploaded files for the current chat session.
 * - Tracks files uploaded during the session
 * - Rebuilds file list from history when session is restored
 * - Deduplicates files by filename
 *
 * Extracted from useAgentLoop.ts as part of ARCH-H1 refactoring.
 */

import { ref, type Ref } from 'vue'
import type { DisplayMessage } from '@/types/chat'

export interface SessionFile {
  filename: string
  content: string
  fileId?: string
}

export interface UseSessionFilesOptions {
  history: Ref<DisplayMessage[]>
}

export function useSessionFiles(options: UseSessionFilesOptions) {
  const { history } = options

  /**
   * Uploaded files for the current session.
   * These are attached to user messages and provided to the LLM as context.
   */
  const sessionUploadedFiles = ref<SessionFile[]>([])

  /**
   * Add a file to the session.
   * Automatically deduplicates by filename.
   */
  function addSessionFile(file: SessionFile) {
    const exists = sessionUploadedFiles.value.some((f) => f.filename === file.filename)
    if (!exists) {
      sessionUploadedFiles.value.push(file)
    }
  }

  /**
   * Rebuilds sessionUploadedFiles from history after a session switch or restore.
   * Call this whenever history is replaced from IndexedDB.
   *
   * This ensures that files from previous messages in the restored session
   * are available for subsequent LLM requests.
   */
  function rebuildSessionFiles() {
    const seen = new Set<string>()
    sessionUploadedFiles.value = []

    for (const msg of history.value) {
      if (msg.attachedFiles) {
        for (const f of msg.attachedFiles) {
          if (!seen.has(f.filename)) {
            seen.add(f.filename)
            sessionUploadedFiles.value.push(f)
          }
        }
      }
    }
  }

  /**
   * Clear all session files.
   * Typically called when starting a new conversation.
   */
  function clearSessionFiles() {
    sessionUploadedFiles.value = []
  }

  /**
   * Get current session files for chat context.
   * Returns undefined if no files uploaded (cleaner than empty array for API).
   */
  function getSessionFilesForChat(): SessionFile[] | undefined {
    return sessionUploadedFiles.value.length > 0 ? [...sessionUploadedFiles.value] : undefined
  }

  return {
    sessionUploadedFiles,
    addSessionFile,
    rebuildSessionFiles,
    clearSessionFiles,
    getSessionFilesForChat,
  }
}
