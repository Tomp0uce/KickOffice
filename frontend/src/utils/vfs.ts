/**
 * Virtual Filesystem (VFS) for KickOffice agents
 *
 * Provides an in-memory filesystem using just-bash/browser that allows:
 * - Users to upload files (images, CSVs, etc.)
 * - Agent to read files via read_file tool
 * - Agent to execute bash commands via bash tool
 *
 * Adapted from Open_Excel open-excel-main/src/lib/vfs/index.ts
 */

import { Bash, InMemoryFs } from 'just-bash/browser';

// Singleton instances (reset on session switch)
let fs: InMemoryFs | null = null;
let bash: Bash | null = null;

/**
 * Get or create the virtual filesystem instance
 */
export function getVfs(): InMemoryFs {
  if (!fs) {
    fs = new InMemoryFs({
      '/home/user/uploads/.keep': '',
      '/home/user/scripts/.keep': '',
    });
  }
  return fs;
}

/**
 * Get or create the Bash instance
 */
export function getBash(): Bash {
  if (!bash) {
    bash = new Bash({
      fs: getVfs(),
      cwd: '/home/user',
    });
  }
  return bash;
}

/**
 * Reset the VFS (clears all files, creates fresh instances)
 * Call this when switching sessions.
 */
export function resetVfs(): void {
  fs = null;
  bash = null;
}

/**
 * Snapshot all files in the VFS as path→Uint8Array pairs.
 * Used for persisting VFS state to IndexedDB.
 */
export async function snapshotVfs(): Promise<{ path: string; data: Uint8Array }[]> {
  const vfs = getVfs();
  const allPaths = vfs.getAllPaths();
  const files: { path: string; data: Uint8Array }[] = [];

  for (const p of allPaths) {
    try {
      const stat = await vfs.stat(p);
      if (stat.isFile) {
        const data = await vfs.readFileBuffer(p);
        files.push({ path: p, data });
      }
    } catch {
      // skip unreadable entries
    }
  }

  return files;
}

/**
 * Restore VFS from a snapshot. Resets existing state and writes all files.
 */
export async function restoreVfs(files: { path: string; data: Uint8Array }[]): Promise<void> {
  resetVfs();

  if (files.length === 0) {
    getVfs(); // Just initialize default VFS
    return;
  }

  const initialFiles: Record<string, Uint8Array | string> = {
    '/home/user/uploads/.keep': '',
    '/home/user/scripts/.keep': '',
  };
  for (const f of files) {
    initialFiles[f.path] = f.data;
  }

  fs = new InMemoryFs(initialFiles);
  bash = null; // will be lazily created with the new fs
}

/**
 * Write a file to the VFS
 */
export async function writeFile(path: string, content: string | Uint8Array): Promise<void> {
  const vfs = getVfs();
  const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;

  // Ensure parent directory exists
  const dir = fullPath.substring(0, fullPath.lastIndexOf('/'));
  if (dir && dir !== '/') {
    try {
      await vfs.mkdir(dir, { recursive: true });
    } catch {
      // Directory might already exist
    }
  }

  await vfs.writeFile(fullPath, content);
}

/**
 * Read a file from the VFS
 */
export async function readFile(path: string): Promise<string> {
  const vfs = getVfs();
  const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;
  return vfs.readFile(fullPath);
}

/**
 * Returns the VFS helper object to expose inside Office.js eval sandboxes.
 *
 * DUP-M1: The readFile / readFileBuffer / writeFile triplet was duplicated
 * verbatim in wordTools, excelTools, and powerpointTools.  Centralised here so
 * any future change (e.g. path resolution logic) is made in one place.
 */
export function getVfsSandboxContext() {
  return {
    readFile,
    readFileBuffer: async (path: string): Promise<Uint8Array> => {
      const vfs = getVfs();
      const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;
      return vfs.readFileBuffer(fullPath);
    },
    writeFile,
  };
}

/**
 * List files in the VFS uploads directory
 */
export async function listUploads(): Promise<string[]> {
  const vfs = getVfs();
  try {
    const entries = await vfs.readdir('/home/user/uploads');
    return entries.filter(e => e !== '.keep');
  } catch {
    return [];
  }
}
