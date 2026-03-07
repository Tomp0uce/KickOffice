# State Management — KickOffice Frontend

## Overview

Frontend state is split across 5 distinct layers. Each layer has a specific scope and lifetime.

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         KickOffice Frontend State                       │
│                                                                         │
│  ┌─────────────┐  ┌──────────────────┐  ┌──────────────────────────┐  │
│  │  Vue refs   │  │   Composables    │  │       IndexedDB          │  │
│  │  (in-mem)   │  │  (shared state)  │  │    (persistent, async)   │  │
│  └─────────────┘  └──────────────────┘  └──────────────────────────┘  │
│  ┌────────────────────────┐  ┌─────────────────────────────────────┐   │
│  │      localStorage      │  │           sessionStorage            │   │
│  │  (persistent, synced)  │  │      (tab-scoped, ephemeral)        │   │
│  └────────────────────────┘  └─────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Layer 1 — Vue Reactive Refs (in-memory)

**Lifetime**: Component mount → unmount
**Location**: `.vue` SFC files (e.g. `ChatInput.vue`, `HomePage.vue`, `SettingsPage.vue`)

| What lives here | Example |
|-----------------|---------|
| UI state (loading, open panels) | `isLoading`, `isSidebarOpen` |
| Current message being typed | `inputText` |
| Transient error/success states | `errorMessage`, `saveSuccess` |
| Local form state | model selection, temperature slider |

**Rule**: Nothing in this layer should survive a page reload.

---

## Layer 2 — Composables (shared reactive state)

**Lifetime**: App lifetime (composables are singletons created once per app instance)
**Location**: `src/composables/`

| Composable | Responsibility |
|------------|---------------|
| `useAgentLoop.ts` | Chat loop execution, tool dispatch, SSE streaming, abort controller |
| `useSessionManager.ts` | Active session (ID, title, message list), session switching |
| `useSessionDB.ts` | IndexedDB read/write facade (wraps Dexie) |

**Rule**: Composables own the shared state that multiple components need simultaneously. They are the single source of truth for active chat state.

---

## Layer 3 — IndexedDB (via Dexie)

**Lifetime**: Persistent across sessions until explicitly cleared
**Location**: `src/composables/useSessionDB.ts`

| Store | Key | What lives here |
|-------|-----|-----------------|
| `sessions` | `sessionId` | Session metadata (title, createdAt, hostType) |
| `messages` | `sessionId + index` | Full chat message history including tool calls and tool results |
| `vfsSnapshots` | `sessionId` | Virtual File System state snapshots (future) |

**Rule**: IndexedDB is the only layer allowed to hold arbitrarily large blobs (message history, snapshots). Never store large data in localStorage.

---

## Layer 4 — localStorage (persistent, encrypted for credentials)

**Lifetime**: Persistent across browser sessions
**Location**: `src/utils/credentialStorage.ts`, `src/utils/constant.ts`, `src/types/enum.ts`

| Key | Set by | What |
|-----|--------|------|
| `ko_cred_litellmUserKey` | `credentialStorage.ts` | AES-GCM encrypted API key |
| `ko_cred_litellmUserEmail` | `credentialStorage.ts` | AES-GCM encrypted user email |
| `ko_encryption_key` | `credentialCrypto.ts` | JWK-serialized AES-GCM key |
| `rememberCredentials` | `credentialStorage.ts` | Boolean preference (default: `true`) |
| `promptOverrides_*` | `constant.ts` | Per-user prompt customizations |
| `selectedModel` | `enum.ts` / UI | Last selected model tier |
| `selectedLanguage` | UI | UI language preference |

**Encryption**: Credentials are encrypted with AES-GCM-256 via the Web Crypto API (see `credentialCrypto.ts`). The encryption key itself is stored in localStorage (or sessionStorage when `rememberCredentials=false`).

**Rule**: Never store plaintext credentials in localStorage. Use `setUserKey()` / `setUserEmail()` from `credentialStorage.ts`.

---

## Layer 5 — sessionStorage (tab-scoped fallback)

**Lifetime**: Current tab only — cleared on tab close or Office Add-in restart
**Location**: `src/utils/credentialStorage.ts`

| Key | What |
|-----|------|
| `litellmUserKey` | Plaintext API key (used when `rememberCredentials=false`) |
| `litellmUserEmail` | Plaintext user email (used when `rememberCredentials=false`) |
| `ko_encryption_key` | AES key when `rememberCredentials=false` |

**Rule**: Office Add-ins MUST default to `rememberCredentials=true` because sessionStorage is wiped on every taskpane restart. sessionStorage is kept as a fallback for security-conscious users who explicitly opt out of persistence.

---

## Data Flow: Where Does What Go?

```
User types API key
      │
      ▼
SettingsPage.vue (Vue ref: inputKey)
      │
      ▼ setUserKey(value)
credentialStorage.ts
      │
      ├─ rememberCredentials=true ──▶ encryptValue() ──▶ localStorage (ko_cred_*)
      │
      └─ rememberCredentials=false ─────────────────▶ sessionStorage (plaintext)

Chat message sent
      │
      ▼
useAgentLoop.ts (Vue ref: messages[])
      │
      ▼ saveMessage()
useSessionDB.ts
      │
      ▼
IndexedDB (messages store)
```

---

## Migration Paths

`credentialStorage.ts` handles automatic migration when `rememberCredentials` changes:

- **false → true**: Reads plaintext from sessionStorage, encrypts, writes to localStorage, clears sessionStorage.
- **true → false**: Decrypts from localStorage, writes plaintext to sessionStorage, clears localStorage.
- **Key migration**: If the encryption key is found in the wrong storage (from a previous session where `rememberCredentials` changed), it is transparently moved to the correct storage.
