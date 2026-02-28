# REPORT: OpenExcel vs KickOffice — Feature Comparison & Gap Analysis

> Date: 28 February 2026
> Scope: Comprehensive comparison of chat/agent UX, conversation management, status feedback, and tool ecosystem
> Sources: OpenExcel `open-excel-main.zip` reference implementation + KickOffice codebase analysis

---

## TABLE OF CONTENTS

1. [Executive Summary](#1-executive-summary)
2. [Chat Agent UX & Status Updates](#2-chat-agent-ux--status-updates)
3. [Conversation Management](#3-conversation-management)
4. [Tool Execution Transparency](#4-tool-execution-transparency)
5. [Thinking Blocks & Reasoning Display](#5-thinking-blocks--reasoning-display)
6. [Stats Bar & Token/Cost Tracking](#6-stats-bar--tokencost-tracking)
7. [File Handling & Attachments](#7-file-handling--attachments)
8. [Configuration & Settings](#8-configuration--settings)
9. [Multi-Host vs Single-Host](#9-multi-host-vs-single-host)
10. [Tool Ecosystem](#10-tool-ecosystem)
11. [Summary Comparison Table](#11-summary-comparison-table)
12. [Recommended Priorities](#12-recommended-priorities)

---

## 1. EXECUTIVE SUMMARY

**OpenExcel** is a single-host (Excel-only) AI add-in built with React and `@mariozechner/pi-agent-core`. It excels in **agent UX polish**: multi-session conversations, granular tool execution status indicators, a dedicated stats bar with token/cost tracking, and smooth streaming with per-part rendering.

**KickOffice** is a multi-host add-in (Word, Excel, PowerPoint, Outlook) built with Vue 3 and a custom agent loop. It has broader Office coverage and a richer tool set (116+ tools vs ~15), but its chat UX is comparatively basic: single conversation per host, minimal status feedback during agent execution, and no token/cost visibility.

**Key takeaway**: KickOffice's architecture is sound — the gaps are primarily in **frontend UX polish and conversation management**, not in core agent capabilities. OpenExcel's patterns can be adopted incrementally.

---

## 2. CHAT AGENT UX & STATUS UPDATES

### How OpenExcel handles agent mode status updates

OpenExcel provides a **rich, real-time feedback system** during agent execution:

1. **Agent Event System**: The agent emits typed events (`message_start`, `message_update`, `message_end`, `tool_execution_start`, `tool_execution_update`, `tool_execution_end`) that drive granular UI updates.

2. **In-Chat Status Updates**: Each tool call appears as a **collapsible block** directly in the chat with:
   - Tool name with human-readable explanation
   - Spinning loader (pending/running) → green checkmark (complete) → red X (error)
   - Expandable details showing arguments, results, images, and modified cell ranges
   - Modified ranges are **clickable links** that navigate to the affected cells

3. **Loading Indicator**: When waiting for the first response token, a "thinking..." spinner appears below the last user message.

4. **Streaming Cursor**: A pulsing `▊` block cursor animates at the end of streaming text, giving a "typing" feel.

### How KickOffice handles agent mode status updates

KickOffice provides **minimal status feedback**:

1. **Single Status String**: A `currentAction` reactive ref shows one-word status at the bottom of the chat:
   - "Analyzing..." (waiting for response)
   - "Reading..." / "Formatting..." / "Running..." (during tool execution, categorized)
   - "Uploading files..." (during file attachment)

2. **Pulsing Dot**: A small animated dot with the status text appears at the bottom of the message list.

3. **No Per-Tool Indicators**: Tool executions are invisible in the chat. The user sees the agent's text responses but never sees which tools were called, their arguments, or their results.

4. **No Streaming Cursor**: Text appears progressively but without a visual cursor indicator.

### Gap Analysis

| Feature | OpenExcel | KickOffice | Priority |
|---------|-----------|------------|----------|
| Per-tool status blocks in chat | Full (pending/running/complete/error) | None | **HIGH** |
| Tool call details (args/results) | Expandable per tool call | Not shown | **HIGH** |
| Spinning loader during tool exec | Per-tool loader icon | Single pulsing dot | MEDIUM |
| Modified range navigation | Clickable links to cells | Not applicable (multi-host) | LOW |
| Streaming cursor animation | Pulsing `▊` block | None | LOW |
| Loading "thinking..." indicator | Dedicated component | Pulsing dot + text | MEDIUM |

### Recommendation

**Priority: HIGH** — Implement per-tool-call status blocks in `ChatMessageList.vue`. Each tool call from the agent should appear as a collapsible element in the chat flow, showing:
- Tool name (human-readable)
- Status icon (spinner → checkmark/error)
- Expandable section with arguments and results

This requires extending `DisplayMessage` to carry tool call metadata and updating `useAgentLoop.ts` to emit tool-level status updates to the message history.

---

## 3. CONVERSATION MANAGEMENT

### How OpenExcel manages conversations

OpenExcel uses **IndexedDB** (`idb` library) for persistent multi-session management:

1. **Per-Workbook Sessions**: Each Excel file has its own set of conversations, identified by a workbook ID stored in `Office.context.document.settings`.

2. **Session Lifecycle**:
   - Auto-created on first load
   - Auto-named from the first user message (truncated at 40 chars)
   - Saved automatically after each agent turn completes (not during streaming)
   - Full message history + VFS file snapshots preserved per session

3. **Session Switcher UI**: A dropdown in the chat header shows:
   - Current session name (truncated at 20 chars)
   - List of all sessions with message counts
   - Current session highlighted with checkmark
   - "New Chat" and "Delete" buttons
   - Disabled during streaming (prevents corruption)

4. **Session Switch Flow**: `reset agent` → `load messages from DB` → `restore VFS files` → `replaceMessages()` → `update UI`

### How KickOffice manages conversations

KickOffice uses **localStorage** with a single-conversation-per-host model:

1. **One Conversation**: `chatHistory_word`, `chatHistory_excel`, etc. — one flat array per Office host.

2. **Max 100 Messages**: Hard limit; oldest messages are pruned when exceeded.

3. **"New Chat" = Page Reload**: `startNewChat()` calls `window.location.reload()`, destroying all UI state.

4. **No Session Persistence**: Previous conversations are **overwritten** (as the user noted — "ecrase" behavior). There is no way to go back to a previous conversation.

### Gap Analysis

| Feature | OpenExcel | KickOffice | Priority |
|---------|-----------|------------|----------|
| Multi-session support | Full (IndexedDB) | None (single per host) | **CRITICAL** |
| Session switcher UI | Dropdown in header | None | **CRITICAL** |
| Session auto-naming | First user message | N/A | HIGH |
| Per-workbook isolation | Via document settings | Per host type only | MEDIUM |
| New Chat preserves history | Creates new session | **Overwrites** current | **CRITICAL** |
| Session metadata | Created/updated dates, msg count | None | MEDIUM |
| VFS file persistence per session | Full | None | LOW |
| Session delete | With confirmation | N/A | LOW |

### Recommendation

**Priority: CRITICAL** — This is the most impactful UX gap. Users lose their conversation history every time they click "New Chat". Implementation plan:

1. **Storage**: Migrate from localStorage to IndexedDB (via `idb` or `Dexie`). Store sessions with `id`, `hostType`, `name`, `messages`, `createdAt`, `updatedAt`.
2. **Session Model**: Each "New Chat" creates a new session instead of overwriting. Old sessions persist and can be re-opened.
3. **Session Switcher**: Add a dropdown to `ChatHeader.vue` listing all sessions for the current host, with message counts and auto-generated names.
4. **Files Affected**: `ChatHeader.vue`, `HomePage.vue`, new `useSessionManager.ts` composable, `types/chat.ts` (new `ChatSession` interface).

---

## 4. TOOL EXECUTION TRANSPARENCY

### OpenExcel approach

Tool calls are **first-class citizens** in the message display:

```
Assistant message:
├── TextPart: "Let me check the data in column A..."
├── ToolCallPart: get_cell_ranges
│   ├── Status: ✓ complete
│   ├── Args: { ranges: ["A1:A100"] }
│   ├── Result: { data: [...] }
│   └── Modified: [clickable range links]
├── TextPart: "I found 50 entries. Now let me format them..."
└── ToolCallPart: set_cell_range
    ├── Status: ⟳ running
    └── Args: { range: "B1:B50", values: [...] }
```

Each tool call has:
- **Wrench icon** + tool name
- **Status badge**: pending (gray spinner), running (indigo spinner), complete (green check), error (red X)
- **Expandable**: Click to see JSON arguments, result text, images, error messages
- **Modified ranges**: Yellow-highlighted clickable links navigating to the affected cells

### KickOffice approach

Tool calls are **completely invisible** to the user:

```
Assistant message:
└── TextPart: "I've formatted the data in column A."
```

The user only sees the agent's final text response. They have no visibility into:
- Which tools were called
- What arguments were used
- Whether a tool succeeded or failed
- How many tool calls occurred in the turn

The only feedback is a generic `currentAction` string ("Reading...", "Formatting...") at the bottom of the chat.

### Recommendation

**Priority: HIGH** — Add a `MessagePart` system to `DisplayMessage`:

```typescript
interface ToolCallPart {
  type: 'toolCall'
  id: string
  name: string
  args: Record<string, unknown>
  status: 'pending' | 'running' | 'complete' | 'error'
  result?: string
  error?: string
}

interface DisplayMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  parts?: (TextPart | ThinkPart | ToolCallPart)[]
  imageSrc?: string
  richHtml?: string
}
```

Then render tool calls inline in `ChatMessageList.vue` as collapsible blocks between text segments.

---

## 5. THINKING BLOCKS & REASONING DISPLAY

### OpenExcel approach

- **Dedicated `ThinkingBlock` component** with brain icon
- **Streaming indicator**: Animated dots `...` pulse while thinking is still in progress
- **Visual design**: Collapsible with `ChevronDown`/`ChevronRight` icons, brain icon, "thinking" label
- **Streaming-aware**: Content updates in real-time as the model thinks

### KickOffice approach

- **Native HTML `<details>` element** with "Thought Process" summary (i18n: `thoughtProcess`)
- **No streaming indicator**: The thinking block appears all at once after parsing
- **Basic styling**: Border, rounded corners, secondary background
- **Expand tracking**: `expandedThoughts` object tracks open/closed state per segment

### Gap Analysis

| Feature | OpenExcel | KickOffice | Priority |
|---------|-----------|------------|----------|
| Thinking block component | Dedicated with icons | Native `<details>` | LOW |
| Streaming indicator in block | Animated dots | None | MEDIUM |
| Brain/icon visual | Brain + chevron icons | Default disclosure triangle | LOW |
| Real-time content update | During streaming | After complete | MEDIUM |

### Recommendation

**Priority: LOW-MEDIUM** — The current `<details>` approach works but lacks polish. A dedicated `ThinkingBlock.vue` component with:
- Brain icon + "Thinking" label
- Animated dots during streaming
- Smooth expand/collapse transition

Would be a quick UX win but is not critical.

---

## 6. STATS BAR & TOKEN/COST TRACKING

### OpenExcel approach

A dedicated **`StatsBar` component** at the bottom of the chat interface:

```
↑1.2k  ↓3.5k  R500  W200          $0.0142
25.3%/128k    anthropic • claude-sonnet-4-20250514 • medium
```

Displays:
- **Input tokens** (↑): Tokens sent to model
- **Output tokens** (↓): Tokens received from model
- **Cache read** (R): Cached tokens reused
- **Cache write** (W): New tokens cached
- **Total cost** ($): Calculated from token counts and model pricing
- **Context usage** (%): Current context window usage vs max
- **Provider + Model**: Active provider and model name
- **Thinking level**: If extended thinking is enabled

Updated after every agent turn via `calculateSessionStats()`.

### KickOffice approach

**No stats bar exists.** The only model-related UI is:
- Model tier dropdown in `ChatInput.vue` (shows "standard", "reasoning", "image")
- Backend health indicator (green/red dot)

No token counts, no cost tracking, no context usage display, no model name display.

### Recommendation

**Priority: MEDIUM** — Add a `StatsBar.vue` component below `ChatInput.vue`:

```vue
<div class="stats-bar text-[10px] border-t px-2 py-1 flex justify-between">
  <span>{{ modelName }}</span>
  <span v-if="tokenStats">↑{{ inputTokens }} ↓{{ outputTokens }}</span>
</div>
```

This requires `useAgentLoop.ts` to capture token usage from streaming response headers/metadata and expose it as a reactive ref. Most OpenAI-compatible APIs return `usage` in the final SSE chunk.

---

## 7. FILE HANDLING & ATTACHMENTS

### OpenExcel approach

- **Drag-and-drop overlay**: Full-window overlay with backdrop blur, upload icon, "Drop files here"
- **File chips**: Each uploaded file shows as a chip with name, size, and remove button
- **VFS (Virtual File System)**: Files are stored in a per-session VFS, persisted in IndexedDB
- **Session-scoped**: Files are restored when switching back to a session
- **Agent access**: Files are injected into the prompt via `<attachments>` tags

### KickOffice approach

- **Drag-and-drop**: Supported in `ChatInput.vue` with visual feedback (border highlight)
- **File chips**: Uploaded files shown as chips with name, size, remove button
- **Backend processing**: Files sent to `/api/upload` endpoint, parsed (PDF, DOCX, XLSX, CSV), text returned
- **No persistence**: Files are not saved between sessions
- **Prompt injection**: Parsed file content added as `<attachments>` in the prompt

### Gap Analysis

| Feature | OpenExcel | KickOffice | Priority |
|---------|-----------|------------|----------|
| Drag-and-drop upload | Full overlay | Border highlight | LOW |
| File chips display | Name + size + remove | Name + size + remove | PARITY |
| File persistence per session | IndexedDB VFS | None | MEDIUM |
| File parsing (PDF, DOCX, etc.) | Client-side | Server-side | PARITY (different approach) |
| Agent file access | VFS `read` tool | `<attachments>` in prompt | PARITY |

---

## 8. CONFIGURATION & SETTINGS

### OpenExcel approach

- **Integrated settings panel**: Tab within the chat interface (chat | settings)
- **Provider selection**: Multiple providers (Anthropic, OpenAI, Google, custom)
- **Model selection**: Per-provider model lists with custom model support
- **API key management**: Per-provider with OAuth support for some providers
- **Thinking level**: Selector for extended thinking depth
- **CORS proxy**: Configurable proxy URL for browser-based API calls
- **Agent skills**: Install/uninstall marketplace for additional capabilities

### KickOffice approach

- **Separate settings page**: `SettingsPage.vue` navigated via router
- **Model tiers**: Standard / Reasoning / Image (backend-configured)
- **LiteLLM credentials**: API key + email (with remember option)
- **Custom system prompts**: Saved prompt library
- **Tool preferences**: Enable/disable individual tools per host
- **Agent config**: Max iterations slider
- **Theme**: Dark/light mode toggle
- **i18n**: Language selector (13 reply languages)

### Comparison

KickOffice actually has **richer settings** in some areas (tool preferences, saved prompts, i18n), but the settings are on a **separate page** rather than integrated into the chat. OpenExcel keeps everything in a single-pane experience.

---

## 9. MULTI-HOST VS SINGLE-HOST

### OpenExcel: Excel Only
- All tools, prompts, and UX optimized for Excel
- ~15 tools focused on cell/range/chart/object manipulation
- No need for host detection or tool filtering

### KickOffice: Word + Excel + PowerPoint + Outlook
- 116+ tools across 4 Office hosts
- Dynamic tool filtering per host
- Host-specific quick actions, prompts, and behaviors
- Host detection via `hostDetection.ts`

**This is KickOffice's strongest competitive advantage.** OpenExcel cannot interact with Word, PowerPoint, or Outlook.

---

## 10. TOOL ECOSYSTEM

| Category | OpenExcel | KickOffice |
|----------|-----------|------------|
| **Excel** | ~15 tools (cell read/write, formatting, charts) | 45 tools + OpenExcel ports |
| **Word** | None | 40 tools (full document manipulation) |
| **PowerPoint** | None | 15 tools (slides, shapes, notes) |
| **Outlook** | None | 14 tools (mail compose/read) |
| **General** | File read, Bash shell | getCurrentDate, calculateMath |
| **Dynamic exec** | None | SES sandbox (eval_officejs, eval_wordjs, etc.) |
| **File upload** | VFS with client parsing | Server-side parsing (PDF, DOCX, XLSX, CSV) |

KickOffice has **7-8x more tools** and covers the entire Office suite. OpenExcel is narrower but deeper in its Excel-specific capabilities.

---

## 11. SUMMARY COMPARISON TABLE

| Feature | OpenExcel | KickOffice | Gap Severity |
|---------|-----------|------------|-------------|
| **Chat & Agent UX** |
| Streaming text | ✓ With cursor animation | ✓ Progressive | LOW |
| Tool call status blocks | ✓ Full (pending/running/ok/error) | ✗ Not shown | **HIGH** |
| Tool call details | ✓ Expandable args/results | ✗ Not shown | **HIGH** |
| Loading indicator | ✓ "thinking..." spinner | ✓ Pulsing dot | LOW |
| Thinking blocks | ✓ Brain icon + streaming dots | ✓ Basic `<details>` | LOW-MEDIUM |
| **Conversation Management** |
| Multi-session | ✓ IndexedDB, unlimited | ✗ Single per host | **CRITICAL** |
| Session switcher | ✓ Header dropdown | ✗ None | **CRITICAL** |
| Session auto-naming | ✓ First user message | ✗ N/A | HIGH |
| New Chat behavior | ✓ Creates new session | ✗ **Overwrites** history | **CRITICAL** |
| Conversation persistence | ✓ Full (IndexedDB) | ✓ Partial (localStorage, 100 msgs) | HIGH |
| **Stats & Monitoring** |
| Token count display | ✓ Input/output/cache | ✗ None | MEDIUM |
| Cost tracking | ✓ Per-session | ✗ None | MEDIUM |
| Context usage | ✓ Percentage bar | ✗ None | MEDIUM |
| Model name display | ✓ In stats bar | ✗ Tier dropdown only | LOW |
| **Multi-Host Support** |
| Excel | ✓ | ✓ (+ OpenExcel ports) | PARITY |
| Word | ✗ | ✓ (40 tools) | **KO ADVANTAGE** |
| PowerPoint | ✗ | ✓ (15 tools) | **KO ADVANTAGE** |
| Outlook | ✗ | ✓ (14 tools) | **KO ADVANTAGE** |
| **Tool Ecosystem** |
| Total tools | ~15 | 116+ | **KO ADVANTAGE** |
| Dynamic code exec | ✗ | ✓ (SES sandbox) | **KO ADVANTAGE** |
| File processing | Client-side VFS | Server-side upload | PARITY |
| **Settings & Config** |
| Multi-provider | ✓ (Anthropic, OpenAI, Google) | ✗ (Single LiteLLM) | MEDIUM |
| Tool preferences | ✗ | ✓ (per-tool toggle) | **KO ADVANTAGE** |
| Saved prompts | ✗ | ✓ (prompt library) | **KO ADVANTAGE** |
| i18n / Languages | ✗ | ✓ (13 reply languages) | **KO ADVANTAGE** |
| Dark/light theme | ✓ | ✓ | PARITY |

---

## 12. RECOMMENDED PRIORITIES

Based on the analysis, here is the recommended implementation order for closing the gaps:

### Priority 1: CRITICAL — Conversation Management (3-5 days)

**The most impactful UX improvement.** Users currently lose all conversation history on "New Chat".

| Task | Files | Effort |
|------|-------|--------|
| Create `useSessionManager.ts` composable | New file | 1-2 days |
| Migrate storage from localStorage to IndexedDB | `HomePage.vue`, new DB module | 1 day |
| Add session switcher dropdown to `ChatHeader.vue` | `ChatHeader.vue` | 1 day |
| Add `ChatSession` interface | `types/chat.ts` | 0.5 day |
| Auto-name sessions from first user message | `useSessionManager.ts` | 0.5 day |

### Priority 2: HIGH — Tool Execution Status Blocks (2-3 days)

**Second most impactful.** Users have no visibility into what the agent is doing.

| Task | Files | Effort |
|------|-------|--------|
| Extend `DisplayMessage` with `parts` array | `types/chat.ts` | 0.5 day |
| Create `ToolCallBlock.vue` component | New file | 1 day |
| Update `useAgentLoop.ts` to emit tool-level events | `useAgentLoop.ts` | 1 day |
| Render tool call blocks in `ChatMessageList.vue` | `ChatMessageList.vue` | 0.5 day |

### Priority 3: MEDIUM — Stats Bar (1-2 days)

| Task | Files | Effort |
|------|-------|--------|
| Create `StatsBar.vue` component | New file | 0.5 day |
| Capture token usage from streaming responses | `useAgentLoop.ts`, `backend.ts` | 1 day |
| Display model name, tokens, context usage | `StatsBar.vue`, `HomePage.vue` | 0.5 day |

### Priority 4: LOW — Polish (1-2 days)

| Task | Files | Effort |
|------|-------|--------|
| Streaming cursor animation | `ChatMessageList.vue` | 0.5 day |
| Enhanced thinking block component | New `ThinkingBlock.vue` | 0.5 day |
| Settings panel integration (tab in chat) | `ChatHeader.vue`, `SettingsPage.vue` | 1 day |

---

## CONCLUSION

KickOffice's **core strength** is its multi-host Office coverage (4 apps, 116+ tools, SES sandbox, i18n) — territory where OpenExcel simply cannot compete. However, the **chat agent UX** is where OpenExcel clearly leads:

1. **Conversation management** is the most critical gap — the current "overwrite on New Chat" behavior causes data loss and frustration.
2. **Tool execution transparency** is the second biggest gap — users deserve to see what the agent is doing, not just the final output.
3. **Stats/monitoring** is a nice-to-have that builds user trust and helps debug issues.

Addressing priorities 1 and 2 (conversation management + tool status blocks) would bring KickOffice's chat UX to parity with OpenExcel while maintaining its multi-host advantage.
