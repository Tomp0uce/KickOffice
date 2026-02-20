# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added

- **Auto-scroll to beginning of AI response**: Automatically scrolls the chat window so the beginning of the AI response remains in view when generating long messages.
- **Excel, Outlook, PowerPoint Tooltips**: Added localized definitions for quick action tooltips.

### Fixed

- **Tool state synchronization**: The interface "Settings" now correctly restricts the tools dynamically passed to the agent.
- **Streaming in Agent Loop**: Replaced synchronous `chatSync` calls with `chatStream` for immediate feedback during basic tool queries.
- **Context Management**: Implemented an intelligent pruning strategy ensuring message context doesn't exceed LLM token limits while keeping function and function call pairs together.
- **Chat History Persistence**: Historize conversations locally via `localStorage` (segmented by Office Host) to avoid data loss on taskpane closure.
- **UI UX translations/texts**: Extracted hardcoded "Thought process" string to `fr.json` & `en.json`.
- **Developer syntax**: Replaced confusing `${text}` / `${language}` syntax with standardized `[TEXT]` / `[LANGUAGE]` indicators in the built-in prompts editor in settings.
- **GitHub Action Permissions**: Grant write permissions to the bumps-version GitHub action workflow.
