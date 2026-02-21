# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added

- **Auto-scroll to beginning of AI response**: Automatically scrolls the chat window so the beginning of the AI response remains in view when generating long messages.
- **Excel, Outlook, PowerPoint Tooltips**: Added localized definitions for quick action tooltips.
- **PowerPoint Visual Quick Action automation**: Automatically triggers image model selection and generates images instantly when 5+ words are selected.
- **Extended PowerPoint Agent Skills**: Added `deleteSlide`, `getShapes`, `deleteShape`, `setShapeFill`, `moveResizeShape`, and `getAllSlidesOverview`.
- **PowerPoint Bullets Rendering**: Fixed an issue where AI-generated bullet points were inserted as raw text. They are now inserted as properly formatted HTML lists to enforce native slide bullets.
- **Translate Quick Action Unification**: Standardized the "Translate" action across Word, Outlook, and PowerPoint to automatically detect the source language (FR/EN) and translate to the other, sharing the same prompt, icon, tooltip, and menu position.

### Changed

- **DESIGN_REVIEW.md rewritten in English**: Full code audit with 35 identified issues organized by severity (3 CRITICAL, 6 HIGH, 16 MEDIUM, 10 LOW). Includes tracking matrix and priority recommendations.

### Fixed

- **Tool state synchronization**: The interface "Settings" now correctly restricts the tools dynamically passed to the agent.
- **Streaming in Agent Loop**: Replaced synchronous `chatSync` calls with `chatStream` for immediate feedback during basic tool queries.
- **Context Management**: Implemented an intelligent pruning strategy ensuring message context doesn't exceed LLM token limits while keeping function and function call pairs together.
- **Chat History Persistence**: Historize conversations locally via `localStorage` (segmented by Office Host) to avoid data loss on taskpane closure.
- **UI UX translations/texts**: Extracted hardcoded "Thought process" string to `fr.json` & `en.json`.
- **Developer syntax**: Replaced confusing `${text}` / `${language}` syntax with standardized `[TEXT]` / `[LANGUAGE]` indicators in the built-in prompts editor in settings.
- **Image Generation Prompts**: Refined prompts to prevent text rejection and improve the corporate presentation aesthetic.
- **Image Clipboard Fallback**: Fixed an issue where users couldn't copy AI-generated images in environments lacking standard clipboard support (Office Webviews).
- **GitHub Action Permissions**: Grant write permissions to the bumps-version GitHub action workflow.
