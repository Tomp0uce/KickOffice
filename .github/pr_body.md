## Describe your changes

This PR addresses multiple UX, UI, and Architectural technical debts:

- **Tool state synchronization**: The interface "Settings" now correctly filters the tools dynamically passed to the agent.
- **Streaming in Agent Loop**: Replaced synchronous `chatSync` calls with `chatStream` for immediate feedback during basic tool queries.
- **Context Management**: Implemented an intelligent pruning strategy ensuring message context doesn't exceed LLM token limits while keeping function and function call pairs together.
- **Chat History Persistence**: Historize conversations locally via `localStorage` (segmented by Office Host) to avoid data loss on taskpane closure.
- **UI UX translations/texts**: Extracted hardcoded "Thought process" string to `fr.json` & `en.json`.
- **Developer syntax**: Replaced confusing `${text}` / `${language}` syntax with standardized `[TEXT]` / `[LANGUAGE]` indicators in the built-in prompts editor in settings.
- **Auto-scroll**: The chat automatically scrolls to the beginning of the AI response when generating long messages.
- **Tooltips**: Added missing localized tooltips for Excel, Outlook, and PowerPoint quick actions.
- **GitHub Action Permissions**: Granted write permissions to the bump-version GitHub action workflow.
- **PowerPoint Visual Quick Action automation**: Automatically triggers image model selection and generates images instantly.
- **Extended PowerPoint Agent Skills**: Added `deleteSlide`, `getShapes`, `deleteShape`, `setShapeFill`, `moveResizeShape`, and `getAllSlidesOverview`.
- **PowerPoint Bullets Rendering**: AI-generated bullet points are now inserted as properly formatted HTML lists to enforce native slide bullets.
- **Translate Quick Action Unification**: Standardized the "Translate" action across Word, Outlook, and PowerPoint to handle bidirectional FR/EN translation automatically.
- **Image Generation and Clipboard Fixes**: Refined the visual agent prompts and fixed an issue where copied images failed in environments lacking standard clipboard support.

## Issue ticket number and link

Addressed several items in `DESIGN_REVIEW.md` and `UX_REVIEW.md`.

## Checklist before requesting a review

- [x] I have performed a self-review of my code
- [x] I have updated `README.md` and `CHANGELOG.md`
- [x] All translations are externalized via `i18n` (`t()`)
