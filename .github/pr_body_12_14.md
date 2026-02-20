### Complete overhaul of Quick Action Prompts, Dynamic Language Matching, and PPT list formatting.

**1. Image Generation Fixes (Phase 12)**

- Clicking the Image generation quick action now actively calls `getOfficeSelection` if the input is empty.
- Enforces a minimum length of 5 words (triggers UI error `imageSelectionTooShort` otherwise).

**2. PowerPoint Formatting & Tooltips (Phase 12)**

- Added missing translation keys (`pptPunchify_tooltip`, etc.) for PPT generic quick actions in EN/FR.
- Upgraded `insertIntoPowerPoint` fallback to cast markdown output into `Office.CoercionType.Html` via `renderOfficeCommonApiHtml`. This physically injects `<ul><li>` preventing PPT from destroying AI bullet structures during insertion.

**3. Complete Review of Quick Action Prompts (Phase 13)**

- **Extract (Outlook)**: Stripped unneeded requests for 'Summary'. Now exclusively returns actionable tasks, owners, and deadlines.
- **Concise (Outlook)**: Removed arbitrary '30-50%' math boundaries favoring formatting directives ('eliminate corporate fluff, use bullet points').
- **Analyze (Excel)**: Prevented the AI from calculating math/statistics blindly, focusing instead on extracting structural patterns and data trends.
- **Speaker Notes (PPT)**: Lowered limit to <100 words per slide for snappy presenter readability.
- **Rewrite/Summary**: Banned overly complex 'AI' phrasing, making instructions strictly conversational and adaptative in length.

**4. Dynamic Output Language Matching (Phase 14)**

- Replaced the hardcoded UI binding `in ${language}` within `constant.ts` text-transformation prompts.
- Injected a strict system check: `Analyze the language of the provided text. You MUST respond in the exact SAME language as the original text, disregarding any other UI language preferences.`
