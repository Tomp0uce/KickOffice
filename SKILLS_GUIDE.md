# KickOffice Skills System Guide

> **Document version**: 2.0 (2026-03-18)
> **Based on**: [Anthropic Skill Creator](https://github.com/anthropics/skills/tree/main/skills/skill-creator)

---

## Table of Contents

1. [What are Skills?](#what-are-skills)
2. [Skills in KickOffice](#skills-in-kickoffice)
3. [Skill File Format](#skill-file-format)
4. [User Skills — Creating Your Own](#user-skills--creating-your-own)
5. [Built-in Skills Reference](#built-in-skills-reference)
6. [Host Skills](#host-skills)
7. [Best Practices](#best-practices)
8. [Troubleshooting](#troubleshooting)

---

## What are Skills?

**Skills** are specialized instructions that guide Claude's behavior for specific tasks. Think of them as expert playbooks that tell Claude:
- What tools to use and in what order
- What style and format to follow
- What constraints to respect
- How to handle edge cases

Skills are defined in **Markdown files** (`.skill.md`) that contain structured instructions Claude reads before executing a task.

### Benefits of Skills
- ✅ **Consistency**: Same instructions every time
- ✅ **Customization**: Easily modify behavior without code changes
- ✅ **Transparency**: Human-readable instructions
- ✅ **Modularity**: Reusable across different contexts
- ✅ **Testing**: Iterate on prompts without rebuilding

---

## Skills in KickOffice

KickOffice uses three types of skills:

### 1. Host Skills (Chat Libre Mode)
Located in: `frontend/src/skills/`

These skills provide context for **free chat** mode (Chat Libre). They document available tools, patterns, and best practices for each Office host.

| Skill File | Purpose | Loaded When |
|------------|---------|-------------|
| `word.skill.md` | Word tool guidelines, formatting workflows | User chats in Word |
| `excel.skill.md` | Excel tool usage, formula patterns, chart extraction | User chats in Excel |
| `powerpoint.skill.md` | PowerPoint tool guidelines, slide manipulation | User chats in PowerPoint |
| `outlook.skill.md` | Outlook tool usage, email patterns | User chats in Outlook |
| `common.skill.md` | Universal guidelines (all hosts) | Always loaded |

### 2. Quick Action Skills (Built-in Workflows)
Located in: `frontend/src/skills/quickactions/`

These skills define specific, user-triggered workflows. When a user clicks a Quick Action button (like "Bullets", "Translate", "Review"), the corresponding skill is loaded to guide Claude's response.

### 3. User Skills (Custom, LLM-created)

User-created skills stored in `localStorage` as JSON (key: `ki_UserSkills_v1`). Each skill has the same `.skill.md` format as built-in skills and can be:
- **Created** via the Skill Creator modal (describe in natural language → LLM generates the skill)
- **Edited** inline in the Settings → Skills Library tab
- **Exported** as `.skill.md` for sharing with other users
- **Imported** from a `.skill.md` file
- **Executed** with one click from the Quick Actions Bar dropdown

| Skill File | Quick Action | Host | Purpose |
|------------|--------------|------|---------|
| `bullets.skill.md` | Bullets | PowerPoint | Transform text into concise bullet points |
| `punchify.skill.md` | Punchify | PowerPoint | Make text more impactful and concise |
| `review.skill.md` | Review | PowerPoint | Provide expert feedback on current slide |
| `ppt-translate.skill.md` | Translate slide | PowerPoint | Translate all text shapes on active slide |
| `ppt-proofread.skill.md` | Proofread slide | PowerPoint | Correct spelling/grammar on active slide |
| `translate.skill.md` | Translate | Word, Outlook | Translate text to target language |
| `formalize.skill.md` | Formalize | Word, Outlook | Transform casual text to professional |
| `concise.skill.md` | Concise | Word, Outlook | Condense wordy text (30-50% reduction) |
| `proofread.skill.md` | Proofread | Word | Fix spelling, grammar, punctuation errors |
| `word-proofread.skill.md` | Proofread (Track Changes) | Word | Grammar/spelling corrections via Track Changes |
| `word-translate.skill.md` | Translate (Track Changes) | Word | Translation via Track Changes, reversible |
| `word-review.skill.md` | Review doc | Word | Document review with suggestions via Track Changes |
| `polish.skill.md` | Polish | Word | General text quality improvement |
| `academic.skill.md` | Academic | Word | Academic/formal writing style |
| `summary.skill.md` | Summary | Word | Summarize document content |
| `ingest.skill.md` | Clean (Ingest) | Excel | Normalize and clean imported data |
| `autograph.skill.md` | Beautify (Autograph) | Excel | Format spreadsheet for presentation |
| `explain-excel.skill.md` | Explain Excel | Excel | Explain formulas and worksheet structure |
| `formula-generator.skill.md` | Formula Generator | Excel | Generate Excel formulas from description |
| `data-trend.skill.md` | Data Trend | Excel | Analyze trends in data |
| `chart-digitizer.skill.md` | Chart Digitizer | Excel | Extract chart data from image and recreate in Excel |
| `pixel-art.skill.md` | Pixel Art | Excel | Convert image to pixel art using cell colors |
| `extract.skill.md` | Extract Tasks | Outlook | Extract action items from email |
| `reply.skill.md` | Smart Reply | Outlook | Generate contextual email reply |

---

## Skill File Format

All skills (built-in and user-created) share the same `.skill.md` format: a YAML frontmatter block followed by free-form Markdown instructions.

### YAML Frontmatter (required)

```markdown
---
name: Reformuler en bullets
description: "Transforme le texte sélectionné en 3-7 bullet points concis et percutants pour PowerPoint. Idéal pour convertir des paragraphes denses en points mémorables."
host: powerpoint
executionMode: immediate
icon: List
actionKey: bullets
---
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `name` | string | ✅ | Short label shown in UI (≤ 5 words, imperative) |
| `description` | string | ✅ | User-facing benefit description (~30 words). Shown in dropdown tooltip. |
| `host` | `word \| excel \| powerpoint \| outlook \| all` | ✅ | Contextual filter — skill only appears for this host |
| `executionMode` | `immediate \| draft \| agent` | ✅ | How the skill executes (see table below) |
| `icon` | string | ✅ | Lucide icon name (e.g., `List`, `Languages`, `Wand2`) |
| `actionKey` | string | ❌ | Built-in skills only — links to the quick action key |

### Execution Modes

| Mode | Behavior | When to use |
|------|----------|-------------|
| `immediate` | Streams response into chat | Text transformation (translate, summarize, rewrite) — no document modification |
| `draft` | Pre-fills the textarea + focus glow | User wants to review before sending (email reply, draft) |
| `agent` | Runs the full agentic loop with Office tools | Any task that modifies the document, inserts content, or requires reading context |

**Rule of thumb**: If the skill touches the document → `agent`. If it returns text in chat → `immediate`.

### Skill Body (free Markdown)

The body follows the frontmatter and is injected as the system prompt. It follows **Theory of Mind** principles (from Anthropic's skill design guidelines):

- **Explain the WHY** behind each rule, not just the WHAT
  - ❌ `CRITICAL: NEVER drop {{PRESERVE_N}} placeholders`
  - ✅ `Preserve {{PRESERVE_N}} markers — they represent embedded images; dropping them breaks the email silently`
- **Use concrete examples** (input/output pairs) rather than abstract rules
- **Structured freely** — use whatever sections serve clarity: tools, approach, edge cases, examples
- **Anticipate real edge cases** only (not hypothetical ones)

### Example: complete skill file

```markdown
---
name: Traduire le texte
description: "Traduit le texte sélectionné entre français et anglais en détectant automatiquement la langue source. Préserve les formatages gras/italique/tableaux et les placeholders d'images {{PRESERVE_N}}."
host: all
executionMode: immediate
icon: Languages
actionKey: translate
---

Traduis le texte reçu vers l'autre langue (FR ↔ EN). Détecte la langue source — ne te fie pas au tag [UI language] pour déterminer la direction.

**Pourquoi la détection automatique ?** L'utilisateur travaille souvent sur des documents multilingues. Une direction fixe casserait la moitié des cas d'usage.

## Préservation du formatage

Maintiens tous les marqueurs autour du texte traduit :
- `**gras**` → `**texte traduit**`
- `[color:#CC0000]texte[/color]` → `[color:#CC0000]texte traduit[/color]`

## Placeholders {{PRESERVE_N}}

Ces marqueurs représentent des images embarquées. Les supprimer casserait l'email de manière invisible.
Positionne-les au même endroit logique dans le texte traduit.

**Output** : retourne UNIQUEMENT le texte traduit. Aucune explication.
```

---

## User Skills — Creating Your Own

### Method 1: Skill Creator (recommended, no code)

1. Click the **`+`** button next to the skills dropdown in the Quick Actions Bar
2. Describe what you want the skill to do (natural language, any language)
3. Select the target host (Word / Excel / PowerPoint / Outlook / All)
4. Click **Generate** — the LLM creates a full `.skill.md` with appropriate tools and instructions
5. Review and edit the generated fields (name, description, execution mode, markdown body)
6. Click **Test on selection** to try it on your currently selected text before saving
7. Click **Save** — the skill appears immediately in the dropdown

### Method 2: Import a `.skill.md` file

1. Go to **Settings → Skills**
2. Click the **Import** button
3. Select a `.skill.md` file (with YAML frontmatter)

### Method 3: Write from scratch (developers)

Create a `.skill.md` file with the format from §3 and import it via the UI, or add it directly to `frontend/src/skills/quickactions/` as a built-in skill (requires code registration in `skills/index.ts`).

### Sharing skills

Export any skill as a `.skill.md` file (Settings → Skills → Export button). The exported file includes YAML frontmatter and can be imported by any KickOffice user.

### Iterating on a skill

Edit the skill inline in Settings → Skills. The markdown body is fully editable. Changes take effect immediately on the next execution.

---

## Built-in Skills Reference

### How Quick Actions Use Skills

When a user clicks a Quick Action button:

1. **Selection**: System captures the selected text (Word/PowerPoint) or email body (Outlook)
2. **Skill Loading**: Corresponding `.skill.md` file is read and injected as system context
3. **Execution**: Claude processes the text following skill instructions
4. **Insertion**: Result is inserted back into the document via `proposeRevision` (Word) or direct replacement

### Quick Action Workflow

```
User clicks "Bullets"
  → System loads `bullets.skill.md`
  → System captures selected text
  → Claude reads skill + selected text
  → Claude outputs bullet points (following skill rules)
  → System inserts result via proposeRevision
  → User sees tracked changes in document
```

### Language Handling

**Critical Rule**: Quick Action skills MUST preserve the input language.

❌ **Wrong**:
```markdown
Input (French): "Le projet a dépassé nos attentes..."
Output (English): "- Project exceeded expectations"  ← WRONG! Changed language
```

✅ **Correct**:
```markdown
Input (French): "Le projet a dépassé nos attentes..."
Output (French): "- Projet a dépassé les attentes"  ← Correct! Preserved French
```

All Quick Action skills include a **Language Preservation** section enforcing this rule.

### Rich Content Preservation (Outlook)

Outlook emails may contain embedded images represented as `{{PRESERVE_N}}` placeholders.

**Skills MUST preserve these placeholders**:
```markdown
Input: "See the chart {{PRESERVE_0}} for details."
Output: "Voir le graphique {{PRESERVE_0}} pour les détails."  ← Placeholder intact
```

See `translate.skill.md` and `formalize.skill.md` for implementation examples.

---

## Host Skills

### Word Skill (`word.skill.md`)

**Purpose**: Guide Claude in using Word tools effectively during free chat.

**Key Sections**:
- **AVAILABLE TOOLS**: Prioritized list of Word tools
- **Decision Trees**: When to use `proposeRevision` vs `insertContent` vs `searchAndFormat`
- **Formatting Workflows**: How to apply styles without breaking Track Changes
- **API Patterns**: Office.js code snippets for common operations

**Example Usage**:
```
User: "Translate this document to French"
Claude (guided by word.skill.md):
  1. Reads word.skill.md guidelines
  2. Sees proposeRevision is PREFERRED for text modification
  3. Calls proposeRevision with French translation
  4. User sees tracked changes (original → French)
```

### Excel Skill (`excel.skill.md`)

**Purpose**: Guide Claude in Excel tool usage, formula patterns, and chart data extraction.

**Key Sections**:
- **Multi-curve chart extraction**: Iterate per series with distinct colors
- **Formula language mapping**: Semicolon (`;`) for fr/de/es vs comma (`,`) for en/zh
- **Tool selection**: When to use `getWorksheetData` vs `getRangeAsCsv`

### PowerPoint Skill (`powerpoint.skill.md`)

**Purpose**: Guide Claude in slide manipulation, visual balance, and presentation structure.

**Key Sections**:
- **Slide tools**: When to use `insertTextOnSlide` vs `searchAndFormatInPresentation`
- **Visual guidelines**: Text density, bullet limits, title formulas
- **Screenshot integration**: Using `screenshotSlide` for context-aware suggestions

### Outlook Skill (`outlook.skill.md`)

**Purpose**: Guide Claude in email composition, smart replies, and rich content handling.

**Key Sections**:
- **Preservation system**: How to handle embedded images (`{{PRESERVE_N}}`)
- **Callback patterns**: Outlook uses callbacks, not async/await
- **Reply modes**: When to use smart-reply vs draft mode

### Common Skill (`common.skill.md`)

**Purpose**: Universal guidelines loaded for all hosts.

**Key Sections**:
- **General tools**: File upload, image generation, feedback
- **Conversation style**: Professional, concise, action-oriented
- **Error handling**: How to report tool failures to users

---

## Best Practices

### 1. Be Specific and Actionable
❌ Bad: "Make the text better"
✅ Good: "Reduce word count by 30-50% while preserving all key facts"

### 2. Provide Examples
Include **concrete before/after examples**. Claude learns patterns from examples.

### 3. State Constraints Clearly
- Length limits: "3-7 bullets maximum"
- Format rules: "NO em-dashes or semicolons"
- Preservation rules: "Keep `{{PRESERVE_N}}` markers unchanged"

### 4. Use Imperative Language
❌ "You should consider using proposeRevision"
✅ "Use `proposeRevision` for text modification"

### 5. Prioritize Rules
When multiple options exist, use **priority order**:
```markdown
Use in this priority order:
1. `searchAndFormat` (PREFERRED for formatting words)
2. `applyTaggedFormatting` (for complex tagged spans)
3. `formatText` (ONLY if user said "format my selection")
```

### 6. Handle Edge Cases Explicitly
Don't assume Claude knows how to handle unusual inputs. Document them:
```markdown
## Edge Cases
- **Already bullets**: Refine structure, don't just return as-is
- **No text selected**: Return error (handled by caller)
- **Mixed languages**: Translate all translatable content
```

### 7. Separate What vs How
- **Purpose/When to Use**: WHAT the skill does (for humans and Claude)
- **Tool Usage/Examples**: HOW to do it (for Claude only)

### 8. Test with Real Data
Don't just test happy paths. Try:
- Empty selections
- Very long text (1000+ words)
- Mixed formatting (bold + italic + colors)
- Multiple languages
- Edge formatting (tables, lists, images)

---

## Troubleshooting

### Claude isn't following the skill instructions

**Possible Causes**:
1. **Skill not loaded**: Check that the skill file exists and is properly registered
2. **Conflicting instructions**: Skill contradicts other system prompts
3. **Too vague**: Instructions lack specificity or examples
4. **Too complex**: Skill is too long or has too many rules

**Solutions**:
- Add more specific constraints and examples
- Simplify: Break into smaller, focused skills
- Use imperative language ("DO X" not "you might consider X")
- Test with minimal skill first, then add complexity

### Skill produces inconsistent results

**Possible Causes**:
1. **Ambiguous language**: "Make it better", "improve the text"
2. **Missing edge case handling**: Unusual inputs not documented
3. **No examples**: Claude has no pattern to follow
4. **Language detection failing**: Not preserving input language

**Solutions**:
- Add concrete examples showing desired output
- Document edge cases explicitly
- Add Language Preservation section with examples
- Use structured Output Requirements (numbered, specific)

### Tool calls failing or wrong tools called

**Possible Causes**:
1. **Missing Tool Usage section**: Claude doesn't know which tools to use
2. **Wrong tool sequence**: Tools called out of order
3. **Missing tool in definition**: Tool doesn't exist for this host

**Solutions**:
- Add explicit Tool Usage section with sequence
- Check tool availability in host's `*Tools.ts` file
- Test tool execution independently in Chat Libre

### Output format is wrong

**Possible Causes**:
1. **Vague Output Requirements**: "Return bullets" (what format? how many?)
2. **Missing format examples**: Claude guesses the structure
3. **Conflicting style rules**: Multiple sections contradict each other

**Solutions**:
- Add precise format specification (use code blocks for examples)
- Show before/after transformations
- Use "Return ONLY X, no preamble" to prevent explanations

---

## Advanced Topics

### Skill Composition

Skills can reference other skills or patterns:
```markdown
## Related Patterns
For complex formatting, combine with `searchAndFormat` workflow (see word.skill.md, Workflow C).
```

### Dynamic Skill Selection

Future enhancement: Load skills dynamically based on user intent.
```typescript
// Pseudo-code for future implementation
const intent = detectIntent(userMessage) // "translation", "formatting", "condensing"
const skill = loadSkill(intent) // translate.skill.md, formalize.skill.md, concise.skill.md
```

### Skill Versioning

Track skill changes in version control:
```markdown
> **Version**: 1.2 (2026-03-14)
> **Changelog**: Added edge case for mixed languages
```

### A/B Testing Skills

Test skill variations by creating alternate versions:
```
bullets-v1.skill.md  ← Conservative (5-7 bullets)
bullets-v2.skill.md  ← Aggressive (3-5 bullets)
```

---

## Appendix: Skill Architecture

### File Structure
```
frontend/src/skills/
├── common.skill.md            # Universal guidelines (all hosts)
├── word.skill.md              # Word host guidelines
├── excel.skill.md             # Excel host guidelines
├── powerpoint.skill.md        # PowerPoint host guidelines
├── outlook.skill.md           # Outlook host guidelines
├── index.ts                   # Skill loader/registry
└── quickactions/              # Quick Action skills (24 files)
    ├── bullets.skill.md       # PowerPoint — concise bullets
    ├── punchify.skill.md      # PowerPoint — make text impactful
    ├── review.skill.md        # PowerPoint — slide feedback
    ├── ppt-translate.skill.md # PowerPoint — translate active slide
    ├── ppt-proofread.skill.md # PowerPoint — proofread active slide
    ├── translate.skill.md     # Word, Outlook — language translation
    ├── formalize.skill.md     # Word, Outlook — casual → professional
    ├── concise.skill.md       # Word, Outlook — word reduction
    ├── proofread.skill.md     # Word — spelling/grammar
    ├── word-proofread.skill.md # Word — spelling/grammar via Track Changes
    ├── word-translate.skill.md # Word — translation via Track Changes
    ├── word-review.skill.md   # Word — document review via Track Changes
    ├── polish.skill.md        # Word — quality improvement
    ├── academic.skill.md      # Word — academic style
    ├── summary.skill.md       # Word — summarization
    ├── ingest.skill.md        # Excel — data cleaning
    ├── autograph.skill.md     # Excel — formatting
    ├── explain-excel.skill.md # Excel — formula explanation
    ├── formula-generator.skill.md  # Excel — formula generation
    ├── data-trend.skill.md    # Excel — trend analysis
    ├── chart-digitizer.skill.md    # Excel — chart image → data + chart
    ├── pixel-art.skill.md     # Excel — image → cell pixel art
    ├── extract.skill.md       # Outlook — task extraction
    └── reply.skill.md         # Outlook — smart reply
```

### Skill Loading Flow
```
1. User clicks Quick Action "Bullets"
2. System looks for `frontend/src/skills/quickactions/bullets.skill.md`
3. File content is read and injected as system message
4. Claude receives:
   - System: [bullets.skill.md content]
   - User: [selected text to transform]
5. Claude outputs following skill guidelines
6. System inserts result via proposeRevision
```

---

## Resources

- **Anthropic's Skill Guide**: [Complete Guide to Building Skills for Claude](https://www.anthropic.com/index/building-effective-agents)
- **Claude API Documentation**: [https://docs.anthropic.com/](https://docs.anthropic.com/)
- **KickOffice Source**: `frontend/src/skills/` directory
- **Example Skills**: See `quickactions/` subdirectory

---

**Questions or need help?**
Check existing skill files for patterns or open an issue in the repository.
