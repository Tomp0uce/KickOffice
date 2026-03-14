# KickOffice Skills System Guide

> **Document version**: 1.0 (2026-03-14)
> **Based on**: [Anthropic's Complete Guide to Building Skills for Claude](https://www.anthropic.com/index/building-effective-agents)

---

## Table of Contents

1. [What are Skills?](#what-are-skills)
2. [Skills in KickOffice](#skills-in-kickoffice)
3. [Skill File Format](#skill-file-format)
4. [Creating Custom Skills](#creating-custom-skills)
5. [Quick Action Skills](#quick-action-skills)
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

KickOffice uses two types of skills:

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

### 2. Quick Action Skills (Predefined Workflows)
Located in: `frontend/src/skills/quickactions/`

These skills define specific, user-triggered workflows. When a user clicks a Quick Action button (like "Bullets", "Translate", "Review"), the corresponding skill is loaded to guide Claude's response.

| Skill File | Quick Action | Host | Purpose |
|------------|--------------|------|---------|
| `bullets.skill.md` | Bullets | PowerPoint | Transform text into concise bullet points |
| `punchify.skill.md` | Punchify | PowerPoint | Make text more impactful and concise |
| `review.skill.md` | Review | PowerPoint | Provide expert feedback on current slide |
| `translate.skill.md` | Translate | Word, Outlook | Translate text to target language |
| `formalize.skill.md` | Formalize | Word, Outlook | Transform casual text to professional |
| `concise.skill.md` | Concise | Word, Outlook | Condense wordy text (30-50% reduction) |
| `proofread.skill.md` | Proofread | Word | Fix spelling, grammar, punctuation errors |
| `polish.skill.md` | Polish | Word | General text quality improvement |
| `academic.skill.md` | Academic | Word | Academic/formal writing style |
| `summary.skill.md` | Summary | Word | Summarize document content |
| `ingest.skill.md` | Clean (Ingest) | Excel | Normalize and clean imported data |
| `autograph.skill.md` | Beautify (Autograph) | Excel | Format spreadsheet for presentation |
| `explain-excel.skill.md` | Explain Excel | Excel | Explain formulas and worksheet structure |
| `formula-generator.skill.md` | Formula Generator | Excel | Generate Excel formulas from description |
| `data-trend.skill.md` | Data Trend | Excel | Analyze trends in data |
| `extract.skill.md` | Extract Tasks | Outlook | Extract action items from email |
| `reply.skill.md` | Smart Reply | Outlook | Generate contextual email reply |

---

## Skill File Format

Skills follow a structured Markdown format with these key sections:

### Required Sections

#### 1. **Title** (H1)
```markdown
# Skill Name
```

#### 2. **Purpose** (H2)
Clear one-sentence description of what the skill does.
```markdown
## Purpose
Transform selected text into concise, impactful bullet points optimized for PowerPoint presentations.
```

#### 3. **When to Use** (H2)
Specific trigger conditions for this skill.
```markdown
## When to Use
- User clicks the "Bullets" Quick Action
- Selected text contains dense paragraphs
- Goal: Presentation-ready bullet points
```

#### 4. **Input Contract** (H2)
What the skill expects as input.
```markdown
## Input Contract
- **Selected text**: The content to transform
- **Language**: Preserve the language of the original text
- **Context**: PowerPoint presentation slide
```

#### 5. **Output Requirements** (H2)
Precise output format and constraints.
```markdown
## Output Requirements
1. **Structure**: Return ONLY the bullet points, no preamble
2. **Format**: Use `- ` for main bullets, `  - ` for sub-bullets
3. **Length**: 3-7 main bullets maximum
4. **Language**: MUST match the original text language exactly
```

### Optional Sections

#### Tool Usage
Specify which Office.js tools to call (if any).
```markdown
## Tool Usage
**MUST execute in this exact order:**
1. `getCurrentSlideIndex` - Get current slide number
2. `screenshotSlide` - Capture visual layout
3. `getAllSlidesOverview` - Understand full presentation context
```

#### Style Guidelines
Specific formatting rules and patterns.
```markdown
## Style Guidelines
- **NO em-dashes (—) or semicolons (;)**
- **Active voice only**
- **Present tense when possible**
```

#### Examples
Concrete before/after transformations.
```markdown
## Example Transformation

### Input:
\`\`\`
The marketing campaign was very successful...
\`\`\`

### Output:
\`\`\`
- 45% increase in website traffic
- 30% boost in social media engagement
\`\`\`
```

#### Edge Cases
How to handle unusual inputs.
```markdown
## Edge Cases
- **Already bullets**: Optimize and refine existing structure
- **Too short (< 20 words)**: Return simplified version
```

---

## Creating Custom Skills

### Step 1: Choose Skill Type

**Quick Action Skill** (recommended for new workflows):
- Predefined, user-triggered workflows
- Consistent, repeatable outputs
- File location: `frontend/src/skills/quickactions/<name>.skill.md`

**Host Skill** (modify existing):
- General guidance for free chat
- File location: `frontend/src/skills/<host>.skill.md`

### Step 2: Define the Skill

Create a new `.skill.md` file following the format above. Use existing skills as templates:

```bash
# Copy an existing skill as a starting point
cp frontend/src/skills/quickactions/bullets.skill.md frontend/src/skills/quickactions/my-action.skill.md

# Edit the file
# Update: Purpose, When to Use, Input/Output, Examples
```

### Step 3: Register the Quick Action (if applicable)

If creating a new Quick Action, register it in `frontend/src/utils/constant.ts`:

```typescript
export const powerPointQuickActions: PowerPointQuickAction[] = [
  // ... existing actions
  {
    key: 'myAction',
    label: 'My Action',
    icon: SomeIcon,
    category: 'transform',
  },
]
```

### Step 4: Test the Skill

1. **Build**: `npm run build` (frontend)
2. **Load add-in**: Refresh your Office application
3. **Trigger**: Click the Quick Action or type in Chat Libre
4. **Verify**: Check that Claude follows the skill instructions

### Step 5: Iterate

Skills are **declarative prompts** — iterate on the Markdown content, not code:
- Too verbose? Tighten the Output Requirements
- Wrong tools called? Update the Tool Usage section
- Inconsistent results? Add more Examples and Edge Cases

---

## Quick Action Skills

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
└── quickactions/              # Quick Action skills (17 files)
    ├── bullets.skill.md       # PowerPoint — concise bullets
    ├── punchify.skill.md      # PowerPoint — make text impactful
    ├── review.skill.md        # PowerPoint — slide feedback
    ├── translate.skill.md     # Word, Outlook — language translation
    ├── formalize.skill.md     # Word, Outlook — casual → professional
    ├── concise.skill.md       # Word, Outlook — word reduction
    ├── proofread.skill.md     # Word — spelling/grammar
    ├── polish.skill.md        # Word — quality improvement
    ├── academic.skill.md      # Word — academic style
    ├── summary.skill.md       # Word — summarization
    ├── ingest.skill.md        # Excel — data cleaning
    ├── autograph.skill.md     # Excel — formatting
    ├── explain-excel.skill.md # Excel — formula explanation
    ├── formula-generator.skill.md  # Excel — formula generation
    ├── data-trend.skill.md    # Excel — trend analysis
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
