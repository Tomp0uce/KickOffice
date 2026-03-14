# Concise Quick Action Skill

## Purpose
Condense text to its essential message, removing redundancy and wordiness while preserving all key information.

## When to Use
- User clicks "Concise" Quick Action in Word or Outlook
- Text is wordy, repetitive, or unnecessarily long
- Goal: Maximum information density with minimum word count

## Input Contract
- **Selected text**: Verbose content to condense
- **Language**: Preserve the language of the original text
- **Context**: Any text-based content (emails, documents, reports)
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders (Outlook)

## Output Requirements
1. **Reduce word count**: Target 30-50% reduction when possible
2. **Preserve all facts**: No information loss
3. **Maintain structure**: Keep original format (bullets, paragraphs, lists)
4. **Keep clarity**: Never sacrifice understanding for brevity
5. **Preserve tone**: Formal stays formal, casual stays casual
6. **Keep placeholders**: `{{PRESERVE_N}}` markers unchanged

## Condensing Techniques

### 1. Remove Redundancy
- "advance planning" → "planning"
- "end result" → "result"
- "past history" → "history"
- "completely eliminate" → "eliminate"
- "absolutely essential" → "essential"

### 2. Cut Filler Phrases
| Wordy → Concise |
|-----------------|
| "in order to" → "to" |
| "due to the fact that" → "because" |
| "at this point in time" → "now" |
| "in the event that" → "if" |
| "with regard to" → "about" |
| "for the purpose of" → "to" |
| "it is important to note that" → (remove, just state the fact) |

### 3. Simplify Structure
- **Passive → Active**: "The report was written by John" → "John wrote the report"
- **Noun phrases → Verbs**: "make a decision" → "decide", "have a discussion" → "discuss"
- **Compound sentences → Simple**: Break overly complex sentences
- **Remove qualifying hedges**: "It seems that", "It appears", "arguably"

### 4. Compress Lists
- Instead of "First, ... Second, ... Third, ..." → use bullets
- Combine related points
- Remove transitional phrases between items

## What NOT to Remove
- **Specific data**: Numbers, dates, names
- **Technical terms**: Keep precise vocabulary
- **Context needed for clarity**: Don't create ambiguity
- **Legal language**: Be cautious with contracts/formal agreements
- **Tone indicators**: Keep courtesy phrases if they're essential

## Tool Usage
**DO NOT** call Office.js tools. Return pure text output.

## Example Transformations

### Example 1: Wordy Business Email
**Before** (98 words):
```
I am writing to inform you that we have completed our comprehensive analysis of the quarterly financial reports. At this point in time, I would like to bring to your attention the fact that there are several areas where we could potentially make improvements in order to increase our overall efficiency. It is important to note that these recommendations are based on our detailed review of the data. I would greatly appreciate it if you could take the time to review these findings at your earliest convenience.
```

**After** (35 words):
```
We completed the quarterly financial analysis. Several areas show potential for efficiency improvements. These recommendations are based on our detailed data review. Please review the findings at your earliest convenience.
```

### Example 2: Redundant Report Paragraph
**Before** (67 words):
```
The project team successfully completed the implementation phase ahead of the originally scheduled deadline. The team worked together collaboratively to overcome various different challenges and obstacles that arose during the course of the project. As a direct result of their combined efforts, we were able to achieve all of the goals and objectives that had been set out in the initial project plan.
```

**After** (26 words):
```
The team completed implementation ahead of schedule. By collaborating effectively to overcome challenges, they achieved all objectives outlined in the project plan.
```

### Example 3: Outlook Email with Image
**Before**:
```
Hi everyone,

I wanted to take a moment to share with you all the latest design mockups {{PRESERVE_0}} that the team has been working on. I think you'll find that they are quite interesting and show a lot of potential. It would be great if we could all try to find some time in our schedules to review them together as a group.
```

**After**:
```
Hi everyone,

Please review the latest design mockups {{PRESERVE_0}}. Let's schedule a group review.
```

### Example 4: Technical Content
**Before**:
```
In order to successfully deploy the application to the production environment, it is necessary to first complete the testing phase and then obtain approval from the QA team before we can proceed with the deployment process.
```

**After**:
```
To deploy to production: complete testing, obtain QA approval, then deploy.
```

## Language-Specific Patterns

### French Condensing
**Before**:
```
Je voulais prendre un moment pour vous informer du fait que nous avons terminé l'analyse complète et détaillée des rapports financiers du trimestre.
```

**After**:
```
Nous avons terminé l'analyse des rapports financiers trimestriels.
```

### English Condensing
- "I wanted to let you know" → "Note that" or remove entirely
- "I am of the opinion that" → "I believe" or just state the opinion
- "The reason why" → "Why" or "Because"

## Preservation Rules (Outlook)
- Keep `{{PRESERVE_N}}` placeholders EXACTLY as-is
- Position logically in condensed text
- Don't break their context

## Edge Cases
- **Already concise**: Make minimal changes, focus on small improvements
- **Legal/contract text**: Be conservative, preserve precise language
- **Very short text**: May not need condensing
- **Lists of specifics**: Don't over-compress if detail is the point

## Quality Check
After condensing, verify:
- ✓ All key facts preserved?
- ✓ Meaning unchanged?
- ✓ Still grammatically correct?
- ✓ Clarity maintained?
- ✓ Appropriate word count reduction?

## Target Audience Consideration
- **Executive summary**: Aggressive condensing (50%+ reduction)
- **Technical documentation**: Conservative (20-30% reduction), preserve precision
- **Business communication**: Moderate (30-40% reduction), keep professional tone
- **Casual emails**: Focus on removing filler, keep conversational flow
