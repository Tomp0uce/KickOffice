# Punchify Quick Action Skill

## Purpose
Make text more impactful, concise, and engaging while preserving key information. Optimize for presentation delivery and audience retention.

## When to Use
- User clicks the "Punchify" Quick Action in PowerPoint
- Text is bland, wordy, or lacks impact
- Goal: Transform into memorable, punchy statements

## Input Contract
- **Selected text**: Bullet points, paragraphs, or slide titles
- **Language**: Preserve the language of the original text
- **Context**: PowerPoint presentation (optimize for verbal delivery)

## Output Requirements
1. **Conciseness**: Reduce word count by 30-50% when possible
2. **Impact**: Use power words, active voice, concrete nouns
3. **Structure**: Maintain original format (bullets stay bullets, paragraphs stay paragraphs)
4. **Clarity**: Never sacrifice clarity for brevity
5. **Language**: MUST match the original text language exactly

## Transformation Techniques
1. **Remove redundancy**: "in order to" → "to", "at this point in time" → "now"
2. **Use strong verbs**: "make improvements to" → "improve", "provide assistance" → "help"
3. **Eliminate weak modifiers**: "very", "really", "quite", "somewhat"
4. **Choose concrete over abstract**: "customer satisfaction metrics" → "happy customers"
5. **Apply parallelism**: Ensure bullets follow consistent grammatical structure

## Style Guidelines (PPT_STYLE_RULES)
- **NO em-dashes (—) or semicolons (;)** — split into separate bullets or use commas
- **Active voice only**: "was implemented by the team" → "team implemented"
- **Present tense when possible**: More immediate and engaging
- **Numbers over words**: "three benefits" → "3 benefits"

## Tool Usage
**DO NOT** call any Office.js tools. Return pure text output that will be inserted via `proposeRevision`.

## Example Transformations

### Example 1: Wordy bullet → Punchy bullet
**Before**:
```
- We are going to make improvements to the customer experience by implementing a new feedback system
```
**After**:
```
- Improve customer experience with new feedback system
```

### Example 2: Weak paragraph → Strong paragraph
**Before**:
```
Our team has been working very hard to try to increase sales performance. We think that by focusing on customer relationships, we can really make a difference in our quarterly results.
```
**After**:
```
Focus on customer relationships to boost quarterly sales. Direct engagement drives measurable results.
```

### Example 3: Bland title → Impactful title
**Before**:
```
An Overview of Our Strategic Planning Process
```
**After**:
```
Strategy That Works: Our Planning Process
```

## Language Preservation
**CRITICAL**: Output language MUST match input language. Never translate.

French example:
```
Before: "Nous allons essayer de faire des améliorations au niveau de l'expérience client"
After: "Améliorer l'expérience client via un nouveau système de feedback"
```

## Edge Cases
- **Already punchy text**: Make minimal changes, focus on polish
- **Technical jargon**: Preserve necessary technical terms, simplify surrounding text
- **Bullet lists**: May reorder for logical flow if it improves impact
