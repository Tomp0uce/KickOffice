# Formalize Quick Action Skill

## Purpose
Transform casual or informal text into professional, business-appropriate language suitable for formal documents and communication.

## When to Use
- User clicks "Formalize" Quick Action in Word or Outlook
- Text contains casual language, slang, or conversational tone
- Goal: Professional transformation while preserving meaning

## Input Contract
- **Selected text**: Informal/casual content to formalize
- **Language**: Preserve the language of the original text
- **Context**: Business correspondence, reports, formal documents
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders (Outlook)

## Output Requirements
1. **Formal register**: Use business-appropriate vocabulary
2. **Complete sentences**: Expand fragments into full sentences
3. **Professional tone**: Remove colloquialisms, slang, emojis
4. **Preserve meaning**: Don't change the core message
5. **Maintain structure**: Keep paragraphs, lists, formatting
6. **Keep placeholders**: Preserve `{{PRESERVE_N}}` markers unchanged

## Transformation Rules

### Vocabulary Changes
| Casual → Formal |
|-----------------|
| "get" → "obtain", "receive", "acquire" |
| "really" → "significantly", "considerably" |
| "a lot of" → "numerous", "substantial", "considerable" |
| "find out" → "determine", "ascertain", "discover" |
| "talk about" → "discuss", "address", "examine" |
| "think about" → "consider", "evaluate", "assess" |
| "show" → "demonstrate", "illustrate", "indicate" |

### Grammar & Structure
- **Contractions**: "don't" → "do not", "we'll" → "we will"
- **Pronouns**: Minimize "I think", "I feel" → state facts directly
- **Passive voice**: Use when appropriate for formality
- **Hedging**: Remove excessive qualifiers ("maybe", "probably", "sort of")

### Remove Informal Elements
- Emojis: 😊 → (remove completely)
- Slang: "gonna", "wanna", "kinda" → "going to", "want to", "somewhat"
- Filler words: "like", "you know", "basically" → (remove)
- Exclamation marks: Use sparingly, prefer periods
- ALL CAPS: Convert to normal case with appropriate emphasis

## Language-Specific Patterns

### French Formalization
- "Salut" → "Bonjour"
- "Merci beaucoup" → "Je vous remercie"
- "Ok" → "D'accord" or "Entendu"
- "T'inquiète" → "Ne vous inquiétez pas"
- "Tutoiement" (tu) → "Vouvoiement" (vous) when appropriate

### English Formalization
- "Hey" → "Hello" or "Dear"
- "Thanks" → "Thank you" or "I appreciate"
- "No problem" → "You're welcome" or "Certainly"
- "ASAP" → "at your earliest convenience" or specific deadline

## Tool Usage
**DO NOT** call Office.js tools. Return pure text output.

## Example Transformations

### Example 1: Casual Email → Formal Email
**Before**:
```
Hey John,

Thanks for getting back to me! I wanted to talk about the project. We've got a couple of issues that we need to fix ASAP. Can we hop on a call tomorrow? Let me know what works for you.

Thanks!
```

**After**:
```
Dear John,

Thank you for your response. I would like to discuss the project status. We have identified several issues that require immediate attention. Would you be available for a call tomorrow? Please let me know your availability.

Best regards,
```

### Example 2: Informal Report → Formal Report
**Before**:
```
So basically, the sales numbers for Q2 are really good. We got a lot more customers than last quarter - like 45% more! The team did an awesome job with the new marketing campaign.
```

**After**:
```
The second quarter sales performance demonstrates significant growth. We achieved a 45% increase in customer acquisition compared to the previous quarter. The marketing team's campaign implementation has proven highly effective.
```

### Example 3: Outlook with Image Placeholder
**Before**:
```
Hey team!

Check out these awesome results {{PRESERVE_0}} from last week's campaign! We totally crushed it 🎉

Let's keep it up!
```

**After**:
```
Dear Team,

Please review the results {{PRESERVE_0}} from last week's campaign. The performance metrics exceeded our expectations.

I look forward to maintaining this momentum.
```

## Preservation Rules (Outlook)
- Keep `{{PRESERVE_N}}` placeholders EXACTLY as-is
- Position them logically in the formalized text
- Do not translate or modify them

## Edge Cases
- **Already formal**: Make minimal changes, focus on polish
- **Technical content**: Preserve technical terminology
- **Mixed formal/informal**: Standardize to formal throughout
- **Very short text**: May simply capitalize or add punctuation
- **Email signatures**: Keep as-is unless explicitly included in selection

## When NOT to Over-Formalize
- **Proper nouns**: Keep names unchanged
- **Established phrases**: "Thank you" is fine, don't force "I extend my gratitude"
- **Industry jargon**: Keep if appropriate for the audience
- **Cultural context**: Some industries prefer slightly less formal tone

## Tone Balance
Aim for "professional yet approachable", not "stiff and robotic". Formal doesn't mean impersonal.

**Too formal** (avoid):
```
"I hereby acknowledge receipt of your electronic correspondence dated..."
```

**Appropriately formal** (target):
```
"Thank you for your email dated..."
```
