# Proofread Quick Action Skill

## Purpose
Identify and correct spelling, grammar, punctuation, and style errors while preserving the author's voice and intent.

## When to Use
- User clicks "Proofread" Quick Action in Word or Outlook
- Text may contain errors, typos, or style inconsistencies
- Goal: Error-free, polished text

## Input Contract
- **Selected text**: Content to proofread (may contain errors)
- **Language**: Auto-detect and preserve original language
- **Context**: Any written content
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders (Outlook)

## Output Requirements
1. **Fix ALL errors**: Spelling, grammar, punctuation, capitalization
2. **Preserve author's voice**: Don't rewrite, just correct
3. **Maintain structure**: Keep original formatting and organization
4. **Keep intent**: Don't change meaning or tone
5. **Return corrected text**: No explanations, just the clean version
6. **Keep placeholders**: `{{PRESERVE_N}}` markers unchanged

## Error Categories

### 1. Spelling Errors
- Typos: "teh" → "the", "recieve" → "receive"
- Homophones: "there/their/they're", "your/you're", "its/it's"
- Commonly misspelled: "accommodate", "occurred", "separate"

### 2. Grammar Errors
- **Subject-verb agreement**: "The team are" → "The team is"
- **Tense consistency**: Maintain consistent tense throughout
- **Pronoun agreement**: "Everyone should bring their" (acceptable) vs formal alternatives
- **Run-on sentences**: Add proper punctuation
- **Sentence fragments**: Complete incomplete sentences (unless intentional style)

### 3. Punctuation Errors
- **Missing commas**: In lists, after introductory phrases, between clauses
- **Comma splices**: Don't join independent clauses with just a comma
- **Apostrophes**: "its" (possessive) vs "it's" (it is)
- **Quotation marks**: Proper placement with other punctuation
- **Hyphens/Dashes**: Compound modifiers, em-dashes for breaks

### 4. Capitalization
- **Proper nouns**: Names, places, brands
- **Titles**: Headline style vs sentence style
- **After punctuation**: Capitalize after periods, not after commas
- **Acronyms**: Keep uppercase (CEO, NASA, FAQ)

### 5. Style & Consistency
- **Number style**: Spell out one-ten, use digits for 11+, or follow document convention
- **Date format**: Consistent throughout (March 14, 2024 vs 14/03/2024)
- **Abbreviations**: Consistent use (e.g. vs eg vs e.g.)
- **Oxford comma**: Use consistently (A, B, and C vs A, B and C)

## What NOT to Change

### Preserve Intent
- **Informal language**: If author wrote casually, keep it casual
- **Technical terms**: Don't "correct" specialized vocabulary
- **Brand names**: Keep unconventional capitalizations (iPhone, eBay)
- **Intentional fragments**: Short sentences for emphasis are okay
- **Style choices**: Author's voice > strict rules

### Regional Variations
- **British vs American English**: Be consistent within document
  - colour/color, realise/realize, centre/center
  - organisation/organization
- **French variations**: Quebec vs France spellings

## Tool Usage
**DO NOT** call Office.js tools. Return pure corrected text.

## Example Corrections

### Example 1: Multiple Error Types
**Before**:
```
The companys new policy is affect all employee's. Their will be a meeting to discuss this on tuesday, march 15th at 10:00 AM in the main conference room. Everyone should bring there laptop and be prepared to take notes.
```

**After**:
```
The company's new policy affects all employees. There will be a meeting to discuss this on Tuesday, March 15th at 10:00 AM in the main conference room. Everyone should bring their laptop and be prepared to take notes.
```

**Corrections made**:
- "companys" → "company's" (apostrophe for possessive)
- "affect" → "affects" (subject-verb agreement)
- "employee's" → "employees" (plural, not possessive)
- "Their" → "There" (correct homophone)
- "tuesday, march 15th" → "Tuesday, March 15th" (capitalize days/months)
- "there" → "their" (possessive pronoun)

### Example 2: Punctuation & Grammar
**Before**:
```
The project was completed ahead of schedule however the team encountered several challenges along the way. Including budget constraints limited resources and tight deadlines.
```

**After**:
```
The project was completed ahead of schedule; however, the team encountered several challenges along the way, including budget constraints, limited resources, and tight deadlines.
```

**Corrections made**:
- Added semicolon before "however"
- Added comma after "however"
- Fixed sentence fragment by connecting to previous sentence
- Added commas in list

### Example 3: Outlook Email with Placeholder
**Before**:
```
Hi John

Thank you for you're email. I've reviewed the document {{PRESERVE_0}} and have a few question's about the budget section. Can we setup a quick call to discus this.

Best regard's
```

**After**:
```
Hi John,

Thank you for your email. I've reviewed the document {{PRESERVE_0}} and have a few questions about the budget section. Can we set up a quick call to discuss this?

Best regards,
```

**Corrections made**:
- Added comma after greeting
- "you're" → "your" (possessive)
- "question's" → "questions" (plural, not possessive)
- "setup" → "set up" (verb phrase)
- "discus" → "discuss" (spelling)
- Added question mark
- "regard's" → "regards," (no apostrophe, add comma)

### Example 4: French Text
**Before**:
```
Bonjour,

J'ai recu votre message et je souhaite vous informez que le projet sera terminer d'ici vendredi. Merci de me tenir informé de tout changements.

Cordialement
```

**After**:
```
Bonjour,

J'ai reçu votre message et je souhaite vous informer que le projet sera terminé d'ici vendredi. Merci de me tenir informé de tout changement.

Cordialement,
```

**Corrections made**:
- "recu" → "reçu" (accent)
- "informez" → "informer" (infinitive after "vous")
- "terminer" → "terminé" (past participle with "sera")
- "changements" → "changement" (singular after "tout")
- Added comma after closing

## Preservation Rules (Outlook)
- Keep `{{PRESERVE_N}}` placeholders EXACTLY as-is
- Ensure proper punctuation around them
- Don't "correct" them

## Edge Cases
- **Already perfect**: Return unchanged (don't introduce errors!)
- **Ambiguous errors**: If unsure, leave as-is
- **Multiple valid corrections**: Choose the most conventional
- **Slang in casual context**: May keep if appropriate to tone
- **Code/technical content**: Don't "fix" syntax that's correct in that context

## Quality Metrics
A good proofread:
- ✓ Fixes all objective errors (spelling, grammar)
- ✓ Maintains author's voice
- ✓ Doesn't introduce new errors
- ✓ Respects language conventions
- ✓ Preserves intended meaning

## When to Be Conservative
- **Legal documents**: Prefer minimal changes
- **Quotes**: Never alter quoted text
- **Poetry/creative writing**: Respect artistic license
- **Technical specifications**: Precision > style
