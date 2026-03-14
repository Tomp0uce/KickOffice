# Academic Quick Action Skill

## Purpose
Transform casual or general writing into formal academic style suitable for research papers, theses, dissertations, and scholarly publications.

## When to Use
- User clicks "Academic" Quick Action in Word
- Text needs to meet academic writing standards
- Goal: Scholarly tone, precise language, formal structure

## Input Contract
- **Selected text**: Content to academicize (may be informal notes, draft sections, or partially formal text)
- **Language**: Preserve the language of the original text
- **Context**: Academic writing (research papers, literature reviews, methodology sections, abstracts)

## Output Requirements
1. **Formal academic tone**: Eliminate casual language, contractions, colloquialisms
2. **Precise terminology**: Use discipline-appropriate scholarly vocabulary
3. **Objective voice**: Remove first-person opinions, use third-person or passive voice appropriately
4. **Structured argumentation**: Present claims with evidence and logical flow
5. **Citation placeholders**: Insert [Author, Year] placeholders where claims need citations
6. **Return academic text**: No explanations, just the transformed version

## Academic Writing Principles

### 1. Tone and Voice
- **Remove contractions**: "don't" → "do not", "it's" → "it is"
- **Eliminate colloquialisms**: "a lot of" → "numerous", "pretty much" → "essentially"
- **Avoid casual expressions**: "thing", "stuff", "kind of", "sort of"
- **Third-person preference**: "We found" → "The findings indicate" (discipline-dependent)
- **Objective stance**: "I think" → "This analysis suggests", "obviously" → "evidently"

### 2. Vocabulary Elevation
| Casual → Academic |
|-------------------|
| "show" → "demonstrate/illustrate/indicate" |
| "use" → "employ/utilize/implement" |
| "look at" → "examine/investigate/analyze" |
| "find out" → "ascertain/determine/establish" |
| "talk about" → "discuss/address/explore" |
| "point out" → "highlight/emphasize/underscore" |
| "make better" → "enhance/improve/optimize" |

### 3. Sentence Structure
- **Complex sentences**: Combine simple sentences with subordinate clauses
- **Nominalizations**: "The economy grew rapidly" → "Rapid economic growth occurred"
- **Passive voice (when appropriate)**: "We tested the hypothesis" → "The hypothesis was tested"
- **Hedging language**: "This proves" → "This suggests/indicates", "always" → "typically/generally"
- **Signposting**: "First", "Furthermore", "Consequently", "In conclusion"

### 4. Citation Integration
Insert citation placeholders where:
- Factual claims require sources
- Definitions from literature are used
- Previous research is referenced
- Theoretical frameworks are mentioned

Format: `[Author, Year]` or `[Author et al., Year]`

### 5. Disciplinary Conventions
- **Sciences**: Emphasize methodology, results, statistical significance
- **Humanities**: Focus on interpretation, textual analysis, argumentation
- **Social Sciences**: Balance empirical data with theoretical frameworks

## What NOT to Change

### Preserve Content
- **Technical terms**: Keep discipline-specific vocabulary accurate
- **Data and facts**: Don't alter numbers, dates, proper nouns
- **Existing citations**: Maintain any references already present
- **Quotes**: Never modify direct quotations
- **Methodology details**: Keep procedural descriptions precise

## Tool Usage
**DO NOT** call Office.js tools. Return pure academic text.

## Example Transformations

### Example 1: Introduction Section
**Before**:
```
Social media has become really popular over the last few years. A lot of researchers think it's affecting how young people interact with each other. This paper looks at whether using Facebook and Instagram makes teenagers feel more anxious or lonely.
```

**After**:
```
Social media platforms have experienced unprecedented growth in adoption rates over the past decade [Author, Year]. Scholarly discourse increasingly emphasizes the potential impact of these technologies on adolescent social development and interpersonal communication patterns [Author et al., Year]. This study examines the correlation between social media usage—specifically Facebook and Instagram—and self-reported anxiety and loneliness levels among adolescent populations.
```

**Transformations**:
- "really popular" → "unprecedented growth in adoption rates"
- Added citation placeholders for claims
- "A lot of researchers think" → "Scholarly discourse increasingly emphasizes"
- "young people" → "adolescent populations"
- "looks at" → "examines"
- "whether...makes teenagers feel" → "correlation between...and self-reported...levels"

### Example 2: Methodology Section
**Before**:
```
We got 200 students to fill out a survey about their social media use. We asked them questions about how often they use it and if they feel stressed. Then we looked at the data to see if there was a connection.
```

**After**:
```
A cohort of 200 undergraduate students (n=200) participated in a cross-sectional survey examining social media consumption patterns and associated psychological outcomes. Participants completed a standardized questionnaire assessing frequency of platform engagement and self-reported stress indicators. Statistical analysis employed correlation coefficients to determine the strength and significance of associations between variables (α=0.05).
```

**Transformations**:
- "We got 200 students" → "A cohort of 200 undergraduate students participated"
- "fill out a survey" → "completed a standardized questionnaire"
- "asked them questions" → "assessing"
- "how often they use it" → "frequency of platform engagement"
- "if they feel stressed" → "self-reported stress indicators"
- "looked at the data" → "Statistical analysis employed"
- "see if there was a connection" → "determine the strength and significance of associations"

### Example 3: Results Section
**Before**:
```
We found out that students who spend more time on social media tend to be more anxious. The numbers show this pretty clearly. About 65% of heavy users said they felt anxious often, compared to only 30% of light users.
```

**After**:
```
The analysis revealed a statistically significant positive correlation between social media usage duration and anxiety levels (r=0.58, p<0.001). Heavy users (>3 hours daily) reported elevated anxiety with notably higher frequency (65%) than light users (<1 hour daily, 30%). This disparity suggests a dose-dependent relationship between platform engagement and psychological distress [Author, Year].
```

**Transformations**:
- "We found out" → "The analysis revealed"
- Added statistical details (r, p-value)
- "tend to be" → "positive correlation between"
- "The numbers show this pretty clearly" → removed, integrated into sentence structure
- Defined user categories precisely (>3 hours, <1 hour)
- "said they felt" → "reported"
- Added interpretation and citation placeholder

### Example 4: Discussion Section
**Before**:
```
These results are important because they show that social media might be bad for mental health. Other studies have found similar things. We think schools and parents should pay attention to this issue and maybe limit how much time kids spend online.
```

**After**:
```
These findings hold significant implications for understanding the nexus between digital technology and adolescent psychological wellbeing. The observed correlations align with previous research demonstrating adverse mental health outcomes associated with excessive social media engagement [Author, Year; Author et al., Year]. The evidence suggests that educational institutions and parental stakeholders should consider implementing evidence-based interventions to regulate adolescent screen time and promote digital literacy [Author, Year].
```

**Transformations**:
- "are important" → "hold significant implications"
- "show that" → "for understanding"
- "might be bad for" → "adverse...outcomes associated with"
- "Other studies have found similar things" → "align with previous research" + citations
- "We think" → "The evidence suggests"
- "schools and parents" → "educational institutions and parental stakeholders"
- "pay attention to this issue" → "consider implementing evidence-based interventions"
- "maybe limit" → "regulate"
- "kids" → "adolescent"

### Example 5: French Academic Text
**Before**:
```
On a remarqué que les étudiants qui utilisent beaucoup les réseaux sociaux ont plus de problèmes de concentration. C'est probablement à cause des notifications qui les dérangent tout le temps.
```

**After**:
```
L'analyse a mis en évidence une corrélation significative entre l'utilisation intensive des plateformes de médias sociaux et les déficits attentionnels observés chez les étudiants universitaires [Auteur, Année]. Cette association pourrait s'expliquer par l'effet perturbateur des notifications récurrentes, lesquelles fragmentent les périodes de concentration soutenue et compromettent l'efficacité cognitive [Auteur et al., Année].
```

## Academic Style Guidelines

### Hedging Language (Appropriate Caution)
- **Avoid overstatement**: "This proves" → "This suggests/indicates/supports"
- **Use modals**: "will" → "may/might/could"
- **Qualifiers**: "This always" → "This typically/generally/often"
- **Tentative verbs**: "shows" → "appears to show/seems to indicate"

### Signposting and Transitions
- **Introduction**: "This paper examines...", "The primary objective is..."
- **Literature review**: "Previous research indicates...", "Conversely, Author (Year) argues..."
- **Methodology**: "The study employed...", "Data were collected via..."
- **Results**: "The findings reveal...", "As illustrated in Table 1..."
- **Discussion**: "These results suggest...", "A potential limitation is..."
- **Conclusion**: "In summary...", "Future research should address..."

### Formal Structures
- **Definition**: "X is defined as..."
- **Classification**: "X can be categorized into..."
- **Comparison**: "While X demonstrates..., Y exhibits..."
- **Causation**: "X contributes to Y", "X is attributed to..."
- **Evidence**: "This is evidenced by...", "As demonstrated by..."

## Edge Cases
- **Already academic**: Refine further, add citation placeholders if missing
- **Technical content**: Maintain precision, don't over-formalize accurate terminology
- **Quotes within text**: Keep quotes verbatim, academicize surrounding text
- **First-person acceptable**: In some disciplines (qualitative research, reflective essays), first-person is standard—preserve if appropriate
- **Data tables/figures**: Don't modify, but ensure textual references are formal

## Quality Check
After academicizing, verify:
- ✓ Eliminated all casual language?
- ✓ Objective and formal tone throughout?
- ✓ Citation placeholders where needed?
- ✓ Discipline-appropriate vocabulary?
- ✓ No factual distortions?

## Academic vs Other Actions
- **Academic** = formal scholarly tone with citations
- **Polish** = refine quality, not necessarily formal
- **Formalize** = professional business tone, not scholarly
- **Proofread** = fix errors, don't change style
