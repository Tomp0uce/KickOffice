---
name: Polir le texte
description: "Améliore la qualité globale du texte : fluidité, style, précision lexicale. Ne change pas le sens ni la structure, seulement l'expression. Retourne le texte poli dans la même langue."
host: word
executionMode: immediate
icon: Sparkles
actionKey: polish
---

# Polish Quick Action Skill

## Purpose

Refine and elevate text quality by improving word choice, flow, clarity, and style while maintaining the author's core message and tone.

## When to Use

- User clicks "Polish" Quick Action in Word
- Text is acceptable but could be more refined, elegant, or professional
- Goal: Higher-quality writing without changing meaning or personality

## Input Contract

- **Selected text**: Content to polish (may be rough draft, informal notes, or already-decent writing)
- **Language**: Preserve the language of the original text
- **Context**: Any written content (reports, emails, articles, notes)

## Output Requirements

1. **Improve word choice**: Replace weak, vague, or repetitive words with stronger, more precise alternatives
2. **Enhance flow**: Improve sentence rhythm and transitions between ideas
3. **Increase clarity**: Make complex ideas easier to understand without dumbing down
4. **Preserve meaning**: Don't change the author's intent, facts, or core message
5. **Maintain tone**: Keep the original level of formality/casualness
6. **Return polished text**: No explanations, just the improved version

## Polishing Techniques

### 1. Word Choice Enhancement

- **Weak verbs → Strong verbs**: "make" → "create/build/forge", "get" → "obtain/acquire/secure"
- **Vague nouns → Specific nouns**: "thing" → specific term, "stuff" → concrete items
- **Overused words → Varied alternatives**: "very good" → "excellent/outstanding/exceptional"
- **Clichés → Fresh expressions**: Avoid "thinking outside the box", "at the end of the day"

### 2. Flow Improvements

- **Add transitions**: Between paragraphs and major ideas
- **Vary sentence length**: Mix short punchy sentences with longer flowing ones
- **Fix choppy rhythm**: Combine or restructure overly short consecutive sentences
- **Improve parallel structure**: Maintain consistency in lists and comparisons

### 3. Clarity Enhancements

- **Resolve ambiguity**: Make pronoun references clear
- **Reduce jargon**: Explain or replace unnecessary technical terms
- **Break complex sentences**: Split when a sentence tries to do too much
- **Add context**: Insert brief clarifications where meaning could be unclear

### 4. Style Refinement

- **Active voice preference**: Convert passive voice when it weakens the writing
- **Eliminate redundancy**: "final conclusion" → "conclusion", "future plans" → "plans"
- **Remove filler**: "in my opinion", "I think that", "it seems like"
- **Strengthen assertions**: "might be important" → "is important" (when warranted)

## What NOT to Change

### Preserve Intent

- **Technical precision**: Don't "improve" specialized vocabulary that's accurate
- **Author's voice**: Keep their personality and writing style
- **Factual content**: Don't alter data, names, dates, or specific claims
- **Intentional style**: Respect deliberate repetition, fragments, or informal choices
- **Regional variations**: Maintain British vs American English, local terminology

## Tool Usage

**DO NOT** call Office.js tools. Return pure polished text.

## Example Transformations

### Example 1: Business Report

**Before**:

```
The project was pretty successful and we got a lot of good results. The team worked really hard and managed to finish everything on time. We think this shows that the approach we used was the right one.
```

**After**:

```
The project achieved significant success, delivering strong results across all metrics. Through dedicated effort, the team completed deliverables on schedule. These outcomes validate our strategic approach and methodology.
```

**Improvements made**:

- "pretty successful" → "achieved significant success"
- "got a lot of good results" → "delivering strong results across all metrics"
- "worked really hard" → "dedicated effort"
- "managed to finish" → "completed"
- "we think" → removed (strengthened assertion)
- "was the right one" → "validate our strategic approach"

### Example 2: Academic Writing

**Before**:

```
This study looks at how social media impacts young people's mental health. We found out that using social media a lot can make anxiety worse. The results show this is a big problem that needs more research.
```

**After**:

```
This study examines social media's influence on adolescent mental health. Our findings reveal a correlation between excessive social media use and heightened anxiety levels. These results underscore the urgency of further investigation into this critical public health concern.
```

**Improvements made**:

- "looks at" → "examines"
- "young people" → "adolescent" (more precise in academic context)
- "impacts" → "influence on"
- "found out" → "findings reveal"
- "using social media a lot" → "excessive social media use"
- "make anxiety worse" → "heightened anxiety levels"
- "big problem" → "critical public health concern"

### Example 3: Email (Professional but Conversational)

**Before**:

```
Hi Sarah,

I wanted to reach out about the budget thing we talked about last week. I think we should probably schedule a meeting to go over the numbers and figure out what to do next. Let me know when you're free.
```

**After**:

```
Hi Sarah,

Following up on our budget discussion from last week—I'd like to schedule a meeting to review the figures and determine our next steps. Please let me know your availability.
```

**Improvements made**:

- "wanted to reach out about" → "Following up on"
- "the budget thing" → "our budget discussion"
- "I think we should probably" → "I'd like to" (more direct)
- "go over the numbers" → "review the figures"
- "figure out what to do next" → "determine our next steps"
- "when you're free" → "your availability"

### Example 4: French Text

**Before**:

```
Le projet a été assez bon et on a eu des résultats intéressants. L'équipe a beaucoup travaillé et a réussi à faire tout ce qu'il fallait. On pense que c'était la bonne méthode.
```

**After**:

```
Le projet s'est avéré remarquablement efficace, générant des résultats significatifs. L'équipe a démontré un engagement exemplaire en accomplissant l'ensemble des objectifs dans les délais impartis. Ces succès confirment la pertinence de notre approche méthodologique.
```

**Improvements made**:

- "assez bon" → "remarquablement efficace"
- "on a eu" → "générant"
- "intéressants" → "significatifs"
- "beaucoup travaillé" → "démontré un engagement exemplaire"
- "réussi à faire tout ce qu'il fallait" → "accomplissant l'ensemble des objectifs dans les délais impartis"
- "on pense" → removed
- "c'était la bonne méthode" → "confirment la pertinence de notre approche méthodologique"

## Edge Cases

- **Already polished**: Make subtle refinements, focus on minor improvements
- **Very informal by design**: Respect casual tone, only fix obvious issues
- **Creative writing**: Be conservative, preserve unique voice and style
- **Technical documentation**: Focus on clarity over elegance
- **Lists/bullet points**: Improve conciseness and parallel structure

## Quality Check

After polishing, verify:

- ✓ Text sounds more professional/refined?
- ✓ Meaning unchanged?
- ✓ Author's voice still recognizable?
- ✓ No new errors introduced?
- ✓ Reads smoothly?

## Polishing vs Other Actions

- **Polish** = improve quality while keeping length similar
- **Concise** = reduce length, focus on brevity
- **Formalize** = shift tone from casual to professional
- **Proofread** = fix errors, don't enhance style

Polish combines some aspects of all three but focuses on **elevation** rather than transformation.
