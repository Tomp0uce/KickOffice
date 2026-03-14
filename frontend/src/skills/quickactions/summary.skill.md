# Summary Quick Action Skill

## Purpose
Distill long documents or sections into concise summaries that capture key points, main arguments, and essential information.

## When to Use
- User clicks "Summary" Quick Action in Word
- Text is lengthy and needs a brief overview
- Goal: Quick-read version capturing core content

## Input Contract
- **Selected text**: Long-form content to summarize (articles, reports, chapters, meeting notes)
- **Language**: Preserve the language of the original text
- **Context**: Any document requiring condensation

## Output Requirements
1. **Capture key points**: Include all main ideas, arguments, and conclusions
2. **Preserve hierarchy**: Maintain the logical structure and flow
3. **Be concise**: Target 15-25% of original length
4. **Remain objective**: Don't interpret or editorialize
5. **Use clear structure**: Organize with headings or bullets when appropriate
6. **Return summary only**: No meta-commentary like "This document discusses..."

## Summary Techniques

### 1. Identify Core Content
- **Main thesis/argument**: What is the central claim?
- **Key supporting points**: What evidence supports the thesis?
- **Conclusions/outcomes**: What are the takeaways?
- **Critical data**: Include essential numbers, dates, names

### 2. Discard Non-Essential
- **Examples (unless critical)**: Remove most illustrative examples
- **Repetition**: Consolidate repeated points
- **Background context**: Reduce lengthy setup to bare minimum
- **Tangents**: Omit digressions from main topic
- **Filler language**: Remove "as mentioned earlier", "it is worth noting"

### 3. Structure Options

**For documents with clear sections:**
```
# Summary

## Section 1 Title
- Key point 1
- Key point 2

## Section 2 Title
- Key point 1
- Key point 2
```

**For narrative/argumentative text:**
```
# Summary

[Opening sentence stating main thesis]

[2-3 paragraphs covering key supporting arguments]

[Closing sentence with conclusion/outcome]
```

**For reports/research:**
```
# Executive Summary

**Purpose**: [1 sentence]
**Key Findings**:
- Finding 1
- Finding 2
- Finding 3
**Recommendations**: [1-2 sentences]
```

## Tool Usage
**DO NOT** call Office.js tools. Return pure summary text.

## Example Transformations

### Example 1: Business Report (600 words → 120 words)
**Before**:
```
The Q4 2025 sales performance exceeded projections across all regions, driven primarily by strong adoption of our new Enterprise Suite product line. North America led with 45% growth, followed by EMEA at 32% and APAC at 28%. The Enterprise Suite, launched in July, generated $12M in revenue—accounting for 38% of total Q4 sales. Customer feedback indicates high satisfaction (NPS: 72), with particular praise for the AI-powered analytics module.

However, operational challenges emerged. Supply chain disruptions delayed hardware shipments by 2-3 weeks in 60% of orders. Customer support ticket volume increased 40% due to onboarding complexity. Three major clients (Acme Corp, GlobalTech, Innovate Inc.) escalated concerns about integration with legacy systems.

Looking forward, Q1 2026 strategy focuses on: (1) streamlining onboarding via new tutorial content, (2) expanding support team capacity by 30%, (3) developing pre-built integration templates for top 5 legacy platforms. Sales pipeline remains robust with $18M in qualified leads. Risk factors include potential economic slowdown and competitive pressure from NewCompetitor's recent product launch.
```

**After**:
```
# Q4 2025 Executive Summary

**Performance**: Sales exceeded projections with 45% growth (North America), 32% (EMEA), 28% (APAC). Enterprise Suite generated $12M (38% of total Q4 sales), achieving NPS of 72.

**Challenges**: Supply chain delays affected 60% of orders (2-3 weeks). Support tickets rose 40% due to onboarding complexity. Three major clients raised integration concerns.

**Q1 2026 Strategy**: Streamline onboarding, expand support team 30%, develop legacy integration templates. Pipeline: $18M in qualified leads. Risks: economic slowdown, NewCompetitor product launch.
```

**Reduction**: 600 words → 120 words (80% reduction)

### Example 2: Academic Article (1200 words → 200 words)
**Before**:
```
[Long academic article about social media's impact on adolescent mental health, including literature review, methodology (survey of 500 students), results showing correlation between usage time and anxiety, discussion of mechanisms, limitations, and future research directions]
```

**After**:
```
# Summary

This study examines the relationship between social media usage and mental health outcomes in adolescents (n=500, ages 13-18).

**Key Findings**: A significant positive correlation (r=0.58, p<0.001) exists between daily social media usage time and self-reported anxiety levels. Heavy users (>3 hours/day, 65%) reported elevated anxiety compared to light users (<1 hour/day, 30%). Instagram and TikTok showed stronger associations than Facebook.

**Mechanisms**: The data suggest three contributing factors: (1) social comparison effects, (2) sleep disruption from evening use, and (3) FOMO (fear of missing out) driven by continuous notification exposure.

**Limitations**: Cross-sectional design prevents causal inference. Self-reported data may be subject to recall bias.

**Implications**: Results support implementing digital literacy programs in schools and evidence-based screen time guidelines for parents. Future longitudinal research should track usage patterns and mental health outcomes over time.
```

**Reduction**: 1200 words → 200 words (83% reduction)

### Example 3: Meeting Notes (800 words → 150 words)
**Before**:
```
[Detailed meeting transcript covering agenda items: project status updates from 4 teams, budget discussion, timeline adjustments, action items, attendee questions, tangential discussions]
```

**After**:
```
# Meeting Summary — Product Launch Planning
**Date**: March 14, 2026 | **Attendees**: 12

**Project Status**:
- Design: Complete, assets delivered to dev team
- Development: 85% complete, on track for April 15 release
- Marketing: Campaign materials in final review, launch event scheduled April 20
- QA: Critical bugs resolved, regression testing underway

**Budget**: Approved $50K additional allocation for post-launch support team.

**Timeline Adjustment**: Soft launch moved to April 10 (5 days earlier) to accommodate press embargo schedule.

**Action Items**:
- Sarah: Finalize vendor contracts by March 20
- Dev team: Complete regression testing by March 25
- Marketing: Submit event invitation list by March 18
- All: Attend launch rehearsal April 8, 2pm

**Next Meeting**: March 28, 10am
```

**Reduction**: 800 words → 150 words (81% reduction)

### Example 4: French Document Summary
**Before**:
```
[Long French report about renewable energy policy, covering current state, challenges, policy recommendations, economic impact analysis, environmental benefits, and implementation timeline — 1000 words]
```

**After**:
```
# Résumé — Politique Énergétique Renouvelable

**État Actuel**: Les énergies renouvelables représentent 32% de la production nationale, en hausse de 8% par rapport à 2024. L'éolien et le solaire dominent avec 65% de la capacité installée.

**Défis**: Coûts d'infrastructure élevés, intermittence de production, résistance réglementaire locale, obsolescence du réseau de distribution.

**Recommandations Politiques**:
1. Augmenter les subventions à l'installation de 15%
2. Moderniser le réseau électrique (investissement: 2,5 milliards €)
3. Assouplir les permis de construction pour projets éoliens
4. Lancer programme de formation (5000 techniciens d'ici 2028)

**Impact Économique**: Création estimée de 12 000 emplois, réduction des importations d'énergie de 20%, économies de 800M€/an d'ici 2030.

**Calendrier**: Mise en œuvre par phases — 2026-2028 (infrastructure), 2028-2030 (déploiement complet).
```

**Reduction**: 1000 words → 180 words (82% reduction)

## Summary Types by Content

### Executive Summary (Business Reports)
- **Length**: 10-15% of original
- **Focus**: Outcomes, financials, decisions, action items
- **Structure**: Purpose, Key Findings, Recommendations

### Abstract (Academic Papers)
- **Length**: 150-250 words (regardless of original length)
- **Focus**: Research question, methodology, key findings, implications
- **Structure**: Background, Methods, Results, Conclusion

### Brief (Legal/Policy Documents)
- **Length**: 15-20% of original
- **Focus**: Key provisions, obligations, deadlines, stakeholders
- **Structure**: Overview, Key Terms, Obligations, Next Steps

### Synopsis (Narrative/Articles)
- **Length**: 15-25% of original
- **Focus**: Main narrative arc, key arguments, conclusion
- **Structure**: Thesis, Supporting Points, Conclusion

## What NOT to Summarize Away

### Always Include
- **Core thesis/argument**: The main point of the document
- **Critical data**: Numbers that are central to the content
- **Decisions made**: Outcomes of meetings or proposals
- **Action items**: Tasks assigned with owners and deadlines
- **Key names**: Principal people, organizations, products
- **Deadlines**: All time-sensitive information

## Edge Cases
- **Already concise (<300 words)**: Provide brief overview (2-3 sentences) highlighting main point
- **Lists/tables**: Convert to prose summary or condensed bullet points
- **Technical content**: Preserve essential technical terms, explain if critical
- **Multiple topics**: Use section headings to organize disparate content
- **No clear structure**: Impose logical organization in summary

## Quality Check
After summarizing, verify:
- ✓ Captures all main points?
- ✓ Omits unnecessary details?
- ✓ Maintains logical flow?
- ✓ Remains objective?
- ✓ Target length achieved (15-25% of original)?

## Summary vs Other Actions
- **Summary** = condense to key points, reduce length by 75-85%
- **Concise** = tighten writing, reduce length by 30-50%
- **Bullets** = restructure into bullet format
- **Proofread** = fix errors, no length change
