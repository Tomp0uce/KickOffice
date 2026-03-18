---
name: Reformuler en bullets
description: "Transforme le texte sélectionné en 3-7 bullet points concis et percutants pour PowerPoint. Préserve la langue d'origine. Idéal pour convertir des paragraphes denses en points mémorables à afficher sur une slide."
host: powerpoint
executionMode: immediate
icon: List
actionKey: bullets
---

# Bullets Quick Action Skill

## Purpose

Transform selected text into concise, impactful bullet points optimized for PowerPoint presentations.

## When to Use

- User clicks the "Bullets" Quick Action in PowerPoint
- Selected text contains dense paragraphs or long-form content
- Goal: Convert to scannable, presentation-ready bullet points

## Input Contract

- **Selected text**: The content to transform (1-5 paragraphs typical)
- **Language**: Preserve the language of the original text
- **Context**: PowerPoint presentation slide

## Output Requirements

1. **Structure**: Return ONLY the bullet points, no preamble
2. **Format**: Use `- ` for main bullets, ` -` for sub-bullets
3. **Length**: 3-7 main bullets maximum
4. **Density**: 1-2 lines per bullet (8-15 words ideal)
5. **Style**: Active voice, parallel structure, no ending punctuation
6. **Language**: MUST match the original text language exactly

## Style Guidelines (PPT_STYLE_RULES)

- **NO em-dashes (—) or semicolons (;)** — use commas, periods, or split into multiple bullets
- **NO complex clauses** — one idea per bullet
- **Start with strong verbs** when possible
- **Remove filler words**: "that", "which", "in order to", etc.
- **Keep parallel structure** within bullet groups

## Tool Usage

**DO NOT** call any Office.js tools. Return pure text output that will be inserted via `proposeRevision`.

## Example Transformation

### Input (dense paragraph):

```
The marketing campaign was very successful and exceeded our initial expectations. We saw a 45% increase in website traffic, a 30% boost in social media engagement, and our email open rates improved by 25%. The team worked together effectively to deliver these results ahead of schedule.
```

### Output (bullets):

```
- 45% increase in website traffic
- 30% boost in social media engagement
- 25% improvement in email open rates
- Delivered ahead of schedule
```

## Error Handling

- If no text selected: Return error message (handled by caller)
- If text is already bullets: Optimize and refine existing structure
- If text is too short (< 20 words): Return simplified version or single bullet

## Language Preservation

**CRITICAL**: If the input is in French, respond in French. If English, respond in English. Never translate unless explicitly requested.

Example French input → French output:

```
Input: "Le projet a dépassé nos attentes avec une hausse de 45% du trafic web..."
Output:
- Hausse de 45% du trafic web
- Amélioration de 30% de l'engagement
```
