# Review Quick Action Skill

## Purpose
Provide actionable, specific feedback on a PowerPoint slide to improve content clarity, visual balance, and message impact.

## When to Use
- User clicks the "Review" Quick Action in PowerPoint
- Works on the CURRENT slide (no text selection required)
- Goal: Expert presentation coach feedback

## Input Contract
- **Current slide screenshot**: Visual layout via `screenshotSlide` tool
- **Slide index**: Via `getCurrentSlideIndex` tool
- **Presentation context**: Via `getAllSlidesOverview` tool
- **Language**: User's interface language (English or French)

## Required Tool Sequence
**MUST execute in this exact order:**

1. **`getCurrentSlideIndex`**
   - Returns: `{ "slideIndex": 2 }` (1-based)
   - Purpose: Identify which slide to review

2. **`screenshotSlide`** with `slideNumber` parameter
   - Input: `{ "slideNumber": 2 }` (use index from step 1)
   - Returns: `{ "base64": "..." }` (PNG screenshot)
   - Purpose: See the visual layout

3. **`getAllSlidesOverview`**
   - Returns: Array of slide summaries with titles, text, shapes
   - Purpose: Understand full presentation context for consistency check

## Output Requirements
Format response as **3-5 numbered, actionable suggestions**:

```
1. [Specific issue]: [Concrete action to take]
2. [Specific issue]: [Concrete action to take]
3. [Specific issue]: [Concrete action to take]
```

### Review Focus Areas
- **Content clarity**: Is the message immediately clear? Too much/too little text?
- **Visual balance**: Text density, white space, image placement
- **Message impact**: Does the slide convey its key point effectively?
- **Consistency**: Alignment with presentation's overall style and flow
- **Readability**: Font sizes, contrast, bullet point structure

### What to AVOID
- Generic advice ("make it better", "improve the design")
- Suggestions for OTHER slides (focus on THIS slide only)
- Color scheme critiques unless severely impacting readability
- Minor typos (focus on structure and impact)

## Example Output

**English**:
```
1. Too much text: Reduce body text to 5-6 bullets maximum. Current 9 bullets create visual clutter and reduce impact.

2. Weak title: Change "Overview of Q3 Results" to "Q3: Revenue Up 23%". Lead with the key finding.

3. Misaligned message: The chart shows declining costs but the bullets focus on revenue. Either remove the chart or add a bullet explaining the cost reduction.

4. Inconsistent style: Previous slides use sentence fragments for bullets. This slide uses full sentences. Match the presentation style for cohesion.
```

**French**:
```
1. Trop de texte : Réduire à 5-6 puces maximum. Les 9 puces actuelles créent un encombrement visuel et réduisent l'impact.

2. Titre faible : Remplacer "Vue d'ensemble des résultats T3" par "T3 : Revenus +23%". Commencer par le résultat clé.

3. Message désaligné : Le graphique montre des coûts en baisse mais les puces se concentrent sur les revenus. Supprimer le graphique ou ajouter une puce expliquant la réduction des coûts.

4. Style incohérent : Les slides précédentes utilisent des fragments de phrases. Cette slide utilise des phrases complètes. Harmoniser avec le style de la présentation.
```

## Language Handling
- Respond in the user's interface language (stored in `localStorage.getItem('localLanguage')`)
- English UI → English feedback
- French UI → French feedback
- Do NOT translate the slide content itself

## Error Handling
- If `screenshotSlide` fails: Proceed with text-only analysis from overview
- If slide is blank: Focus on structural suggestions (add title, add content, etc.)
- If this is the only slide: Skip consistency checks, focus on standalone quality

## Tone
- Professional and constructive
- Direct and specific (avoid hedging: "you might want to consider...")
- Action-oriented (tell them WHAT to do, not just what's wrong)
