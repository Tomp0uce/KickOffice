---
name: Traduire le texte
description: "Traduit le texte sélectionné entre français et anglais en détectant automatiquement la langue source. Préserve les formatages gras/italique/tableaux et les placeholders d'images {{PRESERVE_N}}."
host: all
executionMode: immediate
icon: Languages
actionKey: translate
---

# Translate Quick Action Skill

## Purpose

Translate selected text to the target language while preserving formatting, tone, and embedded content (images, placeholders).

## When to Use

- User clicks "Translate" Quick Action in Word or Outlook
- Selected text is in a different language than desired
- Goal: Accurate translation maintaining original structure

## Input Contract

- **Selected text**: Content to translate (may contain formatting, lists, tables)
- **Target language**: User's interface language (English ↔ French)
- **Context**: Word document or Outlook email body
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders for images/embedded content

## Output Requirements

1. **Translate ALL text** to the target language
2. **Preserve structure**: Bullets stay bullets, paragraphs stay paragraphs, headings stay headings
3. **Maintain tone**: Formal → formal, casual → casual
4. **Keep placeholders**: `{{PRESERVE_0}}`, `{{PRESERVE_1}}` must remain UNCHANGED
5. **No additions**: Don't add explanations, notes, or preambles

## Critical Preservation Rules (Outlook)

**IMAGES**: Text may contain placeholders like `{{PRESERVE_0}}`, `{{PRESERVE_1}}`, etc. These represent embedded images.

**YOU MUST**:

- Keep ALL `{{PRESERVE_N}}` markers EXACTLY as-is
- Do NOT translate placeholder text
- Do NOT remove or modify placeholders
- Position them in the same logical location in the translated text

Example:

```
Input (English):
"Please review the attached diagram {{PRESERVE_0}} and provide feedback."

Output (French):
"Veuillez examiner le diagramme ci-joint {{PRESERVE_0}} et fournir vos commentaires."
```

## Language Detection & Target

**Detect the language of the selected text** — do NOT rely on the `[UI language]` header for direction.

- If the text is **primarily in French** → translate to **English**
- If the text is **primarily in English** → translate to **French**
- If the text is in another language → translate to **French** (default)
- **Preserve untranslatable**: Keep proper nouns, brand names, technical terms as-is when appropriate

**OUTPUT**: Return ONLY the translated text. No explanations, no preambles, no language labels.

## Formatting Preservation

- **Bold**: `**text**` → `**texte**` — keep the `**` markers around the translated word
- **Italic**: `*text*` → `*texte*` — keep the `*` markers around the translated word
- **Underline**: `__text__` → `__texte__` — keep the `__` markers
- **Color**: `[color:#CC0000]text[/color]` → `[color:#CC0000]texte[/color]` — keep the `[color:...]...[/color]` tags and translate only the inner text
- **Lists**: Maintain bullet/number structure
- **Tables**: Translate cell content, keep table structure
- **Links**: Translate link text, keep URL unchanged

**CRITICAL**: Every formatting marker (`**`, `*`, `__`, `[color:...]...[/color]`) that wraps a word in the input MUST wrap the corresponding translated word in the output. Never drop or strip these markers.

## Tone Matching

Match the formality level of the source:

**Formal business** (contracts, official docs):

```
English: "We hereby acknowledge receipt of your correspondence..."
French: "Nous accusons réception de votre correspondance..."
```

**Casual** (internal emails):

```
English: "Hey team, quick update on the project..."
French: "Salut l'équipe, petit update sur le projet..."
```

**Technical** (documentation):

```
English: "Initialize the configuration parameters before executing..."
French: "Initialiser les paramètres de configuration avant d'exécuter..."
```

## Tool Usage

**DO NOT** call Office.js tools. Return pure translated text that will be inserted via `proposeRevision` (Word) or direct replacement (Outlook).

## Example Transformations

### Example 1: Business Email with Image

**Input (French)**:

```
Bonjour,

Voici le rapport trimestriel {{PRESERVE_0}} pour votre révision.

Les points clés:
- Croissance de 15%
- Nouveaux clients: 47
- Satisfaction: 92%

Cordialement,
Marie
```

**Output (English)**:

```
Hello,

Here is the quarterly report {{PRESERVE_0}} for your review.

Key points:
- 15% growth
- New clients: 47
- Satisfaction: 92%

Best regards,
Marie
```

### Example 2: Word Document Paragraph

**Input (English)**:

```
The company has implemented a new remote work policy effective immediately. All employees are required to complete the online training module by Friday. Please contact HR with any questions.
```

**Output (French)**:

```
L'entreprise a mis en place une nouvelle politique de télétravail effective immédiatement. Tous les employés doivent compléter le module de formation en ligne d'ici vendredi. Veuillez contacter les RH pour toute question.
```

## Edge Cases

- **Mixed languages**: Translate all translatable content
- **Already in target language**: Return unchanged or lightly polish if needed
- **Code snippets**: Keep code unchanged, translate only comments/explanations
- **Abbreviations**: Expand if target language convention differs (e.g., "CEO" → "PDG" in French)

## Error Handling

- If source language unclear: Proceed with best guess
- If text is empty: Return error message
- If placeholders are malformed: Preserve them as-is and warn in output
