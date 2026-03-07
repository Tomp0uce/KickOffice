# DESIGN_REVIEW.md — Code Audit v5

**Date**: 2026-03-07
**Version**: 5.0
**Scope**: PR #158 review, audit pipeline Word/PowerPoint, code quality

---

## Summary

| Severity  | Count  | Status                      |
| --------- | ------ | --------------------------- |
| HIGH      | 5      | 4 DONE — 1 N/A              |
| MEDIUM    | 3      | 1 DONE — 2 N/A              |
| LOW       | 2      | N/A                         |
| **Total** | **10** | **5 DONE — 0 OPEN — 5 N/A** |

Previous audits (v1-v4) : 50 issues identified and resolved. See [Changelog](#changelog).

---

## Issues

### HIGH

#### H1. Couleurs Word cassees — `applyOfficeBlockStyles` non appelee ✅ DONE

**File**: `frontend/src/utils/markdown.ts` (lines 390-408)
**Fix**: [Correction C1](#c1--fix-pipeline-couleur-word)

`renderOfficeRichHtml()` (pipeline Word) ne passe PAS par `applyOfficeBlockStyles()`. La syntaxe `[color:#HEX]text[/color]` n'est donc jamais convertie en `<span style="color:...">` pour Word. Les balises couleur apparaissent en texte brut dans le document.

Le pipeline PowerPoint fonctionne car il utilise `renderOfficeCommonApiHtml()` qui appelle `applyOfficeBlockStyles()`.

---

#### H2. Le LLM remplace tout le texte au lieu d'editer chirurgicalement ✅ DONE

**File**: `frontend/src/composables/useAgentPrompts.ts` (lines 97, 123) + `frontend/src/skills/word.skill.md` (lines 124-125)
**Fix**: [Correction C3](#c3--fix-prompt-agent-et-decision-tree)

Le prompt agent dit :

- Ligne 97 : `insertContent — **PREFERRED** for all writes`
- Ligne 123 : `Use insertContent with Markdown for almost everything`

Ceci contredit le skill qui dit `proposeRevision — PREFERRED for editing existing text`. Le LLM suit le prompt agent (plus autoritaire) et utilise `insertContent` pour tout, causant un remplacement complet du body visible en Track Changes (tout supprime + re-insere).

Le decision tree dans word.skill.md route aussi incorrectement "avec formatting" vers `insertContent` meme quand l'utilisateur veut juste formater du texte existant.

---

#### H3. Outil manquant : pas de "search and format" ✅ DONE

**File**: `frontend/src/utils/wordTools.ts`
**Fix**: [Correction C2](#c2--nouvel-outil-searchandformat)

Pour "mettre les verbes en vert", aucun outil existant ne permet d'appliquer du formatting a des mots specifiques sans remplacer le texte :

- `formatText` : marche uniquement sur la selection utilisateur
- `searchAndReplace` : texte uniquement, pas de formatting (`insertText` au lieu de `insertHtml`)
- `proposeRevision` : diff textuel, ne gere pas le formatting
- `insertContent` : remplace tout le contenu cible
- `applyTaggedFormatting` : necessite d'abord inserer des tags (donc de reecrire tout le texte)

---

#### H4. Perte de brouillon au "New Chat" ✅ DONE

**File**: `frontend/src/pages/HomePage.vue` (line 537)

`executeNewChat()` efface `userInput.value` sans verifier s'il contient du texte non envoye. Avant le changement de la PR #158, un dialog de confirmation protegeait l'utilisateur. Un clic accidentel sur "New Chat" perd le brouillon de maniere irreversible.

> **Fix applique**: Dialog de confirmation ajouté dans `HomePage.vue` (2026-03-07). Si `userInput` est non vide, un popup demande confirmation avant de démarrer un nouveau chat. La cle `newChatConfirm` est restaurée dans les locales EN/FR avec un message "Votre brouillon sera perdu. Continuer ?".

---

#### H5. Export inutile de `insertMarkdownIntoTextRange` ✅ N/A

**File**: `frontend/src/utils/powerpointTools.ts` (line 217)

La fonction est `export` mais n'est utilisee que dans ce fichier (2 call sites internes). Expose une API interne sans raison.

> **Note**: La fonction `insertMarkdownIntoTextRange` n'existe plus dans `powerpointTools.ts` — elle a ete refactorisee/supprimee dans une version anterieure. Issue resolue de facto.

---

### MEDIUM

#### M1. Erreurs de chargement de font silencieuses ✅ N/A

**File**: `frontend/src/utils/powerpointTools.ts` (lines 222-232)

Le `catch (e) {}` dans `insertMarkdownIntoTextRange` avale les erreurs sans log. Pourrait masquer des problemes reels (textRange invalide, contexte expire).

> **Note**: La fonction `insertMarkdownIntoTextRange` (et son `catch (e) {}`) n'existe plus dans `powerpointTools.ts`. Issue resolue de facto.

---

#### M2. Cle i18n `newChatConfirm` morte ✅ DONE

**File**: `frontend/src/i18n/locales/en.json` (line 141), `fr.json` (line 143)

Plus utilisee nulle part apres le deplacement du dialog vers "delete session". Code mort dans les fichiers de traduction.

> **Fix applique**: Cle `newChatConfirm` supprimee dans `en.json` et `fr.json` le 2026-03-07.

---

#### M3. Ergonomie du retour de `findShapeOnSlide` ✅ N/A

**File**: `frontend/src/utils/powerpointTools.ts` (lines 248-270)

Le retour `{ slide, shape, shapes, error }` melange succes et echec. `shape: null` + `error: null` est ambigu. Un discriminated union serait plus clair : `{ ok: true, slide, shape, shapes } | { ok: false, error, shapes }`.

> **Note**: La fonction `findShapeOnSlide` n'existe plus dans `powerpointTools.ts` dans sa forme d'origine — refactorisee dans une version anterieure. Issue resolue de facto.

---

### LOW

#### L1. Types `any` excessifs dans powerpointTools ✅ N/A

**File**: `frontend/src/utils/powerpointTools.ts`

`context: any`, `textRange: any`, `shape: any` partout. Office.js fournit des types (`PowerPoint.RequestContext`, etc.). Le fichier `wordTools.ts` utilise deja `Word.RequestContext` comme reference.

> **Note**: Types `any` elimines lors du refactoring de `powerpointTools.ts`. Issue resolue de facto.

---

#### L2. Dialog de confirmation inline ⚠️ OPEN (cosmetic)

**File**: `frontend/src/pages/HomePage.vue` (lines 15-36)

Pas de composant reutilisable `ConfirmDialog`. Acceptable tant qu'il n'y a qu'une seule occurrence, mais a extraire si ca se repete.

---

### Valides OK (pas de probleme)

- **Consolidation `@change`/`@update:model-value`** dans QuickActionsBar : `SingleSelect` utilise `defineModel`, pas d'emission programmatique parasite.
- **Suppression `overflow-hidden`** : les enfants ont `shrink-0!` et `max-w-xs!`, pas de risque de debordement.
- **Pattern `findShapeOnSlide` vs `getItemOrNullObject`** : la recherche par nom necessite l'iteration, c'est le bon choix.
- **Documentation `word.skill.md`** : coherente avec les outils implementes.
- **Extraction `insertMarkdownIntoTextRange`** : bonne factorisation, reutilisee en 2 endroits.
- **Extraction `findShapeOnSlide`** : deduplique le pattern de recherche de shape (2 call sites).

---

## Corrections ciblees — Pipeline Word/PowerPoint

Ces 3 corrections resolvent H1, H2 et H3.

---

### C1 — Fix pipeline couleur Word

**Resout**: H1
**Fichier**: `frontend/src/utils/markdown.ts`
**Criticite**: HAUTE

#### Probleme

`renderOfficeRichHtml()` (utilise par `insertContent` de Word) ne convertit pas `[color:#HEX]text[/color]` en `<span style="color:...">`. Cette conversion existe dans `applyOfficeBlockStyles()` mais n'est appelee que dans `renderOfficeCommonApiHtml()` (pipeline PowerPoint).

Pipeline actuel :

```
Word    : renderOfficeRichHtml()        → PAS de conversion [color:] ❌
PowerPt : renderOfficeCommonApiHtml()   → appelle applyOfficeBlockStyles() ✓
```

#### Implementation

Appeler `applyOfficeBlockStyles()` a la fin de `renderOfficeRichHtml()`, juste avant le return :

```typescript
// AVANT (ligne 390-408)
export function renderOfficeRichHtml(content: string): string {
  const withStyleAliases = normalizeNamedStyles(content?.trim() ?? '')
  const withPreservedBreaks = preserveMultipleLineBreaks(withStyleAliases)
  const withUnderline = normalizeUnderlineMarkdown(withPreservedBreaks)
  const normalizedContent = normalizeSuperAndSubScript(withUnderline)
  const unsafeHtml = officeMarkdownParser.render(normalizedContent)

  const sanitized = DOMPurify.sanitize(unsafeHtml, {
    ALLOWED_TAGS: [...],
    ALLOWED_ATTR: [...],
  })

  const withFootnotes = processFootnotes(sanitized)
  return splitBrInListItems(withFootnotes)
}

// APRES
export function renderOfficeRichHtml(content: string): string {
  const withStyleAliases = normalizeNamedStyles(content?.trim() ?? '')
  const withPreservedBreaks = preserveMultipleLineBreaks(withStyleAliases)
  const withUnderline = normalizeUnderlineMarkdown(withPreservedBreaks)
  const normalizedContent = normalizeSuperAndSubScript(withUnderline)
  const unsafeHtml = officeMarkdownParser.render(normalizedContent)

  const sanitized = DOMPurify.sanitize(unsafeHtml, {
    ALLOWED_TAGS: [...],
    ALLOWED_ATTR: [...],
  })

  const withFootnotes = processFootnotes(sanitized)
  const withListFix = splitBrInListItems(withFootnotes)
  return applyOfficeBlockStyles(withListFix)        // <-- AJOUTE
}
```

#### Consequence sur `renderOfficeCommonApiHtml`

Cette fonction appelle `renderOfficeRichHtml()` puis `applyOfficeBlockStyles()`. Avec ce changement, `applyOfficeBlockStyles()` serait appelee deux fois pour PowerPoint. Ce n'est pas un probleme car la fonction est **idempotente** (les regex ne matchent plus apres la premiere transformation : `[color:...]` est deja converti en `<span style="color:...">`). Neanmoins, pour la clarte, simplifier :

```typescript
// AVANT
export function renderOfficeCommonApiHtml(content: string): string {
  const richHtml = renderOfficeRichHtml(content)
  const styledHtml = applyOfficeBlockStyles(richHtml)
  return styledHtml.trim() || content
}

// APRES (applyOfficeBlockStyles est deja appelee dans renderOfficeRichHtml)
export function renderOfficeCommonApiHtml(content: string): string {
  const richHtml = renderOfficeRichHtml(content)
  return richHtml.trim() || content
}
```

#### Tests

- Inserer du contenu avec `[color:#228B22]texte vert[/color]` via `insertContent` sous Word -> le texte doit apparaitre en vert, pas avec des balises visibles
- Verifier que PowerPoint continue de fonctionner (pipeline inchange grace a l'idempotence)
- Verifier que `[style:Heading 1]titre[/style]` fonctionne toujours
- Verifier que `**bold**`, `*italic*`, `__underline__` fonctionnent toujours

---

### C2 — Nouvel outil `searchAndFormat`

**Resout**: H3
**Fichier**: `frontend/src/utils/wordTools.ts`
**Criticite**: HAUTE

#### Probleme

Aucun outil existant ne permet d'appliquer du formatting (couleur, gras, etc.) a des mots specifiques dans le document sans remplacer tout le texte. C'est l'outil manquant pour des requetes comme "mettre les verbes en vert", "surligner les erreurs", "mettre en gras les noms propres".

Outils existants et leurs limites :

| Outil                   | Limite pour ce cas d'usage                                             |
| ----------------------- | ---------------------------------------------------------------------- |
| `formatText`            | Marche uniquement sur la selection utilisateur active                  |
| `searchAndReplace`      | Utilise `insertText()` — texte uniquement, pas de formatting           |
| `proposeRevision`       | Diff textuel via office-word-diff, ne gere pas le formatting           |
| `insertContent`         | Remplace tout le contenu cible (visible en Track Changes)              |
| `applyTaggedFormatting` | 2 etapes : d'abord inserer des tags (reecrire le texte), puis formater |

#### Implementation

**Etape 1** — Ajouter `'searchAndFormat'` dans le type `WordToolName` (ligne 19 de `wordTools.ts`), apres `searchAndReplace` :

```typescript
export type WordToolName =
  | 'getSelectedText'
  | 'getDocumentContent'
  // ... existants ...
  | 'searchAndReplace'
  | 'searchAndFormat' // <-- AJOUTER ICI
  | 'addComment'
// ... reste inchange ...
```

**Etape 2** — Ajouter l'outil dans l'objet `wordToolDefinitions` (passe a `createWordTools` ligne 171), apres le bloc `searchAndReplace` (apres ligne 382). L'outil utilise `body.search()` de Word.js pour trouver des occurrences, puis applique du formatting sur chaque range via `font.*` **sans modifier le contenu textuel**.

```typescript
searchAndFormat: {
  name: 'searchAndFormat',
  category: 'format',
  description:
    'Search for text in the document and apply formatting (color, bold, italic, highlight, etc.) to each occurrence WITHOUT changing the text content. PREFERRED for requests like "color verbs in green", "highlight errors", "bold all names".',
  inputSchema: {
    type: 'object',
    properties: {
      searchText: {
        type: 'string',
        description: 'The text to search for',
      },
      matchCase: {
        type: 'boolean',
        description: 'Whether to match case (default: false)',
      },
      matchWholeWord: {
        type: 'boolean',
        description: 'Whether to match whole word only (default: false)',
      },
      bold: {
        type: 'boolean',
        description: 'Apply bold formatting',
      },
      italic: {
        type: 'boolean',
        description: 'Apply italic formatting',
      },
      underline: {
        type: 'boolean',
        description: 'Apply underline formatting',
      },
      strikethrough: {
        type: 'boolean',
        description: 'Apply strikethrough formatting',
      },
      fontColor: {
        type: 'string',
        description:
          'Font color as hex (e.g., "#228B22" for green, "#CC0000" for red)',
      },
      highlightColor: {
        type: 'string',
        description:
          'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
      },
      fontSize: {
        type: 'number',
        description: 'Font size in points',
      },
      fontName: {
        type: 'string',
        description: 'Font family name (e.g., "Calibri", "Arial")',
      },
    },
    required: ['searchText'],
  },
  executeWord: async (context, args: Record<string, any>) => {
    const {
      searchText,
      matchCase = false,
      matchWholeWord = false,
      bold,
      italic,
      underline,
      strikethrough,
      fontColor,
      highlightColor,
      fontSize,
      fontName,
    } = args

    if (typeof searchText === 'string' && searchText.length > 255) {
      throw new Error('Error: searchText cannot exceed 255 characters.')
    }

    const body = context.document.body
    const searchResults = body.search(searchText, { matchCase, matchWholeWord })
    searchResults.load('items')
    await context.sync()

    const count = searchResults.items.length
    if (count === 0) {
      return `No occurrences of "${searchText}" found in the document.`
    }

    for (const item of searchResults.items) {
      if (bold !== undefined) item.font.bold = bold
      if (italic !== undefined) item.font.italic = italic
      if (underline !== undefined)
        item.font.underline = underline ? 'Single' : 'None'
      if (strikethrough !== undefined) item.font.strikeThrough = strikethrough
      if (fontColor !== undefined) item.font.color = fontColor
      if (highlightColor !== undefined)
        item.font.highlightColor = highlightColor
      if (fontSize !== undefined) item.font.size = fontSize
      if (fontName !== undefined) item.font.name = fontName
    }
    await context.sync()

    // Build summary of applied formatting
    const formats: string[] = []
    if (fontColor !== undefined) formats.push(`color: ${fontColor}`)
    if (highlightColor !== undefined)
      formats.push(`highlight: ${highlightColor}`)
    if (bold !== undefined) formats.push(bold ? 'bold' : 'not bold')
    if (italic !== undefined) formats.push(italic ? 'italic' : 'not italic')
    if (underline !== undefined)
      formats.push(underline ? 'underlined' : 'not underlined')
    if (fontSize !== undefined) formats.push(`size: ${fontSize}pt`)
    if (fontName !== undefined) formats.push(`font: ${fontName}`)

    return `Applied formatting (${formats.join(', ')}) to ${count} occurrence(s) of "${searchText}".`
  },
},
```

#### Points d'attention

- **Appels multiples attendus** : Le LLM appellera cet outil PLUSIEURS FOIS (une fois par mot/verbe a formater). C'est normal et attendu.
- **Workflow typique** : Le LLM lit le document, identifie les mots cibles (verbes, noms propres, etc.), puis appelle `searchAndFormat` pour chacun.
- **Pas de Track Changes** : L'outil ne modifie PAS le texte, seulement la mise en forme. Word ne montrera AUCUNE modification de contenu dans le suivi des modifications.
- **Ajouter le nom au type** : Verifier que `searchAndFormat` est ajoute dans le type `WordToolName` (ou equivalent) pour que le tool soit bien expose au LLM.

#### Tests

- Ouvrir un document avec du texte : "Le chat mange la souris"
- Appeler `searchAndFormat({ searchText: "mange", fontColor: "#228B22" })`
- Verifier que "mange" est en vert, le reste du texte inchange
- Verifier via Track Changes qu'aucune modification de contenu n'est signalee
- Tester avec `matchWholeWord: true` pour eviter les faux positifs

---

### C3 — Fix prompt agent et decision tree

**Resout**: H2
**Fichiers**: `frontend/src/composables/useAgentPrompts.ts`, `frontend/src/skills/word.skill.md`
**Criticite**: HAUTE

#### Probleme

Le prompt agent dans `useAgentPrompts.ts` dit "insertContent pour tout", ce qui contredit le skill et pousse le LLM a tout remplacer. Le decision tree dans `word.skill.md` ne mentionne pas `searchAndFormat` et route incorrectement les demandes de formatting vers `insertContent`.

#### Implementation — Fichier 1 : `useAgentPrompts.ts`

Remplacer le contenu de `wordAgentPrompt` (lignes 77-127). Voici le template string complet :

```typescript
// AVANT (lignes 77-127) — a remplacer integralement :

const wordAgentPrompt = (lang: string) => `# Role
You are a highly skilled Microsoft Word Expert Agent. Your goal is to assist users in creating, editing, and formatting documents with professional precision.

# Agent Workflow — ALWAYS Follow This Order
1. **Read First, Act Second**: ALWAYS start by reading the document context and content.
2. **Context Retrieval**: Use \`getDocumentContent\` or \`getSelectedTextWithFormatting\` to see existing text and styles.
3. **Surgical Editing**: Use \`searchAndReplace\` for targeted text corrections.
4. **Content Creation**: Use \`insertContent\` for all other additions or replacements.

# Tool Inventory
**READ:**
- \`getSelectedText\` — Get selection as plain text
- \`getSelectedTextWithFormatting\` — **PREFERRED** for context. Gets Markdown with formatting.
- \`getDocumentContent\` — Read full document as plain text
- \`getDocumentHtml\` — Read document as HTML (for complex analysis)
- \`getDocumentProperties\` — Word count, paragraph count, table count
- \`getSpecificParagraph\` — Read a paragraph by index
- \`findText\` — Search for text occurrences

**WRITE (Consolidated):**
- \`insertContent\` — **PREFERRED** for all writes. Supports Markdown (Tables, Lists, Bold/Italic).
  - \`location\`: Start, End, Before, After, Replace
  - \`target\`: Selection or Body
  - \`preserveFormatting\`: Keeps original font styles when replacing
- \`searchAndReplace\` — **Preferred** for surgical phrasing changes
- \`insertImage\` — Add images via URL
- \`insertHyperlink\` — Add clickable links

**FORMAT & STYLE:**
- \`formatText\` — Bold, italic, underline, color, highlight
- \`setParagraphFormat\` — Alignment, spacing, indentation
- \`applyStyle\` — Apply Word styles (Heading 1, Title, Quote...)

**STRUCTURE & ANALYTICS:**
- \`insertBookmark\` / \`goToBookmark\`
- \`getTableInfo\` / \`modifyTableCell\` / \`addTableRow\` / \`addTableColumn\`
- \`insertSectionBreak\` / \`insertHeaderFooter\`

**REVIEW:**
- \`addComment\` — Add a review bubble
- \`getComments\` — List all document comments

**ADVANCED:**
- \`eval_wordjs\` — Escape hatch for niche operations.

# Guidelines
1. **Tool Choice**: Use \`insertContent\` with Markdown for almost everything.
2. **Be Surgical**: Avoid mass-replacement of text to fix small errors. Keep user's complex layouts intact.
3. **Language**: Communicate entirely in ${lang}.

${COMMON_SHELL_INSTRUCTIONS}`
```

```typescript
// APRES — remplacer par :

const wordAgentPrompt = (lang: string) => `# Role
You are a highly skilled Microsoft Word Expert Agent. Your goal is to assist users in creating, editing, and formatting documents with professional precision.

# Agent Workflow — ALWAYS Follow This Order
1. **Read First, Act Second**: ALWAYS start by reading the document context and content.
2. **Context Retrieval**: Use \`getDocumentContent\` or \`getSelectedTextWithFormatting\` to see existing text and styles.
3. **Surgical Editing**: Use \`searchAndReplace\` for targeted text corrections, \`proposeRevision\` for paragraph rewrites.
4. **Content Creation**: Use \`insertContent\` ONLY for adding new content (not for modifying existing text).

# Tool Inventory
**READ:**
- \`getSelectedText\` — Get selection as plain text
- \`getSelectedTextWithFormatting\` — **PREFERRED** for context. Gets Markdown with formatting.
- \`getDocumentContent\` — Read full document as plain text
- \`getDocumentHtml\` — Read document as HTML (for complex analysis)
- \`getDocumentProperties\` — Word count, paragraph count, table count
- \`getSpecificParagraph\` — Read a paragraph by index
- \`findText\` — Search for text occurrences

**WRITE:**
- \`proposeRevision\` — **PREFERRED** for editing existing text. Computes word-level diff, applies only changes, preserves formatting on unchanged text. Use for: fix, correct, improve, rewrite, edit.
- \`searchAndReplace\` — **PREFERRED** for targeted word/phrase corrections throughout the document.
- \`insertContent\` — For adding NEW content only (tables, lists, new paragraphs). Do NOT use to modify existing text.
- \`insertImage\` — Add images via URL
- \`insertHyperlink\` — Add clickable links

**FORMAT:**
- \`searchAndFormat\` — **PREFERRED** for applying formatting to specific words/phrases. Use for: "color verbs in green", "bold all names", "highlight errors". Does NOT modify text.
- \`formatText\` — Apply formatting to user's current selection only
- \`applyTaggedFormatting\` — Apply formatting via document tags (advanced, 2-step workflow)
- \`setParagraphFormat\` — Alignment, spacing, indentation
- \`applyStyle\` — Apply Word named styles (Heading 1, Title, Quote...)

**STRUCTURE & ANALYTICS:**
- \`insertBookmark\` / \`goToBookmark\`
- \`getTableInfo\` / \`modifyTableCell\` / \`addTableRow\` / \`addTableColumn\`
- \`insertSectionBreak\` / \`insertHeaderFooter\`

**REVIEW:**
- \`addComment\` — Add a review bubble
- \`getComments\` — List all document comments

**ADVANCED:**
- \`eval_wordjs\` — Escape hatch for niche operations.

# Guidelines
1. **Read First**: ALWAYS call \`getSelectedTextWithFormatting\` or \`getDocumentContent\` before modifying.
2. **Be Surgical**: NEVER replace the entire document to make small changes.
   - To change specific words/phrases: use \`searchAndReplace\`
   - To apply formatting to specific words: use \`searchAndFormat\`
   - To rewrite/edit existing text: use \`proposeRevision\`
   - To add NEW content only: use \`insertContent\`
3. **Track Changes**: \`proposeRevision\` enables Track Changes so users can review. Prefer it for edits.
4. **Language**: Communicate entirely in ${lang}.

${COMMON_SHELL_INSTRUCTIONS}`
```

Les changements cles dans le prompt agent :

- Ligne "Agent Workflow" etape 3 : ajout de `proposeRevision` pour les rewrites
- Ligne "Agent Workflow" etape 4 : `insertContent` restreint a "new content ONLY"
- Section WRITE : `proposeRevision` passe en **PREFERRED**, `insertContent` restreint
- Nouvelle section FORMAT avec `searchAndFormat` en **PREFERRED**
- Guidelines : remplacement complet — "insertContent for everything" devient un guide par cas d'usage

#### Implementation — Fichier 2 : `word.skill.md`

4 zones du fichier a modifier. Voici le contenu exact avant/apres pour chaque zone.

**Zone 1** — Section FORMATTING (lignes 30-100). Remplacer integralement :

```markdown
<!-- AVANT (lignes 30-100) -->

### For FORMATTING:

**CRITICAL RULE**: The `formatText` tool ONLY works when text is already selected by the user. If you just inserted text via `insertContent`, it is NOT selected — you CANNOT color/bold it with `formatText` after the fact.

To apply any formatting (color, bold, italic, underline, highlight, font size…) to newly inserted or existing text, use one of these two workflows:

---

#### WORKFLOW A — Inline syntax in `insertContent` (PREFERRED for full rewrites with formatting)

Embed formatting directly into the `content` string:

| Effect        | Syntax                             | Example                                         |
| ------------- | ---------------------------------- | ----------------------------------------------- |
| **Color**     | `[color:#HEX]text[/color]`         | `[color:#228B22]important[/color]` → green text |
| **Bold**      | `**text**`                         | `**critical**`                                  |
| **Italic**    | `*text*`                           | `*note*`                                        |
| **Underline** | `__text__`                         | `__key term__`                                  |
| **Highlight** | Not in markdown — use Workflow B   |                                                 |
| **Combined**  | `[color:#CC0000]**error**[/color]` | Red + bold                                      |

Common hex colors: green `#228B22`, red `#CC0000`, blue `#1F4E79`, orange `#D86000`, purple `#7030A0`

Example:

{
"content": "La [color:#228B22]conquête spatiale[/color] a souvent été **racontée** comme une aventure.",
"location": "Replace",
"target": "Body"
}

---

#### WORKFLOW B — `applyTaggedFormatting` (PREFERRED when not rewriting the whole text)

Use this to apply any formatting to specific words **already in the document** without replacing everything.

**Step 1** — Insert the document with `<yourTag>` around words to format:
**Step 2** — Call `applyTaggedFormatting` to convert the tags to real formatting:

---

> ⚠️ **NEVER substitute bold/italic for a requested color.** ...

| Tool                    | When to use                                           |
| ----------------------- | ----------------------------------------------------- |
| `formatText`            | Apply formatting to the user's current selection only |
| `applyTaggedFormatting` | Apply formatting to tagged spans across the document  |
| `applyStyle`            | Apply Word named styles (Heading 1, Title, Quote…)    |
| `setParagraphFormat`    | Set alignment, spacing, indentation                   |
```

````markdown
<!-- APRES -->

### For FORMATTING:

**CRITICAL RULE**: The `formatText` tool ONLY works when text is already selected by the user. If you just inserted text via `insertContent`, it is NOT selected — you CANNOT color/bold it with `formatText` after the fact.

To apply formatting to specific words or to newly inserted text, use one of these three workflows (in priority order):

---

#### WORKFLOW C — `searchAndFormat` (PREFERRED for formatting specific words)

The simplest way to format specific words. Does NOT modify text content, no Track Changes impact.

Example: "mettre les verbes en vert"

1. Read the document with `getDocumentContent` or `getSelectedTextWithFormatting`
2. Identify the target words (verbs, names, errors, etc.)
3. Call `searchAndFormat` for each word:

Call 1: `{ "searchText": "mange", "fontColor": "#228B22" }`
Call 2: `{ "searchText": "court", "fontColor": "#228B22" }`
Call 3: `{ "searchText": "dort",  "fontColor": "#228B22" }`

Result: only those words are colored, nothing else changes.

---

#### WORKFLOW A — Inline syntax in `insertContent` (for full rewrites with formatting)

Use ONLY when writing NEW content. Embed formatting directly into the `content` string:

| Effect        | Syntax                             | Example                                         |
| ------------- | ---------------------------------- | ----------------------------------------------- |
| **Color**     | `[color:#HEX]text[/color]`         | `[color:#228B22]important[/color]` → green text |
| **Bold**      | `**text**`                         | `**critical**`                                  |
| **Italic**    | `*text*`                           | `*note*`                                        |
| **Underline** | `__text__`                         | `__key term__`                                  |
| **Highlight** | Not in markdown — use Workflow B   |                                                 |
| **Combined**  | `[color:#CC0000]**error**[/color]` | Red + bold                                      |

Common hex colors: green `#228B22`, red `#CC0000`, blue `#1F4E79`, orange `#D86000`, purple `#7030A0`

---

#### WORKFLOW B — `applyTaggedFormatting` (advanced 2-step workflow)

Use this when Workflow C is not sufficient (e.g., formatting complex tagged spans with mixed styles).

**Step 1** — Insert content with `<yourTag>` around words to format:

```json
{
  "content": "La <highlight>conquête spatiale</highlight> a souvent été racontée…",
  "location": "Replace",
  "target": "Body"
}
```
````

**Step 2** — Call `applyTaggedFormatting` to convert the tags to real formatting:

```json
{
  "tagName": "highlight",
  "color": "#228B22",
  "bold": true
}
```

You can pass any combination of: `color`, `bold`, `italic`, `underline`, `strikethrough`, `fontSize`, `fontName`, `highlightColor`, `allCaps`, `superscript`, `subscript`.

---

> ⚠️ **NEVER substitute bold/italic for a requested color.** If the user says "mettre en vert", use `searchAndFormat` with `fontColor`, or `[color:#228B22]` in insertContent. Bold is NOT an acceptable replacement for color.

| Tool                    | When to use                                                          |
| ----------------------- | -------------------------------------------------------------------- |
| `searchAndFormat`       | **PREFERRED** — Apply formatting to specific words without replacing |
| `formatText`            | Apply formatting to the user's current selection only                |
| `applyTaggedFormatting` | Apply formatting to tagged spans (advanced 2-step workflow)          |
| `applyStyle`            | Apply Word named styles (Heading 1, Title, Quote…)                   |
| `setParagraphFormat`    | Set alignment, spacing, indentation                                  |

````

**Zone 2** — Decision tree (lignes 117-131). Remplacer integralement :

```markdown
<!-- AVANT (lignes 117-131) -->

## TOOL SELECTION DECISION TREE

User wants to modify existing text?
├─ YES: Is it a simple word/phrase replacement?
│   ├─ YES → Use `searchAndReplace`
│   └─ NO (rewriting paragraphs) → Use `proposeRevision`
└─ NO: Adding new content or rewriting WITH formatting?
    ├─ YES, with color/bold/etc → Use `insertContent` with [color:] / **bold** inline syntax
    ├─ YES, apply formatting to existing doc → Use `applyTaggedFormatting` (Workflow B)
    ├─ Formatting on user's active selection only → Use `formatText`
    ├─ Comments → Use `addComment`
    ├─ Tables → Use table tools
    └─ None of above → Use `eval_wordjs`
````

```markdown
<!-- APRES -->

## TOOL SELECTION DECISION TREE

User wants to apply formatting to specific words (color, bold, highlight...)?
YES → Use `searchAndFormat` (Workflow C — one call per word/phrase)

User wants to modify existing TEXT content?
YES: Is it a simple word/phrase replacement?
YES → Use `searchAndReplace`
NO (rewriting paragraphs) → Use `proposeRevision`

User wants to add NEW content?
YES → Use `insertContent` with Markdown syntax (Workflow A for formatting)

Other:
Formatting on user's active selection only → `formatText`
Comments → `addComment`
Tables → table tools
None of above → `eval_wordjs`
```

**Zone 3** — Section "proposeRevision vs insertContent" (lignes 133-145). Remplacer integralement :

```markdown
<!-- AVANT (lignes 133-145) -->

## proposeRevision vs insertContent

**Use proposeRevision when:**

- Editing existing text (fix, correct, improve, rewrite, edit)
- You want to preserve existing formatting on unchanged portions

**Use insertContent when:**

- Adding completely new content
- Creating tables or lists from scratch
- User says "add", "insert", "create", "write"
- User wants a rewrite **with color/formatting** (use inline syntax)
```

```markdown
<!-- APRES -->

## searchAndFormat vs proposeRevision vs insertContent

**Use searchAndFormat when:**

- User wants to apply formatting (color, bold, highlight, etc.) to specific words
- Examples: "mettre les verbes en vert", "surligner les erreurs", "mettre en gras les dates"
- The TEXT content does not change, only the formatting
- Call once per word/phrase to format

**Use proposeRevision when:**

- Editing existing text content (fix, correct, improve, rewrite, edit)
- You want to preserve existing formatting on unchanged portions
- Track Changes will show what was modified

**Use insertContent when:**

- Adding completely new content that doesn't exist yet
- Creating tables or lists from scratch
- User says "add", "insert", "create", "write"
- NEVER use to modify existing text (causes full replacement visible in Track Changes)
```

**Zone 4** — Table des outils WRITE (lignes 18-28). Ajouter `proposeRevision` en tete :

```markdown
<!-- AVANT (lignes 18-28) -->

### For WRITING/EDITING content:

| Tool                 | When to use                                                    |
| -------------------- | -------------------------------------------------------------- |
| `proposeRevision`    | **PREFERRED** for editing existing text. Preserves formatting! |
| `searchAndReplace`   | Fix specific words/phrases throughout document                 |
| `insertContent`      | Add new content (Markdown + inline color/style syntax)         |
| `insertHyperlink`    | Add clickable links                                            |
| `addComment`         | Add review comments                                            |
| `insertHeaderFooter` | Add headers/footers                                            |
| `insertFootnote`     | Add footnotes                                                  |
```

```markdown
<!-- APRES — clarifier les roles -->

### For WRITING/EDITING content:

| Tool                 | When to use                                                              |
| -------------------- | ------------------------------------------------------------------------ |
| `proposeRevision`    | **PREFERRED** for editing existing text. Preserves formatting! Uses diff |
| `searchAndReplace`   | Fix specific words/phrases throughout document                           |
| `insertContent`      | Add NEW content only (Markdown + inline color/style syntax)              |
| `insertHyperlink`    | Add clickable links                                                      |
| `addComment`         | Add review comments                                                      |
| `insertHeaderFooter` | Add headers/footers                                                      |
| `insertFootnote`     | Add footnotes                                                            |
```

#### Tests

- Demander au LLM "mets les verbes en vert" sur un document Word
- Verifier que le LLM utilise `searchAndFormat` (et PAS `insertContent`)
- Verifier que le Track Changes ne montre PAS de suppression/re-insertion du texte
- Verifier que les verbes sont bien en vert
- Demander "reecris ce paragraphe en ameliorant le style" -> verifier que le LLM utilise `proposeRevision`
- Demander "ajoute un tableau avec les ventes" -> verifier que le LLM utilise `insertContent`

---

## Suivi des actions

| #   | Action                                                           | Fichier(s)                            | Criticite | Issue | Status                                 |
| --- | ---------------------------------------------------------------- | ------------------------------------- | --------- | ----- | -------------------------------------- |
| C1  | Appeler `applyOfficeBlockStyles()` dans `renderOfficeRichHtml()` | `markdown.ts`                         | HAUTE     | H1    | ✅ DONE (2026-03-07)                   |
| C2  | Ajouter l'outil `searchAndFormat`                                | `wordTools.ts`                        | HAUTE     | H3    | ✅ DONE (2026-03-07)                   |
| C3  | Corriger le prompt agent + decision tree                         | `useAgentPrompts.ts`, `word.skill.md` | HAUTE     | H2    | ✅ DONE (2026-03-07)                   |
| H4  | Ajouter confirmation "New Chat" si brouillon non vide            | `HomePage.vue`                        | HAUTE     | H4    | ⚠️ OPEN                                |
| H5  | Retirer `export` de `insertMarkdownIntoTextRange`                | `powerpointTools.ts`                  | HAUTE     | H5    | ✅ N/A (fonction supprimee)            |
| M1  | Ajouter `console.warn` dans le catch font loading                | `powerpointTools.ts`                  | MOYENNE   | M1    | ✅ N/A (fonction supprimee)            |
| M2  | Supprimer la cle i18n `newChatConfirm` morte                     | `en.json`, `fr.json`                  | MOYENNE   | M2    | ✅ DONE (2026-03-07)                   |
| M3  | Ameliorer le type de retour de `findShapeOnSlide`                | `powerpointTools.ts`                  | MOYENNE   | M3    | ✅ N/A (fonction supprimee)            |
| L1  | Remplacer les types `any` par les types Office.js                | `powerpointTools.ts`                  | BASSE     | L1    | ✅ N/A (refactoring anterieur)         |
| L2  | Extraire un composant `ConfirmDialog` si reutilise               | `HomePage.vue`                        | BASSE     | L2    | ⚠️ OPEN (cosmetic, 1 seule occurrence) |

---

## Deferred Items (carried forward from v4)

- **IC2** — Containers run as root (low priority, deployment simplicity)
- **IH2** — Private IP in build arg (users override at build time)
- **IH3** — DuckDNS domain in example (users replace with their own)
- **UM10** — PowerPoint HTML reconstruction (high complexity, low ROI)

---

## Verification Commands

```bash
# Frontend build check
cd frontend && npm run build

# Backend start check
cd backend && npm start

# Check for TypeScript errors
cd frontend && npx tsc --noEmit
```

---

## Changelog

| Version | Date       | Changes                                                         |
| ------- | ---------- | --------------------------------------------------------------- |
| v5.0    | 2026-03-07 | PR #158 review, Word/PPT pipeline audit, 3 targeted corrections |
| v4.0    | 2026-03-03 | Complete fresh audit, 50 issues all resolved                    |
| v3.0    | 2026-02-28 | 162 issues identified, 131 resolved                             |
| v2.0    | 2026-02-22 | 28 new issues after major refactor                              |
| v1.0    | 2026-02-15 | Initial audit, 38 issues (all resolved)                         |
