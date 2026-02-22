# Rapport d'analyse : Mise en forme texte et gestion des puces
## Redink vs KickOffice - Recommandations d'amelioration

**Date** : 2026-02-22
**Objectif** : Identifier dans le projet Redink les approches de mise en forme de texte et de gestion des puces qui pourraient resoudre les problemes actuels de KickOffice (puces PowerPoint/Word, mise en forme texte).

---

## 1. RESUME EXECUTIF

KickOffice souffre de problemes de mise en forme dans Word et PowerPoint car :
1. Le pipeline Markdown -> HTML est trop simpliste et ne gere pas les specificites des hotes Office
2. Les prompts systeme ne donnent pas d'instructions de formatage suffisamment precises au LLM
3. L'insertion dans PowerPoint n'exploite pas les capacites de l'API moderne pour les puces
4. Il manque une couche d'adaptation du HTML aux exigences specifiques de Word (styles inline, espacement, polices)

Redink a resolu ces problemes via :
- Un pipeline Markdown -> HTML -> CF_HTML enrichi avec styles inline herites du document
- Une gestion fine des `<br>` dans les listes et paragraphes
- Des prompts systeme tres detailles sur le format de sortie attendu
- L'utilisation de HtmlAgilityPack pour manipuler l'arbre HTML avant insertion

---

## 2. ANALYSE DETAILLEE DE REDINK

### 2.1 Pipeline de conversion texte (SharedMethods.TextHandling.vb)

Le coeur du systeme Redink est un pipeline en 8 etapes :

**Etape 1 - Pre-traitement Markdown** (lignes 54-74)
- Normalisation des sauts de ligne multiples
- Insertion de `&nbsp;` entre les sauts de ligne consecutifs pour eviter que Word les fusionne
- C'est un probleme que KickOffice rencontre aussi : les paragraphes vides sont avales

**Etape 2 - Conversion Markdown -> HTML via Markdig** (lignes 77-96)
- Utilise un pipeline Markdig tres complet avec extensions :
  - `UsePipeTables()` - tableaux pipe
  - `UseGridTables()` - tableaux grille
  - `UseSoftlineBreakAsHardlineBreak()` - sauts de ligne simples = `<br>`
  - `UseListExtras()` - **listes avancees** (a, b, c / i, ii, iii)
  - `UseDefinitionLists()` - listes de definitions
  - `UseTaskLists()` - cases a cocher
  - `UseAdvancedExtensions()` - toutes les extensions avancees
  - `UseGenericAttributes()` - attributs personnalises

**Etape 3 - Nettoyage des retours a la ligne residuels** (lignes 99-102)
- Suppression de tous les `\r\n`, `\r`, `\n` restants dans le HTML

**Etape 4 - Eclatement des `<br>` en `<p>` separes** (lignes 143-177)
- **Point cle** : Dans les `<p>`, chaque `<br>` est converti en un element `<p>` distinct
- Dans les `<li>`, chaque segment separe par `<br>` devient un `<p>` enfant du `<li>`
- Cela garantit que Word interprete correctement chaque paragraphe

**Etape 5 - Heritage des styles du document** (lignes 181-216)
- **Innovation majeure** : Redink lit les proprietes de formatage actuelles du curseur Word :
  - Nom de police, taille, gras, italique, couleur
  - Espacement avant/apres paragraphe
  - Interligne (simple, 1.5, double, multiple, exact, minimum)
- Ces proprietes sont converties en CSS inline

**Etape 6 - Application des styles inline** (lignes 224-249)
- Chaque `<p>` et `<li>` recoit les styles CSS herites
- Les headings (`<h1>` a `<h6>`) ne recoivent que la police et la couleur (pas la taille)
- **Resultat** : Le texte insere respecte automatiquement le style du document en cours

**Etape 7 - Construction du paquet CF_HTML** (lignes 252-284)
- Construction d'un paquet HTML complet avec `<html><head><meta charset>`
- Calcul des offsets UTF-8 byte-level pour le format CF_HTML (clipboard)
- Les marqueurs `<!--StartFragment-->` et `<!--EndFragment-->` delimitent le contenu

**Etape 8 - Insertion via clipboard** (lignes 288-358)
- Sauvegarde du clipboard existant
- Ecriture du HTML au clipboard en format CF_HTML
- Paste dans Word avec `PasteAndFormat(wdFormatOriginalFormatting)`
- Restauration du clipboard original

### 2.2 Systeme de prompts Redink

Les prompts Redink (promptlib.txt) sont tres structures avec des instructions de formatage explicites :
- Format de sortie specifie (Markdown, texte brut, tableaux)
- Instructions claires sur les puces : "do not make any paragraphs or bullet points" quand ce n'est pas voulu
- Utilisation de tags XML pour delimiter le contenu (`<TEXTTOPROCESS>`, `<COMPARE>`, etc.)

### 2.3 Nettoyage HTML (SimplifyHtml / CleanHtmlNode)

Redink maintient une whitelist stricte de tags HTML autorises :
```
Tags autorises : b, strong, i, em, u, font, span, p, ul, ol, li, br
Attributs autorises : style, class
```
Tout tag non autorise est "deploye" (remplace par son contenu enfant).

---

## 3. ANALYSE DE KICKOFFICE - PROBLEMES IDENTIFIES

### 3.1 Pipeline officeRichText.ts - Lacunes

**Probleme 1 : Pas d'heritage des styles du document**
```typescript
// KickOffice - officeRichText.ts ligne 110-122
function applyOfficeBlockStyles(html: string): string {
  return html
    .replace(/<h1>/gi, '<h1 style="margin:0 0 8px 0; font-size:2em; font-weight:700;">')
    .replace(/<p>/gi, '<p style="margin:0 0 6px 0;">')
    .replace(/<ul>/gi, '<ul style="margin:0 0 6px 0; padding-left:1.25em;">')
    .replace(/<li>/gi, '<li style="margin:0 0 2px 0;">')
}
```
- Les styles sont hardcodes (tailles fixes, marges fixes)
- Aucune lecture de la police/taille/couleur actuelle du document
- Resultat : le texte insere ne correspond pas au style du document en cours

**Probleme 2 : Gestion des `<br>` dans les listes insuffisante**
- KickOffice ne fait pas la conversion `<br>` -> `<p>` dans les `<li>` que fait Redink
- Les `<br>` dans les elements de liste creent des problemes d'espacement

**Probleme 3 : Markdown-it au lieu de Markdig - extensions manquantes**
- KickOffice utilise markdown-it qui est bien, mais sans les extensions de listes avancees
- Pas de `UseListExtras` equivalent (listes alphabetiques, romaines)
- Pas de `UseDefinitionLists` equivalent

**Probleme 4 : Pas de normalisation des sauts de ligne multiples**
- Redink insere des `&nbsp;` entre les sauts de ligne consecutifs pour preserver les paragraphes vides
- KickOffice perd ces paragraphes vides

### 3.2 PowerPoint - powerpointTools.ts - Lacunes

**Probleme 5 : Description du tool trompeuse / auto-limitante**
```typescript
// powerpointTools.ts lignes 190-191
description: 'Replace the currently selected PowerPoint text with new text.
IMPORTANT: PowerPoint API does NOT support inserting HTML, Markdown, or applying
text formatting (bold, italics). ONLY plain text can be inserted.'
```
- Cette description dit au LLM que PowerPoint ne supporte PAS le formatage
- **Or, le code fait deja de l'insertion HTML** (via `insertIntoPowerPoint` qui utilise `insertHtml` pour l'API moderne et `CoercionType.Html` pour le fallback)
- Le LLM ne va donc jamais essayer de formater car le prompt lui dit que c'est impossible
- C'est une contradiction majeure entre la description du tool et son implementation

**Probleme 6 : Puces PowerPoint - mauvaise strategie**
```typescript
// powerpointTools.ts ligne 63
export function normalizePowerPointListText(text: string): string {
  const normalizedNewlines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  return stripMarkdownListMarkers(normalizedNewlines)
}
```
- `stripMarkdownListMarkers` supprime les marqueurs de liste (`-`, `*`, `1.`)
- Cela detruit la structure de liste au lieu de la preserver
- La fonction est censee "jouer bien avec les shapes nativement bulleted" mais en pratique elle perd l'information

**Probleme 7 : insertTextBox utilise stripRichFormattingSyntax**
```typescript
// powerpointTools.ts ligne 432
const shape = slide.shapes.addTextBox(stripRichFormattingSyntax(text))
```
- Tout le formatage riche est strip avant insertion dans un nouveau textbox
- On perd donc systematiquement gras, italique, etc.

### 3.3 Prompts - Lacunes

**Probleme 8 : Prompt PowerPoint quasi-vide**
```typescript
// useAgentPrompts.ts ligne 46
const powerPointAgentPrompt = (lang: string) =>
  `# Role\nYou are a PowerPoint presentation expert.\n# Guidelines\n5. **Language**: entirely in ${lang}.`
```
- Comparons avec le prompt Word qui fait 15 lignes et donne des instructions de formatage
- Le prompt PowerPoint ne donne AUCUNE instruction sur le formatage du texte
- Pas de guideline sur comment structurer les puces, le gras, etc.

**Probleme 9 : Prompt Word - instruction HTML contradictoire**
```typescript
// useAgentPrompts.ts ligne 44
'3. **Formatting**: When generating or modifying document content, ALWAYS format
your response using semantic HTML tags (e.g., <b>, <i>, <u>, <h1> to <h6>, <p>,
<ul>, <li>, <br>) instead of plain text or markdown'
```
- Le prompt demande au LLM de generer du HTML brut
- Mais le pipeline `renderOfficeCommonApiHtml` attend du Markdown qu'il convertit en HTML
- Si le LLM envoie du HTML directement, le parser Markdown peut le corrompre
- Il faudrait soit accepter le HTML direct, soit demander du Markdown au LLM

**Probleme 10 : GLOBAL_STYLE_INSTRUCTIONS trop limite**
```typescript
export const GLOBAL_STYLE_INSTRUCTIONS = `
CRITICAL INSTRUCTIONS FOR ALL GENERATIONS:
- NEVER use em-dashes (—).
- NEVER use semicolons (;).
- Keep the sentence structure natural and highly human-like.`
```
- Aucune instruction sur le formatage des puces
- Aucune instruction sur la hierarchie des listes
- Aucune instruction sur l'utilisation du Markdown standard pour les listes

---

## 4. RECOMMANDATIONS D'IMPLEMENTATION

### R1. Heritage des styles du document (Priorite : HAUTE)

**Origine Redink** : `InsertTextWithFormat` lignes 181-216

**Strategie pour KickOffice** :
Avant d'inserer du HTML dans Word, lire les proprietes de formatage au point d'insertion via l'API Word.js et les injecter en CSS inline dans le HTML.

**Implementation** :
```typescript
// Nouveau fichier ou extension de officeRichText.ts
async function getInsertionPointStyles(): Promise<InlineStyles> {
  return Word.run(async (context) => {
    const range = context.document.getSelection()
    range.load(['font/name', 'font/size', 'font/bold', 'font/italic',
                'font/color', 'paragraphFormat/spaceAfter',
                'paragraphFormat/spaceBefore', 'paragraphFormat/lineSpacing'])
    await context.sync()
    return {
      fontFamily: range.font.name,
      fontSize: `${range.font.size}pt`,
      fontWeight: range.font.bold ? 'bold' : 'normal',
      fontStyle: range.font.italic ? 'italic' : 'normal',
      color: range.font.color,
      marginTop: `${range.paragraphFormat.spaceBefore}pt`,
      marginBottom: `${range.paragraphFormat.spaceAfter}pt`,
      lineHeight: `${range.paragraphFormat.lineSpacing}pt`,
    }
  })
}

function applyInheritedStyles(html: string, styles: InlineStyles): string {
  const cssBlock = `font-family:'${styles.fontFamily}'; font-size:${styles.fontSize}; color:${styles.color};`
  // Appliquer a chaque <p> et <li>
  return html
    .replace(/<p>/gi, `<p style="${cssBlock} margin:${styles.marginTop} 0 ${styles.marginBottom} 0;">`)
    .replace(/<li>/gi, `<li style="${cssBlock}">`)
}
```

**Fichiers a modifier** :
- `frontend/src/utils/officeRichText.ts` - ajouter la logique d'heritage
- `frontend/src/api/common.ts` (ou equivalent) - appeler `getInsertionPointStyles` avant insertion

---

### R2. Corriger les descriptions des tools PowerPoint (Priorite : HAUTE)

**Probleme** : Les descriptions disent que le formatage est impossible alors que le code le supporte.

**Implementation** :
```typescript
// powerpointTools.ts - replaceSelectedText
description: 'Replace the currently selected PowerPoint text with new content. '
  + 'The text will be rendered from Markdown to HTML before insertion. '
  + 'You can use Markdown formatting: **bold**, *italic*, bullet lists (- item), '
  + 'numbered lists (1. item), and headings (## Heading). '
  + 'Indented sub-items are supported for nested lists.',

// powerpointTools.ts - insertTextBox
description: 'Insert a text box into a specific slide. '
  + 'Content supports Markdown formatting for rich text rendering.',
```

**Fichiers a modifier** :
- `frontend/src/utils/powerpointTools.ts` lignes 190 et 390

---

### R3. Ameliorer le prompt PowerPoint (Priorite : HAUTE)

**Origine Redink** : Prompts detailles avec instructions de format explicites

**Implementation** :
```typescript
const powerPointAgentPrompt = (lang: string) => `# Role
You are a highly skilled Microsoft PowerPoint Expert Agent.

# Capabilities
- You can interact with the presentation using provided tools.
- You understand slide design, visual hierarchy, and presentation best practices.

# Guidelines
1. **Tool First**: Use tools for any slide modification.
2. **Formatting**: When generating text content for slides, use Markdown:
   - Use **bold** for emphasis and key terms
   - Use bullet lists (- item) for main points
   - Use indented bullets (  - sub-item) for details
   - Use numbered lists (1. item) for sequential steps
   - Keep bullet text concise (max 8-10 words per point)
3. **Bullet Hierarchy**: Structure content with clear visual hierarchy:
   - Level 1: Main points (- Main point)
   - Level 2: Supporting details (  - Detail)
   - Level 3: Sub-details (    - Sub-detail)
4. **Conciseness**: Slides should be scannable. Avoid full sentences.
5. **Language**: Communicate entirely in ${lang}.

# Safety
Do not delete slides unless explicitly instructed.`
```

**Fichiers a modifier** :
- `frontend/src/composables/useAgentPrompts.ts` ligne 46

---

### R4. Gestion des `<br>` dans les listes (Priorite : MOYENNE)

**Origine Redink** : `InsertTextWithFormat` lignes 143-177

**Strategie** : Ajouter une etape de post-traitement HTML qui eclate les `<br>` dans les `<p>` et `<li>`.

**Implementation** :
```typescript
// officeRichText.ts - nouvelle fonction
function splitBrInListItems(html: string): string {
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')

  doc.querySelectorAll('li, p').forEach(node => {
    const innerHTML = node.innerHTML
    if (!/<br\s*\/?>/i.test(innerHTML)) return

    const segments = innerHTML.split(/<br\s*\/?>/i).map(s => s.trim()).filter(Boolean)
    if (segments.length <= 1) return

    if (node.tagName === 'LI') {
      node.innerHTML = segments.map(seg => `<p>${seg}</p>`).join('')
    } else if (node.tagName === 'P') {
      const parent = node.parentNode
      if (!parent) return
      segments.forEach(seg => {
        const newP = document.createElement('p')
        newP.innerHTML = seg
        // Copier le style du noeud original
        if (node.getAttribute('style')) {
          newP.setAttribute('style', node.getAttribute('style')!)
        }
        parent.insertBefore(newP, node)
      })
      parent.removeChild(node)
    }
  })

  return doc.body.innerHTML
}
```

**Fichiers a modifier** :
- `frontend/src/utils/officeRichText.ts` - ajouter `splitBrInListItems` dans le pipeline `renderOfficeCommonApiHtml`

---

### R5. Preserver les sauts de ligne multiples (Priorite : MOYENNE)

**Origine Redink** : `InsertTextWithMarkdown` lignes 54-74

**Implementation** :
```typescript
// officeRichText.ts - dans renderOfficeRichHtml, avant le parsing Markdown
function preserveMultipleLineBreaks(content: string): string {
  // Remplacer les sauts de ligne multiples par des paragraphes vides
  // pour que le parser Markdown ne les avale pas
  return content.replace(/(\n\s*\n)/g, (match) => {
    const breaks = match.split(/\n/).filter(Boolean)
    if (breaks.length <= 1) return match
    return breaks.join('\n&nbsp;\n')
  })
}
```

**Fichiers a modifier** :
- `frontend/src/utils/officeRichText.ts` - integrer dans `renderOfficeRichHtml`

---

### R6. Harmoniser le prompt Word (HTML vs Markdown) (Priorite : HAUTE)

**Probleme** : Le prompt demande du HTML, le pipeline attend du Markdown.

**Solution recommandee** : Demander du Markdown au LLM (car le pipeline le convertit) :

```typescript
const wordAgentPrompt = (lang: string) => `# Role
You are a highly skilled Microsoft Word Expert Agent.

# Guidelines
1. **Tool First**: Prioritize using tools for document operations.
2. **Direct Actions**: For formatting requests, execute changes with tools.
3. **Formatting**: When generating content for insertion into the document,
   use standard Markdown syntax:
   - **bold** for emphasis
   - *italic* for nuance
   - __underline__ for highlighting
   - # Heading 1, ## Heading 2, etc. for headings
   - - item or * item for bullet lists
   - 1. item for numbered lists
   - Indentation with 2 spaces for nested lists
   Do NOT use raw HTML tags. Use Markdown exclusively.
4. **Bullet Lists**: When creating lists:
   - Use - for unordered lists
   - Use 1. 2. 3. for ordered lists
   - Indent with 2 spaces for sub-levels
   - Each list item should be a complete but concise thought
5. **Accuracy**: Ensure changes are precise.
6. **Conciseness**: Provide brief explanations.
7. **Language**: Communicate in ${lang}.

# Safety
Do not perform destructive actions unless explicitly instructed.`
```

**Fichiers a modifier** :
- `frontend/src/composables/useAgentPrompts.ts` ligne 44

---

### R7. Ameliorer les extensions Markdown-it (Priorite : BASSE)

**Origine Redink** : Pipeline Markdig avec extensions avancees

**Implementation** : Ajouter des plugins markdown-it equivalents :

```bash
npm install markdown-it-task-lists markdown-it-deflist markdown-it-footnote
```

```typescript
// officeRichText.ts
import markdownItTaskLists from 'markdown-it-task-lists'
import markdownItDeflist from 'markdown-it-deflist'

const officeMarkdownParser = new MarkdownIt({ breaks: true, html: true, linkify: true })
  .use(markdownItTaskLists)
  .use(markdownItDeflist)
```

**Fichiers a modifier** :
- `frontend/src/utils/officeRichText.ts` lignes 4-9
- `package.json` - nouvelles dependances

---

### R8. Ne plus stripper le formatage dans insertTextBox (Priorite : MOYENNE)

**Implementation** :
```typescript
// powerpointTools.ts - insertTextBox (ligne 432)
// AVANT :
const shape = slide.shapes.addTextBox(stripRichFormattingSyntax(text))

// APRES : Utiliser l'insertion HTML si l'API le supporte
const shape = slide.shapes.addTextBox(text)
// Puis appliquer le formatage via textRange.insertHtml si disponible
try {
  shape.textFrame.textRange.insertHtml(
    renderOfficeCommonApiHtml(text), 'Replace'
  )
} catch {
  // Fallback : texte brut si HTML non supporte
  shape.textFrame.textRange.text = stripRichFormattingSyntax(text)
}
```

**Fichiers a modifier** :
- `frontend/src/utils/powerpointTools.ts` ligne 432

---

### R9. Enrichir GLOBAL_STYLE_INSTRUCTIONS (Priorite : MOYENNE)

**Implementation** :
```typescript
export const GLOBAL_STYLE_INSTRUCTIONS = `
CRITICAL INSTRUCTIONS FOR ALL GENERATIONS:
- NEVER use em-dashes (\u2014).
- NEVER use semicolons (;).
- Keep the sentence structure natural and highly human-like.
- When creating bullet lists, use standard Markdown syntax:
  - Use "-" for unordered lists (not "*" or "+")
  - Use "1." "2." "3." for numbered lists
  - Use 2-space indentation for nested sub-items
  - Each bullet should be a concise, standalone point
- For emphasis, use **bold** (not CAPS or underlining)
- For document structure, use Markdown headings (# ## ###)`
```

**Fichiers a modifier** :
- `frontend/src/utils/constant.ts` lignes 17-21

---

## 5. PLAN D'IMPLEMENTATION PAR PRIORITE

### Phase 1 - Corrections rapides (impact immediat)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R2 | Corriger descriptions tools PowerPoint | powerpointTools.ts | 15 min |
| R3 | Ameliorer prompt PowerPoint | useAgentPrompts.ts | 30 min |
| R6 | Harmoniser prompt Word (Markdown) | useAgentPrompts.ts | 30 min |
| R9 | Enrichir GLOBAL_STYLE_INSTRUCTIONS | constant.ts | 15 min |

**Impact** : Le LLM generera du contenu correctement formate car il aura les bonnes instructions. C'est la correction la plus importante car elle resout le probleme a la source.

### Phase 2 - Ameliorations du pipeline HTML

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R4 | Gestion des `<br>` dans les listes | officeRichText.ts | 1-2h |
| R5 | Preserver sauts de ligne multiples | officeRichText.ts | 30 min |
| R8 | HTML dans insertTextBox PowerPoint | powerpointTools.ts | 1h |

**Impact** : Le rendu HTML sera plus fidele et les puces/listes s'afficheront correctement.

### Phase 3 - Heritage des styles (amelioration profonde)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R1 | Heritage styles du document Word | officeRichText.ts + common.ts | 2-3h |
| R7 | Extensions Markdown-it supplementaires | officeRichText.ts + package.json | 1h |

**Impact** : Le texte insere sera visuellement coherent avec le document existant (meme police, taille, couleur).

---

## 6. ELEMENTS REDINK NON APPLICABLES A KICKOFFICE

Les elements suivants de Redink ne sont **pas** transposables :

1. **Insertion via clipboard CF_HTML** : Redink utilise le clipboard Windows natif pour coller du HTML dans Word. KickOffice etant une web app (Office Add-in), il doit passer par les APIs Office.js (`setSelectedDataAsync`, `insertHtml`). Cette approche n'est pas applicable.

2. **VBA Interop direct** : L'acces direct aux objets COM Word (Range.Font, ParagraphFormat) n'est pas disponible dans Office.js avec la meme granularite. L'API Word.js offre neanmoins un sous-ensemble suffisant.

3. **ConvertMarkupToRTF** : La generation RTF directe n'est pas pertinente pour un add-in web.

4. **HtmlAgilityPack** : Cote serveur .NET. En remplacement, `DOMParser` cote navigateur ou une librairie comme `cheerio` ferait l'equivalent.

---

## 7. CONCLUSION

Les problemes de mise en forme de KickOffice sont principalement dus a :

1. **Des prompts qui ne guident pas correctement le LLM** sur le format de sortie attendu (surtout pour PowerPoint et les listes) - **c'est le probleme principal**
2. **Des descriptions de tools contradictoires** qui disent au LLM que le formatage est impossible alors que le code le supporte
3. **Un pipeline HTML qui n'adapte pas les styles** au contexte du document

La Phase 1 (corrections de prompts et descriptions) devrait resoudre la majorite des problemes visibles en moins de 2 heures de travail, sans modification structurelle du code. Les Phases 2 et 3 apporteront des ameliorations plus profondes de qualite de rendu.
