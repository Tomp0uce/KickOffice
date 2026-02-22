# Rapport d'analyse : Mise en forme texte et gestion des puces
## Recommandations d'amelioration KickOffice

**Date** : 2026-02-22
**Objectif** : Identifier les approches de reference en matiere de mise en forme de texte et de gestion des puces qui pourraient resoudre les problemes actuels de KickOffice (puces PowerPoint/Word, mise en forme texte).

---

## 1. RESUME EXECUTIF

KickOffice souffre de problemes de mise en forme dans Word et PowerPoint car :
1. Le pipeline Markdown -> HTML est trop simpliste et ne gere pas les specificites des hotes Office
2. Les prompts systeme ne donnent pas d'instructions de formatage suffisamment precises au LLM
3. L'insertion dans PowerPoint n'exploite pas les capacites de l'API moderne pour les puces
4. Il manque une couche d'adaptation du HTML aux exigences specifiques de Word (styles inline, espacement, polices)

La solution de reference a resolu ces problemes via :
- Un pipeline Markdown -> HTML -> CF_HTML enrichi avec styles inline herites du document
- Une gestion fine des `<br>` dans les listes et paragraphes
- Des prompts systeme tres detailles sur le format de sortie attendu
- L'utilisation de HtmlAgilityPack pour manipuler l'arbre HTML avant insertion

---

## 2. ANALYSE DETAILLEE DE LA SOLUTION DE REFERENCE

### 2.1 Pipeline de conversion texte (SharedMethods.TextHandling.vb)

Le coeur du systeme de reference est un pipeline en 8 etapes :

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
- **Innovation majeure** : La solution de reference lit les proprietes de formatage actuelles du curseur Word :
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

### 2.2 Systeme de prompts de reference

Les prompts de reference (promptlib.txt) sont tres structures avec des instructions de formatage explicites :
- Format de sortie specifie (Markdown, texte brut, tableaux)
- Instructions claires sur les puces : "do not make any paragraphs or bullet points" quand ce n'est pas voulu
- Utilisation de tags XML pour delimiter le contenu (`<TEXTTOPROCESS>`, `<COMPARE>`, etc.)

### 2.3 Nettoyage HTML (SimplifyHtml / CleanHtmlNode)

La solution de reference maintient une whitelist stricte de tags HTML autorises :
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
- KickOffice ne fait pas la conversion `<br>` -> `<p>` dans les `<li>` que fait la solution de reference
- Les `<br>` dans les elements de liste creent des problemes d'espacement

**Probleme 3 : Markdown-it au lieu de Markdig - extensions manquantes**
- KickOffice utilise markdown-it qui est bien, mais sans les extensions de listes avancees
- Pas de `UseListExtras` equivalent (listes alphabetiques, romaines)
- Pas de `UseDefinitionLists` equivalent

**Probleme 4 : Pas de normalisation des sauts de ligne multiples**
- La solution de reference insere des `&nbsp;` entre les sauts de ligne consecutifs pour preserver les paragraphes vides
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

**Origine reference** : `InsertTextWithFormat` lignes 181-216

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

**Origine reference** : Prompts detailles avec instructions de format explicites

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

**Origine reference** : `InsertTextWithFormat` lignes 143-177

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

**Origine reference** : `InsertTextWithMarkdown` lignes 54-74

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

**Origine reference** : Pipeline Markdig avec extensions avancees

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

## 6. ELEMENTS DE LA SOLUTION DE REFERENCE NON APPLICABLES A KICKOFFICE

Les elements suivants de la solution de reference ne sont **pas** transposables :

1. **Insertion via clipboard CF_HTML** : La solution de reference utilise le clipboard Windows natif pour coller du HTML dans Word. KickOffice etant une web app (Office Add-in), il doit passer par les APIs Office.js (`setSelectedDataAsync`, `insertHtml`). Cette approche n'est pas applicable.

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

---

## 8. ANALYSE APPROFONDIE : DOUBLE PUCES ET INDENTATION

### 8.1 Cause racine des doubles puces dans PowerPoint

Le probleme de double puces dans PowerPoint vient d'un **conflit entre les puces natives des shapes et les puces HTML/Markdown**.

**Scenario type :**
1. L'utilisateur a un textbox PowerPoint avec des puces natives (bullet points configures dans le theme du slide)
2. Le LLM genere du Markdown avec des marqueurs de liste : `- Point 1\n- Point 2`
3. Le pipeline `renderOfficeCommonApiHtml` convertit ca en `<ul><li>Point 1</li><li>Point 2</li></ul>`
4. Le HTML est insere via `insertHtml` ou `CoercionType.Html`
5. **PowerPoint ajoute ses puces natives AU-DESSUS des puces HTML** = double puces

**Preuve dans le code :**
```typescript
// powerpointTools.ts:97-115 - insertIntoPowerPoint
const htmlContent = renderOfficeCommonApiHtml(normalizedNewlines)
// ... tente insertHtml avec l'API moderne
textRange.insertHtml(htmlContent, 'Replace')
// ... puis fallback CoercionType.Html
```
Le HTML contient `<ul><li>` qui genere ses propres puces. Mais si le shape cible a deja un style de paragraphe avec puces natives, on obtient puce native + puce HTML.

**Le fallback texte brut a aussi un probleme :**
```typescript
// powerpointTools.ts:142-156 - fallbackToText
const fallbackText = stripRichFormattingSyntax(text, true) // true = strip list markers
```
Quand `stripListMarkers=true`, la fonction `stripMarkdownListMarkers` retire les `- ` et `1. ` mais garde l'indentation comme tabs. Resultat : texte indente sans puce visible si le shape n'a pas de puces natives, ou puces natives mais pas de hierarchie visible.

### 8.2 Cause racine des doubles puces dans Word

Dans Word, le probleme est similaire mais se manifeste differemment :

**Scenario type :**
1. Le LLM genere du contenu avec `<ul><li>` via le pipeline `renderOfficeRichHtml`
2. Word.js `insertHtml()` insere le HTML
3. Word interprete `<ul><li>` et cree un paragraphe avec le style "List Paragraph" + puce native
4. Mais le HTML peut AUSSI contenir un caractere "bullet" unicode (U+2022) si le LLM a melange Markdown et texte brut
5. Ou bien si le point d'insertion etait deja dans une liste, Word herite du style de liste et ajoute une puce supplementaire

**Le tool `insertList` de Word confirme le probleme :**
```typescript
// wordTools.ts:454-467 - insertList
const markdownList = listType === 'bullet'
  ? items.map((i: string) => `* ${i}`).join('\n')
  : items.map((i: string, idx: number) => `${idx + 1}. ${i}`).join('\n')
range.insertHtml(renderOfficeRichHtml(markdownList), 'After')
```
Ce tool genere du Markdown (`* item`) puis le convertit en HTML (`<ul><li>`). C'est correct en soi, mais :
- Il n'y a **aucun nettoyage du contexte de destination** (le curseur est-il deja dans une liste ?)
- Il n'y a **aucune gestion de l'indentation/niveaux de liste** (le tool n'accepte que des items plats)
- Les sous-listes sont impossibles via ce tool

### 8.3 Cause racine de l'impossibilite d'indenter les puces

**PowerPoint :**
Le HTML `<ul>` standard ne supporte qu'un niveau. Pour les sous-listes, il faut des `<ul>` imbriques :
```html
<ul>
  <li>Niveau 1
    <ul>
      <li>Niveau 2</li>
    </ul>
  </li>
</ul>
```
Le pipeline `renderOfficeRichHtml` produit bien cette structure depuis du Markdown indente.
MAIS le probleme est que :
1. Les descriptions de tools disent au LLM que le formatage est impossible (voir Probleme 5)
2. Le prompt PowerPoint ne mentionne pas la hierarchie des listes (voir Probleme 8)
3. Donc le LLM ne genere jamais de listes imbriquees dans le bon format

**Word :**
L'indentation via HTML fonctionne nativement avec `<ul>` imbrique dans Word.js `insertHtml`.
Le probleme est que le prompt Word demande du HTML brut (Probleme 9) mais le pipeline attend du Markdown.
Si le LLM genere du Markdown avec indentation :
```markdown
- Point principal
  - Sous-point
    - Sous-sous-point
```
Le parser markdown-it le convertit correctement en HTML imbrique. Le probleme est donc **uniquement au niveau des prompts** qui ne guident pas le LLM.

### 8.4 Comment la solution de reference gere les listes (comparaison)

La solution de reference a une approche tres differente car elle passe par le clipboard CF_HTML :

1. **Conversion Markdown -> HTML** : Markdig avec `UseListExtras()` genere du HTML avec `<ol>` et `<ul>` standard
2. **Pas de conflit avec les puces natives** : Comme la solution de reference utilise `PasteAndFormat(wdFormatOriginalFormatting)`, Word interprete le HTML comme un collage depuis un navigateur. Les listes HTML deviennent des paragraphes Word avec le style "List Paragraph" et les puces correspondantes. Il n'y a PAS de duplication car le paste remplace completement le formatage.
3. **Indentation preservee** : Les `<ul>` imbriques dans le HTML sont correctement convertis en niveaux de liste Word avec indentation progressive.

**Point cle applicable a KickOffice** : Le vrai probleme n'est pas dans le HTML genere mais dans la **methode d'insertion**. `insertHtml` de Word.js fait essentiellement la meme chose que le paste de la solution de reference. Le probleme vient du fait que le LLM ne genere pas le bon Markdown en amont.

---

## 9. ANALYSE OUTLOOK : POURQUOI CA MARCHE MIEUX

### 9.1 Ce qui fonctionne dans Outlook

Outlook utilise le **meme pipeline** `renderOfficeRichHtml` que Word et PowerPoint, mais avec moins de problemes. Voici pourquoi :

**1. Insertion coherente via HTML :**
```typescript
// outlookTools.ts:216-224 - insertTextAtCursor
const html = renderOfficeRichHtml(text)
mailbox.item.body.setSelectedDataAsync(
  html,
  { coercionType: getOfficeCoercionType().Html },
  ...
)
```
- Le HTML est insere directement via `setSelectedDataAsync` avec `CoercionType.Html`
- Outlook (etant un client email HTML) interprete nativement le HTML avec `<ul>`, `<li>`, `<strong>`, etc.
- **Pas de conflit avec des styles natifs** : les emails n'ont pas de "shapes" ou de "styles de liste document" comme Word/PowerPoint

**2. Les tools Outlook n'ont PAS de mensonge sur les capacites :**
```typescript
// outlookTools.ts - insertTextAtCursor
description: 'Insert plain text at the current cursor position...'
// Malgre le nom "plain text", le code fait du HTML via renderOfficeRichHtml
```
Le nom est trompeur ("plain text") mais au moins il ne dit PAS au LLM que le formatage est impossible.

**3. Le prompt Outlook est certes minimal mais ne contredit rien :**
```typescript
// useAgentPrompts.ts:47
const outlookAgentPrompt = (lang: string) =>
  `# Role\nYou are a Microsoft Outlook Email Expert Agent.\n# Guidelines\n4. **Language**: ${lang}.`
```
Pas d'instruction contradictoire HTML vs Markdown. Le LLM genere du Markdown naturellement, le pipeline le convertit.

### 9.2 Ce qui pourrait etre ameliore dans Outlook

Malgre le fonctionnement correct, Outlook a les memes lacunes potentielles :

1. **Pas de style inline herite** : Le HTML insere utilise des styles CSS hardcodes (`applyOfficeBlockStyles`). Si l'email a une police Calibri 11pt et le HTML insere a des styles differents, il y aura un decalage visuel.

2. **Prompt trop vide** : Le prompt Outlook ne mentionne pas le formatage. Ca fonctionne par defaut car Markdown -> HTML -> email est un chemin naturel, mais des instructions explicites sur la mise en forme des listes amelioreraient la coherence.

3. **Tool `insertTextAtCursor` vs `insertHtmlAtCursor` duplique** : Il y a deux tools qui font presque la meme chose :
   - `insertTextAtCursor` : appelle `renderOfficeRichHtml(text)` puis insere en HTML
   - `insertHtmlAtCursor` : insere du HTML brut sans passer par le pipeline Markdown

   C'est confusant pour le LLM. Il devrait y avoir un seul tool d'insertion clair.

---

## 10. STRATEGIE D'UNIFORMISATION WORD / POWERPOINT / OUTLOOK

### 10.1 Principe directeur

L'objectif est de faire passer les 3 hotes par le **meme pipeline de formatage** avec des adaptations minimales par hote. Outlook fonctionne comme reference car son chemin est le plus simple et le plus efficace.

### 10.2 Pipeline unifie propose

```
[LLM genere du Markdown standard]
         |
         v
[renderOfficeRichHtml()] -- Pipeline commun unique
  1. normalizeNamedStyles()       -- Tags de style personnalises
  2. normalizeUnderlineMarkdown() -- __text__ -> <u>
  3. normalizeSuperAndSubScript() -- ^text^ -> <sup>
  4. markdown-it.render()         -- Markdown -> HTML brut
  5. DOMPurify.sanitize()         -- Nettoyage securite
         |
         v
[applyOfficeBlockStyles()] -- Styles CSS inline pour les blocs
  + NEW: splitBrInListItems()  -- Eclater <br> dans les <li>
  + NEW: normalizeListNesting() -- S'assurer que les <ul> imbriques sont propres
         |
         v
   +-----+-----+-----+
   |     |           |
   v     v           v
 [WORD] [POWERPOINT] [OUTLOOK]
```

### 10.3 Adaptations par hote

**Word** (`wordFormatter.ts` + `wordTools.ts`) :
- Utiliser `insertHtml()` de Word.js (deja fait)
- **NOUVEAU** : Avant insertion, detecter si le curseur est dans une liste existante. Si oui, ne PAS envoyer de `<ul>/<ol>` wrapper, envoyer juste les `<li>` pour eviter les doubles puces
- **NOUVEAU** : Pour le tool `insertList`, supporter les sous-niveaux

**PowerPoint** (`powerpointTools.ts`) :
- Utiliser `insertHtml()` de l'API moderne (deja fait pour `replaceSelectedText`)
- **NOUVEAU** : Pour `insertTextBox`, creer le textbox PUIS inserer le HTML dans son textRange
- **CRITIQUE** : Corriger les descriptions de tools pour ne plus mentir sur les capacites
- **NOUVEAU** : Avant insertion HTML dans un shape existant, verifier si le shape a des puces natives. Si oui, utiliser le fallback text SANS strip des marqueurs pour eviter les doubles puces

**Outlook** (`outlookTools.ts`) :
- Garder l'approche actuelle (fonctionne bien)
- **NOUVEAU** : Fusionner `insertTextAtCursor` et `insertHtmlAtCursor` ou clarifier leur role dans les descriptions
- **NOUVEAU** : Ajouter le style inline herite si possible via `getAsync` sur le body HTML

### 10.4 Detection des puces natives pour eviter les doubles puces (R10 - NOUVEAU)

C'est la recommandation la plus critique pour resoudre le probleme de double puces.

**Strategie pour PowerPoint :**
```typescript
// powerpointTools.ts - Nouvelle fonction helper
async function hasNativeBullets(context: any, textRange: any): Promise<boolean> {
  // Charger les proprietes de paragraphe du textRange
  try {
    const paragraphs = textRange.paragraphs
    paragraphs.load('items')
    await context.sync()
    if (paragraphs.items.length > 0) {
      const firstPara = paragraphs.items[0]
      firstPara.load('bulletFormat/visible')
      await context.sync()
      return firstPara.bulletFormat?.visible === true
    }
  } catch {
    // API non disponible, on assume pas de puces natives
  }
  return false
}

// Dans insertIntoPowerPoint, adapter l'insertion :
async function insertIntoPowerPoint(text: string, useHtml = true): Promise<void> {
  const normalizedNewlines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')

  // Verifier si le shape cible a des puces natives
  let targetHasNativeBullets = false
  try {
    await PowerPoint.run(async (context: any) => {
      const textRange = context.presentation.getSelectedTextRanges().getItemAt(0)
      targetHasNativeBullets = await hasNativeBullets(context, textRange)
    })
  } catch {}

  if (targetHasNativeBullets) {
    // Shape avec puces natives : inserer en texte brut AVEC les retours ligne
    // mais SANS marqueurs Markdown pour eviter les doubles puces
    // Les puces natives du shape s'appliqueront automatiquement a chaque ligne
    const plainText = stripRichFormattingSyntax(normalizedNewlines, true)
    // Fallback texte brut
    await setSelectedDataAsText(plainText)
    return
  }

  // Shape sans puces natives : inserer en HTML (le HTML apporte ses propres puces)
  const htmlContent = renderOfficeCommonApiHtml(normalizedNewlines)
  // ... insertion HTML comme avant
}
```

**Strategie pour Word :**
```typescript
// wordTools.ts - Dans insertText et replaceSelectedText
// Detecter si le curseur est dans un style de liste
async function isInsertionInList(context: Word.RequestContext): Promise<boolean> {
  const selection = context.document.getSelection()
  const para = selection.paragraphs.getFirst()
  para.load('style,listItem')
  await context.sync()
  try {
    // Si listItem existe et n'est pas null, on est dans une liste
    para.listItem.load('level')
    await context.sync()
    return true
  } catch {
    return false
  }
}
```

### 10.5 Tableau recapitulatif des problemes de puces et corrections

| Probleme | Cause | Hote | Correction |
|----------|-------|------|------------|
| Double puces | Shape avec puces natives + HTML `<ul><li>` | PowerPoint | R10: Detecter puces natives, fallback texte brut |
| Double puces | Curseur dans une liste + insertion `<ul>` | Word | R10: Detecter contexte liste, adapter HTML |
| Pas d'indentation | Le LLM ne genere pas de sous-listes | PowerPoint + Word | R2+R3+R6: Corriger prompts et descriptions tools |
| Pas d'indentation | `insertList` tool plat sans niveaux | Word | R11: Ajouter support sous-niveaux dans insertList |
| Puces sans texte | `stripMarkdownListMarkers` retire les marqueurs | PowerPoint | R10: Ne strip que si puces natives detectees |
| Formatage perdu | `stripRichFormattingSyntax` dans insertTextBox | PowerPoint | R8: Utiliser HTML dans insertTextBox |

---

## 11. RECOMMANDATIONS ADDITIONNELLES

### R10. Detection des puces natives avant insertion (Priorite : TRES HAUTE)

Voir section 10.4 ci-dessus. C'est LA correction qui eliminera les doubles puces.

**Fichiers a modifier** :
- `frontend/src/utils/powerpointTools.ts` - ajouter `hasNativeBullets`, modifier `insertIntoPowerPoint`
- `frontend/src/utils/wordTools.ts` - ajouter `isInsertionInList`, adapter `insertText` et `replaceSelectedText`

### R11. Supporter les sous-niveaux dans le tool insertList de Word (Priorite : MOYENNE)

**Probleme actuel** : Le tool `insertList` de Word n'accepte qu'un tableau plat d'items.

**Implementation** :
```typescript
// wordTools.ts - insertList ameliore
inputSchema: {
  type: 'object',
  properties: {
    items: {
      type: 'array',
      description: 'Array of list items. Each item can be a string or an object {text, children} for nested sub-items.',
      items: {
        oneOf: [
          { type: 'string' },
          {
            type: 'object',
            properties: {
              text: { type: 'string' },
              children: { type: 'array', items: { type: 'string' } }
            }
          }
        ]
      },
    },
    listType: { type: 'string', enum: ['bullet', 'number'] },
  },
  required: ['items', 'listType'],
},
executeWord: async (context, args) => {
  const { items, listType } = args
  const markdownLines: string[] = []
  for (const item of items) {
    if (typeof item === 'string') {
      markdownLines.push(listType === 'bullet' ? `- ${item}` : `${markdownLines.length + 1}. ${item}`)
    } else {
      markdownLines.push(listType === 'bullet' ? `- ${item.text}` : `${markdownLines.length + 1}. ${item.text}`)
      if (item.children) {
        for (const child of item.children) {
          markdownLines.push(listType === 'bullet' ? `  - ${child}` : `  1. ${child}`)
        }
      }
    }
  }
  range.insertHtml(renderOfficeRichHtml(markdownLines.join('\n')), 'After')
  await context.sync()
  return `Successfully inserted ${listType} list with ${items.length} items`
}
```

**Fichiers a modifier** :
- `frontend/src/utils/wordTools.ts` - tool `insertList`

### R12. Uniformiser les prompts des 3 hotes avec instructions de formatage communes (Priorite : HAUTE)

**Probleme** : Chaque hote a un prompt radicalement different sur le formatage. Il faut un bloc commun.

**Implementation** :
```typescript
// useAgentPrompts.ts - Bloc commun pour tous les hotes
const COMMON_FORMATTING_INSTRUCTIONS = `
## Output Formatting Rules
When generating content that will be inserted into the document:
- Use standard Markdown syntax exclusively. Do NOT use raw HTML tags.
- **Bold**: Use **text** for emphasis
- *Italic*: Use *text* for nuance
- Bullet lists: Use "- " prefix. Each item on its own line.
- Numbered lists: Use "1. ", "2. ", etc.
- Nested sub-items: Indent with exactly 2 spaces before the marker:
  - Level 1: "- Item"
  - Level 2: "  - Sub-item"
  - Level 3: "    - Sub-sub-item"
- Headings: Use # for level 1, ## for level 2, etc.
- NEVER mix bullet symbols. Use "-" consistently, never "*" or "+".
- NEVER put an empty line between consecutive list items of the same level.`

// Puis dans chaque prompt hote :
const wordAgentPrompt = (lang: string) => `...
${COMMON_FORMATTING_INSTRUCTIONS}
...`

const powerPointAgentPrompt = (lang: string) => `...
${COMMON_FORMATTING_INSTRUCTIONS}
...`

const outlookAgentPrompt = (lang: string) => `...
${COMMON_FORMATTING_INSTRUCTIONS}
...`
```

**Fichiers a modifier** :
- `frontend/src/composables/useAgentPrompts.ts` - ajouter `COMMON_FORMATTING_INSTRUCTIONS` et l'integrer dans les 3 prompts

### R13. Clarifier les tools Outlook (insertion texte vs HTML) (Priorite : BASSE)

**Probleme** : Deux tools qui font quasi la meme chose avec des noms confusants.
- `insertTextAtCursor` : nom dit "text" mais insere du HTML (via renderOfficeRichHtml)
- `insertHtmlAtCursor` : insere du HTML brut

**Solution** : Renommer et clarifier les descriptions :
```typescript
insertTextAtCursor: {
  description: 'Insert Markdown-formatted text at the cursor. '
    + 'The text is automatically converted to rich HTML. '
    + 'Use Markdown: **bold**, *italic*, - bullets, 1. numbers.',
  // ...
},
insertHtmlAtCursor: {
  description: 'Insert raw HTML at the cursor. Only use this for '
    + 'pre-formatted HTML content. For normal text, prefer insertTextAtCursor.',
  // ...
},
```

**Fichiers a modifier** :
- `frontend/src/utils/outlookTools.ts`

---

## 12. PLAN D'IMPLEMENTATION MIS A JOUR

### Phase 1 - Corrections de prompts (impact immediat, ~2h)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R2 | Corriger descriptions tools PowerPoint | powerpointTools.ts | 15 min |
| R3 | Ameliorer prompt PowerPoint | useAgentPrompts.ts | 30 min |
| R6 | Harmoniser prompt Word (Markdown) | useAgentPrompts.ts | 30 min |
| R9 | Enrichir GLOBAL_STYLE_INSTRUCTIONS | constant.ts | 15 min |
| R12 | Bloc commun de formatage pour les 3 hotes | useAgentPrompts.ts | 30 min |
| R13 | Clarifier tools Outlook | outlookTools.ts | 15 min |

### Phase 2 - Elimination des doubles puces (~3-4h)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R10 | Detection puces natives PowerPoint | powerpointTools.ts | 2h |
| R10 | Detection contexte liste Word | wordTools.ts | 1h |
| R8 | HTML dans insertTextBox PowerPoint | powerpointTools.ts | 1h |

### Phase 3 - Ameliorations du pipeline HTML (~3h)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R4 | Gestion des `<br>` dans les listes | officeRichText.ts | 1-2h |
| R5 | Preserver sauts de ligne multiples | officeRichText.ts | 30 min |
| R11 | Sous-niveaux dans insertList Word | wordTools.ts | 1h |

### Phase 4 - Heritage des styles (amelioration profonde, ~4h)

| # | Action | Fichier | Effort |
|---|--------|---------|--------|
| R1 | Heritage styles du document Word | officeRichText.ts + common.ts | 2-3h |
| R7 | Extensions Markdown-it supplementaires | officeRichText.ts + package.json | 1h |

---

## 13. CONCLUSION FINALE

Les problemes de puces et de mise en forme dans KickOffice ont **3 couches de causes** :

1. **Couche prompts** (80% du probleme visible) : Les prompts ne guident pas le LLM sur le format Markdown attendu. Le prompt PowerPoint est quasi-vide. Le prompt Word demande du HTML mais le pipeline attend du Markdown. Les descriptions de tools PowerPoint mentent sur les capacites. **La Phase 1 resout ca.**

2. **Couche insertion** (15% du probleme) : Pas de detection des puces natives des shapes/paragraphes avant insertion. Le HTML `<ul><li>` s'ajoute aux puces natives au lieu de les remplacer. **La Phase 2 resout ca.**

3. **Couche pipeline HTML** (5% du probleme) : Styles CSS hardcodes au lieu d'heriter du document, gestion incomplete des `<br>` dans les listes. **Les Phases 3 et 4 resolvent ca.**

**Outlook fonctionne mieux** car les emails HTML n'ont pas de concept de "puces natives de shape" et le chemin HTML -> email est le plus naturel. Il sert de reference pour uniformiser Word et PowerPoint.

L'approche de reference confirme que le pipeline Markdown -> HTML enrichi avec styles inline est la bonne direction. La difference cle est que la solution de reference a un meilleur controle sur le HTML final grace a la manipulation DOM (HtmlAgilityPack) et l'heritage des styles du document, deux choses que KickOffice peut reproduire avec `DOMParser` et l'API Word.js/PowerPoint.js.

---

## 14. FONCTIONNALITES DE MISE EN FORME MANQUANTES DANS KICKOFFICE

### 14.1 Rendu HTML natif Word avec DOM traversal (HTMLToWord.vb)

**Approche de reference** :
La solution de reference a un moteur complet `ParseHtmlNode` / `RenderInline` qui traverse l'arbre DOM HTML noeud par noeud et genere du contenu Word natif avec l'API COM Interop. Chaque element HTML est converti en formatage Word reel :

| Element HTML | Rendu Word natif reference | KickOffice |
|---|---|---|
| `<strong>/<b>` | `range.Font.Bold = True` | Via `insertHtml` (indirect) |
| `<em>/<i>` | `range.Font.Italic = True` | Via `insertHtml` (indirect) |
| `<u>` | `range.Font.Underline = wdUnderlineSingle` | Via `insertHtml` (indirect) |
| `<del>/<s>` | `range.Font.StrikeThrough = True` | **ABSENT** |
| `<sub>` | `range.Font.Subscript = True` | Via custom Markdown `~text~` |
| `<sup>` | `range.Font.Superscript = True` | Via custom Markdown `^text^` |
| `<code>` inline | Courier New 10pt + fond gris | **Style CSS seulement** |
| `<pre>` code block | Courier New 10pt + indent 14pt | **Style CSS seulement** |
| `<br>` | Saut de ligne souple `ChrW(11)` (Shift+Enter) | **Saut de paragraphe dur** |
| `<img>` | `InlineShapes.AddPicture` (local ou download URL) | **ABSENT dans le rendu HTML** |
| `<a href>` | Hyperlink Word natif | Via `insertHtml` (indirect) |
| `<h1>` a `<h6>` | Styles Word builtin `wdStyleHeading1-6` | **Taille CSS seulement, pas de style Word** |
| `<blockquote>` | Paragraphe avec indent 0.75cm | **Margin CSS seulement** |
| `<hr>` | Bordure de paragraphe basse | **ABSENT** |
| `<table>` | `Tables.Add` avec cells recursive | Tool `insertTable` separe |
| `<input checkbox>` | ContentControl checkbox natif Word | **ABSENT** |
| `<dl>/<dt>/<dd>` | Term en gras + indent definition | **ABSENT** |
| Footnotes `<sup><a>` | Bookmark + Hyperlink interne | **ABSENT** |
| Emoji | Police "Segoe UI Emoji" + fond bleu | **ABSENT** |

**Impact pour KickOffice** :
La methode `insertHtml()` de Word.js fait la plupart de ce travail automatiquement car Word interprete le HTML. Mais certaines choses ne passent pas bien via `insertHtml` :
- Les **headings** ne deviennent pas des styles Word builtin (juste du texte avec une grande taille)
- Le **`<br>`** cree un paragraphe dur au lieu d'un saut de ligne souple
- Les **code blocks** n'ont pas le bon formatage (police monospace + fond)
- Les **checkboxes** sont rendues en caracteres Unicode au lieu de ContentControls

**Recommandation R14** : Pour les headings, post-traiter l'insertion HTML avec Word.js pour appliquer les styles builtin sur les paragraphes correspondants. Cela resoudrait le probleme de TOC (table des matieres) qui ne fonctionne pas avec du texte simplement en gros caracteres.

### 14.2 Sauvegarde et restauration du formatage (FormatSaveAndRestore.vb)

**Approche de reference** :
Avant de modifier un texte (correction, traduction, etc.), la solution de reference capture l'integralite du formatage du paragraphe :
- Style Word (Normal, Heading1, etc.)
- Police, taille, gras, italique, souligne, couleur
- Format de liste (template, niveau)
- Alignement, espacement, interligne
- SpaceBeforeAuto, SpaceAfterAuto, DisableLineHeightGrid

Apres la modification LLM, la solution de reference restaure ce formatage sur chaque paragraphe, meme si le LLM a renvoye du texte brut.

**Ce que KickOffice ne fait PAS** :
- Quand le LLM modifie un texte via `replaceSelectedText`, il y a un `preserveFormatting` qui restaure juste police/taille/couleur mais PAS :
  - Le style de paragraphe (Heading1, etc.)
  - Le format de liste (bullets natifs, niveau)
  - L'espacement inter-paragraphe
  - L'alignement

**Recommandation R15** : Enrichir `replaceSelectedText` dans `wordTools.ts` pour capturer et restaurer le style de paragraphe complet (au minimum `styleBuiltIn` et `listItem`).

### 14.3 Conversion Markdown -> Word avec preservation du formatage (ConvertRangeToMarkdown)

**Approche de reference** :
Quand la solution de reference envoie du texte au LLM pour correction/traduction, il convertit d'abord le formatage Word en Markdown :
- Gras -> `**text**`
- Italique -> `*text*`
- Souligne -> `__text__`
- Barre -> `~~text~~`
- Surligne -> tags HTML inline
- Headings -> `# heading`
- Listes -> `- item` / `1. item`

Le LLM recoit donc du Markdown structure et peut le modifier tout en preservant le formatage. Quand le resultat revient, Redink re-convertit le Markdown en formatage Word natif.

**Ce que KickOffice ne fait PAS** :
Quand KickOffice envoie le texte selectionne au LLM, il envoie du texte brut (via `getSelectedText` -> `range.text`). Tout le formatage est perdu. Le LLM renvoie du texte sans savoir qu'il y avait du gras, de l'italique, etc.

**Recommandation R16** : Ajouter un mode "selection avec formatage" qui convertit la selection Word en Markdown avant envoi au LLM. Utiliser `getDocumentHtml` (tool existant) puis un convertisseur HTML -> Markdown cote client.

### 14.4 Diff visuel avec tracked changes (Outlook - CompareAndInsertText)

**Approche de reference pour Outlook** :
Apres correction d'un email, Redink genere un **diff visuel** avec :
- Texte insere en **bleu souligne**
- Texte supprime en **rouge barre**

Cela utilise soit `Word.CompareDocuments` (compare complete) soit `DiffPlex` (diff word-level).

**Ce que KickOffice ne fait PAS** :
KickOffice remplace le texte directement sans montrer les differences. L'utilisateur ne peut pas voir ce qui a change.

**Recommandation R17** : Implementer un mode "diff visuel" optionnel pour les corrections dans Outlook et Word. Utiliser une librairie JS comme `diff-match-patch` pour generer le diff, puis encoder les insertions/suppressions en HTML colore (`<span style="color:blue;text-decoration:underline">` / `<span style="color:red;text-decoration:line-through">`).

### 14.5 DocStyle - Application intelligente de styles Word

**Approche de reference** :
Le systeme DocStyle de Redink peut :
1. **Extraire** les styles d'un document modele (police, taille, espacement, listes, bordures) en JSON
2. **Appliquer** ces styles a un autre document en utilisant le LLM pour mapper les paragraphes aux styles
3. **Gerer les numbering restarts** (reprises de numerotation dans les listes)

**Ce que KickOffice ne fait PAS** :
KickOffice n'a aucune fonctionnalite d'application de styles de document. Le formatage est soit hardcode en CSS (via `applyOfficeBlockStyles`), soit laisse a l'interpretation du LLM.

**Recommandation R18** : A terme, permettre au LLM d'appliquer des styles Word builtin via un tool. Le tool `setParagraphFormat` existe mais ne gere pas les styles builtin. Ajouter un parametre `styleBuiltIn` au tool, ou creer un tool `applyStyle`.

### 14.6 Gestion des notes de bas de page (Footnotes)

**Approche de reference** :
Le moteur HTMLToWord gere les footnotes Markdown (extension Markdig) :
- Detecte les `<li id="fn:1">` (definitions de footnotes)
- Cree des bookmarks Word
- Insere des hyperlinks internes (`<sup><a href="#fn:1">`)
- Genere des references croisees navigables

**Ce que KickOffice ne fait PAS** :
Le tool `insertFootnote` existe dans `wordTools.ts` mais :
- Il insere des footnotes "brutes" sans formatage
- Il n'y a pas de gestion des footnotes dans le pipeline Markdown -> HTML
- markdown-it ne genere pas de footnotes par defaut (necessite le plugin `markdown-it-footnote`)

**Recommandation R19** : Ajouter le plugin `markdown-it-footnote` au pipeline `renderOfficeRichHtml` pour que les footnotes Markdown soient correctement rendues en HTML, puis en footnotes Word via `insertHtml`.

### 14.7 Excel - Insertion structuree dans les cellules

**Approche de reference** :
L'insertion dans Excel suit un protocole structure :
- Le LLM repond avec des directives `[Cell: A1] [Formula: =SUM(B1:B10)]` ou `[Value: Hello]`
- Le parser `ParseLLMResponse` extrait ces directives
- `ApplyLLMInstructions` applique formules/valeurs/commentaires cellule par cellule
- Gestion de la localisation des formules (separateurs `,` vs `;`)
- Sauvegarde/restauration de l'etat des cellules (undo)
- Nettoyage RTF/JSON des valeurs

**Ce que KickOffice ne fait PAS** :
KickOffice n'a **aucun support Excel**. Il n'y a pas de `excelTools.ts`, pas de prompts Excel, pas de tools d'insertion dans les cellules.

**Recommandation R20** : Si Excel est dans la roadmap, s'inspirer du protocole structuree `[Cell:][Formula:][Value:]` de Redink. Implementer un `excelTools.ts` avec les tools :
- `getCellValue` / `setCellValue` / `setCellFormula`
- `getSelectedRange` / `insertData` (tableau 2D)
- `addComment` / `formatCells`

### 14.8 Checkboxes et task lists

**Approche de reference** :
- Markdown task lists (`- [x] Done`, `- [ ] Todo`) sont rendues en :
  - Soit des symboles Unicode (☑ / ☐) avec texte
  - Soit des **ContentControl Checkboxes** Word natifs (cochables)

**Ce que KickOffice ne fait PAS** :
- Les task lists Markdown ne sont pas reconnues (pas de plugin `markdown-it-task-lists`)
- Pas de rendu en checkboxes Word

**Recommandation R21** : Ajouter `markdown-it-task-lists` au pipeline. Les checkboxes HTML generees seront interpretees par Word.js `insertHtml` comme des symboles.

### 14.9 Horizontal rules (`<hr>`)

**Approche de reference** :
Les `---` Markdown sont rendus comme un paragraphe vide avec une bordure basse (trait horizontal natif Word).

**Ce que KickOffice ne fait PAS** :
Les `<hr>` HTML sont generes par markdown-it mais `insertHtml` de Word.js les interprete de maniere incoherente (parfois un trait, parfois rien).

**Recommandation R22** : Ajouter un post-traitement dans `applyOfficeBlockStyles` pour convertir les `<hr>` en `<p style="border-bottom:1px solid #000; margin:8px 0;">&nbsp;</p>`.

### 14.10 Code blocks avec fond colore

**Approche de reference** :
Les blocs de code (`<pre>`) recoivent Courier New 10pt + indentation + fond gris (`Shading.BackgroundPatternColor`).

**Ce que KickOffice ne fait PAS** :
Les `<pre>` dans le CSS de `applyOfficeBlockStyles` n'ont pas de style specifique. Le code apparait en police normale sans differenciation visuelle.

**Recommandation R23** : Ajouter dans `applyOfficeBlockStyles` :
```typescript
.replace(/<pre>/gi, '<pre style="font-family:Consolas,monospace; font-size:10pt; background:#f4f4f4; padding:8px; margin:6px 0;">')
.replace(/<code>/gi, '<code style="font-family:Consolas,monospace; font-size:0.9em; background:#f0f0f0; padding:1px 4px;">')
```

---

## 15. TABLEAU COMPARATIF COMPLET SOLUTION DE REFERENCE vs KICKOFFICE

### 15.1 Fonctionnalites de mise en forme

| Fonctionnalite | Redink | KickOffice Word | KickOffice PPT | KickOffice Outlook | Action |
|---|---|---|---|---|---|
| Gras/Italique/Souligne | Natif Word | Via insertHtml | Partiel | Via insertHtml | OK |
| Barre (strikethrough) | Natif Word | **ABSENT du pipeline** | **ABSENT** | **ABSENT** | R14 |
| Subscript/Superscript | Natif Word | Via Markdown custom | Via Markdown custom | Via Markdown custom | OK |
| Headings -> Styles Word | Styles builtin | **Taille CSS seulement** | N/A | N/A | R14 |
| Code inline (fond gris) | Natif Word | **CSS basique** | **ABSENT** | **CSS basique** | R23 |
| Code block (Courier+indent) | Natif Word | **CSS basique** | **ABSENT** | **CSS basique** | R23 |
| Listes a puces | Natif Word + niveaux | Via insertHtml | Double puces | Via insertHtml | R10 |
| Listes numerotees | Natif Word + niveaux | Via insertHtml | **ABSENT** | Via insertHtml | R10 |
| Listes imbriquees | Multi-niveau Word | Possible mais pas guide | **ABSENT** | Possible | R2+R3 |
| Definition lists `<dl>` | Natif Word | **ABSENT** | **ABSENT** | **ABSENT** | R7 |
| Task lists (checkboxes) | ContentControl | **ABSENT** | **ABSENT** | **ABSENT** | R21 |
| Footnotes | Bookmarks + liens | Tool basique | N/A | N/A | R19 |
| Horizontal rule `<hr>` | Bordure paragraphe | **Inconsistant** | N/A | N/A | R22 |
| Tables | Natif Word | Tool insertTable | N/A | Via insertHtml | OK |
| Images inline | InlineShapes | Tool insertImage | N/A | N/A | OK |
| Hyperlinks | Natif Word | Via insertHtml | N/A | Via insertHtml | OK |
| Emoji (police dediee) | Segoe UI Emoji | **ABSENT** | **ABSENT** | **ABSENT** | Optionnel |
| Blockquote (indent) | Paragraphe indente | Via CSS margin | N/A | Via CSS margin | OK |

### 15.2 Fonctionnalites de workflow

| Fonctionnalite | Redink | KickOffice | Action |
|---|---|---|---|
| Heritage styles document | Font/taille/couleur/espacement | **ABSENT** | R1 |
| Sauvegarde/restauration formatage | Style+font+liste+spacing complet | Partiel (font name/size/color) | R15 |
| Selection avec formatage (Markdown) | ConvertRangeToMarkdown | Texte brut seulement | R16 |
| Diff visuel (corrections) | Bleu insere / Rouge supprime | **ABSENT** | R17 |
| DocStyle (templates de style) | Extraction + application JSON/LLM | **ABSENT** | R18 |
| `<br>` -> saut souple (Shift+Enter) | `ChrW(11)` | **Saut dur de paragraphe** | R14 |
| Support Excel | Complet (formules, valeurs, commentaires) | **ABSENT** | R20 |

---

## 16. PLAN D'IMPLEMENTATION FINAL (TOUTES RECOMMANDATIONS)

### Phase 1 - Corrections de prompts (impact immediat, ~2h)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| R2 | Corriger descriptions tools PowerPoint | powerpointTools.ts | 15 min |
| R3 | Ameliorer prompt PowerPoint | useAgentPrompts.ts | 30 min |
| R6 | Harmoniser prompt Word (Markdown) | useAgentPrompts.ts | 30 min |
| R9 | Enrichir GLOBAL_STYLE_INSTRUCTIONS | constant.ts | 15 min |
| R12 | Bloc commun de formatage 3 hotes | useAgentPrompts.ts | 30 min |
| R13 | Clarifier tools Outlook | outlookTools.ts | 15 min |

### Phase 2 - Elimination des doubles puces (~3-4h)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| R10 | Detection puces natives PowerPoint | powerpointTools.ts | 2h |
| R10 | Detection contexte liste Word | wordTools.ts | 1h |
| R8 | HTML dans insertTextBox PowerPoint | powerpointTools.ts | 1h |

### Phase 3 - Ameliorations du pipeline HTML (~4h)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| R4 | Gestion des `<br>` dans les listes | officeRichText.ts | 1-2h |
| R5 | Preserver sauts de ligne multiples | officeRichText.ts | 30 min |
| R11 | Sous-niveaux dans insertList Word | wordTools.ts | 1h |
| R23 | Styles code blocks + code inline | officeRichText.ts | 30 min |
| R22 | Horizontal rules | officeRichText.ts | 15 min |
| R21 | Plugin task lists | officeRichText.ts + package.json | 30 min |

### Phase 4 - Heritage des styles + formatage avance (~5h)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| R1 | Heritage styles du document Word | officeRichText.ts + common.ts | 2-3h |
| R7 | Extensions Markdown-it (deflist, footnotes) | officeRichText.ts + package.json | 1h |
| R14 | Headings -> styles Word builtin | wordTools.ts | 1h |
| R15 | Restauration formatage complete | wordTools.ts | 1h |

### Phase 5 - Fonctionnalites avancees (optionnel, ~8h)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| R16 | Selection avec formatage Markdown | wordTools.ts + useOfficeSelection.ts | 2h |
| R17 | Diff visuel pour corrections | Nouveau composable + lib diff | 3h |
| R19 | Footnotes Markdown completes | officeRichText.ts + package.json | 1h |
| R18 | Tool applyStyle | wordTools.ts | 1h |
| R20 | Support Excel basique | Nouveau excelTools.ts | 4h+ |

---

## 17. CONCLUSION MISE A JOUR

Au-dela des problemes de puces et de mise en forme de base, l'analyse de Redink revele **3 categories de fonctionnalites manquantes** dans KickOffice :

### A. Fonctionnalites de mise en forme immediatement applicables (Phases 1-3)
- Corrections de prompts et descriptions de tools (R2, R3, R6, R9, R12, R13)
- Detection des puces natives (R10)
- Styles code blocks, horizontal rules, task lists (R21, R22, R23)

### B. Fonctionnalites de qualite de rendu (Phase 4)
- Heritage des styles du document (R1)
- Headings en styles Word builtin au lieu de tailles CSS (R14)
- Restauration complete du formatage apres modification (R15)
- Plugins Markdown avances - definition lists, footnotes (R7, R19)

### C. Fonctionnalites avancees (Phase 5)
- Envoi du formatage au LLM via Markdown (R16) - le LLM pourra preserver gras/italique
- Diff visuel pour les corrections (R17) - comme Redink dans Outlook
- Application intelligente de styles (R18)
- Support Excel (R20) - si dans la roadmap

Les Phases 1 a 3 resolvent les problemes actuels. La Phase 4 eleve la qualite au niveau de Redink. La Phase 5 ajoute des fonctionnalites differenciantes.
