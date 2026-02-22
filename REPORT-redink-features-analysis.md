# RAPPORT : Fonctionnalites Redink a reprendre dans KickOffice (hors mise en forme)

> Date : 22 fevrier 2026
> Scope : 6 fonctionnalites selectionnees parmi ~50 identifiees dans Redink
> Hotes : Word, PowerPoint, Outlook, Excel

---

## TABLE DES MATIERES

1. [Freestyle avance (modes prefixes)](#1-freestyle-avance-modes-prefixes)
2. [Generation de slides depuis Word](#2-generation-de-slides-depuis-word)
3. [Traduction/Correction de document entier](#3-traductioncorrection-de-document-entier)
4. [CSV Analyzer (Excel)](#4-csv-analyzer-excel)
5. [Suggest Titles (Word)](#5-suggest-titles-word)
6. [Sum-up Revisions (Word)](#6-sum-up-revisions-word)
7. [Tableau comparatif](#7-tableau-comparatif)
8. [Plan d'implementation](#8-plan-dimplementation)

---

## 1. FREESTYLE AVANCE (MODES PREFIXES)

### 1.1 Ce que fait Redink

Le Freestyle de Redink est un systeme universel de prompting avec **14 prefixes de sortie** et **15+ triggers inline**. L'utilisateur tape son prompt avec un prefixe optionnel qui determine ou et comment le resultat est affiche.

#### Prefixes de sortie (debut du prompt)

| Prefixe | Comportement |
|---------|-------------|
| `Replace:` | Remplace le texte selectionne par le resultat LLM |
| `Append:` / `Add:` | Ajoute le resultat apres la selection (original preserve) |
| `Markup:` | Genere du Track Changes Word montrant les differences |
| `MarkupDiff:` | Genere un diff unifie |
| `MarkupWord:` | Force le Track Changes natif Word |
| `Clipboard:` / `Clip:` | Place le resultat dans le presse-papier (fenetre separee) |
| `Newdoc:` | Cree un nouveau document avec le resultat |
| `Pane:` | Affiche le resultat dans un panneau lateral (lecture seule) |
| `Bubbles:` | Insere le resultat comme commentaires Word |
| `Reply:` / `Pushback:` | Repond aux commentaires existants |
| `Slides:` | Ajoute le resultat comme slides PowerPoint |
| `Chart:` | Cree un diagramme draw.io |
| `File:` / `Files:` | Traite des fichiers au lieu du texte du document |
| `Pure:` | Utilise le prompt directement comme system prompt (sans traitement) |

#### Triggers inline (dans le prompt)

| Trigger | Effet |
|---------|-------|
| `(all)` | Selectionne tout le document au lieu de la selection |
| `{doc}` | L'utilisateur choisit un fichier a inclure (txt, docx, pdf, pptx...) |
| `{dir}` | L'utilisateur choisit un repertoire ; tous les fichiers supportes sont charges |
| `{[chemin]}` | Inclut un fichier/repertoire au chemin specifie |
| `(mystyle)` | Applique un profil de style d'ecriture selectionne |
| `(lib)` | Recherche dans une bibliotheque de connaissances (RAG) |
| `(net)` | Recherche internet avec confirmation utilisateur |
| `(iterate)` | Traitement par morceaux (l'utilisateur specifie le nombre de paragraphes par chunk) |
| `(bubbles)` | Extrait le texte des commentaires existants et l'inclut dans le prompt |
| `(noformat)` / `(nf)` | Desactive la preservation du formatage |
| `(keepformat)` / `(kf)` | Force la preservation du formatage caractere |
| `(keepparaformat)` / `(kpf)` | Force la preservation du formatage paragraphe |
| `(file)` / `(clip)` | Attache un fichier ou le contenu du presse-papier a la requete LLM |
| `(multimodel)` | Execute le prompt sur plusieurs modeles en parallele |
| `(adddoc)` | Inclut le contenu d'autres documents Word ouverts |

#### Inclusion de fichiers externes

Le systeme supporte 25+ extensions : `.txt`, `.rtf`, `.doc`, `.docx`, `.pdf`, `.pptx`, `.csv`, `.json`, `.xml`, `.html`, `.md`, `.py`, `.js`, `.ts`, `.java`, `.cpp`, `.sql`, `.yaml`...

Les fichiers sont encapsules en XML pour le LLM :
```xml
<document1 filename="example.txt">
[contenu du fichier]
</document1>
<document2 directory="C:\MyDocs" filename="report.pdf">
[contenu du fichier]
</document2>
```

Pour les PDF, une detection OCR est proposee (par fichier ou globale pour un repertoire).

#### Interface utilisateur

Quand du texte est selectionne, 3 boutons rapides sont proposes : `Clipboard:`, `Pane:`, `MarkupDiff:`. Sans selection : `Clipboard:`, `Pane:`.

Un **prompt library** est integre : si le prompt est vide et la bibliotheque activee, un selecteur de prompts s'affiche.

### 1.2 Ce que KickOffice a deja

- **Modes draft/immediate** sur les quick actions Outlook et Excel (`src/types/chat.ts` lignes 18-31)
- **Mode draft** : pre-remplit le champ de saisie avec un prefixe
- **Mode immediate** : execute directement avec un system prompt custom

**Ce qui manque** :
- Aucun routage de sortie (clipboard, pane, append, markup, slides, chart)
- Aucune inclusion de fichiers externes (`{doc}`, `{dir}`)
- Aucun trigger inline (`(all)`, `(iterate)`, `(net)`, `(lib)`)
- Aucune bibliotheque de prompts reutilisables
- Pas de mode "traitement par morceaux" pour les longs documents

### 1.3 Recommandations pour KickOffice

**F1-A : Modes de sortie multiples (priorite haute)**

Ajouter un systeme de prefixes dans le champ de saisie du chat. Implementer au minimum :

| Prefixe KickOffice | Equivalent Redink | Implementation |
|---|---|---|
| `replace:` | `Replace:` | Utiliser `replaceSelectedText` existant |
| `append:` | `Append:` | Inserer apres la selection via `insertTextAfterSelection` |
| `comment:` | `Bubbles:` | Utiliser `addComment` existant |
| `slides:` | `Slides:` | Enchainer avec tools PowerPoint existants |
| `pane:` | `Pane:` | Afficher dans le panel lateral du chat (deja la) |

**Fichiers concernes** : `useAgentLoop.ts` (detection du prefixe), `useAgentPrompts.ts` (adaptation du system prompt selon le prefixe)

**F1-B : Inclusion de fichiers (priorite moyenne)**

Le plus impactant serait de permettre l'inclusion de fichiers dans le prompt. Via l'API Office.js, on peut lire le contenu de fichiers ouverts. Pour les fichiers externes, un mecanisme d'upload serait necessaire.

**F1-C : Traitement par morceaux / iterate (priorite moyenne)**

Ajouter un mode qui decoupe le document en chunks de N paragraphes et traite chaque chunk sequentiellement. Utile pour les longs documents (corrections, traductions).

**Fichiers concernes** : Nouveau composable `useChunkedProcessing.ts`

---

## 2. GENERATION DE SLIDES DEPUIS WORD

### 2.1 Ce que fait Redink

Redink peut generer une presentation PowerPoint complete a partir d'un document Word. Le workflow :

1. **Extraction de la structure existante** : Si un PPTX existe, extraction des metadonnees en JSON (slides, layouts, placeholders, dimensions)
2. **Envoi au LLM** : Le JSON de la presentation + le texte du document sont envoyes au LLM
3. **Plan d'action JSON** : Le LLM repond avec un plan structure :

```json
{
  "actions": [
    {
      "op": "add_slide",
      "anchor": { "mode": "at_end" },
      "layoutRelId": "rId2",
      "elements": [
        { "type": "title", "text": "Introduction", "style": { "fontSize": 44, "bold": true } },
        { "type": "bullet_text", "bullets": [
          { "text": "Point 1", "level": 0 },
          { "text": "Sous-point", "level": 1 }
        ]},
        { "type": "shape", "shapeType": "rectangle", "transform": { "x": 0.1, "y": 0.2, "width": 0.8, "height": 0.3 }, "fill": { "type": "solid", "color": "#FF0000" } },
        { "type": "svg_icon", "svg": "<svg>...</svg>", "transform": { "x": 0.1, "y": 0.1, "width": 0.1, "height": 0.1 } }
      ],
      "notes": "Notes du presentateur"
    }
  ]
}
```

4. **Application OpenXML** : Chaque action est executee directement sur le fichier PPTX :
   - Clonage de layouts templates avec fallback intelligent
   - Creation d'elements (titres, texte, puces multi-niveaux, formes geometriques, icones SVG)
   - Notes du presentateur
   - Coordonnees en pourcentage ou en EMU

#### Elements supportes

| Type | Description |
|------|-------------|
| `title` | Titre de la slide (placeholder Title/CenteredTitle) |
| `text` | Texte sans puces dans un placeholder |
| `bullet_text` | Liste a puces multi-niveaux (0-8 niveaux) |
| `shape` | 20+ formes geometriques (rectangle, ovale, fleches, flowchart, chevron...) avec remplissage et contour |
| `svg_icon` | Icone SVG embarquee |

#### Gestion des templates

Resolution de layout par hierarchie : ID de relation > URI > nom lisible. Fallback : layout cover-like > layout par defaut > exception.

### 2.2 Ce que KickOffice a deja

- **Tools PowerPoint** : `addSlide`, `insertTextBox`, `insertImage`, `getSlideContent`, `getSlideCount`, `getAllSlidesOverview`
- **Quick actions** : bullets, speakerNotes, punchify, proofread, visual

**Ce qui manque** :
- Aucune generation de presentation multi-slides a partir de contenu
- Pas de clonage de layouts
- Pas de creation de formes geometriques
- Pas d'icones SVG
- Pas de systeme d'ancrage (inserer apres slide X)
- Les tools existants sont unitaires (une slide a la fois)

### 2.3 Recommandations pour KickOffice

**F2-A : Quick action "Generate Presentation" (priorite haute)**

Ajouter une quick action Word qui :
1. Lit le contenu du document Word (via `getDocumentContent`)
2. Envoie au LLM avec un prompt specifique demandant un plan de presentation
3. Le LLM utilise les tools `addSlide` + `insertTextBox` existants en sequence pour creer chaque slide

**Avantage** : Aucun nouveau tool necessaire. Le LLM orchestre les tools existants.

**Fichiers concernes** : `constant.ts` (nouvelle quick action), `useAgentPrompts.ts` (nouveau prompt)

**F2-B : Tool `createPresentation` enrichi (priorite moyenne)**

Pour aller plus loin, creer un tool qui prend un JSON de plan de presentation et genere toutes les slides en une seule operation. Cela serait plus rapide et plus fiable qu'une orchestration tool-by-tool.

**Fichiers concernes** : `powerpointTools.ts` (nouveau tool)

**F2-C : Formes et icones (priorite basse)**

Ajouter des tools pour creer des formes geometriques et inserer des icones SVG. L'API Office.js `slide.shapes.addGeometricShape()` le supporte.

---

## 3. TRADUCTION/CORRECTION DE DOCUMENT ENTIER

### 3.1 Ce que fait Redink

Redink peut traduire ou corriger un document Word **entier** paragraphe par paragraphe avec **100% de preservation du formatage**.

#### Architecture

1. **Manipulation OpenXML directe** : Seuls les noeuds `<w:t>` (texte) sont modifies. Tout le reste est intact :
   - Styles (`<w:pStyle>`, `<w:rStyle>`)
   - Proprietes de police (`<w:rPr>`)
   - Proprietes de paragraphe (`<w:pPr>`)
   - Champs, signets, formes, mise en page

2. **Traitement par lots** :
   - **Taille de lot** : 10 paragraphes ou 15 000 caracteres max
   - **Fenetre de contexte** : 3 paragraphes precedents (deja traduits) + 2 paragraphes suivants (non traduits) envoyes comme contexte
   - **Ajustement dynamique** : Le lot est reduit si le nombre de caracteres depasse la limite

3. **Distribution proportionnelle** : Quand le texte est reparti sur plusieurs runs Word (ex: un mot en gras au milieu d'une phrase), la traduction est redistribuee proportionnellement sur les runs en preservant les frontieres de formatage.

4. **Preservation des espaces** : Attribut `xml:space="preserve"` sur les noeuds texte avec espaces en debut/fin.

5. **Scope** :
   - Document unique (`.doc` converti en `.docx` automatiquement)
   - Repertoire entier (recursif optionnel, deduplication)
   - Traite aussi : `comments.xml`, headers, footers, footnotes, endnotes

6. **Mode correction** : Genere un document Word Compare montrant toutes les modifications via `CompareDocuments()` natif.

### 3.2 Ce que KickOffice a deja

- **Quick action `translate`** dans Word : traduction de la selection uniquement
- **Quick action `proofread`** : correction via commentaires (pas de remplacement)
- **Tool `replaceSelectedText`** : remplacement avec `preserveFormatting` partiel (police/taille/couleur)
- **Tool `getDocumentContent`** : lecture du contenu complet

**Ce qui manque** :
- Aucun traitement batch paragraphe par paragraphe
- Aucune traduction de document entier
- Pas de fenetre de contexte entre les lots
- Pas de preservation du formatage a 100% (styles, listes, espacement perdus)
- Pas de generation de document Compare

### 3.3 Recommandations pour KickOffice

**F3-A : Tool `translateFullDocument` (priorite haute)**

Implementer un workflow batch dans un nouveau tool ou composable :

1. Lire tous les paragraphes via `context.document.body.paragraphs`
2. Grouper en lots de 10 paragraphes / 15K caracteres
3. Pour chaque lot, envoyer au LLM avec :
   - System prompt de traduction
   - 3 paragraphes de contexte precedent (deja traduits)
   - 2 paragraphes de contexte suivant (originaux)
4. Parser la reponse indexee
5. Remplacer chaque paragraphe via `paragraph.insertText(translatedText, Word.InsertLocation.replace)`

**Challenge API Office.js** : `insertText` avec `replace` preserve certains formatages de paragraphe (style, alignement) mais pas les formatages inline (gras sur un mot). Pour une preservation a 100%, il faudrait manipuler les runs individuellement, ce qui est plus complexe avec Office.js qu'avec OpenXML.

**Fichiers concernes** : Nouveau composable `useDocumentBatchProcessing.ts`, `wordTools.ts` (nouveau tool), `constant.ts` (nouvelle quick action)

**F3-B : Fenetre de contexte (priorite haute)**

Le pattern "3 avant + 2 apres" est essentiel pour la qualite de traduction. Sans contexte, le LLM fait des choix de terminologie inconsistants d'un lot a l'autre.

**F3-C : Mode correction avec Compare (priorite basse)**

Word.js ne supporte pas `CompareDocuments()`. Alternative : utiliser le diff visuel (insertions en bleu, suppressions en rouge) directement dans le document via `insertHtml`.

---

## 4. CSV ANALYZER (EXCEL)

### 4.1 Ce que fait Redink

Le CSV Analyzer de Redink permet d'analyser des fichiers CSV/TXT volumineux via le LLM avec des resultats structures dans Excel.

#### Workflow

1. **Chargement du fichier** :
   - Selection de fichier via dialog
   - Detection d'encodage (UTF-8, UTF-16, BOM)
   - Parsing CSV avec support des champs quotes et separateurs custom
   - Lecture efficace (FileStream 1MB buffer, comptage de lignes sans charger tout en memoire)

2. **Configuration** :
   - **Separateur** : configurable (`,`, `;`, `\t`, etc.)
   - **Taille de chunk** : 50 lignes par defaut (configurable)
   - **Selection de colonnes** : l'utilisateur choisit un sous-ensemble de colonnes a analyser
   - **Plage de lignes** : debut/fin optionnels

3. **Traitement par chunks** :
   - Chaque chunk inclut un header synthetique (noms de colonnes)
   - Format : `LineInFile | Column1 | Column2 | ...`
   - Prompt enveloppe dans des tags `<LINESTOPROCESS>...</LINESTOPROCESS>`
   - Retry avec backoff (750ms * numero de retry)

4. **Format de reponse LLM** :
   - `line@@result§§§line@@result` (separateurs specifiques)
   - Token `[NORESULT]` pour chunks sans resultats

5. **Sortie Excel structuree** :
   - En-tete du rapport (titre, metadata, modele utilise, date)
   - Colonnes : "Line(s)" | "Result"
   - Texte wrappé, alignement haut-gauche
   - Pied de page avec statistiques

### 4.2 Ce que KickOffice a deja

- **Quick action `analyze`** pour Excel : analyse les cellules selectionnees uniquement
- Aucun mecanisme d'upload de fichier
- Aucun outil de parsing CSV

### 4.3 Recommandations pour KickOffice

**F4-A : Quick action "Analyze CSV" (priorite moyenne)**

Puisque KickOffice est un add-in Office.js, l'approche la plus naturelle serait :
1. L'utilisateur ouvre le CSV dans Excel (Excel sait deja ouvrir les CSV)
2. Une quick action "Analyze Data" lit la plage selectionnee ou tout le worksheet
3. Le LLM analyse les donnees par chunks
4. Les resultats sont ecrits dans un nouveau worksheet

**Avantage** : Pas besoin d'implementer un parser CSV ni un systeme d'upload de fichier. Excel fait le travail.

**Fichiers concernes** : `constant.ts` (nouvelle quick action Excel), `excelTools.ts` (nouveaux tools si necessaire)

**F4-B : Traitement par chunks pour Excel (priorite moyenne)**

Le pattern de chunking est le meme que pour la traduction de document (F3-A). Un composable generique `useChunkedProcessing.ts` pourrait servir aux deux cas.

---

## 5. SUGGEST TITLES (WORD)

### 5.1 Ce que fait Redink

Genere plusieurs suggestions de titres pour le texte selectionne dans differents registres :
- Formel / professionnel
- Blog / informatif
- Informel / creatif
- Humoristique

Les resultats sont affiches en mode "analyse" (dans des commentaires/bulles Word ou un panneau), pas en remplacement du texte.

Le prompt systeme utilise est `SP_SuggestTitles` (interpole au runtime).

### 5.2 Ce que KickOffice a deja

- Aucune fonctionnalite de suggestion de titres
- Le tool `addComment` existe et pourrait servir pour afficher les suggestions
- Le systeme de quick actions est parfaitement adapte pour ajouter cette feature

### 5.3 Recommandations pour KickOffice

**F5-A : Quick action "Suggest Titles" (priorite haute - tres simple)**

C'est la fonctionnalite la plus simple a implementer : une seule quick action avec un bon prompt.

```typescript
// Dans constant.ts, section WORD_QUICK_ACTIONS
{
  id: 'suggestTitles',
  label: 'Suggest titles',
  icon: 'lightbulb',
  prompt: {
    system: `You are a professional editor. Based on the selected text, suggest 5 title options in different styles:
1. **Formal/Professional** - Suitable for reports and official documents
2. **Descriptive/Informative** - Clear and factual
3. **Creative/Engaging** - Catchy and memorable
4. **Short/Punchy** - Maximum 6 words
5. **Question-based** - Formulated as a compelling question

Present each title with its category label. Respond in the same language as the input text.`,
    user: 'Suggest titles for this text:\n\n{selection}'
  }
}
```

Le resultat s'affiche dans le chat (pas de remplacement necessaire).

**Fichiers concernes** : `constant.ts` uniquement

---

## 6. SUM-UP REVISIONS (WORD)

### 6.1 Ce que fait Redink

Resume intelligemment tous les tracked changes (revisions) d'un document Word.

#### Workflow

1. **Portee** : Selection ou document entier
2. **Filtre de date** : L'utilisateur specifie la date la plus ancienne (defaut : 7 jours)
3. **Extraction des revisions** :
   - Filtre les revisions purement cosmetiques (changements de style, proprietes de paragraphe)
   - Extrait les revisions substantielles :
     - Insertions : `<ins>texte</ins>`
     - Suppressions : `<del>texte</del>`
     - Deplacements : `<del>[moved from:]texte</del>` / `<ins>[moved to:]texte</ins>`
   - Tri par position dans le document

4. **Collecte des commentaires** :
   - Si filtre de date actif : commentaires a partir de cette date
   - Si pas de filtre : commentaires dans les 60 minutes precedant la premiere revision (capture les commentaires contextuels)
   - Format : `<comment author="Nom" scope="TexteCible">TexteCommentaire</comment>`

5. **Analyse LLM** : Revisions + commentaires envoyes au LLM qui genere un resume structure en Markdown

6. **Affichage** : Markdown converti en HTML et affiche dans une fenetre avec metadonnees (plage de dates, portee)

### 6.2 Ce que KickOffice a deja

- **Tool `addComment`** : peut ajouter des commentaires
- **Tool `getDocumentContent`** : lit le contenu mais PAS les revisions
- Aucun acces aux tracked changes via les tools actuels

**Ce qui manque** :
- Aucun tool pour lire les tracked changes Word
- Aucun tool pour lire les commentaires existants
- Aucune quick action pour resumer les revisions

### 6.3 Recommandations pour KickOffice

**F6-A : Tool `getTrackedChanges` (priorite haute)**

L'API Word.js supporte la lecture des tracked changes via `document.body.getRange().getTrackedChanges()` (API disponible dans WordApi 1.6+).

```typescript
// Dans wordTools.ts
{
  name: 'getTrackedChanges',
  description: 'Get all tracked changes (revisions) in the document or selection',
  parameters: { scope: { type: 'string', enum: ['selection', 'document'] } },
  execute: async (params, context) => {
    const range = params.scope === 'selection'
      ? context.document.getSelection()
      : context.document.body.getRange();
    const changes = range.trackedChanges;
    changes.load('items');
    await context.sync();
    // Format changes for LLM
  }
}
```

**Note** : L'API `TrackedChanges` de Word.js est relativement recente. Verifier la compatibilite avec les versions Word ciblees.

**F6-B : Tool `getComments` (priorite haute)**

```typescript
// Dans wordTools.ts
{
  name: 'getComments',
  description: 'Get all comments in the document',
  execute: async (params, context) => {
    const comments = context.document.body.getComments();
    // ...
  }
}
```

**F6-C : Quick action "Summarize Changes" (priorite haute)**

Une fois les tools F6-A et F6-B en place, la quick action est simple :

```typescript
{
  id: 'sumUpRevisions',
  label: 'Summarize changes',
  icon: 'history',
  prompt: {
    system: `You are a document change analyst. Analyze the tracked changes and comments provided and create a structured summary. Group changes by:
1. Content additions (new text)
2. Content deletions (removed text)
3. Content modifications (changed text)
4. Moved sections
Include who made each change if available. Respond in the same language as the document.`,
    user: 'Summarize the tracked changes in this document.'
  }
}
```

Le LLM appellerait `getTrackedChanges` + `getComments` via les tools, puis genererait le resume.

**Fichiers concernes** : `wordTools.ts` (2 nouveaux tools), `constant.ts` (nouvelle quick action)

---

## 7. TABLEAU COMPARATIF

| # | Fonctionnalite | Redink | KickOffice actuel | Effort d'implementation |
|---|---|---|---|---|
| F1 | Freestyle prefixes | 14 prefixes + 15 triggers | Draft/Immediate basique | Moyen (3-5 jours) |
| F2 | Generate Slides depuis Word | Full OpenXML + formes + SVG | Tools unitaires | Faible (quick action) a Moyen (tool batch) |
| F3 | Traduction document entier | 100% formatage, batch, contexte | Selection seulement | Eleve (5-8 jours) |
| F4 | CSV Analyzer | Upload + chunks + Excel output | Rien | Faible si CSV ouvert dans Excel |
| F5 | Suggest Titles | Multi-registre + bulles | Rien | Tres faible (1h) |
| F6 | Sum-up Revisions | Revisions + commentaires + date | Rien | Moyen (2-3 jours pour tools + QA) |

---

## 8. PLAN D'IMPLEMENTATION

### Phase 1 - Gains rapides (1-2 jours)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| F5-A | Quick action "Suggest Titles" | `constant.ts` | 1h |
| F2-A | Quick action "Generate Presentation" (Word -> PPT via tools existants) | `constant.ts`, `useAgentPrompts.ts` | 2-3h |
| F4-A | Quick action "Analyze Data" (CSV ouvert dans Excel) | `constant.ts` | 2-3h |

### Phase 2 - Nouveaux tools (3-5 jours)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| F6-A | Tool `getTrackedChanges` | `wordTools.ts` | 4-6h |
| F6-B | Tool `getComments` | `wordTools.ts` | 2-3h |
| F6-C | Quick action "Summarize Changes" | `constant.ts` | 1h |
| F1-A | Systeme de prefixes de sortie (replace, append, comment) | `useAgentLoop.ts`, `useAgentPrompts.ts` | 1-2 jours |

### Phase 3 - Traitement batch (5-8 jours)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| F3-A | Composable `useDocumentBatchProcessing` | Nouveau fichier | 2-3 jours |
| F3-B | Tool `translateFullDocument` | `wordTools.ts`, `constant.ts` | 2-3 jours |
| F1-C | Mode iterate (traitement par chunks) | `useAgentLoop.ts` | 1-2 jours |

### Phase 4 - Enrichissements (optionnel)

| # | Action | Fichier(s) | Effort |
|---|--------|---------|--------|
| F2-B | Tool `createPresentation` batch | `powerpointTools.ts` | 2-3 jours |
| F2-C | Formes geometriques + SVG | `powerpointTools.ts` | 1-2 jours |
| F1-B | Inclusion de fichiers externes | `useAgentLoop.ts`, UI upload | 3-5 jours |

---

## 9. CONCLUSION

Sur les 6 fonctionnalites retenues :

- **F5 (Suggest Titles)** est implementable en 1h avec une simple quick action
- **F2-A (Generate Slides)** et **F4-A (CSV Analyzer)** sont realisables rapidement en s'appuyant sur les tools existants
- **F6 (Sum-up Revisions)** necessite 2 nouveaux tools Word.js mais a une forte valeur ajoutee
- **F1 (Freestyle prefixes)** transformerait l'UX en donnant a l'utilisateur le controle sur la destination des resultats
- **F3 (Traduction document entier)** est le plus complexe mais aussi le plus impactant pour les utilisateurs travaillant sur de longs documents

L'ordre de priorite suggere (rapport effort/impact) : **F5 > F2-A > F6 > F4-A > F1-A > F3**
