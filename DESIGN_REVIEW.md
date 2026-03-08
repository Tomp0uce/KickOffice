# DESIGN_REVIEW.md — Code Audit v9.0

**Date**: 2026-03-08
**Version**: 9.0
**Scope**: Revue utilisateur globale — Bugs bloquants, problèmes de performance, UX et IA

---

## État de Santé Global (V9.1 — 2026-03-08)

Tous les points critiques et majeurs identifiés dans la V9.0 ont été audités, vérifiés ou corrigés. L'application est maintenant stable sur les flux transversaux (fichiers, PowerPoint et Excel).

---

## ITEMS CRITIQUES (Bloquants ou impactant lourdement l'application)

### GEN-C1 — L'ajout de fichier est cassé [VÉRIFIÉ — OK]

**Statut :** Fonctionnel. Le composant `ChatInput.vue` capte correctement les fichiers et les transmet à `useAgentLoop.ts` via l'événement `submit`. L'extraction est gérée par `uploadFile`.
**Stratégie d'implémentation :**

1. Inspecter le composant de téléchargement (ex: `FileUpload.vue` ou la zone de drag-and-drop).
2. Vérifier que l'événement `@change` de l'input `<input type="file">` capte bien le fichier sélectionné.
3. S'assurer que les objets de fichiers sont correctement convertis (Base64/File) et ajoutés au store (`Pinia`) ou à l'état global contrôlant les pièces jointes (`attachments`).
4. Vérifier la console pour toute erreur de sécurité (CORS) ou limitation de taille qui bloquerait l'ajout.

### PPT-C2 — Boucle infinie lors de la création d'une slide depuis une image [FIXÉ — ARCHITECTURE OK]

**Criticité :** Critique
**Statut :** [FIXÉ — ARCHITECTURE OK]
**Fix :** Implémentation d'un registre d'images (`powerpointImageRegistry`) et d'une persistance de session (`sessionUploadedFiles/Images`). Le LLM ne manipule plus de Base64 mais des noms de fichiers. Le `tokenManager` a été corrigé pour ne plus compter le poids des images Base64 comme du texte. Les fichiers PDF/Docx sont persistés durant toute la session.

**Stratégie d'implémentation :**

1. Modifier `getAllSlidesOverview` (`powerpointTools.ts`) pour qu'il inclue les images dans son résumé sans essayer d'en lire le `textFrame` (ex: ajouter `[Image (Width x Height)]` au texte de la slide).
2. Ajouter une directive dans les instructions système spécifiques à PowerPoint (`useAgentPrompts.ts`) : "Ne vérifiez pas le résultat avec getAllSlidesOverview suite à un appel réussi de insertImageOnSlide. Considérez que l'image a été insérée."

### GEN-C3 — Lenteur excessive après 5 appels d'outils [VÉRIFIÉ — OK]

**Statut :** Optimisé. L'élagage du contexte est implémenté via `prepareMessagesForContext` dans `tokenManager.ts` avec une limite de 1.2M de caractères, préservant les derniers messages et tronquant les résultats d'outils volumineux.
**Stratégie d'implémentation :**

1. **Élagage du contexte (Context Pruning)** : Le contexte LLM grossit de manière exponentielle car les retours des appels d'outils (ex: JSON d'état) sont accumulés dans les prompts successifs. Implémenter une troncature ou un résumé automatique des `tool_results` trop anciens pour alléger la taille du prompt envoyé.
2. Analyser le code de la boucle d'agent (`useAgentLoop.ts`) pour tout processus synchrone ou timeout artificiel s'accumulant après plusieurs outils.

### WD-C4 — Plante sur texte inséré depuis un PDF sans sélection [VÉRIFIÉ — OK]

**Statut :** Fixé. L'outil `insertContent` (`wordTools.ts`) effectue désormais un `insertedRange.select()` automatique après l'insertion, garantissant une sélection valide pour les appels de formatage ultérieurs.
**Stratégie d'implémentation :**

1. Modifier l'outil d'insertion de texte (`insertContent` dans `wordTools.ts` par exemple) pour qu'il retourne et **sélectionne automatiquement** la plage de texte (`Range`) qu'il vient de créer via l'API Office (`range.select()`).
2. Indiquer dans `word.skill.md` que l'agent doit explicitement s'assurer d'avoir un texte sélectionné _avant_ tout appel à `formatText`, ou bien utiliser un outil comme `searchAndFormat` spécifique.

---

## ITEMS MAJEURS (Ergonomie, Comportement inattendu, Problèmes UI)

### XL-M1 — Génération de graphes Excel traite l'axe X comme série de données [FIXÉ — OK]

**Statut :** Amélioré. `excelTools.ts` définit désormais explicitement l'Axe X (Category Axis) via `setCategoryNames` en utilisant la première ligne/colonne de la plage fournie.
**Stratégie d'implémentation :**

1. Modifier la fonction `manageObject` ou `createChart` dans `excelTools.ts`.
2. Au lieu de laisser `Excel.ChartSeriesBy.auto`, passer explicitement les paramètres pour que la 1ère colonne (ou ligne) soit utilisée comme axe des catégories (Axe X) (`chart.axes.categoryAxis.setCategoryNames(...)`).
3. Retirer la plage correspondante des séries de données (DataSeries) pour éviter le doublon visuel.

### PPT-M2 — L'agent n'utilise pas les boîtes (titre, corps) des templates PowerPoint [VÉRIFIÉ — OK]

**Statut :** Fonctionnel. `addSlide` inspecte `placeholderFormat` pour injecter le contenu dans les zones natives du template au lieu de créer des zones flottantes.
**Stratégie d'implémentation :**

1. Étendre les paramètres de l'outil `addSlide` pour récupérer les "Shapes" de type placeholder une fois la slide créée.
2. Injecter les paramètres texte (`title`, `body`) directement dans les placeholders au lieu de créer de nouvelles zones de texte flottantes "sauvages".

### GEN-M3 — Taille du Taskpane bloquée à 300px au lieu de 450px [FIXÉ — OK]

**Statut :** Manifestes régénérés. La balise `<RequestedWidth>450</RequestedWidth>` est présente dans les templates et les manifestes finaux ont été mis à jour via `generate-manifests.js`.

**Stratégie d'implémentation :**

1. Documenter ou ajouter un script pre-start (`npm run build:manifests`) pour s'assurer que `generate-manifests.js` est systématiquement exécuté avant `npm start`.
2. Lancer manuellement `node scripts/generate-manifests.js` pour mettre à jour les manifestes finaux.

### GEN-M4 — Espace et heure apparaissant prématurément [FIXÉ — OK]

**Statut :** Résolu. La condition d'affichage du timestamp dans `ChatMessageList.vue` est désormais alignée sur la visibilité du contenu du message, évitant l'heure "fantôme" avant la réponse.

**Stratégie d'implémentation :**

1. Aligner la condition d'affichage du timestamp sur celle du corps du message dans `ChatMessageList.vue` (ou bien n'afficher le timestamp que si `text`, `toolCalls` ou `imageSrc` sont présents).
2. Vérifier l'espacement généré par le flex container pour éviter l'apparition d'un espace fantôme.

### PPT-M5 — L'action rapide "Notes d'orateur" ne s'insère pas d'elle-même [FIXÉ — OK]

**Statut :** Amélioré. `setCurrentSlideSpeakerNotes` gère désormais les erreurs explicitement (API 1.5+) et avertit l'utilisateur en cas d'échec ou d'incompatibilité.

**Stratégie d'implémentation :**

1. Modifier `setCurrentSlideSpeakerNotes` (dans `powerpointTools.ts`) pour qu'il affiche un message clair ou lance une exception gérée si la fonction échoue.
2. Dans le cas Web où la zone de notes n'existerait pas par défaut, essayer de l'initialiser ou avertir l'utilisateur d'une incompatibilité.

---

## ITEMS MINEURS (Amélioration UI cosmétique, Ajustements de Prompts)

### PPT-L1 — L'action "Impact" n'est pas adaptée à PowerPoint [FIXÉ — OK]

**Statut :** Amélioré. Le prompt `punchify` dans `constant.ts` a été assoupli pour permettre des slogans percutants ou des listes courtes selon le contexte, au lieu de forcer systématiquement des puces.

**Criticité :** Mineure

**Problème :** Produit parfois un texte dense ou force obligatoirement des puces, sans tenir compte de la meilleure manière de présenter le contenu pour PowerPoint.
**Cause Racine :** Le prompt système défini dans `frontend/src/utils/constant.ts` pour l'action `punchify` ("Max 6-7 bullets total") force le LLM à transformer quoi qu'il arrive le contenu en bullet points ("puces"), ignorant le fait qu'une phrase courte pourrait être plus percutante.

**Stratégie d'implémentation :**

1. Modifier `constant.ts` en ajustant le prompt `punchify`. Il faut lui retirer la contrainte stricte des bullet points systématiques.
2. Nouveau prompt `punchify` ciblé : "Évaluez la meilleure façon de présenter cela pour un support visuel (PowerPoint/Keynote). Utilisez soit une courte phrase très percutante, soit 3 à 5 très courtes puces, soit une combinaison équilibrée. Ne forcez pas les puces si un court slogan ou texte direct est plus puissant."

### PPT-L2 — L'image générée est uniquement carrée et tronquée [FIXÉ — OK]

**Statut :** Résolu. Bien que la génération reste en 1024x1024 (pour compatibilité/coût), l'outil `insertImageOnSlide` positionne désormais l'image intelligemment (centrée) sans la tronquer, respectant son ratio carré sur les slides 16:9.

**Stratégie d'implémentation :**

1. **Format Image :** Conserver volontairement la génération d'images en résolution standard **1024x1024** pour des raisons de simplicité, de coût et de compatibilité.
2. **Côté PowerPoint (`powerpointTools.ts` ou Prompt) :** Mettre à jour l'implémentation de `insertImageOnSlide` (ou la consigne de positionnement du LLM) pour intégrer intelligemment l'image au format 1:1 sur la slide sans forcer de redimensionnement qui l'étirerait. Ex: la centrer avec des marges ou la positionner à droite du texte, en respectant son ratio carré naturel.

### GEN-L3 — Utilité des boutons/check-box de formatage sous le champ chat [FIXÉ — OK]

**Statut :** Implémenté (Phantom Context). Les cases à cocher ont été supprimées de l'UI. Le texte sélectionné (ou le contenu de la slide/feuille par défaut) est désormais envoyé automatiquement à l'agent. Le prompt système a été mis à jour avec le pattern "Modificateur Intelligent".

**Stratégie d'implémentation :**

1. **Nettoyage UI (`ChatInput.vue`) :** Supprimer purement et simplement les éléments de case à cocher pour `"Formatage Word"` et `"Inclure la sélection"`. Les `v-model` (`useWordFormatting`, `useSelectedText`) peuvent toujours subsister en code mais être figés à `true` par défaut.
2. **Context Fantôme Permanent (`useAgentLoop.ts` / `useAgentPrompts.ts`) :**
   - Transmettre systématiquement le texte/contenu sélectionné (que ce soit une sélection Word, des cellules Excel, du texte de slide PPT, ou un corps de courriel Outlook) en tant que contexte dans la portion cachée du prompt.
3. **Mise à jour du Prompt Système global :**
   - Ajouter une directive claire sur l'exploitation de ce contexte selon l'hôte (Excel, PPT, Word, Outlook).
   - "Le système vous joint ci-dessous la sélection ou le contexte actuel de l'utilisateur (texte, slide, cellules excel, email). Pensez `Modificateur intelligent` : Si la demande de l'utilisateur implique de modifier son brouillon de travail actuel, basez-vous sur ce contexte envoyé pour effectuer l'action (en utilisant les outils d'édition/écriture). Si la demande est d'ordre général, exploitez cette sélection uniquement à titre informatif, comme contexte."
   - L'objectif est que la bascule de "Dois-je utiliser le texte de la slide ?" se fasse dynamiquement par le LLM.
