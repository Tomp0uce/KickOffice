# Revue UX Complète (Février 2026)

**Ce document repart de zéro pour consolider les problèmes UX restants par ordre de criticité.**

## Résumé Exécutif

L'interface de KickOffice s'est professionnalisée, notamment avec la gestion i18n bien prise en charge et le système de statuts du backend. Toutefois, plusieurs workflows critiques présentent des frictions importantes pour un usage fluide (streaming, boutons sans libellés, flux d'images brisés).

---

## Liste des Points par Ordre de Criticité

### 🔴 CRITIQUE (Bloque ou détériore fortement l'usage core)

1. **UX-C1 : Pas de streaming dans la boucle conversationnelle (Agent Mode)**
   - **Problème** : Contrairement aux actions rapides, le chat manuel avec l'agent est 100% synchrone. L'utilisateur tape un message et reste face à un UI vide pendant 10 à 30 secondes le temps que l'IA réfléchisse et exécute.
   - **Solution** : Implémenter le retour asynchrone progressif des tokens (streaming) de bout en bout.

2. **UX-C2 : Flux de création d'image "Visual" brisé (PowerPoint)**
   - **Problème** : Le bouton "Visual" génère un prompt textuel que l'utilisateur doit ensuite copier/coller manuellement vers un onglet "Image mode" pour générer la véritable image. C'est un processus lourd en 5 étapes.
   - **Solution** : Action qui route automatiquement vers le composant générateur d'images.

3. **UX-C3 : Perte totale du contexte sans confirmation ("Nouveau Chat")**
   - **Problème** : En cliquant sur le bouton de "clear" conversationnel, tout l'historique de chat s'efface immédiatement sans modale de validation. L'historique n'est d'ailleurs pas sauvegardé.
   - **Solution** : Fenêtre de dialogue de confirmation, voire historique sauvegardé.

### 🟠 ÉLEVÉE (Dégrade la découverte ou provoque de l'incompréhension)

4. **UX-H1 : Boutons d'actions rapides sans texte ("Icon-only")**
   - **Problème** : Les icônes supérieures n'ont aucun texte accolé. Les utilisateurs ne peuvent deviner la différence entre la brosse (Polish), le livre (Academic), etc.
   - **Solution** : Ajouter de courts libellés textuels ou un tooltip plus persistant/évident (surtout pour le touch/tablette).

5. **UX-H2 : Boutons d'action ("Replace/Copy/Append") trop intrusifs**
   - **Problème** : À chaque réponse de l'IA, les 3 boutons d'action d'insertion prennent visuellement beaucoup d'espace. Sur une longue conversation, c'est lourd.
   - **Solution** : N'afficher les boutons que sur le survol (Hover), ou uniquement sur la dernière réponse de l'agent.

6. **UX-H3 : Explication floue des options de sélecteur de modèle**
   - **Problème** : Les options "Nano", "Standard", "Raisonnement" relèvent du jargon de développeur.
   - **Solution** : Renommer pour cibler le besoin : "Réponse basique/rapide", "Réponse Qualité", "Réflexion profonde".

### 🟡 MOYENNE (Frictions mineures et jargon)

7. **UX-M1 : Impossible de régénérer un message (Retry) ou d'éditer**
   - **Problème** : En cas de réponse non-satisfaisante, l'utilisateur doit retaper entièrement son prompt. Pas de bouton "Regenerate".

8. **UX-M2 : Section "Built-in Prompts" dans les paramètres inégale et technique**
   - **Problème** : N'inclut que les prompts de Word et Excel. Jargon d'interpolation `${language}` imposé.

9. **UX-M3 : Affichage du "Thought process" toujours en Anglais**
   - **Problème** : Dans la balise `<summary>` de `ChatMessageList.vue`, le texte est écrit en dur en anglais, cassant l'immersion i18n.

10. **UX-M4 : Indicateurs de clics (Checkbox) trop petits**
    - **Problème** : Les conteneurs CSS des checkboxes réduisent la taille de l'élément (`h-3.5`). Clic difficile.

### 🟢 FAIBLE (Confort)

11. **UX-L1 : État vide ("Empty State") inerte**
    - Pas de suggestions de prompt cliquables quand la fenêtre de chat est vierge.

12. **UX-L2 : Pas d'indicateur "l'IA est en train d'écrire..."**
    - Au-delà du texte de statut, une animation classique de trois petits points sauterait aux yeux et rassurerait.
