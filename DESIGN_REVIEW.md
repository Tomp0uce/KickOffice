# Revue de Code & Audit

**Date**: 2026-02-20
**Périmètre**: Architecture de KickOffice, Code Quality, et état global post-refactoring.

---

## 1. Résumé Exécutif

L'architecture actuelle (frontend Vue 3 + Tailwind, backend via endpoints `/api`) a fait l'objet d'un refactoring réussi récent qui a réglé la majorité des goulots d'étranglement qui ralentissaient la maintenabilité (ex. fichier central `useAgentLoop.ts` éclaté en composables distincts).

Les points critiques historiques ont été traités :

- La gestion de l'état (LocalStorage) des outils a été unifiée.
- Les hardcodes des chaînes côté Typescript / UI (ex: Analyse de la demande) ont été externalisés vers le système de traduction `i18n` (vérifié dans `fr.json`).
- L'injection de texte au format Markdown dans Word / PowerPoint fonctionne désormais via les méthodes adaptées (`insertHtml` / HTML Coercion).

## 2. Points Récemment Résolus (Fixed)

Voici les problèmes majeurs qui ont été traités avec succès :

- ✅ **Désynchronisation de l'état des outils (Feature Toggle)** : L'interface "Settings" filtre désormais correctement les outils utilisés dynamiquement par l'agent.
- ✅ **Absence de Streaming Complet dans la boucle de l'Agent** : Les appels synchrones (`chatSync`) ont été remplacés par `chatStream`, offrant un retour en temps réel y compris lors de l'appel d'outils.
- ✅ **Persistance de l'Historique des Conversations** : L'historique est désormais sauvegardé de manière persistante via `localStorage` (isolé par Host Office).
- ✅ **Gestion fine du contexte du Prompt (Pruning)** : Implémentation d'une fenêtre de contexte intelligente garantissant de ne pas dépasser la limite de tokens tout en préservant l'intégrité des appels d'outils.
- ✅ **Traductions en dur** : Le label "Thought process" a été externalisé vers le système `i18n`. Les descriptions manquantes pour les info-bulles (tooltips) des actions rapides Excel/PPT/Outlook ont été ajoutées.
- ✅ **Exposition de syntaxe développeur** : Remplacement de la syntaxe obscure `${text}` par des balises plus intuitives comme `[TEXT]` dans l'interface de paramétrage.
- ✅ **Interface (Amélioration UX)** : Ajout du défilement automatique (auto-scroll) qui maintient lisible le début du message généré par l'IA lors des réponses longues.

## 3. Nouveaux Points Bloquants & Dette Technique Restante (À Prioriser)

Bien que l'expérience utilisateur de base soit désormais robuste, voici les nouveaux points bloquants et problématiques techniques à traiter :

### CRITIQUE (Performance & Stabilité du Build)

1. **Avertissements sur la taille des "Chunks" JavaScript** :
   - _Problème_ : Le processus de build (Vite) signale des fichiers volumineux (ex. plus de 500kB pour certains utilitaires). Cela peut ralentir significativement le chargement initial du Taskpane Web.
   - _Solution proposée_ : Implémenter le "Code-Splitting" manuel (chargement paresseux des composants Vue via `defineAsyncComponent`) et revoir l'arborescence des dépendances Vite (`manualChunks`).

### ÉLEVÉE (Compatibilité de l'Environnement)

1. **Avertissement de Version Node.js** :
   - _Problème_ : L'environnement de compilation génère des erreurs de compatibilité liés à l'usage ciblé de Node.js (actuellement `20.17.0` détecté, la doc recommande `20.19+` ou `22.12+`).
   - _Solution proposée_ : Harmoniser les pré-requis `engines` dans le `package.json` ou mettre à jour la CI GitHub Actions (comme fait précédemment pour les permissions Git).

### MOYENNE (Qualité Code & Finitions)

1. **Tests automatisés E2E manquants** :
   - _Problème_ : L'application KickOffice dépend de plusieurs Hosts différents (Word, Excel, PowerPoint, Outlook). L'absence de tests E2E augmente le temps de validation manuelle pour chaque release.
   - _Solution proposée_ : Ajouter un framework comme Playwright ou Cypress (bien que plus complexe avec le runtime Office.js).
