# Audit Exhaustif des Skills & Veille Technologique (KickOffice vs Marché)

**Fait avec un comparatif direct sur :**

- KickOffice (existant)
- MS-Office-AI (Menahishayan)
- RedInk (LawDigital)
- OutlookLLM (fgblanch)

---

## 1. L'état actuel des compétences de KickOffice

| Application    | Outils Actuels                                       | Manques Immédiats (API Office.js)                                                                  |
| -------------- | ---------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| **Word**       | 37 outils (Formats, Styles, R/W complet, navigation) | Application de styles personnalisés, génération de la hiérarchie TDM.                              |
| **Excel**      | 39 outils (R/W Plages, Formatting, Filtres, Charts)  | Outils de suppression de plages, lecture des propriétés de charts.                                 |
| **PowerPoint** | 8 outils (Lecture texte, notes, insert basic)        | **Carence Critique** : Modification/suppression de slides et de formes, lecture des notes, design. |
| **Outlook**    | 13 outils (R/W basique du compose/read)              | Attachement de fichiers dynamique, extraction de la priorité / importance.                         |

---

## 2. Benchmark et Analyse Concurrentielle des Autres Dépôts

### A. MS-Office-AI (Menahishayan)

- **Modèles Supportés** : Initialement sur Gemini (KickOffice gère également Gemini via divers modèles).
- **Architecture de la mémoire** : Intègre un début de "Context Window Memory" pour limiter la surcharge de l'historique au moment des appels LLM (KickOffice manque de cette dynamique d'élagage intelligent).
- **Historique** : Ils travaillent persistance visuelle des chats sur le UI.

### B. Redink (LawDigital) — Le Concurrent Le Plus Riche

Fournit une suite massive pour Word, PPT et Excel :

- **Audio de présentation (Google/OpenAI Audio)** : Redink lit les speaker-notes et génère le fichier audio associé pour convertir PPT en "vidéo parlée".
- **Génération ET Édition d'images** : L'intégration d'Image in-place est bien plus puissante que KickOffice.
- **Anonymisation locale (Data Privacy)** : Suppression des noms de code avant l'envoi au LLM via des scripts locaux (listes), puis réinsertion des noms en retour. (Énorme skill manquant pour l'entreprise !).
- **Moteur de validation de scripts contractuels** : Redink applique des check-lists pré-rédigées aux documents pour faire un document de "Review report" automatique. Très utile sur Word (contrats).
- **Web Agent & Recherche Web** : Ils permettent au LLM d'aller sur Internet pour croiser la demande avec des cas réels avant d'écrire.
- **Tone Matching** : L'IA est fine-tunée localement sur les brouillons de l'utilisateur pour cloner l'écriture avant de rédiger des emails.

### C. OutlookLLM (fgblanch)

- **Localisation Complète** : Backend fait sur TensorRT-LLM pour faire tourner Mistral ou Llama2 localement sans partager la donnée sur le cloud. (Un excellent axe d'évolution pour KickOffice en mode "Confidentialité Absolue").
- **RAG pour Q&A de Boite Mail** : Indexation locale vectorielle de la boîte mail (via Vector DB Python) pour que l'IA puisse répondre à : "Quand a eu lieu mon rendez-vous avec Paul l'année dernière ?".

---

## 3. Plan d'Évolution & Nouveaux "Skills" Proposés pour KickOffice

Sur la base de cet audit (état Office.js + analyse repos), voici les Skills à implémenter par criticité et retour utilisateur potentiel :

### PRIORITÉ 1 : Combler la dette PowerPoint (Office.js pur)

- `deleteSlide`, `getShapes`, `deleteShape`, `setShapeFill`, `moveResizeShape`, `getAllSlidesOverview`.
- _Objectif_ : Donner enfin à l'agent un véritable panel de maniabilité sur PPT.

### PRIORITÉ 2 : S'inspirer de Redink pour l'aspect de Sécurité (Corporate)

- **Agent Skill : `AnonymizeData` & `DeanonymizeData`**
  - Mettre en place un outil tampon local qui remplace les noms propres, dates sensibles ou entités par des variables `[ENTITY-01]` avant d'appeler l'IA, puis replace le vrai nom ensuite. Cela permettrait à KickOffice d'être vendu comme une solution "Privacy-first".

### PRIORITÉ 3 : ConnectIvité Web (S'inspirer de Redink)

- **Agent Skill : `WebSearch` / `FetchIntranetUrl`**
  - Permettre à l'Agent de faire des requêtes vers le Web (Wikipedia, météo, cours de la bourse, Jurisprudence) pour injecter dans Word/Excel des données qu'il n'hallucine pas.

### PRIORITÉ 4 : Augmentation de l'Outlook Context (S'inspirer d'OutlookLLM)

- **Agent Skill : `KnowledgeBaseRAG` (Calendar & Inbox)**
  - Plutôt qu'interagir uniquement avec le mail actuellement ouvert, il faut un Endpoint Backend capable de stocker une version vectorisée du thread de mails ou du calendrier pour fournir des réponses "connaissant" tout le contexte d'un client.

### PRIORITÉ 5 : Génération Audio & Video (Création de contenu)

- **Agent Skill : `GenerateSpeechFromNotes`**
  - Comme Redink, pouvoir créer une voix de synthèse associée à un slide de PowerPoint (Via un call TTS d'OpenAI ou Google ElevenLabs).

## Conclusion de l'Audit

KickOffice dispose d'une exécution de base robuste ("Foundation") et UX très soignée (Vite, Vue, Tailwind). Mais **Redink** gagne brutalement sur la **profondeur des Features applicatives spécifiques** (Anonymisation, scripts de contrats, Audio). KickOffice doit urgemment développer la maturité de l'agent vers ces use-cases transverses : Privacy Local, Web Agent, et combler son retard en modélisation sur PowerPoint.
