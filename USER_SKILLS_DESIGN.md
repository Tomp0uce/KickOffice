# User Skills — Design & Implementation Spec

**Statut** : Prêt à implémenter
**Date** : 2026-03-18
**Référence Anthropic** : [Skill Creator](https://github.com/anthropics/skills/tree/main/skills/skill-creator)

---

## Décisions prises

| Décision | Choix |
|----------|-------|
| Migration custom prompts | Remplacement complet + migration automatique |
| Stockage user skills | localStorage `ki_UserSkills_v1` + export/import `.skill.md` |
| UX QuickActionsBar | Dropdown select → exécution immédiate (1 clic) |
| Format built-in skills | Retrofit YAML frontmatter (format unifié built-in ↔ user) |
| Corps des skills | Markdown libre guidé (pas de template rigide) |
| Skill creator flow | Décrire → Générer → Review → **Tester** → Sauvegarder |

---

## 1. Contexte & Problème actuel

### Ce qui existe

**Custom Prompts** (`savedPrompts.ts` + `PromptsTab.vue`) :
- Stocke `{ id, name, systemPrompt, userPrompt }` en localStorage (clé : `savedPrompts`)
- Affiché dans un `SingleSelect` dans `QuickActionsBar.vue`
- Quand sélectionné → `loadSelectedPrompt()` dans `useHomePage.ts` (l.255) :
  - `customSystemPrompt.value = prompt.systemPrompt`
  - `userInput.value = prompt.userPrompt` (pré-remplit la textarea)
- L'utilisateur doit encore **modifier la textarea et appuyer sur Envoyer** — 3 étapes minimum
- `customSystemPrompt` est injecté dans `useAgentLoop.ts` (l.749) : `const systemPrompt = customSystemPrompt.value || agentPrompt(lang)`

**Built-in Skills** (`skills/quickactions/*.skill.md`) :
- 24 fichiers markdown bruts, sans frontmatter
- Importés via Vite `?raw` dans `skills/index.ts`
- Map hardcodée `quickActionSkillMap` : clé → contenu brut
- Injectés en priority-1 dans `useQuickActions.ts` (l.652) : `getQuickActionSkill(actionKey)`

### Ce qui change

Les **User Skills** remplacent les Custom Prompts. Différence fondamentale de comportement :
- **Ancien** : prompt sélectionné → pré-remplit textarea → l'utilisateur tape → envoie
- **Nouveau** : skill sélectionné → **exécution immédiate** (comme une quick action)

Les User Skills et les Built-in Skills partagent le **même format `.skill.md`** avec frontmatter YAML.

---

## 2. Format Unifié `.skill.md`

### Inspiré d'Anthropic (Progressive Disclosure)

Source : [SKILL.md Anthropic](https://github.com/anthropics/skills/blob/main/skills/skill-creator/SKILL.md)

> Trois niveaux de chargement :
> 1. **Metadata** (~50 mots) — toujours disponible dans l'UI (dropdown, library)
> 2. **Corps de la skill** — chargé à l'exécution comme system prompt
> 3. **Ressources additionnelles** — (non implémenté en v1, réservé)

### Format du fichier

```markdown
---
name: Reformuler en bullets
description: Transforme le texte sélectionné en liste de 3-5 bullet points percutants et concis. Idéal pour convertir un paragraphe dense en points mémorables pour une présentation ou un email. Préserve la langue d'origine.
host: powerpoint
executionMode: immediate
icon: List
actionKey: bullets
---

[Corps libre en markdown — cf. Section 4]
```

### Champs frontmatter

| Champ | Type | Requis | Description |
|-------|------|--------|-------------|
| `name` | string | ✅ | Nom court (≤ 5 mots, verbe impératif) |
| `description` | string | ✅ | "Pushy" — bénéfice utilisateur + quand utiliser. Affiché dans le dropdown. |
| `host` | `word \| excel \| powerpoint \| outlook \| all` | ✅ | Filtre contextuel |
| `executionMode` | `immediate \| draft \| agent` | ✅ | Mode d'exécution |
| `icon` | string | ✅ | Nom d'icône Lucide (ex: `List`, `Wand2`, `Languages`) |
| `actionKey` | string | ❌ | **Built-in uniquement** — lien vers la quick action existante |

### Choix du `executionMode`

| Mode | Quand l'utiliser | Comportement |
|------|-----------------|--------------|
| `immediate` | Transformation de texte en chat (traduire, reformuler, résumer) | `chatStream` → résultat en streaming dans le chat |
| `draft` | L'utilisateur veut relire avant d'envoyer (réponse email, brouillon) | Pré-remplit la textarea + glow — l'utilisateur envoie |
| `agent` | Modification du document, plusieurs opérations séquentielles | `runAgentLoop` avec accès aux outils Office |

**Règle simple** : Dès que la skill touche le document → `agent`. Si elle retourne du texte dans le chat → `immediate`.

---

## 3. Exemple Complet : translate.skill.md Avant/Après

### Avant (format actuel — pas de frontmatter, sections rigides)

```markdown
# Translate Quick Action Skill

## Purpose
Translate selected text...

## When to Use
- User clicks "Translate"...

## Critical Preservation Rules (Outlook)
**YOU MUST**:
- Keep ALL {{PRESERVE_N}} markers EXACTLY as-is
...
```

### Après (format unifié)

```markdown
---
name: Traduire le texte
description: Traduit le texte sélectionné entre français et anglais en détectant automatiquement la langue source. Préserve les formatages gras/italique/tableaux, les placeholders d'images {{PRESERVE_N}}, et adapte le registre au contexte (formel, casual, technique).
host: all
executionMode: immediate
icon: Languages
actionKey: translate
---

Traduis le texte reçu vers l'autre langue (FR ↔ EN). Détecte la langue source — ne te fie pas au tag `[UI language]` pour déterminer la direction, seulement pour la langue de réponse.

**Pourquoi la détection automatique ?** L'utilisateur travaille souvent sur des documents multilingues. Une direction fixe casserait la moitié des cas d'usage en traduisant dans la mauvaise langue.

## Préservation du formatage

Maintiens tous les marqueurs autour de chaque mot traduit :
- `**gras**` → `**texte traduit**`
- `*italique*` → `*texte traduit*`
- `[color:#CC0000]texte[/color]` → `[color:#CC0000]texte traduit[/color]`
- Tableaux : traduis le contenu des cellules, préserve la structure
- Liens : traduis le texte d'affichage, garde l'URL intacte

## Placeholders {{PRESERVE_N}}

Le texte peut contenir des marqueurs `{{PRESERVE_0}}`, `{{PRESERVE_1}}`, etc. représentant des images embarquées dans Outlook.

**Pourquoi ne pas les toucher ?** Ces placeholders sont remplacés par les images réelles après traitement. Les supprimer ou déplacer casserait l'email de manière invisible.

Règle : positionne-les au même endroit logique dans le texte traduit.

Exemple :
```
Entrée FR : "Voici le rapport {{PRESERVE_0}} pour révision."
Sortie EN : "Here is the report {{PRESERVE_0}} for review."
```

## Détection de la langue

- Texte principalement **français** → traduire en **anglais**
- Texte principalement **anglais** → traduire en **français**
- Autre langue → traduire en **français** (défaut)
- Noms propres, marques, termes techniques : garde-les si non-traduisibles

## Ton

Adapte le registre de la source : formel → formel, casual → casual, technique → technique.

**Output** : retourne UNIQUEMENT le texte traduit. Aucune explication, aucun préambule.
```

---

## 4. Principes de Rédaction d'une Skill (Theory of Mind)

Source : Principe "Theory of Mind" d'Anthropic — expliquer *pourquoi* plutôt que d'imposer avec `MUST`.

1. **Explique le POURQUOI** : `Préserve les placeholders {{PRESERVE_N}} — les supprimer casserait l'email de manière invisible` > `CRITICAL: NEVER drop {{PRESERVE_N}}`

2. **La description est "pushy" et orientée bénéfice** : Dans le dropdown UI, l'utilisateur doit comprendre *exactement* ce que fait la skill et quand l'utiliser en lisant la description.

3. **Structure libre** : Pas de sections obligatoires. Utilise ce qui sert la compréhension : sous-titres, tableaux, exemples input/output. 1-2 exemples concrets valent mieux qu'une liste de règles abstraites.

4. **Anticipe les vrais cas limites** : "Si le texte est déjà dans la langue cible" est utile. "Si le fichier est corrompu" ne l'est pas dans ce contexte.

5. **Choisis les bons outils** (pour les skills `agent`) : Explique explicitement quels outils utiliser et pourquoi.

---

## 5. Migration des Built-in Skills — Format Unifié

### 24 fichiers à modifier + `skills/index.ts` à refactorer

**Mapping frontmatter pour les built-in skills :**

| Fichier | name | host | executionMode | icon | actionKey |
|---------|------|------|---------------|------|-----------|
| `bullets.skill.md` | Reformer en bullets | powerpoint | immediate | List | bullets |
| `punchify.skill.md` | Punchifier le texte | powerpoint | immediate | Zap | punchify |
| `review.skill.md` | Réviser la slide | powerpoint | agent | Eye | review |
| `translate.skill.md` | Traduire le texte | all | immediate | Languages | translate |
| `formalize.skill.md` | Formaliser le ton | outlook | immediate | Briefcase | formalize |
| `concise.skill.md` | Rendre concis | outlook | immediate | Scissors | concise |
| `proofread.skill.md` | Corriger le texte | all | immediate | CheckSquare | proofread |
| `ppt-proofread.skill.md` | Corriger la slide | powerpoint | agent | CheckCircle | ppt-proofread |
| `ppt-translate.skill.md` | Traduire la slide | powerpoint | agent | Globe | ppt-translate |
| `word-translate.skill.md` | Traduire (Word) | word | agent | Languages | word-translate |
| `word-proofread.skill.md` | Corriger (Word) | word | agent | CheckSquare | word-proofread |
| `word-review.skill.md` | Réviser le doc | word | agent | Eye | word-review |
| `polish.skill.md` | Polir le texte | word | immediate | Sparkles | polish |
| `academic.skill.md` | Style académique | word | immediate | GraduationCap | academic |
| `summary.skill.md` | Résumer | word | immediate | AlignLeft | summary |
| `extract.skill.md` | Extraire les actions | outlook | immediate | ListChecks | extract |
| `reply.skill.md` | Rédiger une réponse | outlook | draft | Reply | reply |
| `ingest.skill.md` | Nettoyer les données | excel | agent | Database | ingest |
| `autograph.skill.md` | Embellir le tableau | excel | agent | Palette | autograph |
| `chart-digitizer.skill.md` | Numériser un graphique | excel | agent | BarChart | digitizeChart |
| `pixel-art.skill.md` | Pixel art Excel | excel | agent | Grid3X3 | pixelArt |
| `explain-excel.skill.md` | Expliquer les formules | excel | immediate | HelpCircle | explain |
| `formula-generator.skill.md` | Générer une formule | excel | agent | Function | formulaGenerator |
| `data-trend.skill.md` | Analyser les tendances | excel | agent | TrendingUp | dataTrend |

### Nouveau `frontend/src/utils/skillParser.ts` (à créer)

```typescript
/**
 * skillParser.ts
 *
 * Parses .skill.md files with YAML frontmatter.
 * Lightweight parser — no external dependency.
 * Supports: string values (quoted or unquoted), no nested objects.
 */

export type SkillHost = 'word' | 'excel' | 'powerpoint' | 'outlook' | 'all'
export type SkillExecutionMode = 'immediate' | 'draft' | 'agent'

export interface SkillMetadata {
  name: string
  description: string
  host: SkillHost
  executionMode: SkillExecutionMode
  icon: string
  actionKey?: string         // built-in skills only
}

export interface ParsedSkill {
  metadata: SkillMetadata
  body: string               // markdown body without frontmatter
  raw: string                // full file content (for system prompt injection)
}

const VALID_HOSTS: SkillHost[] = ['word', 'excel', 'powerpoint', 'outlook', 'all']
const VALID_EXECUTION_MODES: SkillExecutionMode[] = ['immediate', 'draft', 'agent']

/**
 * Parse a raw .skill.md string into metadata + body.
 * Falls back to default metadata if frontmatter is missing/invalid.
 */
export function parseSkill(raw: string, fallbackActionKey?: string): ParsedSkill {
  const match = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/)

  if (!match) {
    // No frontmatter — legacy file
    return {
      metadata: {
        name: fallbackActionKey ?? 'Unnamed Skill',
        description: '',
        host: 'all',
        executionMode: 'immediate',
        icon: 'Zap',
        actionKey: fallbackActionKey,
      },
      body: raw,
      raw,
    }
  }

  const frontmatterStr = match[1]
  const body = match[2].trim()
  const metadata = parseFrontmatter(frontmatterStr, fallbackActionKey)

  return { metadata, body, raw }
}

/**
 * Minimal YAML parser for flat key: value pairs.
 * Handles quoted strings, strips inline comments.
 */
function parseFrontmatter(yaml: string, fallbackActionKey?: string): SkillMetadata {
  const lines = yaml.split('\n')
  const obj: Record<string, string> = {}

  for (const line of lines) {
    const colonIdx = line.indexOf(':')
    if (colonIdx === -1) continue
    const key = line.slice(0, colonIdx).trim()
    let val = line.slice(colonIdx + 1).trim()
    // Remove surrounding quotes
    if ((val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1)
    }
    obj[key] = val
  }

  return {
    name: obj.name ?? fallbackActionKey ?? 'Unnamed',
    description: obj.description ?? '',
    host: VALID_HOSTS.includes(obj.host as SkillHost) ? (obj.host as SkillHost) : 'all',
    executionMode: VALID_EXECUTION_MODES.includes(obj.executionMode as SkillExecutionMode)
      ? (obj.executionMode as SkillExecutionMode)
      : 'immediate',
    icon: obj.icon ?? 'Zap',
    actionKey: obj.actionKey || fallbackActionKey,
  }
}

/**
 * Serialize a UserSkill to .skill.md format.
 */
export function serializeSkillToMd(skill: {
  name: string
  description: string
  host: SkillHost
  executionMode: SkillExecutionMode
  icon: string
  skillContent: string
  actionKey?: string
}): string {
  const lines = [
    '---',
    `name: ${skill.name}`,
    `description: "${skill.description.replace(/"/g, '\\"')}"`,
    `host: ${skill.host}`,
    `executionMode: ${skill.executionMode}`,
    `icon: ${skill.icon}`,
  ]
  if (skill.actionKey) lines.push(`actionKey: ${skill.actionKey}`)
  lines.push('---', '', skill.skillContent)
  return lines.join('\n')
}
```

### Nouveau `frontend/src/skills/index.ts` (refactoré)

```typescript
/**
 * Skills Loader — Refactored to be metadata-driven.
 *
 * Built-in skills are imported via Vite ?raw at build-time.
 * Frontmatter is parsed at module load (one-time cost, not per-request).
 */

import { logService } from '@/utils/logger'
import { parseSkill } from '@/utils/skillParser'
import type { ParsedSkill, SkillHost } from '@/utils/skillParser'

// Host skills (unchanged — no frontmatter needed, they're always injected by host)
import commonSkill from './common.skill.md?raw'
import wordSkill from './word.skill.md?raw'
import excelSkill from './excel.skill.md?raw'
import powerpointSkill from './powerpoint.skill.md?raw'
import outlookSkill from './outlook.skill.md?raw'

// Quick Action skills — raw imports
import bulletsSkillRaw from './quickactions/bullets.skill.md?raw'
import punchifySkillRaw from './quickactions/punchify.skill.md?raw'
import reviewSkillRaw from './quickactions/review.skill.md?raw'
import translateSkillRaw from './quickactions/translate.skill.md?raw'
// ... (tous les autres imports identiques avec le suffixe Raw)

// Map: actionKey → raw string (for parsing)
const rawSkillFiles: Record<string, string> = {
  bullets: bulletsSkillRaw,
  punchify: punchifySkillRaw,
  review: reviewSkillRaw,
  translate: translateSkillRaw,
  // ... (même clés qu'avant)
}

// Parse all skills at module load (build-time cost)
const parsedSkills: Map<string, ParsedSkill> = new Map(
  Object.entries(rawSkillFiles).map(([key, raw]) => [key, parseSkill(raw, key)])
)

const hostSkillMap: Record<string, string> = {
  Word: wordSkill,
  Excel: excelSkill,
  PowerPoint: powerpointSkill,
  Outlook: outlookSkill,
}

export function getSkillForHost(host: string): string {
  const hostSkill = hostSkillMap[host]
  if (!hostSkill) {
    logService.warn(`[Skills] Unknown host: ${host}, using common skills only`)
    return commonSkill
  }
  return `${commonSkill}\n\n---\n\n${hostSkill}`
}

export function getCommonSkill(): string { return commonSkill }

/** Get the raw .skill.md content for a quick action (for system prompt injection). */
export function getQuickActionSkill(actionKey: string): string | undefined {
  return parsedSkills.get(actionKey)?.raw
}

/** Get only the metadata of a quick action skill (for UI display). */
export function getQuickActionSkillMetadata(actionKey: string) {
  return parsedSkills.get(actionKey)?.metadata
}

/** Get metadata for all built-in quick action skills. */
export function getAllBuiltInSkillsMetadata() {
  return Array.from(parsedSkills.values()).map(s => s.metadata)
}

export function hasQuickActionSkill(actionKey: string): boolean {
  return parsedSkills.has(actionKey)
}

export function getAvailableQuickActionSkills(): string[] {
  return Array.from(parsedSkills.keys())
}

export type { SkillHost, SkillExecutionMode, ParsedSkill, SkillMetadata } from '@/utils/skillParser'
```

---

## 6. Data Model — User Skills

### Types (`frontend/src/types/userSkill.ts` — à créer)

```typescript
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser'

export interface UserSkill {
  id: string                     // "skill_1710766200000"
  name: string                   // "Reformuler en bullets"
  description: string            // Affiché dans le dropdown et la library
  host: SkillHost
  executionMode: SkillExecutionMode
  icon: string                   // Nom icône Lucide, défaut: "Zap"
  skillContent: string           // Corps markdown (sans frontmatter)
  createdAt: number              // timestamp ms
  updatedAt: number              // timestamp ms
}

export const SKILL_STORAGE_KEY = 'ki_UserSkills_v1'
export const SKILL_MIGRATION_KEY = 'ki_UserSkillsMigrated_v1'
```

---

## 7. Composable `useUserSkills.ts`

**Fichier** : `frontend/src/composables/useUserSkills.ts`

### Interface complète

```typescript
import { ref, computed } from 'vue'
import { logService } from '@/utils/logger'
import { serializeSkillToMd, parseSkill } from '@/utils/skillParser'
import type { UserSkill } from '@/types/userSkill'
import { SKILL_STORAGE_KEY, SKILL_MIGRATION_KEY } from '@/types/userSkill'
import type { SkillHost } from '@/utils/skillParser'

export function useUserSkills() {
  const skills = ref<UserSkill[]>([])

  // ── Load ──────────────────────────────────────────────────────────────────

  function loadFromStorage(): void {
    const stored = localStorage.getItem(SKILL_STORAGE_KEY)
    if (!stored) { skills.value = []; return }
    try {
      const parsed = JSON.parse(stored)
      skills.value = Array.isArray(parsed) ? parsed.filter(isValidUserSkill) : []
    } catch (err) {
      logService.warn('[UserSkills] Failed to parse stored skills', err)
      skills.value = []
    }
  }

  function saveToStorage(): void {
    try {
      localStorage.setItem(SKILL_STORAGE_KEY, JSON.stringify(skills.value))
    } catch (e) {
      if (e instanceof DOMException && e.name === 'QuotaExceededError') {
        logService.warn('[UserSkills] localStorage quota exceeded')
      } else throw e
    }
  }

  // ── CRUD ──────────────────────────────────────────────────────────────────

  function addSkill(skill: Omit<UserSkill, 'id' | 'createdAt' | 'updatedAt'>): UserSkill {
    const newSkill: UserSkill = {
      ...skill,
      id: `skill_${Date.now()}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    }
    skills.value.push(newSkill)
    saveToStorage()
    return newSkill
  }

  function updateSkill(id: string, updates: Partial<Omit<UserSkill, 'id' | 'createdAt'>>): void {
    const idx = skills.value.findIndex(s => s.id === id)
    if (idx === -1) return
    skills.value[idx] = { ...skills.value[idx], ...updates, updatedAt: Date.now() }
    saveToStorage()
  }

  function deleteSkill(id: string): void {
    skills.value = skills.value.filter(s => s.id !== id)
    saveToStorage()
  }

  // ── Export / Import ───────────────────────────────────────────────────────

  function exportSkillToFile(skill: UserSkill): void {
    const content = serializeSkillToMd({
      name: skill.name,
      description: skill.description,
      host: skill.host,
      executionMode: skill.executionMode,
      icon: skill.icon,
      skillContent: skill.skillContent,
    })
    const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${skill.name.toLowerCase().replace(/\s+/g, '-')}.skill.md`
    a.click()
    URL.revokeObjectURL(url)
  }

  async function importSkillFromFile(file: File): Promise<UserSkill | null> {
    try {
      const text = await file.text()
      const parsed = parseSkill(text)
      // Generate new id to avoid collisions
      return addSkill({
        name: parsed.metadata.name,
        description: parsed.metadata.description,
        host: parsed.metadata.host,
        executionMode: parsed.metadata.executionMode,
        icon: parsed.metadata.icon ?? 'Zap',
        skillContent: parsed.body,
      })
    } catch (err) {
      logService.warn('[UserSkills] Failed to import skill file', err)
      return null
    }
  }

  // ── Filtered view ─────────────────────────────────────────────────────────

  function skillsForHost(host: SkillHost) {
    return computed(() =>
      skills.value.filter(s => s.host === host || s.host === 'all')
    )
  }

  // ── Migration from Custom Prompts ─────────────────────────────────────────

  function checkAndMigrateOldPrompts(): boolean {
    // Returns true if migration is available (caller shows the dialog)
    if (localStorage.getItem(SKILL_MIGRATION_KEY)) return false
    const stored = localStorage.getItem('savedPrompts')
    if (!stored) {
      localStorage.setItem(SKILL_MIGRATION_KEY, 'done')
      return false
    }
    try {
      const prompts = JSON.parse(stored)
      if (!Array.isArray(prompts) || prompts.length === 0) {
        localStorage.setItem(SKILL_MIGRATION_KEY, 'done')
        return false
      }
      // Filter out the default empty prompt
      const realPrompts = prompts.filter((p: any) =>
        p.name !== 'Default' || p.systemPrompt || p.userPrompt
      )
      return realPrompts.length > 0
    } catch { return false }
  }

  function migrateOldPrompts(): void {
    const stored = localStorage.getItem('savedPrompts')
    if (!stored) return
    try {
      const prompts = JSON.parse(stored) as Array<{
        id: string; name: string; systemPrompt: string; userPrompt: string
      }>
      for (const p of prompts) {
        if (p.name === 'Default' && !p.systemPrompt && !p.userPrompt) continue
        const lines = ['## Instructions système']
        if (p.systemPrompt) lines.push(p.systemPrompt)
        if (p.userPrompt) {
          lines.push('', '## Message type', p.userPrompt)
        }
        addSkill({
          name: p.name,
          description: p.name,
          host: 'all',
          executionMode: 'immediate',
          icon: 'Zap',
          skillContent: lines.join('\n'),
        })
      }
    } catch (err) {
      logService.warn('[UserSkills] Migration failed', err)
    }
    localStorage.setItem(SKILL_MIGRATION_KEY, 'done')
    // NOTE: Ne pas supprimer 'savedPrompts' ici — laisser la logique de
    // nettoyage dans le composant après confirmation de l'utilisateur
  }

  function confirmMigrationDone(): void {
    localStorage.removeItem('savedPrompts')
    localStorage.setItem(SKILL_MIGRATION_KEY, 'done')
  }

  // ── Init ──────────────────────────────────────────────────────────────────

  loadFromStorage()

  return {
    skills,
    skillsForHost,
    addSkill,
    updateSkill,
    deleteSkill,
    exportSkillToFile,
    importSkillFromFile,
    checkAndMigrateOldPrompts,
    migrateOldPrompts,
    confirmMigrationDone,
  }
}

// ── Validation ────────────────────────────────────────────────────────────────

function isValidUserSkill(item: unknown): item is UserSkill {
  if (!item || typeof item !== 'object') return false
  const o = item as Record<string, unknown>
  return (
    typeof o.id === 'string' &&
    typeof o.name === 'string' &&
    typeof o.skillContent === 'string' &&
    ['word', 'excel', 'powerpoint', 'outlook', 'all'].includes(o.host as string) &&
    ['immediate', 'draft', 'agent'].includes(o.executionMode as string)
  )
}
```

---

## 8. Exécution des User Skills

### Nouveau composable `useUserSkillExecution.ts`

**Principe** : Les user skills s'exécutent exactement comme des quick actions (même pipeline). Pas de pré-remplissage de textarea, exécution immédiate.

```typescript
/**
 * useUserSkillExecution.ts
 *
 * Exécute les user skills avec le même pipeline que les quick actions.
 * - immediate → chatStream
 * - agent → runAgentLoop
 * - draft → pré-remplit la textarea (isDraftFocusGlowing)
 */
import { nextTick } from 'vue'
import type { Ref } from 'vue'
import type { UserSkill } from '@/types/userSkill'
import type { ChatMessage } from '@/api/backend'
import { chatStream, categorizeError } from '@/api/backend'
import { GLOBAL_STYLE_INSTRUCTIONS } from '@/utils/constant'
import { message as messageUtil } from '@/utils/message'
import { logService } from '@/utils/logger'
import type { ModelTier } from '@/types'
import type { DisplayMessage } from '@/types/chat'

export interface UseUserSkillExecutionOptions {
  t: (key: string) => string
  history: Ref<DisplayMessage[]>
  userInput: Ref<string>
  loading: Ref<boolean>
  abortController: Ref<AbortController | null>
  inputTextarea: Ref<HTMLTextAreaElement | undefined>
  isDraftFocusGlowing: Ref<boolean>
  getOfficeSelection: (opts?: any) => Promise<string>
  runAgentLoop: (messages: ChatMessage[], modelTier: ModelTier) => Promise<void>
  resolveChatModelTier: () => ModelTier
  createDisplayMessage: (role: DisplayMessage['role'], content: string) => DisplayMessage
  adjustTextareaHeight: () => void
  scrollToBottom: () => Promise<void>
  scrollToMessageTop?: () => Promise<void>
}

export function useUserSkillExecution(options: UseUserSkillExecutionOptions) {
  const {
    t, history, userInput, loading, abortController,
    inputTextarea, isDraftFocusGlowing,
    getOfficeSelection, runAgentLoop, resolveChatModelTier,
    createDisplayMessage, adjustTextareaHeight, scrollToBottom, scrollToMessageTop,
  } = options

  async function executeUserSkill(skill: UserSkill): Promise<void> {
    if (loading.value || abortController.value) {
      messageUtil.warning(t('requestInProgress') || 'A request is already in progress.')
      return
    }

    const lang = localStorage.getItem('localLanguage') === 'en' ? 'English' : 'Français'

    // Build the system prompt from skill content
    const systemMsg = [
      skill.skillContent,
      '',
      GLOBAL_STYLE_INSTRUCTIONS,
    ].join('\n')

    // Draft mode: pre-fill textarea, don't execute
    if (skill.executionMode === 'draft') {
      userInput.value = ''
      adjustTextareaHeight()
      isDraftFocusGlowing.value = true
      setTimeout(() => { isDraftFocusGlowing.value = false }, 1500)
      await nextTick()
      const el = inputTextarea.value
      if (el) { el.focus(); el.setSelectionRange(0, 0) }
      return
    }

    // Get selected text (optional for agent skills — they call their own tools)
    let selectedText = ''
    try {
      selectedText = await getOfficeSelection()
    } catch {
      // Agent skills can proceed without selection (they use tools)
      if (skill.executionMode !== 'agent') {
        messageUtil.error(t('selectTextPrompt') || 'Please select some text first.')
        return
      }
    }

    if (!selectedText && skill.executionMode !== 'agent') {
      messageUtil.error(t('selectTextPrompt') || 'Please select some text first.')
      return
    }

    const userMsg = selectedText
      ? `[UI language: ${lang}]\n\n<document_content>\n${selectedText}\n</document_content>`
      : `[UI language: ${lang}]`

    history.value.push(createDisplayMessage('user', `[${skill.name}] ${selectedText.substring(0, 100)}${selectedText.length > 100 ? '...' : ''}`))

    const messages: ChatMessage[] = [
      { role: 'system', content: systemMsg },
      { role: 'user', content: userMsg },
    ]

    if (skill.executionMode === 'agent') {
      loading.value = true
      abortController.value = new AbortController()
      try {
        await runAgentLoop(messages, resolveChatModelTier())
      } catch (err: unknown) {
        if (!(err instanceof Error) || err.name !== 'AbortError') {
          logService.error('[UserSkillExecution] agent skill failed', err)
          const last = history.value[history.value.length - 1]
          if (last?.role === 'assistant') {
            const errInfo = categorizeError(err)
            last.content = t(errInfo.i18nKey)
          }
        }
      } finally {
        loading.value = false
        abortController.value = null
      }
      return
    }

    // immediate: chatStream
    history.value.push(createDisplayMessage('assistant', ''))
    await scrollToMessageTop?.()
    loading.value = true
    abortController.value = new AbortController()
    try {
      await chatStream({
        messages,
        modelTier: resolveChatModelTier(),
        onStream: async (text: string) => {
          const last = history.value[history.value.length - 1]
          last.role = 'assistant'
          last.content = text
          await scrollToBottom()
        },
        abortSignal: abortController.value.signal,
      })
      const last = history.value[history.value.length - 1]
      if (!last?.content?.trim()) last.content = t('noModelResponse')
    } catch (err: unknown) {
      if (!(err instanceof Error) || err.name !== 'AbortError') {
        logService.error('[UserSkillExecution] stream failed', err)
        const last = history.value[history.value.length - 1]
        const errInfo = categorizeError(err)
        last.content = t(errInfo.i18nKey)
      }
    } finally {
      loading.value = false
      abortController.value = null
    }
  }

  return { executeUserSkill }
}
```

---

## 9. Mise à jour `QuickActionsBar.vue`

### Changements

- Le `SingleSelect` actuel passe de "Custom Prompts" → "User Skills"
- `@select` déclenche `executeUserSkill` directement (plus de `loadSelectedPrompt`)
- Ajout d'un bouton `+` pour ouvrir le skill creator
- Les options du select proviennent de `skillsForHost(currentHost).value`

```vue
<template>
  <div class="flex w-full flex-wrap items-center justify-center gap-2 rounded-md">
    <!-- User Skills dropdown -->
    <div class="flex flex-1 items-center gap-1 max-w-xs">
      <SingleSelect
        :key-list="userSkillsForHost.map(s => s.id)"
        :placeholder="t('mySkills') || 'Mes skills...'"
        title=""
        :fronticon="false"
        class="flex-1! bg-surface! text-xs!"
        @update:model-value="(id) => $emit('execute-user-skill', String(id))"
      >
        <template #item="{ item }">
          {{ userSkillsForHost.find(s => s.id === item)?.name || item }}
        </template>
      </SingleSelect>
      <CustomButton
        :icon="Plus"
        text=""
        :title="t('createSkill') || 'Créer un skill'"
        type="secondary"
        :icon-size="14"
        class="shrink-0! bg-surface! p-1.5!"
        @click="$emit('open-skill-creator')"
      />
    </div>

    <!-- Built-in quick action buttons (unchanged) -->
    <CustomButton
      v-for="action in quickActions"
      :key="action.key"
      :title="$t(action.tooltipKey || action.key + '_tooltip')"
      text=""
      :icon="action.icon"
      type="secondary"
      :icon-size="16"
      class="shrink-0! bg-surface! p-1.5!"
      :disabled="loading"
      :aria-label="action.label"
      @click="$emit('apply-action', action.key)"
    />
  </div>
</template>

<script lang="ts" setup>
import { Plus } from 'lucide-vue-next'
import CustomButton from '@/components/CustomButton.vue'
import SingleSelect from '@/components/SingleSelect.vue'
import type { QuickAction } from '@/types/chat'
import type { UserSkill } from '@/types/userSkill'

defineProps<{
  quickActions: QuickAction[]
  loading: boolean
  userSkillsForHost: UserSkill[]   // remplace savedPrompts
}>()

defineEmits<{
  (e: 'apply-action', key: string): void
  (e: 'execute-user-skill', id: string): void   // remplace 'load-prompt'
  (e: 'open-skill-creator'): void                // nouveau
}>()
</script>
```

---

## 10. Modifications dans `HomePage.vue`

### Supprimer

```typescript
// SUPPRIMER ces lignes
import { type SavedPrompt } from '@/utils/savedPrompts'
const selectedPromptId = ref('')        // ligne ~117
const customSystemPrompt = ref('')      // ligne ~118
// et les références dans useHomePage / useAgentLoop / useHomePageContext
```

### Ajouter

```typescript
import { useUserSkills } from '@/composables/useUserSkills'
import { useUserSkillExecution } from '@/composables/useUserSkillExecution'
import type { UserSkill } from '@/types/userSkill'

const { skills, skillsForHost, checkAndMigrateOldPrompts, migrateOldPrompts,
        confirmMigrationDone } = useUserSkills()

// Déterminer le host courant (Word/Excel/PPT/Outlook)
const currentHostLower = computed(() =>
  hostIsWord ? 'word' : hostIsExcel ? 'excel' : hostIsPowerPoint ? 'powerpoint' : 'outlook'
)
const userSkillsForHost = skillsForHost(currentHostLower.value)

const { executeUserSkill } = useUserSkillExecution({
  t, history, userInput, loading, abortController,
  inputTextarea, isDraftFocusGlowing,
  getOfficeSelection, runAgentLoop, resolveChatModelTier,
  createDisplayMessage, adjustTextareaHeight, scrollToBottom, scrollToMessageTop,
})

function handleExecuteUserSkill(id: string): void {
  const skill = skills.value.find(s => s.id === id)
  if (skill) executeUserSkill(skill)
}

// Migration check on mount
onMounted(() => {
  if (checkAndMigrateOldPrompts()) {
    showMigrationDialog.value = true
  }
})
```

Dans le template `<QuickActionsBar>` :
```vue
<QuickActionsBar
  :quick-actions="quickActions"
  :loading="loading"
  :user-skills-for-host="userSkillsForHost"   <!-- remplace :saved-prompts -->
  @apply-action="applyQuickAction"
  @execute-user-skill="handleExecuteUserSkill"  <!-- remplace @load-prompt -->
  @open-skill-creator="showSkillCreator = true"
/>
```

### Supprimer dans `useHomePage.ts`

- `loadSelectedPrompt()` (l.255-262)
- Paramètres `customSystemPrompt` et `selectedPromptId` dans `UseHomePageOptions`
- Les réfs correspondantes retournées par `useHomePage`

### Supprimer dans `useAgentLoop.ts`

- Paramètre `customSystemPrompt: Ref<string>` (l.77)
- Usage l.749 : `const systemPrompt = customSystemPrompt.value || agentPrompt(lang)` → remplacer par `const systemPrompt = agentPrompt(lang)`

### Supprimer dans `useHomePageContext.ts`

- `customSystemPrompt: Ref<string>` (l.32)
- `selectedPromptId: Ref<string>` (l.33)
- `loadSelectedPrompt: () => void` (l.68)

---

## 11. Skill Creator — Backend

### Nouveau fichier `backend/src/routes/skillCreator.js`

```javascript
import { Router } from 'express'
import Anthropic from '@anthropic-ai/sdk'
import { ensureLlmApiKey, ensureUserCredentials } from '../middleware/auth.js'
import { logAndRespond } from '../utils/http.js'
import logger from '../utils/logger.js'
import { SKILL_CREATOR_SYSTEM_PROMPT } from '../config/skillCreatorPrompt.js'

const skillCreatorRouter = Router()

skillCreatorRouter.post('/', ensureLlmApiKey, ensureUserCredentials, async (req, res) => {
  const { description, host } = req.body

  if (!description || typeof description !== 'string' || description.trim().length < 5) {
    return logAndRespond(res, 400, { error: 'description is required (min 5 chars)' }, 'POST /api/skill-creator')
  }

  const validHosts = ['word', 'excel', 'powerpoint', 'outlook', 'all']
  if (host && !validHosts.includes(host)) {
    return logAndRespond(res, 400, { error: 'invalid host' }, 'POST /api/skill-creator')
  }

  const client = new Anthropic({ apiKey: req.llmApiKey })

  const userMessage = [
    `L'utilisateur souhaite créer un skill pour : ${description.trim()}`,
    host && host !== 'all' ? `Host cible : ${host}` : 'Host cible : à déterminer depuis la description',
  ].filter(Boolean).join('\n')

  let attempt = 0
  let lastError = null

  // Retry once on invalid JSON
  while (attempt < 2) {
    try {
      const response = await client.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 4096,
        system: SKILL_CREATOR_SYSTEM_PROMPT,
        messages: [{ role: 'user', content: userMessage }],
      })

      const rawText = response.content
        .filter(b => b.type === 'text')
        .map(b => b.text)
        .join('')

      // Extract JSON from response (handles markdown code blocks)
      const jsonMatch = rawText.match(/```json\s*([\s\S]*?)\s*```/) ||
                        rawText.match(/\{[\s\S]*\}/)
      const jsonStr = jsonMatch ? (jsonMatch[1] || jsonMatch[0]) : rawText.trim()

      const parsed = JSON.parse(jsonStr)

      // Validate required fields
      const required = ['name', 'description', 'host', 'executionMode', 'icon', 'skillContent']
      for (const field of required) {
        if (!parsed[field]) throw new Error(`Missing field: ${field}`)
      }
      if (!validHosts.includes(parsed.host)) parsed.host = host || 'all'
      if (!['immediate', 'draft', 'agent'].includes(parsed.executionMode)) {
        parsed.executionMode = 'immediate'
      }

      return res.json(parsed)
    } catch (err) {
      lastError = err
      attempt++
      logger.warn('POST /api/skill-creator JSON parse failed, retrying', { err })
    }
  }

  logger.error('POST /api/skill-creator failed after retries', { error: lastError })
  return logAndRespond(res, 500, { error: 'Failed to generate skill' }, 'POST /api/skill-creator')
})

export { skillCreatorRouter }
```

### Enregistrement dans `backend/src/server.js`

```javascript
import { skillCreatorRouter } from './routes/skillCreator.js'
// ...
app.use('/api/skill-creator', skillCreatorRateLimit, skillCreatorRouter)
```

### System Prompt `backend/src/config/skillCreatorPrompt.js`

```javascript
export const SKILL_CREATOR_SYSTEM_PROMPT = `Tu es un expert en création de skills pour KickOffice, un assistant IA intégré dans Microsoft Office (Word, Excel, PowerPoint, Outlook).

Les skills que tu crées sont des instructions markdown injectées comme system prompt lors de l'exécution d'une action rapide. Elles guident le comportement de l'agent IA pour une tâche spécifique sur un document Office.

## Ce qu'est une bonne skill (Theory of Mind)

1. Explique le POURQUOI derrière chaque contrainte — le modèle comprend mieux les compromis que les injonctions.
   Mauvais : "CRITICAL: NEVER drop formatting markers"
   Bon : "Préserve les marqueurs de formatage **...** — les supprimer casserait la mise en forme du document de manière invisible."

2. La description est orientée bénéfice utilisateur ("pushy") : dit exactement ce que fait la skill et quand l'utiliser.

3. Corps libre en markdown : utilise la structure qui sert la compréhension. 1-2 exemples input/output concrets valent mieux qu'une liste de règles abstraites.

4. Anticipe seulement les vrais cas limites probables.

5. Pour les skills agent, explique explicitement quels outils utiliser et pourquoi.

## Modes d'exécution

- `immediate` : Réponse en streaming dans le chat. Aucun accès aux outils Office. Utiliser pour les transformations de texte (traduire, reformuler, résumer, corriger). Le contenu sélectionné est passé en entrée entre balises <document_content>.

- `draft` : Pré-remplit la textarea de l'utilisateur. Utiliser quand l'utilisateur veut réviser avant d'envoyer (réponse email, brouillon de document). La skill reçoit peu de contexte — son rôle est de guider la rédaction.

- `agent` : Boucle agentique complète avec accès aux outils Office. Utiliser dès que la tâche nécessite de modifier le document, insérer du contenu, lire le contexte, ou faire plusieurs opérations séquentielles.

RÈGLE SIMPLE : Si la skill doit toucher le document → `agent`. Si elle retourne du texte dans le chat → `immediate`.

## Outils disponibles par host

### Word
getSelectedText, getDocumentContent, getDocumentHtml, insertContent, formatText (bold/italic/fontSize/color sur une plage), searchAndReplace, applyTaggedFormatting, setParagraphFormat (alignement/indentation/espacement), addComment, getComments, proposeRevision (avec Track Changes — RECOMMANDÉ pour corrections), proposeDocumentRevision, editDocumentXml (OOXML direct — pour transformations complexes), acceptAiChanges, rejectAiChanges, insertOoxml, getDocumentOoxml, insertHyperlink, modifyTableCell, addTableRow, addTableColumn, deleteTableRowColumn, formatTableCell, insertHeaderFooter, insertFootnote, setPageSetup, applyStyle, findText, getSpecificParagraph, insertSectionBreak, getSelectedTextWithFormatting, eval_wordjs (dernier recours).

Choix d'outil Word : corrections chirurgicales → proposeRevision | remplacement direct → searchAndReplace | transformations OOXML → editDocumentXml

### Excel
getSelectedCells, getWorksheetData, setCellRange, formatRange (couleur/bordures/alignement/format numérique), createTable, modifyStructure (insérer/supprimer lignes-colonnes), sortRange, applyConditionalFormatting, getConditionalFormattingRules, searchAndReplace, findData, getAllObjects (charts/pivots), screenshotRange, getRangeAsCsv, detectDataHeaders, importCsvToSheet, imageToSheet, extract_chart_data (extrait données d'une image de graphique), getNamedRanges, setNamedRange, protectWorksheet, addWorksheet, modifyWorkbookStructure, clearRange, getWorksheetInfo, eval_officejs (dernier recours).

### PowerPoint
getSelectedText, replaceSelectedText, proposeShapeTextRevision (propose révision sans modifier directement), searchAndReplaceInShape, replaceShapeParagraphs, getSpeakerNotes, setSpeakerNotes, getSlideContent, getAllSlidesOverview, addSlide, deleteSlide, duplicateSlide, reorderSlide, getShapes, editSlideXml (OOXML), screenshotSlide, searchIcons, insertIcon, insertImageOnSlide, searchAndFormatInPresentation, verifySlides, getCurrentSlideIndex, eval_powerpointjs (dernier recours).

### Outlook
getEmailBody, writeEmailBody, getEmailSubject, setEmailSubject, getEmailRecipients, addRecipient, getEmailSender, addAttachment, eval_outlookjs (dernier recours).

## Format de réponse OBLIGATOIRE

Réponds UNIQUEMENT avec un objet JSON valide (sans bloc markdown autour) :

{
  "name": "string — verbe à l'infinitif ou impératif, 5 mots max",
  "description": "string — 30-50 mots, bénéfice utilisateur clair, 'pushy' sur quand l'utiliser",
  "host": "word|excel|powerpoint|outlook|all",
  "executionMode": "immediate|draft|agent",
  "icon": "string — nom d'icône Lucide : List, Wand2, FileText, Languages, CheckSquare, TrendingUp, Mail, Table, Scissors, Eye, Sparkles, BookOpen, Reply, Database, BarChart, Function, Grid3X3, Globe, HelpCircle, Briefcase, AlignLeft, ListChecks, CheckCircle, GraduationCap, Palette, Zap",
  "skillContent": "string — corps markdown de la skill (sans frontmatter), \n pour les sauts de ligne"
}
`
```

---

## 12. Skill Creator — Frontend

### Composable `useSkillCreator.ts`

```typescript
import { ref } from 'vue'
import { logService } from '@/utils/logger'
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser'

export interface SkillCreatorResult {
  name: string
  description: string
  host: SkillHost
  executionMode: SkillExecutionMode
  icon: string
  skillContent: string
}

export function useSkillCreator() {
  const generating = ref(false)
  const error = ref<string | null>(null)

  async function generateSkill(
    description: string,
    host: SkillHost
  ): Promise<SkillCreatorResult | null> {
    generating.value = true
    error.value = null
    try {
      const res = await fetch('/api/skill-creator', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ description, host }),
      })
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: 'Unknown error' }))
        throw new Error(err.error || `HTTP ${res.status}`)
      }
      return await res.json() as SkillCreatorResult
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown error'
      logService.error('[SkillCreator] Generation failed', err)
      error.value = msg
      return null
    } finally {
      generating.value = false
    }
  }

  return { generating, error, generateSkill }
}
```

### Composant `SkillCreatorModal.vue`

**Structure en 4 étapes** :

```
Step 1 (describe):  Textarea description + Select host
Step 2 (generating): Spinner
Step 3 (review):    Champs éditables + textarea markdown + boutons Test/Regen
Step 4 (testing):   Aperçu de l'exécution sur le texte sélectionné
```

**Pseudo-code du composant** :

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { useSkillCreator } from '@/composables/useSkillCreator'
import { useUserSkills } from '@/composables/useUserSkills'
import type { SkillHost, SkillExecutionMode } from '@/utils/skillParser'
import type { SkillCreatorResult } from '@/composables/useSkillCreator'

const props = defineProps<{ open: boolean }>()
const emit = defineEmits<{
  (e: 'close'): void
  (e: 'skill-created'): void
}>()

type Step = 'describe' | 'generating' | 'review' | 'testing'
const step = ref<Step>('describe')

// Step 1 — Form
const descriptionInput = ref('')
const selectedHost = ref<SkillHost>('all')

// Step 3 — Editable result
const editableName = ref('')
const editableDescription = ref('')
const editableHost = ref<SkillHost>('all')
const editableExecutionMode = ref<SkillExecutionMode>('immediate')
const editableIcon = ref('Zap')
const editableContent = ref('')

const { generating, error, generateSkill } = useSkillCreator()
const { addSkill } = useUserSkills()

async function handleGenerate() {
  if (!descriptionInput.value.trim()) return
  step.value = 'generating'
  const result = await generateSkill(descriptionInput.value, selectedHost.value)
  if (!result) {
    step.value = 'describe'   // show error, let user retry
    return
  }
  // Populate editable fields
  editableName.value = result.name
  editableDescription.value = result.description
  editableHost.value = result.host
  editableExecutionMode.value = result.executionMode
  editableIcon.value = result.icon
  editableContent.value = result.skillContent
  step.value = 'review'
}

async function handleTest() {
  step.value = 'testing'
  // Le composant parent gère l'exécution du test via un emit
  // (le test s'exécute sur le texte sélectionné dans Office)
  emit('test-skill', {
    name: editableName.value,
    skillContent: editableContent.value,
    executionMode: editableExecutionMode.value,
  })
}

function handleSave() {
  addSkill({
    name: editableName.value,
    description: editableDescription.value,
    host: editableHost.value,
    executionMode: editableExecutionMode.value,
    icon: editableIcon.value,
    skillContent: editableContent.value,
  })
  emit('skill-created')
  emit('close')
  // Reset
  step.value = 'describe'
  descriptionInput.value = ''
}

function handleRegenerate() {
  step.value = 'describe'
}
</script>
```

**Note sur le "Test rapide"** : Quand l'utilisateur clique "Tester", le composant `HomePage.vue` écoute l'event `test-skill` et exécute le skill temporaire (sans le sauvegarder) via `executeUserSkill` avec un `UserSkill` éphémère. Le résultat apparaît dans le chat normalement. Après le test, le modal revient à l'étape Review avec le skill visible pour ajustements.

---

## 13. Composant `SkillLibraryTab.vue`

Remplace `PromptsTab.vue` dans `SettingsPage.vue`.

### Structure

```vue
<template>
  <div class="flex h-full w-full flex-col gap-2 p-2">
    <!-- Header -->
    <div class="flex items-center justify-between">
      <h3 class="text-sm font-semibold text-main">{{ t('mySkills') || 'Mes Skills' }}</h3>
      <div class="flex gap-1">
        <CustomButton :icon="Upload" text="" :title="t('importSkill')" @click="triggerImport" />
        <CustomButton :icon="Plus" text="" :title="t('createSkill')" @click="openCreator" />
      </div>
    </div>

    <!-- Filtres par host -->
    <div class="flex gap-1">
      <button v-for="h in hostFilters" :key="h.value"
        :class="activeFilter === h.value ? 'bg-accent text-white' : 'bg-surface'"
        @click="activeFilter = h.value">{{ h.label }}</button>
    </div>

    <!-- Liste des skills -->
    <div class="flex-1 overflow-auto">
      <div v-if="filteredSkills.length === 0" class="text-center text-secondary text-sm py-4">
        {{ t('noSkillsYet') || 'Aucun skill. Créez-en un !' }}
      </div>
      <div v-for="skill in filteredSkills" :key="skill.id"
        class="rounded-md border border-border bg-surface p-3 mb-2">
        <div class="flex items-start justify-between">
          <div>
            <span class="text-sm font-semibold">{{ skill.name }}</span>
            <span class="ml-2 text-xs text-secondary bg-bg-secondary px-1 rounded">
              {{ skill.host }} · {{ skill.executionMode }}
            </span>
          </div>
          <div class="flex gap-1">
            <button @click="exportSkillToFile(skill)" title="Exporter"><Download /></button>
            <button @click="startEdit(skill)" title="Modifier"><Edit2 /></button>
            <button @click="deleteSkill(skill.id)" title="Supprimer"><Trash2 /></button>
          </div>
        </div>
        <p class="mt-1 text-xs text-secondary">
          {{ skill.description.substring(0, 120) }}{{ skill.description.length > 120 ? '...' : '' }}
        </p>
        <!-- Edit form (inline) -->
        <div v-if="editingId === skill.id" class="mt-3 border-t pt-3">
          <!-- Champs nom, description, host, executionMode, icon -->
          <!-- Textarea pour skillContent (markdown) -->
          <div class="flex gap-2 mt-2">
            <CustomButton type="primary" :text="t('save')" @click="saveEdit" />
            <CustomButton type="secondary" :text="t('cancel')" @click="cancelEdit" />
          </div>
        </div>
      </div>
    </div>

    <!-- Input caché pour l'import -->
    <input ref="fileInput" type="file" accept=".md,.skill.md" class="hidden"
      @change="handleImport" />
  </div>
</template>
```

---

## 14. Migration dans `SettingsPage.vue`

```vue
<!-- Remplacer -->
<PromptsTab v-if="currentTab === 'prompts'" />
<!-- Par -->
<SkillLibraryTab v-if="currentTab === 'skills'" />

<!-- Et mettre à jour la tabList -->
// Avant : { id: 'prompts', label: 'prompts', defaultLabel: 'Prompts', icon: MessageSquare }
// Après : { id: 'skills', label: 'skills', defaultLabel: 'Skills', icon: Zap }

// Supprimer aussi l'onglet builtinPrompts si l'équipe décide de le supprimer
// (question ouverte — voir Section 16)
```

---

## 15. Dialog de Migration (`MigrationDialog.vue`)

À afficher au premier lancement si des custom prompts existent. Simple dialog de confirmation.

```vue
<template>
  <div v-if="visible" class="migration-dialog">
    <p>{{ t('migrationTitle') || 'Vos prompts personnalisés ont été remplacés par les Skills' }}</p>
    <p>{{ t('migrationSubtitle', { count: promptCount }) }}</p>
    <div class="flex gap-2 mt-4">
      <CustomButton type="primary" :text="t('convertToSkills')" @click="handleMigrate" />
      <CustomButton type="secondary" :text="t('ignore')" @click="handleIgnore" />
    </div>
  </div>
</template>
```

---

## 16. Décisions complémentaires

| Point | Décision |
|-------|---------|
| **BuiltinPromptsTab** | **Supprimer dans la même passe** — les utilisateurs créent un User Skill à la place |
| **Rate limiting `/api/skill-creator`** | **10 appels / heure / IP** — erreur 429 avec message explicite |
| **Limite de skills** | Illimité (pas de cap pour la beta) |
| **Icône selector UI** | Champ texte libre dans le creator (la valeur est proposée par le LLM) |
| **i18n** | Ajouter les clés listées en Section 19 |

### Impact sur le Lot 5B (Suppression du code mort)

En plus de `PromptsTab.vue`, ajouter à la liste :
- [ ] Supprimer `frontend/src/components/settings/BuiltinPromptsTab.vue`
- [ ] Supprimer les clés i18n `builtinPrompts`, `resetBuiltinPrompts`, `builtinPromptSaved`, etc.
- [ ] Supprimer les entrées localStorage `ki_Settings_BuiltInPrompts_*` si présentes (nettoyage optionnel)
- [ ] Supprimer l'onglet `builtinPrompts` de la `tabList` dans `SettingsPage.vue`

### Rate limiting dans `backend/src/server.js`

```javascript
import rateLimit from 'express-rate-limit'

const skillCreatorRateLimit = rateLimit({
  windowMs: 60 * 60 * 1000,  // 1 heure
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Trop de générations de skills. Réessayez dans une heure.' },
  keyGenerator: (req) => req.ip,
})

app.use('/api/skill-creator', skillCreatorRateLimit, skillCreatorRouter)
```

---

## 17. Plan d'implémentation — Phases et Lots

### Phase 1 — Format unifié (built-in skills) `~2-3h`

**Lot 1A — Utilitaires**
- [ ] Créer `frontend/src/utils/skillParser.ts` (tel que défini en Section 5)
- [ ] Créer `frontend/src/types/userSkill.ts` (Section 6)
- [ ] Test unitaire de `parseSkill()` : avec frontmatter, sans frontmatter, frontmatter malformé

**Lot 1B — Retrofit des built-in skills**
- [ ] Ajouter le frontmatter YAML aux 24 fichiers `quickactions/*.skill.md` (tableau Section 5)
  - Attention : le corps des skills ne change PAS, seulement le frontmatter est ajouté en tête
  - Format exact : `---\nname: ...\ndescription: "..."\nhost: ...\nexecutionMode: ...\nicon: ...\nactionKey: ...\n---\n\n[corps existant]`
- [ ] Refactorer `frontend/src/skills/index.ts` :
  - Remplacer les noms des imports `xxxSkill` → `xxxSkillRaw` (convention)
  - Remplacer la `const quickActionSkillMap` par `const parsedSkills: Map<string, ParsedSkill>`
  - Ajouter `getQuickActionSkillMetadata()` et `getAllBuiltInSkillsMetadata()`
  - Garder `getQuickActionSkill()` avec la même signature (rétro-compat `useQuickActions.ts`)

**Validation Phase 1** : `npm run build` sans erreur + les quick actions built-in fonctionnent comme avant.

---

### Phase 2 — Fondations User Skills `~2-3h`

**Lot 2A — Composable**
- [ ] Créer `frontend/src/composables/useUserSkills.ts` (Section 7)
  - `loadFromStorage`, `saveToStorage`, `addSkill`, `updateSkill`, `deleteSkill`
  - `exportSkillToFile`, `importSkillFromFile`
  - `skillsForHost`
  - `checkAndMigrateOldPrompts`, `migrateOldPrompts`, `confirmMigrationDone`

**Lot 2B — SkillLibraryTab**
- [ ] Créer `frontend/src/components/settings/SkillLibraryTab.vue` (Section 13)
  - Affichage liste avec filtre par host
  - Édition inline (nom, description, mode, icône, corps markdown)
  - Boutons export / import / delete
  - Import via `<input type="file">` caché
- [ ] Mettre à jour `SettingsPage.vue` :
  - Remplacer l'import et l'usage de `PromptsTab` par `SkillLibraryTab`
  - Mettre à jour la `tabList` : `'prompts'` → `'skills'`

**Validation Phase 2** : Ouvrir les settings → onglet Skills visible → créer/modifier/supprimer/exporter/importer un skill.

---

### Phase 3 — Intégration QuickActionsBar `~1-2h`

**Lot 3A — Exécution**
- [ ] Créer `frontend/src/composables/useUserSkillExecution.ts` (Section 8)

**Lot 3B — Mise à jour HomePage + QuickActionsBar**
- [ ] Mettre à jour `QuickActionsBar.vue` : remplacer `savedPrompts` / `selectedPromptId` / `load-prompt` par `userSkillsForHost` / `execute-user-skill` / `open-skill-creator` (Section 9)
- [ ] Mettre à jour `HomePage.vue` : ajouter `useUserSkills` + `useUserSkillExecution`, supprimer les refs `customSystemPrompt` / `selectedPromptId` (Section 10)
- [ ] Mettre à jour `useHomePage.ts` : supprimer `loadSelectedPrompt`, `customSystemPrompt`, `selectedPromptId` des paramètres et du retour
- [ ] Mettre à jour `useAgentLoop.ts` : supprimer le paramètre `customSystemPrompt: Ref<string>` et l.749
- [ ] Mettre à jour `useHomePageContext.ts` : supprimer les types correspondants

**Validation Phase 3** : Sélectionner un user skill dans le dropdown → exécution immédiate sans pré-remplissage de la textarea.

---

### Phase 4 — Skill Creator `~3-4h`

**Lot 4A — Backend**
- [ ] Créer `backend/src/config/skillCreatorPrompt.js` (system prompt Section 11)
- [ ] Créer `backend/src/routes/skillCreator.js` (Section 11)
- [ ] Enregistrer la route dans `backend/src/server.js`
- [ ] Tester manuellement l'endpoint : `POST /api/skill-creator` avec `{ description: "Reformuler en 5 bullets", host: "powerpoint" }`

**Lot 4B — Frontend**
- [ ] Créer `frontend/src/composables/useSkillCreator.ts` (Section 12)
- [ ] Créer `frontend/src/components/skills/SkillCreatorModal.vue` (Section 12)
  - Step 1 : Describe (textarea + host select)
  - Step 2 : Generating (spinner)
  - Step 3 : Review (champs éditables + markdown textarea)
  - Step 4 : Testing (voir ci-dessous)
- [ ] Connecter le bouton `+` dans `QuickActionsBar.vue` → ouvre `SkillCreatorModal`
- [ ] Gérer le test rapide dans `HomePage.vue` : écouter l'event `test-skill` du modal → `executeUserSkill` avec skill éphémère → résultat dans le chat
- [ ] Affiner le system prompt du creator après 5-10 tests réels (itérer sur la qualité)

**Validation Phase 4** : Créer un skill via le creator → tester → sauvegarder → exécuter depuis le dropdown.

---

### Phase 5 — Migration & Nettoyage `~1h`

**Lot 5A — Migration**
- [ ] Créer `frontend/src/components/MigrationDialog.vue` (Section 15)
- [ ] Intégrer dans `HomePage.vue` : `onMounted` → `checkAndMigrateOldPrompts()` → afficher dialog
- [ ] Tester le flux de migration avec des custom prompts existants en localStorage

**Lot 5B — Suppression du code mort**
- [ ] Supprimer `frontend/src/utils/savedPrompts.ts`
- [ ] Supprimer `frontend/src/components/settings/PromptsTab.vue`
- [ ] Décider du sort de `BuiltinPromptsTab.vue` (cf. Section 16)
- [ ] Supprimer les clés i18n orphelines liées aux custom prompts

**Validation Phase 5** : `npm run build` sans warning sur les imports non utilisés. Les custom prompts sont effacés du codebase.

---

## 18. Risques & Mitigations

| Risque | Impact | Mitigation |
|--------|--------|-----------|
| Skill creator génère un JSON invalide | Moyen | Retry x1 côté backend + validation champs requis |
| Skill creator génère un skill de mauvaise qualité | Fort | System prompt très prescriptif + étape de test avant sauvegarde + bouton Régénérer |
| Parser YAML frontmatter fragile (quotes, caractères spéciaux) | Moyen | Parser défensif avec fallback + tests unitaires couvrant les edge cases |
| localStorage plein (QuotaExceededError) | Faible | Déjà géré dans le pattern existant, à répliquer dans `useUserSkills` |
| Import d'un `.skill.md` malformé par un tiers | Faible | Validation + warning UI + corps brut en fallback |
| `customSystemPrompt` utilisé ailleurs (recherche à faire) | Potentiel | `grep -rn "customSystemPrompt"` avant de supprimer — exactement 7 occurrences connues |
| Breaking change `QuickActionsBar` props | Moyen | Les props changent (savedPrompts → userSkillsForHost) — faire en un seul commit pour éviter l'état intermédiaire cassé |

---

## 19. Clés i18n à ajouter

```json
{
  "mySkills": "Mes Skills",
  "createSkill": "Créer un skill",
  "importSkill": "Importer un skill",
  "noSkillsYet": "Aucun skill. Créez-en un !",
  "skills": "Skills",
  "skillCreator": "Créateur de skills",
  "describeSkill": "Décrivez ce que vous voulez faire...",
  "generating": "Génération en cours...",
  "testSkill": "Tester sur la sélection",
  "regenerate": "Régénérer",
  "saveSkill": "Sauvegarder le skill",
  "migrationTitle": "Évolution : les Prompts deviennent des Skills",
  "migrationSubtitle": "Vous avez {count} prompts personnalisés. Voulez-vous les convertir ?",
  "convertToSkills": "Convertir en skills",
  "skillExported": "Skill exporté",
  "skillImported": "Skill importé",
  "skillCreated": "Skill créé et ajouté à votre bibliothèque"
}
```
