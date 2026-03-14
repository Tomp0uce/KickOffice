# Extract Quick Action Skill

## Purpose
Extract structured information (action items, deadlines, key points, decisions) from unstructured email content and present it in an organized, scannable format.

## When to Use
- User clicks "Extract" Quick Action in Outlook
- Email contains information that needs to be actioned or tracked
- Goal: Pull out actionable items and key facts from email body

## Input Contract
- **Selected text**: Email content (may contain action items, dates, decisions, requests buried in prose)
- **Language**: Preserve the language of the original text
- **Context**: Email (meeting notes, project updates, requests, discussions)
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders (Outlook)

## Output Requirements
1. **Extract action items**: Who needs to do what by when
2. **Extract deadlines**: All dates and time-sensitive information
3. **Extract key decisions**: Conclusions reached or approvals given
4. **Extract questions**: Outstanding questions requiring answers
5. **Structure clearly**: Use headings and bullets for scannability
6. **Return extraction only**: No preamble like "Here's what I found"
7. **Keep placeholders**: `{{PRESERVE_N}}` markers unchanged

## Extraction Categories

### 1. Action Items
Format:
```
## Action Items
- **[Person Name]**: [Action] — [Deadline if mentioned]
- **[Person Name]**: [Action] — [Status: pending/in progress/completed]
```

Extract:
- Explicit tasks ("Please send", "Can you review", "You need to")
- Implicit tasks ("We should", "It would be good to")
- Delegations ("Sarah will", "The team needs to")

### 2. Deadlines & Dates
Format:
```
## Deadlines
- **[Date]**: [What is due]
- **[Date]**: [Event/meeting]
```

Extract:
- Explicit dates ("by March 20", "due Friday")
- Relative dates ("end of week", "next month")
- Convert relative to absolute when possible (e.g., "tomorrow" → specific date if today's date is known from context)

### 3. Decisions Made
Format:
```
## Decisions
- [Decision statement]
- [Approval/rejection with context]
```

Extract:
- Approvals ("approved", "green light", "let's proceed")
- Rejections ("not moving forward", "declined")
- Choices made ("we'll go with Option B")

### 4. Questions / Open Items
Format:
```
## Questions
- [Question from email]
- [Unresolved issue requiring clarification]
```

Extract:
- Direct questions
- Implicit uncertainties ("not sure if", "we need to clarify")

### 5. Key Information
Format:
```
## Key Points
- [Important fact or update]
- [Critical information for context]
```

Extract:
- Budget figures
- Project status updates
- Important changes or announcements

## Tool Usage
**DO NOT** call Office.js tools. Return pure structured text.

## Example Extractions

### Example 1: Project Update Email
**Input**:
```
Hi team,

Quick update on the Q1 website redesign project. We've completed the wireframes and the client approved them yesterday—great news! However, they want to see high-fidelity mockups by March 25 before we start development.

Sarah, can you coordinate with the design team to get those mockups done? Also, we need to finalize the hosting provider decision. I'm leaning towards AWS but want to hear your thoughts.

The development sprint will start April 1, assuming mockups are approved. John, please update the project timeline in Jira to reflect this.

One concern: the client mentioned they might want to add a blog section, but they're not sure yet. We should follow up on this next week.

Let me know if any issues!
```

**Output**:
```
## Action Items
- **Sarah**: Coordinate with design team for high-fidelity mockups — Due: March 25
- **John**: Update project timeline in Jira to reflect April 1 dev sprint start
- **Team**: Follow up with client about potential blog section — Next week

## Deadlines
- **March 25**: High-fidelity mockups due
- **April 1**: Development sprint starts (pending mockup approval)

## Decisions
- Wireframes approved by client

## Questions
- Hosting provider decision: AWS vs alternatives? (Input needed)
- Will client add blog section? (Pending confirmation next week)

## Key Points
- Q1 website redesign project on track
- Client approval required before dev sprint begins
```

### Example 2: Meeting Follow-up Email
**Input**:
```
Hi everyone,

Thanks for joining today's product roadmap meeting. Here's a quick recap:

We agreed to prioritize the mobile app over the desktop version for Q2. Engineering will start the iOS build first (targeting June release), followed by Android in Q3. Marketing needs to update the go-to-market plan to reflect this change.

Budget-wise, we're approved for $150K for mobile development. However, if we want to add push notifications, we'll need an additional $20K—waiting for CFO approval on that.

Open question: Do we want to support offline mode in v1? Alex mentioned it could delay launch by 3 weeks. Let's discuss in next week's planning session (March 21, 2pm).

David, can you share the competitor analysis by Friday? That will help inform our feature decisions.

Thanks,
Lisa
```

**Output**:
```
## Action Items
- **Engineering**: Start iOS mobile app build — Target: June release
- **Engineering**: Begin Android build after iOS — Target: Q3
- **Marketing**: Update go-to-market plan to reflect mobile-first strategy
- **David**: Share competitor analysis — Due: Friday

## Deadlines
- **Friday**: Competitor analysis due
- **March 21, 2pm**: Planning session to discuss offline mode
- **June**: iOS app release target
- **Q3**: Android app release target

## Decisions
- Prioritize mobile app over desktop for Q2
- iOS build starts first, Android follows
- $150K budget approved for mobile development

## Questions
- Support offline mode in v1? (Could delay launch by 3 weeks) — Discuss March 21
- Push notifications feature? (Requires additional $20K, pending CFO approval)

## Key Points
- Mobile-first strategy for Q2
- Budget: $150K approved, +$20K for push notifications pending
```

### Example 3: Client Request Email
**Input**:
```
Good morning,

Following up on our call yesterday. We'd like to move forward with the proposal but need a few modifications:

1. Can you extend the support period from 6 months to 12 months? We're fine with the additional cost.
2. The training sessions should be on-site rather than virtual—our team learns better in person.
3. We need the project completed by August 15 due to our fiscal year-end. Is that feasible?

Also, please send the revised contract by next Tuesday so our legal team can review. They typically need 2 weeks for approval.

Let me know if you have any concerns with these changes.

Best,
Michael
```

**Output**:
```
## Action Items
- **Us**: Extend support period from 6 months to 12 months (calculate additional cost)
- **Us**: Change training sessions from virtual to on-site
- **Us**: Send revised contract — Due: Next Tuesday
- **Us**: Confirm feasibility of August 15 project completion

## Deadlines
- **Next Tuesday**: Revised contract due to client
- **August 15**: Project completion deadline (client requirement)
- **~2 weeks after contract sent**: Legal approval expected

## Decisions
- Client wants to proceed with proposal (with modifications)
- Client accepts additional cost for extended support

## Questions
- Can we meet August 15 deadline? (Requires internal assessment)

## Key Points
- Support period: 6 months → 12 months
- Training: virtual → on-site
- Legal review requires 2 weeks
- Driven by client's fiscal year-end
```

### Example 4: French Email with Image
**Input**:
```
Bonjour,

Suite à notre réunion, voici les points à retenir:

Le budget est approuvé (50K€) mais il faut réduire les coûts de 10% sur les frais généraux. Marie, peux-tu revoir les estimations d'ici vendredi?

Le planning prévoit une livraison le 30 avril. Le client a demandé une démo intermédiaire le 15 avril—est-ce faisable?

Voir le diagramme joint {{PRESERVE_0}} pour l'architecture technique. Paul, tu peux valider l'approche proposée?

Point en suspens: faut-il intégrer l'API externe dès la v1 ou reporter à la v2? Décision à prendre cette semaine.

Cordialement,
```

**Output**:
```
## Actions
- **Marie**: Revoir les estimations pour réduire coûts de 10% — Échéance: Vendredi
- **Paul**: Valider l'approche technique (voir diagramme {{PRESERVE_0}})
- **Équipe**: Confirmer faisabilité de la démo du 15 avril
- **Équipe**: Décider intégration API externe (v1 vs v2) — Cette semaine

## Échéances
- **Vendredi**: Estimations révisées dues
- **15 avril**: Démo intermédiaire client
- **30 avril**: Livraison finale

## Décisions
- Budget approuvé: 50K€
- Réduction de 10% requise sur frais généraux

## Questions
- Démo du 15 avril faisable?
- API externe en v1 ou v2? (Décision cette semaine)

## Points Clés
- Diagramme architecture: {{PRESERVE_0}}
```

## Preservation Rules (Outlook)
- Keep `{{PRESERVE_N}}` placeholders EXACTLY as-is
- Reference them naturally in extracted content
- Don't "extract" or describe them

## Edge Cases
- **No action items**: Still extract decisions, dates, key points if present
- **Informal email**: Extract even if phrasing is casual ("can u send me...")
- **Implicit information**: Infer action items from context ("We should..." → action item)
- **Multiple topics**: Use additional headings to separate unrelated items
- **FYI emails**: Focus on "Key Points" section if no actions

## Quality Check
After extracting, verify:
- ✓ All action items identified with owners?
- ✓ All deadlines captured?
- ✓ Decisions clearly stated?
- ✓ Questions/open items highlighted?
- ✓ Scannable format (bullets, headings)?

## Extract vs Other Actions
- **Extract** = pull out structured info (actions, dates, decisions)
- **Concise** = shorten the entire email text
- **Summary** = create prose overview
- **Proofread** = fix errors
