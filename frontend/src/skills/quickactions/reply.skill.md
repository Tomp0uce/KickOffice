# Reply Quick Action Skill

## Purpose

Generate an appropriate email reply based on user's intent and the original email context, maintaining professional tone and addressing all points.

## When to Use

- User clicks "Reply" Quick Action in Outlook (smart-reply mode)
- User has specified reply intent in the text field (e.g., "accept the meeting" or "decline politely")
- Goal: Draft a complete, contextually appropriate reply

## Input Contract

- **Reply intent**: User's instruction for what the reply should say (from text field after clicking Reply)
- **Original email context**: The email being replied to (injected by system)
- **Language**: Match the language of the original email
- **Rich content**: May contain `{{PRESERVE_N}}` placeholders if original email has images

## Identifying the Correct Recipient

**CRITICAL — always identify the sender of the MOST RECENT message in the thread.**

Email threads display messages in reverse-chronological order: the **most recent message is at the top** of the body. The original sender is at the bottom.

**Parsing rule:**
1. Look for the FIRST occurrence of "**De :**", "**From:**", or "**De:**" in the email body
2. Extract the sender name from that first block — that is the person who sent the latest message
3. Address the greeting to that person (e.g. "Hi Nathan," not the original thread starter)
4. If the thread is a group email and the reply is general, use "Hi all," or "Hi team,"

**Example thread structure:**
```
De : Nathan Olff <nathan@...>   ← MOST RECENT — reply to Nathan
Envoyé : 12 mars 2026 09:30
...Nathan's message...

---

De : Eric Maussion <eric@...>   ← older
Envoyé : 12 mars 2026 09:27
...Eric's message...

---

De : Esteban Francou <...>      ← even older
Date : 11 mars 2026
...original message...
```
→ In this case, always reply to **Nathan** (most recent sender), NOT Eric or Esteban.

## Output Requirements

1. **Address the intent**: Accomplish what the user requested
2. **Correct recipient**: Address the reply to the sender of the MOST RECENT message (see above)
3. **Reference original email**: Acknowledge key points from the message being replied to
4. **Professional tone**: Courteous, clear, appropriately formal
5. **Complete response**: Include greeting, body, closing
6. **Match language**: Reply in the same language as the original email
7. **Actionable if needed**: Include next steps, confirmations, or questions
8. **Return email text only**: No meta-commentary like "Here's a suggested reply"

## Reply Types

### 1. Acceptance

User intent examples: "accept", "yes", "I'll attend", "sounds good"

Response should:

- Clearly confirm acceptance
- Acknowledge key details (date, time, location)
- Express enthusiasm (when appropriate)
- Ask clarifying questions if needed

### 2. Decline

User intent examples: "decline politely", "can't make it", "need to pass"

Response should:

- Politely decline
- Provide brief reason (if user indicated one)
- Suggest alternative (if applicable)
- Maintain positive relationship

### 3. Request Information

User intent examples: "ask for more details", "need clarification on budget", "what's the deadline?"

Response should:

- Acknowledge the original message
- List specific questions clearly
- Explain why information is needed (if relevant)
- Provide context for the request

### 4. Confirmation

User intent examples: "confirm receipt", "acknowledge", "got it thanks"

Response should:

- Confirm understanding
- Summarize key points (if complex)
- State next steps or expected actions

### 5. Follow-up

User intent examples: "follow up on status", "checking in", "any updates?"

Response should:

- Reference previous communication
- Politely request update
- Provide context for urgency (if applicable)

## Tool Usage

After generating the reply text, call `writeEmailBody` with `mode: "Prepend"` to insert it **before** the quoted thread history (standard email convention). Never use `"Append"` for replies — it places the text after the thread, which is incorrect.

## Example Replies

### Example 1: Meeting Acceptance

**Original Email**:

```
Hi John,

Would you be available for a project kickoff meeting this Thursday, March 16 at 2pm? We'll be in Conference Room B and expect the session to last about an hour.

Please confirm your attendance.

Best,
Sarah
```

**User Intent**: "accept"

**Generated Reply**:

```
Hi Sarah,

Yes, I'll be there! Thursday, March 16 at 2pm in Conference Room B works perfectly for me.

Looking forward to the kickoff.

Best,
John
```

---

### Example 2: Polite Decline with Alternative

**Original Email**:

```
Hello,

We'd love to have you speak at our annual conference on April 10-12 in Boston. Your session would be on April 11 at 10am. We can cover travel and accommodation expenses.

Let me know if you're interested!

Thanks,
David
```

**User Intent**: "decline politely, I have a conflict but suggest they check with my colleague Maria"

**Generated Reply**:

```
Hi David,

Thank you so much for the invitation—I'm honored you thought of me. Unfortunately, I have a prior commitment during that time and won't be able to attend.

However, I'd like to suggest my colleague Maria Rodriguez, who has deep expertise in this area and is an excellent speaker. If you'd like, I can introduce you.

I hope the conference is a great success!

Best regards,
[Your name]
```

---

### Example 3: Request for Information

**Original Email**:

```
Team,

The client has requested we accelerate the timeline for the website redesign. They'd like to launch by end of Q2 instead of Q3. This means moving everything up by about 6 weeks.

Thoughts?

Alex
```

**User Intent**: "ask what this means for the budget and if we have capacity"

**Generated Reply**:

```
Hi Alex,

Thanks for the heads up. Before we commit to the accelerated timeline, I have a couple of questions:

1. **Budget**: Will the client approve additional budget for the expedited timeline? We'll likely need overtime or additional contractor support.

2. **Team capacity**: Our current team is at capacity through mid-May. Do we have flexibility to shift other projects, or should we plan to bring in external resources?

Once we have clarity on these points, I can assess feasibility from our side.

Best,
[Your name]
```

---

### Example 4: Confirmation with Summary

**Original Email**:

```
Hi Lisa,

Just confirming our discussion from today's call:

1. You'll send the revised proposal by Friday
2. We'll schedule a follow-up meeting next Tuesday at 3pm
3. Legal will need 2 weeks to review the contract

Sound right?

Michael
```

**User Intent**: "confirm everything is correct"

**Generated Reply**:

```
Hi Michael,

Yes, that's all correct! To recap:

1. Revised proposal will be sent by Friday
2. Follow-up meeting scheduled for next Tuesday at 3pm
3. Legal review will take approximately 2 weeks

I'll send over the proposal by end of week. Looking forward to our Tuesday call.

Best,
Lisa
```

---

### Example 5: Follow-up on Pending Item

**Original Email** (sent 1 week ago):

```
Hi team,

Please review the attached Q1 budget draft and send me your feedback by next Monday.

Thanks,
Jennifer
```

**User Intent**: "follow up, politely ask for the feedback"

**Generated Reply**:

```
Hi Jennifer,

Just following up on the Q1 budget draft you sent last week. I wanted to make sure I didn't miss your feedback or any updates.

If you still need more time to review, please let me know—happy to adjust the timeline if needed.

Thanks!
[Your name]
```

---

### Example 6: French Reply

**Original Email**:

```
Bonjour,

Je vous confirme notre rendez-vous pour le 20 mars à 14h dans nos bureaux. Pourriez-vous apporter les documents relatifs au projet dont nous avons discuté?

Cordialement,
Marie
```

**User Intent**: "confirmer et demander l'adresse exacte"

**Generated Reply**:

```
Bonjour Marie,

Je confirme ma présence le 20 mars à 14h. J'apporterai tous les documents relatifs au projet.

Pourriez-vous me communiquer l'adresse exacte de vos bureaux?

Merci et à bientôt,
[Votre nom]
```

## Reply Tone Guidelines

### Professional but Warm

- Use "Thanks" not "Thx"
- "Looking forward" not "Can't wait"
- "I appreciate" not "You're awesome"
- Keep exclamation marks to 1-2 maximum

### Match Formality

- If original email is formal ("Dear Sir/Madam"), mirror that formality
- If original is casual ("Hey!"), you can be slightly casual (but still professional)

### Language Matching

- **CRITICAL**: Reply in the same language as the original email
- French email → French reply
- Spanish email → Spanish reply
- English email → English reply

## Edge Cases

- **Vague user intent**: "respond appropriately" → Ask user for clarification via error message (system should handle this)
- **Conflicting instructions**: If user says "accept but mention I might be late" → Handle both ("Yes, I'll attend, though I may arrive 10 minutes late")
- **Multiple questions in original**: Address all points unless user intent is very specific
- **Informal original email**: Mirror tone but stay professional (don't be overly casual)

## Quality Check

After generating reply, verify:

- ✓ Addresses user's intent?
- ✓ References key points from original email?
- ✓ Professional tone?
- ✓ Correct language?
- ✓ Complete (greeting + body + closing)?
- ✓ Clear next steps (if applicable)?

## Reply vs Other Actions

- **Reply** = generate new email in response to original (smart-reply mode)
- **Translate** = convert existing text to another language
- **Concise** = shorten existing text
- **Proofread** = fix errors in existing text
- **Extract** = pull out action items from email
