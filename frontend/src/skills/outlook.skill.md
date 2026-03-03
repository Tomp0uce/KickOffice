# Outlook Office.js Skill

## CRITICAL OUTLOOK-SPECIFIC RULES

### Rule 1: Outlook uses the Common API pattern differently

Outlook doesn't use `Outlook.run()`. Instead, use the mailbox API directly:

```javascript
const item = Office.context.mailbox.item;
```

### Rule 2: Body content can be HTML or text

**Read body as text:**
```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Text,
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log(result.value);
    }
  }
);
```

**Read body as HTML:**
```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Html,
  (result) => {
    console.log(result.value);  // Full HTML
  }
);
```

### Rule 3: Writing uses setAsync with coercion type

**Write text:**
```javascript
Office.context.mailbox.item.body.setAsync(
  'Plain text content',
  { coercionType: Office.CoercionType.Text },
  (result) => { /* handle result */ }
);
```

**Write HTML:**
```javascript
Office.context.mailbox.item.body.setAsync(
  '<p>HTML <b>content</b></p>',
  { coercionType: Office.CoercionType.Html },
  (result) => { /* handle result */ }
);
```

### Rule 4: Prepend/Append instead of Replace when possible

**Safer — preserves existing content:**
```javascript
Office.context.mailbox.item.body.prependAsync(
  '<p>New content at start</p>',
  { coercionType: Office.CoercionType.Html },
  (result) => { }
);
```

### Rule 5: Reply in the SAME language as the original email

**CRITICAL**: When the user asks you to reply to an email:
1. Read the existing email body first
2. Detect the language
3. Reply in THAT language, not the user's interface language

### Rule 6: Callback pattern (not async/await)

Outlook uses callbacks, not Promises:

```javascript
// WRONG — This won't work
const body = await Office.context.mailbox.item.body.getAsync(...);

// CORRECT — Use callback
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Text,
  (result) => {
    // Handle result here
  }
);

// Or wrap in Promise:
const body = await new Promise((resolve, reject) => {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Text,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    }
  );
});
```

## AVAILABLE TOOLS

### For READING:
| Tool | When to use |
|------|-------------|
| `getEmailBody` | Get full email body |
| `getEmailSubject` | Get subject line |
| `getEmailRecipients` | Get To/CC/BCC |
| `getEmailSender` | Get sender info |

### For WRITING:
| Tool | When to use |
|------|-------------|
| `writeEmailBody` | **PREFERRED** — Write with mode: Append/Insert/Replace |
| `setEmailSubject` | Update subject |
| `addRecipient` | Add To/CC/BCC recipients |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
| `eval_outlookjs` | Attachments, HTML manipulation, metadata |

## COMMON PATTERNS

### Read email content
```javascript
const item = Office.context.mailbox.item;

// Subject
item.subject.getAsync((result) => {
  console.log('Subject:', result.value);
});

// Body
item.body.getAsync(Office.CoercionType.Text, (result) => {
  console.log('Body:', result.value);
});

// Sender
console.log('From:', item.from.displayName, item.from.emailAddress);
```

### Write email body
```javascript
const content = `
Dear Colleague,

Thank you for your email.

Best regards
`;

Office.context.mailbox.item.body.setAsync(
  content,
  { coercionType: Office.CoercionType.Text },
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Body updated');
    }
  }
);
```

### Add recipient
```javascript
Office.context.mailbox.item.to.addAsync(
  [{ displayName: 'John Doe', emailAddress: 'john@example.com' }],
  (result) => { }
);
```

## COMPOSE vs READ MODE

Outlook items have different capabilities based on mode:

**Compose mode** (writing new email):
- Can modify subject, body, recipients
- Full write access

**Read mode** (viewing received email):
- Read-only access to content
- Can reply/forward but not modify original
