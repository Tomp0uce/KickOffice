import DiffMatchPatch from 'diff-match-patch'

import DOMPurify from 'dompurify'

import { executeOfficeAction } from './officeAction'
import { renderOfficeRichHtml } from './officeRichText'
import { sandboxedEval } from './sandbox'

import { generateVisualDiff } from './common'

export type OutlookToolName =
  | 'getEmailBody'
  | 'getSelectedText'
  | 'setEmailBody'
  | 'insertTextAtCursor'
  | 'setEmailBodyHtml'
  | 'getEmailSubject'
  | 'setEmailSubject'
  | 'getEmailRecipients'
  | 'addRecipient'
  | 'getEmailSender'
  | 'getEmailDate'
  | 'getAttachments'
  | 'insertHtmlAtCursor'
  | 'eval_outlookjs'

type OutlookToolDefinition = WordToolDefinition

type RecipientField = 'to' | 'cc' | 'bcc'

function getMailbox(): any | null {
  return (window as any).Office?.context?.mailbox ?? null
}

function getOfficeAsyncStatus(): any {
  return (window as any).Office?.AsyncResultStatus
}

function getOfficeCoercionType(): any {
  return (window as any).Office?.CoercionType
}

const runOutlook = <T>(action: () => Promise<T>): Promise<T> =>
  executeOfficeAction(action)

type OutlookToolTemplate = Omit<OutlookToolDefinition, 'execute'> & {
  executeOutlook: (mailbox: any | null, args: Record<string, any>) => Promise<string>
}

function createOutlookTools(definitions: Record<OutlookToolName, OutlookToolTemplate>): Record<OutlookToolName, OutlookToolDefinition> {
  return Object.fromEntries(
    Object.entries(definitions).map(([name, definition]) => [
      name,
      {
        ...definition,
        execute: async (args: Record<string, any> = {}) => runOutlook(async () => {
          return Promise.race([
            definition.executeOutlook(getMailbox(), args),
            new Promise<string>(resolve => setTimeout(() => resolve('Error: Outlook API request timed out after 3 seconds.'), 3000))
          ])
        }),
      },
    ]),
  ) as unknown as Record<OutlookToolName, OutlookToolDefinition>
}

function resolveAsyncResult(result: any, onSuccess: (value: any) => string): string {
  if (result.status === getOfficeAsyncStatus()?.Succeeded) {
    return onSuccess(result.value)
  }
  return `Error: ${result.error?.message || 'unknown error'}`
}

function normalizeRecipient(recipient: any): { displayName: string; emailAddress: string } {
  if (!recipient) {
    return { displayName: '', emailAddress: '' }
  }

  if (typeof recipient === 'string') {
    return { displayName: '', emailAddress: recipient.trim() }
  }

  return {
    displayName: recipient.displayName || recipient.name || '',
    emailAddress: recipient.emailAddress || recipient.address || '',
  }
}

function normalizeRecipientsInput(recipients: any): any[] {
  if (Array.isArray(recipients)) {
    return recipients
      .map(normalizeRecipient)
      .filter(r => !!r.emailAddress)
  }

  if (typeof recipients === 'string') {
    return recipients
      .split(',')
      .map(email => normalizeRecipient(email))
      .filter(r => !!r.emailAddress)
  }

  if (recipients && typeof recipients === 'object') {
    const normalized = normalizeRecipient(recipients)
    return normalized.emailAddress ? [normalized] : []
  }

  return []
}

function getRecipientField(field: unknown): RecipientField {
  const value = String(field || 'to').toLowerCase()
  if (value === 'cc' || value === 'bcc') return value
  return 'to'
}

const outlookToolDefinitions = createOutlookTools({
  getEmailBody: {
    name: 'getEmailBody',
    category: 'read',
    description:
      'Get the full body text of the current email. Works in both read and compose mode.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'
      return new Promise<string>((resolve) => {
        mailbox.item.body.getAsync(
          getOfficeCoercionType().Text,
          (result: any) => {
            resolve(resolveAsyncResult(result, (value) => value || ''))
          },
        )
      })
    },
  },

  getSelectedText: {
    name: 'getSelectedText',
    category: 'read',
    description:
      'Get the currently selected text in the email compose window. Returns empty string if nothing is selected or not in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return ''
      if (typeof mailbox.item.getSelectedDataAsync !== 'function') {
        return 'Selection reading is not available in this context.'
      }
      return new Promise<string>((resolve) => {
        mailbox.item.getSelectedDataAsync(
          getOfficeCoercionType().Text,
          (result: any) => {
            if (result.status === getOfficeAsyncStatus()?.Succeeded && result.value?.data) {
              resolve(result.value.data)
            } else {
              resolve('')
            }
          },
        )
      })
    },
  },

  setEmailBody: {
    name: 'setEmailBody',
    category: 'write',
    description:
      'Replace the entire email body with the provided text. Only works in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to set as the email body',
        },
      },
      required: ['text'],
    },
    executeOutlook: async (mailbox, args) => {
      const { text } = args
      if (!mailbox?.item?.body?.setAsync) {
        return 'Cannot set email body: compose mode is not available.'
      }
      return new Promise<string>((resolve) => {
        mailbox.item.body.setAsync(
          text,
          { coercionType: getOfficeCoercionType().Text },
          (result: any) => {
            resolve(resolveAsyncResult(result, () => 'Successfully set email body.'))
          },
        )
      })
    },
  },

  insertTextAtCursor: {
    name: 'insertTextAtCursor',
    category: 'write',
    description: 'Insert Markdown-formatted text at the cursor position in the email body (compose mode). The text is automatically converted to rich HTML before insertion. Use Markdown: **bold**, *italic*, - bullets, 1. numbered lists, indented sub-items (2 spaces).',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to insert at the cursor position',
        },
        diffTracking: {
          type: 'boolean',
          description: 'R17: When true, compares the inserted text with the currently selected text and shows a visual diff (insertions in blue/underline, deletions in red/strikethrough). Useful for proofreading corrections in compose mode. Default: false.',
        },
        originalText: {
          type: 'string',
          description: 'Required when diffTracking is true: the original text to compare against.',
        },
      },
      required: ['text'],
    },
    executeOutlook: async (mailbox, args) => {
      const { text, diffTracking = false, originalText = '' } = args
      if (!mailbox?.item?.body?.setSelectedDataAsync) {
        return 'Cannot insert text at cursor: compose mode is not available.'
      }

      // R17: if diffTracking is requested, generate a visual diff instead of raw text
      const html = diffTracking && originalText
        ? generateVisualDiff(originalText, text)
        : renderOfficeRichHtml(text)

      return new Promise<string>((resolve) => {
        mailbox.item.body.setSelectedDataAsync(
          html,
          { coercionType: getOfficeCoercionType().Html },
          (result: any) => {
            resolve(resolveAsyncResult(result, () => diffTracking
              ? 'Inserted visual diff: insertions in blue/underline, deletions in red/strikethrough.'
              : 'Successfully inserted text at cursor.'))
          },
        )
      })
    },
  },

  setEmailBodyHtml: {
    name: 'setEmailBodyHtml',
    category: 'write',
    description: 'Replace the full email body with HTML content (compose mode).',
    inputSchema: {
      type: 'object',
      properties: {
        html: {
          type: 'string',
          description: 'The HTML content to set as the whole email body',
        },
      },
      required: ['html'],
    },
    executeOutlook: async (mailbox, args) => {
      const { html } = args
      if (!mailbox?.item?.body?.setAsync) {
        return 'Cannot set HTML email body: compose mode is not available.'
      }

      return new Promise<string>((resolve) => {
        mailbox.item.body.setAsync(
          DOMPurify.sanitize(html),
          { coercionType: getOfficeCoercionType().Html },
          (result: any) => {
            resolve(resolveAsyncResult(result, () => 'Successfully set HTML email body.'))
          },
        )
      })
    },
  },

  getEmailSubject: {
    name: 'getEmailSubject',
    category: 'read',
    description: 'Get the current email subject in read or compose mode.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'

      if (mailbox.item.subject && typeof mailbox.item.subject.getAsync === 'function') {
        return new Promise<string>((resolve) => {
          mailbox.item.subject.getAsync((result: any) => {
            resolve(resolveAsyncResult(result, (value) => value || ''))
          })
        })
      }

      return mailbox.item.subject || ''
    },
  },

  setEmailSubject: {
    name: 'setEmailSubject',
    category: 'write',
    description: 'Set the email subject in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: {
          type: 'string',
          description: 'The new email subject line',
        },
      },
      required: ['subject'],
    },
    executeOutlook: async (mailbox, args) => {
      const { subject } = args
      if (!mailbox?.item?.subject?.setAsync) {
        return 'Cannot set email subject: compose mode is not available.'
      }

      return new Promise<string>((resolve) => {
        mailbox.item.subject.setAsync(subject, (result: any) => {
          resolve(resolveAsyncResult(result, () => 'Successfully updated email subject.'))
        })
      })
    },
  },

  getEmailRecipients: {
    name: 'getEmailRecipients',
    category: 'read',
    description: 'Get the current To, Cc, and Bcc recipients of the email.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'

      const item = mailbox.item

      const getField = (field: RecipientField) => {
        const fieldObject = item[field]
        if (fieldObject && typeof fieldObject.getAsync === 'function') {
          return new Promise<any[]>((resolve) => {
            fieldObject.getAsync((result: any) => {
              if (result.status === getOfficeAsyncStatus()?.Succeeded && Array.isArray(result.value)) {
                resolve(result.value.map(normalizeRecipient))
              } else {
                resolve([])
              }
            })
          })
        }

        if (Array.isArray(fieldObject)) {
          return Promise.resolve(fieldObject.map(normalizeRecipient))
        }

        return Promise.resolve([])
      }

      const [to, cc, bcc] = await Promise.all([
        getField('to'),
        getField('cc'),
        getField('bcc'),
      ])

      return JSON.stringify({ to, cc, bcc })
    },
  },

  addRecipient: {
    name: 'addRecipient',
    category: 'write',
    description: 'Add recipient(s) to To, Cc, or Bcc in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {
        field: {
          type: 'string',
          enum: ['to', 'cc', 'bcc'],
          description: 'Recipient field to update (to, cc, bcc). Defaults to to.',
        },
        recipients: {
          type: 'string',
          description:
            'Recipient(s) to add. Provide a single email address or a comma-separated list of emails.',
        },
      },
      required: ['recipients'],
    },
    executeOutlook: async (mailbox, args) => {
      if (!mailbox?.item) return 'No email item available.'

      const field = getRecipientField(args.field)
      const recipients = normalizeRecipientsInput(args.recipients)
      if (recipients.length === 0) {
        return 'No valid recipients provided.'
      }

      const fieldObject = mailbox.item[field]
      if (!fieldObject || typeof fieldObject.addAsync !== 'function') {
        return `Cannot add recipients to ${field}: compose mode is not available.`
      }

      return new Promise<string>((resolve) => {
        fieldObject.addAsync(recipients, (result: any) => {
          resolve(resolveAsyncResult(result, () => `Successfully added ${recipients.length} recipient(s) to ${field}.`))
        })
      })
    },
  },

  getEmailSender: {
    name: 'getEmailSender',
    category: 'read',
    description: 'Get sender information for the current email.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'

      const sender = mailbox.item.from || mailbox.item.sender
      if (!sender) return ''

      return JSON.stringify(normalizeRecipient(sender))
    },
  },

  getEmailDate: {
    name: 'getEmailDate',
    category: 'read',
    description: 'Get creation date/time for the current email item (read mode).',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'

      const value = mailbox.item.dateTimeCreated
      if (!value) {
        return 'Email creation date is not available in this context.'
      }

      const date = value instanceof Date ? value : new Date(value)
      return Number.isNaN(date.getTime()) ? String(value) : date.toISOString()
    },
  },

  getAttachments: {
    name: 'getAttachments',
    category: 'read',
    description: 'List attachments of the current email.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeOutlook: async (mailbox) => {
      if (!mailbox?.item) return 'No email item available.'

      if (typeof mailbox.item.getAttachmentsAsync === 'function') {
        return new Promise<string>((resolve) => {
          mailbox.item.getAttachmentsAsync((result: any) => {
            if (result.status === getOfficeAsyncStatus()?.Succeeded && Array.isArray(result.value)) {
              resolve(JSON.stringify(result.value))
            } else {
              resolve(`Error listing attachments: ${result.error?.message || 'unknown error'}`)
            }
          })
        })
      }

      if (Array.isArray(mailbox.item.attachments)) {
        return JSON.stringify(mailbox.item.attachments)
      }

      return 'Attachments are not available in this context.'
    },
  },

  insertHtmlAtCursor: {
    name: 'insertHtmlAtCursor',
    category: 'write',
    description: 'Insert raw HTML at the cursor position in the email body (compose mode). Only use this for pre-formatted HTML content. For normal text or Markdown content, prefer insertTextAtCursor.',
    inputSchema: {
      type: 'object',
      properties: {
        html: {
          type: 'string',
          description: 'The HTML content to insert at the cursor position',
        },
      },
      required: ['html'],
    },
    executeOutlook: async (mailbox, args) => {
      const { html } = args
      if (!mailbox?.item?.body?.setSelectedDataAsync) {
        return 'Cannot insert HTML at cursor: compose mode is not available.'
      }

      return new Promise<string>((resolve) => {
        mailbox.item.body.setSelectedDataAsync(
          DOMPurify.sanitize(html),
          { coercionType: getOfficeCoercionType().Html },
          (result: any) => {
            resolve(resolveAsyncResult(result, () => 'Successfully inserted HTML at cursor.'))
          },
        )
      })
    },
  },

  eval_outlookjs: {
    name: 'eval_outlookjs',
    category: 'write',
    description: "Execute arbitrary Office.js code within the Outlook mailbox context. Use this as an escape hatch when existing tools don't cover your use case. The code runs inside an async context with `mailbox` available, representing `Office.context.mailbox`. Return a value to get it back as the result.",
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: "JavaScript code to execute. Has access to `mailbox` (Office.context.mailbox). Must be valid async code. Return a value to get it as result. Example: `return new Promise((resolve) => { mailbox.item.subject.getAsync((res) => resolve(res.value)); })`",
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does',
        },
      },
      required: ['code'],
    },
    executeOutlook: async (mailbox, args) => {
      const { code } = args
      try {
        const result = await sandboxedEval(code, { mailbox, Office: typeof (window as any).Office !== 'undefined' ? (window as any).Office : undefined })
        return JSON.stringify({ success: true, result: result ?? null }, null, 2)
      } catch (err: any) {
        return JSON.stringify({ success: false, error: err.message || String(err) }, null, 2)
      }
    },
  },
})

export function getOutlookToolDefinitions(): OutlookToolDefinition[] {
  return Object.values(outlookToolDefinitions)
}

export { outlookToolDefinitions }
