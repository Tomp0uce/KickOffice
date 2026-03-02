import type { WordToolDefinition, OutlookToolDefinition } from '@/types'
import DiffMatchPatch from 'diff-match-patch'

import DOMPurify from 'dompurify'

import { executeOfficeAction } from './officeAction'
import { renderOfficeRichHtml } from './markdown'
import { sandboxedEval } from './sandbox'

import { generateVisualDiff } from './common'

export type OutlookToolName =
  | 'getEmailBody'
  | 'writeEmailBody'
  | 'getEmailSubject'
  | 'setEmailSubject'
  | 'getEmailRecipients'
  | 'addRecipient'
  | 'getEmailSender'
  | 'eval_outlookjs'



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
  throw new Error(`Error: ${result.error?.message || 'unknown error'}`)
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

  writeEmailBody: {
    name: 'writeEmailBody',
    category: 'write',
    description: 'The PREFERRED tool for modifying the email body. Supports Markdown (bold, italic, lists). Can replace the whole body, append to the end, or insert at the cursor. Only works in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'The content to write in Markdown format.',
        },
        mode: {
          type: 'string',
          description: 'How to write the content: "Replace" (full overwrite), "Append" (add to end), or "Insert" (add at cursor). Default: "Append".',
          enum: ['Replace', 'Append', 'Insert'],
        },
        diffTracking: {
          type: 'boolean',
          description: 'When mode is "Insert", shows a visual diff comparing content with current selection. Requires originalText. Default: false.',
        },
        originalText: {
          type: 'string',
          description: 'Required if diffTracking is true: the original text to compare against.',
        },
      },
      required: ['content'],
    },
    executeOutlook: async (mailbox, args: Record<string, any>) => {
      const { content, mode = 'Append', diffTracking = false, originalText = '' } = args
      if (!mailbox?.item?.body) return 'Cannot write email body: compose mode is not available.'

      const html = diffTracking && mode === 'Insert' && originalText
        ? generateVisualDiff(originalText, content)
        : renderOfficeRichHtml(content)

      return new Promise<string>((resolve) => {
        const body = mailbox.item.body
        
        if (mode === 'Replace') {
          body.setAsync(html, { coercionType: getOfficeCoercionType().Html }, (res: any) => {
            resolve(resolveAsyncResult(res, () => 'Successfully replaced email body.'))
          })
        } else if (mode === 'Append') {
          body.getAsync(getOfficeCoercionType().Html, {}, (getResult: any) => {
            if (getResult.status !== getOfficeAsyncStatus()?.Succeeded) {
              resolve('Error: Could not read body to append.')
              return
            }
            const existing = getResult.value || ''
            const separator = existing.trim() ? '<br/><br/>' : ''
            const newBody = existing + separator + DOMPurify.sanitize(html)
            body.setAsync(newBody, { coercionType: getOfficeCoercionType().Html }, (setResult: any) => {
              resolve(resolveAsyncResult(setResult, () => 'Successfully appended to email body.'))
            })
          })
        } else {
          // Insert at cursor
          body.setSelectedDataAsync(html, { coercionType: getOfficeCoercionType().Html }, (res: any) => {
            resolve(resolveAsyncResult(res, () => diffTracking ? 'Inserted visual diff.' : 'Successfully inserted at cursor.'))
          })
        }
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
    executeOutlook: async (mailbox, args: Record<string, any>) => {
      const { subject } = args as Record<string, any>
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
    executeOutlook: async (mailbox, args: Record<string, any>) => {
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
    executeOutlook: async (mailbox, args: Record<string, any>) => {
      const { code } = args as Record<string, any>
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
