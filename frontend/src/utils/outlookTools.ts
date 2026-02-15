export type OutlookToolName =
  | 'getEmailBody'
  | 'getSelectedText'
  | 'setEmailBody'

type OutlookToolDefinition = WordToolDefinition

function getMailbox(): any | null {
  return (window as any).Office?.context?.mailbox ?? null
}

const outlookToolDefinitions: Record<OutlookToolName, OutlookToolDefinition> = {
  getEmailBody: {
    name: 'getEmailBody',
    description:
      'Get the full body text of the current email. Works in both read and compose mode.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      const mailbox = getMailbox()
      if (!mailbox?.item) return 'No email item available.'
      return new Promise<string>((resolve) => {
        mailbox.item.body.getAsync(
          (window as any).Office.CoercionType.Text,
          (result: any) => {
            if (result.status === (window as any).Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || '')
            } else {
              resolve(`Error reading email body: ${result.error?.message || 'unknown error'}`)
            }
          },
        )
      })
    },
  },

  getSelectedText: {
    name: 'getSelectedText',
    description:
      'Get the currently selected text in the email compose window. Returns empty string if nothing is selected or not in compose mode.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      const mailbox = getMailbox()
      if (!mailbox?.item) return ''
      if (typeof mailbox.item.getSelectedDataAsync !== 'function') {
        return 'Selection reading is not available in this context.'
      }
      return new Promise<string>((resolve) => {
        mailbox.item.getSelectedDataAsync(
          (window as any).Office.CoercionType.Text,
          (result: any) => {
            if (result.status === (window as any).Office.AsyncResultStatus.Succeeded && result.value?.data) {
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
    execute: async (args) => {
      const { text } = args
      const mailbox = getMailbox()
      if (!mailbox?.item?.body?.setAsync) {
        return 'Cannot set email body: compose mode is not available.'
      }
      return new Promise<string>((resolve) => {
        mailbox.item.body.setAsync(
          text,
          { coercionType: (window as any).Office.CoercionType.Text },
          (result: any) => {
            if (result.status === (window as any).Office.AsyncResultStatus.Succeeded) {
              resolve('Successfully set email body.')
            } else {
              resolve(`Error setting email body: ${result.error?.message || 'unknown error'}`)
            }
          },
        )
      })
    },
  },
}

export function getOutlookToolDefinitions(): OutlookToolDefinition[] {
  return Object.values(outlookToolDefinitions)
}

export function getOutlookTool(name: OutlookToolName): OutlookToolDefinition | undefined {
  return outlookToolDefinitions[name]
}

export { outlookToolDefinitions }
