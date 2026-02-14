import { evaluate } from 'mathjs'

export type GeneralToolName = 'getCurrentDate' | 'calculateMath'

export interface GeneralToolDefinition {
  name: GeneralToolName
  description: string
  inputSchema: ToolInputSchema
  execute: (args: Record<string, any>) => Promise<string>
}

const generalToolDefinitions: GeneralToolDefinition[] = [
  {
    name: 'getCurrentDate',
    description:
      'Returns the current date and time. Useful for adding timestamps, dates to documents, or understanding temporal context.',
    inputSchema: {
      type: 'object',
      properties: {
        format: {
          type: 'string',
          description: 'Format: "full" (date and time), "date" (date only), "time" (time only), "iso" (ISO 8601)',
          enum: ['full', 'date', 'time', 'iso'],
        },
      },
      required: [],
    },
    execute: async (args) => {
      const { format = 'full' } = args
      const now = new Date()

      switch (format) {
        case 'date':
          return now.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
          })
        case 'time':
          return now.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
          })
        case 'iso':
          return now.toISOString()
        case 'full':
        default:
          return now.toLocaleString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
          })
      }
    },
  },
  {
    name: 'calculateMath',
    description:
      'Evaluates mathematical expressions safely. Supports basic arithmetic (+, -, *, /), parentheses, and common math functions.',
    inputSchema: {
      type: 'object',
      properties: {
        expression: {
          type: 'string',
          description: 'The mathematical expression to evaluate (e.g., "2 + 2 * 3")',
        },
      },
      required: ['expression'],
    },
    execute: async (args) => {
      const { expression } = args
      try {
        const result = evaluate(expression)

        if (typeof result !== 'number' && typeof result !== 'bigint') {
          return `Calculation completed, but result is not a simple number: ${result}`
        }

        return `${expression} = ${result}`
      } catch (error: any) {
        return `Error evaluating expression: ${error.message}`
      }
    },
  },
]

export function getGeneralToolDefinitions(): GeneralToolDefinition[] {
  return generalToolDefinitions
}

export function getEnabledGeneralTools(enabledTools?: GeneralToolName[]): GeneralToolDefinition[] {
  if (!enabledTools || enabledTools.length === 0) {
    return generalToolDefinitions
  }
  return generalToolDefinitions.filter(def => enabledTools.includes(def.name))
}
