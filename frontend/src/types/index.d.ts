type IStringKeyMap = Record<string, any>

type insertTypes = 'replace' | 'append' | 'newLine' | 'NoAction'

type ModelTier = 'nano' | 'standard' | 'reasoning' | 'image'

interface ModelInfo {
  id: string
  label: string
  type: 'chat' | 'image'
}

interface ToolInputSchema {
  type: 'object'
  properties: Record<string, ToolProperty>
  required?: string[]
}

interface ToolProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object'
  description?: string
  enum?: string[]
  items?: ToolProperty
  default?: any
}

interface WordToolDefinition {
  name: string
  description: string
  inputSchema: ToolInputSchema
  execute: (args: Record<string, any>) => Promise<string>
}

// Excel tools reuse the same interface structure
type ExcelToolDefinition = WordToolDefinition

type OfficeHostType = 'Word' | 'Excel' | 'Unknown'
