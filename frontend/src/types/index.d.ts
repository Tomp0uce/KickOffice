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

interface ToolDefinition {
  name: string
  description: string
  inputSchema: ToolInputSchema
  execute: (args: Record<string, any>) => Promise<string>
}

type WordToolDefinition = ToolDefinition
type ExcelToolDefinition = ToolDefinition
type OutlookToolDefinition = ToolDefinition

type OfficeHostType = 'Word' | 'Excel' | 'Unknown'
