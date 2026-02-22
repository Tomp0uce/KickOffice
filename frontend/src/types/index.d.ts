// markdown-it-task-lists ships without TypeScript declarations
declare module 'markdown-it-task-lists'

type IStringKeyMap = Record<string, any>

type insertTypes = 'replace' | 'append' | 'newLine' | 'NoAction'

type ModelTier = 'standard' | 'reasoning' | 'image'

interface ModelInfo {
  id: string
  label: string
  type: 'chat' | 'image'
}

interface ToolInputSchema {
  type: 'object'
  properties: Record<string, ToolProperty>
  required?: string[]
  [key: string]: unknown
}

interface ToolProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object'
  description?: string
  enum?: string[]
  items?: ToolProperty
  default?: any
}

type ToolCategory = 'read' | 'write' | 'format'

interface ToolDefinition {
  name: string
  category: ToolCategory
  description: string
  inputSchema: ToolInputSchema
  execute: (args: Record<string, any>) => Promise<string>
}

type WordToolDefinition = ToolDefinition
type ExcelToolDefinition = ToolDefinition
type PowerPointToolDefinition = ToolDefinition
type OutlookToolDefinition = ToolDefinition

type OfficeHostType = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook' | 'Unknown'
