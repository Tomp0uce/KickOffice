// markdown-it-task-lists ships without TypeScript declarations
declare module 'markdown-it-task-lists'
declare module 'markdown-it-deflist'
declare module 'markdown-it-footnote'
declare module 'turndown' {
  interface Options {
    headingStyle?: 'setext' | 'atx'
    bulletListMarker?: '-' | '+' | '*'
    codeBlockStyle?: 'indented' | 'fenced'
    fence?: string
  }
  interface Rule {
    filter: string | string[] | ((node: any) => boolean)
    replacement: (content: string, node: any, options: Options) => string
  }
  class TurndownService {
    constructor(options?: Options)
    turndown(html: string): string
    addRule(key: string, rule: Rule): this
    use(plugin: (service: TurndownService) => void): this
  }
  export = TurndownService
}
declare module 'diff-match-patch' {
  class diff_match_patch {
    diff_main(text1: string, text2: string): [number, string][]
    diff_cleanupSemantic(diffs: [number, string][]): void
  }
  export = diff_match_patch
}

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
