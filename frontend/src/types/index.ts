export type IStringKeyMap = Record<string, any>;

export type InsertType = 'replace' | 'append' | 'newLine' | 'NoAction';

export type ModelTier = 'standard' | 'reasoning' | 'image';

export interface ModelInfo {
  id: string;
  label: string;
  type: 'chat' | 'image';
  contextWindow?: number;
}

export interface ToolInputSchema {
  type: 'object';
  properties: Record<string, ToolProperty>;
  required?: string[];
  [key: string]: unknown;
}

export interface ToolProperty {
  type?: 'string' | 'number' | 'boolean' | 'array' | 'object' | 'null';
  description?: string;
  enum?: string[];
  items?: ToolProperty;
  properties?: Record<string, ToolProperty>;
  required?: string[];
  default?: any;
  anyOf?: ToolProperty[];
}

export type ToolCategory = 'read' | 'write' | 'format';

export interface ToolDefinition {
  name: string;
  category: ToolCategory;
  description: string;
  inputSchema: ToolInputSchema;
  execute: (args: Record<string, any>) => Promise<string>;
}

export type OfficeHostType = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook' | 'Unknown';
