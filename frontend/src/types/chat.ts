import type { Component } from 'vue'

export interface ToolCallPart {
  id: string
  name: string
  args: Record<string, any>
  status: 'pending' | 'running' | 'complete' | 'error'
  result?: string
  screenshotSrc?: string
}

export interface DisplayMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  imageSrc?: string
  richHtml?: string
  toolCalls?: ToolCallPart[]
  rawMessages?: any[]
  timestamp?: number
  attachedFiles?: Array<{ filename: string; content: string; fileId?: string }>
}

export interface QuickAction {
  key: string
  label: string
  icon: Component
  executeWithAgent?: boolean
  tooltipKey?: string
}

export interface ExcelQuickAction extends QuickAction {
  mode: 'immediate' | 'draft'
  prefix?: string
  systemPrompt?: string
}

export interface PowerPointQuickAction extends QuickAction {
  mode: 'immediate' | 'draft'
  systemPrompt?: string
}

export interface OutlookQuickAction extends QuickAction {
  mode?: 'immediate' | 'draft' | 'smart-reply'
  prefix?: string
}

export interface RenderSegment {
  type: 'text' | 'think'
  text: string
}
