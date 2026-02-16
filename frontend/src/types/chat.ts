import type { Component } from 'vue'

export interface DisplayMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  imageSrc?: string
}

export interface QuickAction {
  key: string
  label: string
  icon: Component
}

export interface ExcelQuickAction extends QuickAction {
  mode: 'immediate' | 'draft'
  prefix?: string
  systemPrompt?: string
}

export interface PowerPointQuickAction extends QuickAction {
  mode: 'immediate' | 'draft'
}

export interface RenderSegment {
  type: 'text' | 'think'
  text: string
}
