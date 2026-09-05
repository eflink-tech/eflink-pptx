// AI 模块类型定义（照 process-on 模式）
export interface AISettings {
  baseUrl: string
  apiKey: string
  model: string
}

export interface ConversationMessage {
  role: 'user' | 'assistant'
  content: string
  /** 关联产物（如已应用的大纲/页面 JSON） */
  payload?: unknown
  time: number
}

export interface Conversation {
  id: string
  title: string
  messages: ConversationMessage[]
  createdAt: number
  updatedAt: number
}

export interface Template {
  id: string
  name: string
  builtin: boolean
  /** 提示词模板：{topic} 等占位符 */
  content: string
}

export class AIError extends Error {
  readonly cause?: unknown

  constructor(message: string, cause?: unknown) {
    super(message)
    this.name = 'AIError'
    this.cause = cause
  }
}

/** AI 大纲结构 */
export interface OutlineItem {
  title: string
  points: string[]
}

export interface OutlineResult {
  topic: string
  slides: OutlineItem[]
}

/** AI 单页生成结构（内部 schema 子集） */
export interface AIGenElement {
  type: 'text' | 'shape'
  role?: 'title' | 'body' | 'accent'
  x: number
  y: number
  w: number
  h: number
  text: string
  fontSize?: number
  fill?: string
}

export interface AIGenSlide {
  title: string
  elements: AIGenElement[]
}

/** 判断错误是否为用户取消 */
export function isAbortError(error: unknown): boolean {
  if (error instanceof DOMException && error.name === 'AbortError') return true
  const message = error instanceof Error ? error.message : ''
  return /aborted/i.test(message)
}
