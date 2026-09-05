// AI 数据持久化（Dexie）
import Dexie, { type Table } from 'dexie'
import type { AISettings, Conversation, Template } from './types'

export class AIDb extends Dexie {
  conversations!: Table<Conversation, string>
  settings!: Table<AISettings & { key: string }, string>
  templates!: Table<Template, string>

  constructor() {
    super('eflink-pptx-ai')
    this.version(1).stores({
      conversations: 'id, updatedAt',
      settings: 'key',
      templates: 'id, builtin',
    })
  }
}

export const aiDb = new AIDb()

const SETTINGS_KEY = 'main'
const DEFAULT_SETTINGS: AISettings = { baseUrl: '', apiKey: '', model: '' }

export const aiSettingsStore = {
  async getSettings(): Promise<AISettings> {
    const row = await aiDb.settings.get(SETTINGS_KEY)
    if (!row) return { ...DEFAULT_SETTINGS }
    return { baseUrl: row.baseUrl, apiKey: row.apiKey, model: row.model }
  },

  async saveSettings(s: AISettings): Promise<void> {
    await aiDb.settings.put({ key: SETTINGS_KEY, ...s })
  },

  async isConfigured(): Promise<boolean> {
    const s = await this.getSettings()
    return Boolean(s.baseUrl && s.apiKey && s.model)
  },
}

export const BUILTIN_TEMPLATES: Template[] = [
  {
    id: 'tpl-outline', name: '生成演示大纲', builtin: true,
    content: '请围绕主题「{topic}」生成一份演示文稿大纲，JSON 格式：{"topic":"主题","slides":[{"title":"页标题","points":["要点1","要点2"]}]}，共 {count} 页，第一页为封面（标题+副标题要点），最后一页为总结。只输出 JSON。',
  },
  {
    id: 'tpl-polish', name: '润色文本', builtin: true,
    content: '请润色以下文本，使其更精炼有力，保持原意，直接输出润色后的文本，不要解释：\n{text}',
  },
  {
    id: 'tpl-expand', name: '扩写文本', builtin: true,
    content: '请扩写以下文本，补充细节与论据，直接输出扩写后的文本，不要解释：\n{text}',
  },
  {
    id: 'tpl-translate', name: '翻译为英文', builtin: true,
    content: '请将以下文本翻译为英文，直接输出译文，不要解释：\n{text}',
  },
]
