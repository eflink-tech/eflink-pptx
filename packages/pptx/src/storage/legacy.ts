/**
 * DocuLuma 品牌迁移：把 efLink 时期的本地数据一次性搬到 doculuma 命名空间。
 * - localStorage：真改名（新键已有值时只丢弃旧值），迁移后旧键消失。
 * - IndexedDB：只读旧库并导入新库，**永不删除旧库**；仅当新表为空时导入，避免覆盖本次会话的数据。
 * 所有失败都静默吞掉：迁移绝不能阻塞编辑器启动或读写。
 */
import { migrateLocalStorageKey, readLegacyIndexedDb } from '@eflink-tech/office-kit'
import {
  LAST_DOC_KEY,
  LEGACY_AI_DB_NAME,
  LEGACY_FILES_DB_NAME,
  LEGACY_LAST_DOC_KEY,
  LEGACY_MIRROR_KEY,
  MIRROR_KEY,
} from './keys'
import { pptxDb, type PPTDocRecord } from '../files/db'
import { aiDb } from '../ai/db'
import type { AISettings, Conversation, Template } from '../ai/types'

/** 一次性把旧品牌的 localStorage 键改名到 doculuma 命名空间（模块级调用一次即可）。 */
export function migrateLegacyLocalStorage(): void {
  migrateLocalStorageKey(LEGACY_LAST_DOC_KEY, LAST_DOC_KEY)
  migrateLocalStorageKey(LEGACY_MIRROR_KEY, MIRROR_KEY)
}

/** 文档库：eflink-pptx-files → doculuma-pptx-files（旧库保留，只在新库为空时导入）。 */
async function importLegacyFilesDb(): Promise<void> {
  const rows = await readLegacyIndexedDb(LEGACY_FILES_DB_NAME, ['documents'])
  const legacy = rows?.documents as PPTDocRecord[] | undefined
  if (!legacy?.length) return
  if (await pptxDb.documents.count() > 0) return
  await pptxDb.documents.bulkPut(legacy)
}

/** AI 库：eflink-pptx-ai → doculuma-pptx-ai（旧库保留，只在新库为空时导入）。 */
async function importLegacyAiDb(): Promise<void> {
  const rows = await readLegacyIndexedDb(LEGACY_AI_DB_NAME, ['conversations', 'settings', 'templates'])
  if (!rows) return

  const conversations = rows.conversations as Conversation[] | undefined
  if (conversations?.length && (await aiDb.conversations.count()) === 0) {
    await aiDb.conversations.bulkPut(conversations)
  }
  const settings = rows.settings as Array<AISettings & { key: string }> | undefined
  if (settings?.length && (await aiDb.settings.count()) === 0) {
    await aiDb.settings.bulkPut(settings)
  }
  const templates = rows.templates as Template[] | undefined
  if (templates?.length && (await aiDb.templates.count()) === 0) {
    await aiDb.templates.bulkPut(templates)
  }
}

let filesImport: Promise<void> | null = null
let aiImport: Promise<void> | null = null

/** 惰性导入旧文档库（结果记忆化，失败不抛）。Dexie 读写路径起点 await 它。 */
export function ensureLegacyImport(): Promise<void> {
  filesImport ??= importLegacyFilesDb().catch(() => {
    /* 迁移失败：新库照常工作，旧库数据保留，用户可稍后重试 */
  })
  return filesImport
}

/** 惰性导入旧 AI 库（结果记忆化，失败不抛）。AI 设置/会话的首次读取前 await 它。 */
export function ensureAiLegacyImport(): Promise<void> {
  aiImport ??= importLegacyAiDb().catch(() => {
    /* 同上：失败不阻塞 AI 面板 */
  })
  return aiImport
}
