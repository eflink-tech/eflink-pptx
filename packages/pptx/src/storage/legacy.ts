/**
 * DocuLuma 品牌迁移：把 efLink 时期的本地数据一次性搬到 doculuma 命名空间。
 * - localStorage：真改名（新键已有值时只丢弃旧值），迁移后旧键消失。
 * - IndexedDB：只读旧库并导入新库，**永不删除旧库**；每表落一次性导入标记，
 *   导入成功或确认跳过后不再回填——否则用户删光数据后，旧数据会在下次启动"复活"。
 * 所有失败都静默吞掉：迁移绝不能阻塞编辑器启动或读写；失败不落标记，下次启动自动重试。
 */
import {
  isLegacyImportDone,
  markLegacyImportDone,
  migrateLocalStorageKey,
  readLegacyIndexedDb,
} from '@eflink-tech/office-kit'
import type { Table } from 'dexie'
import {
  AI_CONVERSATIONS_IMPORT_DONE_KEY,
  AI_SETTINGS_IMPORT_DONE_KEY,
  AI_TEMPLATES_IMPORT_DONE_KEY,
  FILES_IMPORT_DONE_KEY,
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

/** 文档库：eflink-pptx-files → doculuma-pptx-files（旧库保留，按表一次性导入）。 */
async function importLegacyFilesDb(): Promise<void> {
  const rows = await readLegacyIndexedDb(LEGACY_FILES_DB_NAME, ['documents'])
  if (!rows) return // 旧库不存在/读取失败：不落标记，下次启动重试
  await importTable(pptxDb.documents, rows.documents as PPTDocRecord[] | undefined, FILES_IMPORT_DONE_KEY)
}

/** AI 库：eflink-pptx-ai → doculuma-pptx-ai（旧库保留，按表一次性导入）。 */
async function importLegacyAiDb(): Promise<void> {
  const rows = await readLegacyIndexedDb(LEGACY_AI_DB_NAME, ['conversations', 'settings', 'templates'])
  if (!rows) return // 同上
  await Promise.all([
    importTable(aiDb.conversations, rows.conversations as Conversation[] | undefined, AI_CONVERSATIONS_IMPORT_DONE_KEY),
    importTable(aiDb.settings, rows.settings as Array<AISettings & { key: string }> | undefined, AI_SETTINGS_IMPORT_DONE_KEY),
    importTable(aiDb.templates, rows.templates as Template[] | undefined, AI_TEMPLATES_IMPORT_DONE_KEY),
  ])
}

/**
 * 单表一次性导入：目标表非空则跳过（以新库为准）。
 * 无论导入成功还是跳过都落一次性标记：否则用户删光数据后，
 * 下次启动空表判据再次命中，旧数据会"复活"且永远删不干净。
 * bulkPut 单事务原子写入，失败抛给调用方且不落标记，下个会话自然重试。
 */
async function importTable<T>(
  table: Table<T, string>,
  rows: T[] | undefined,
  markerKey: string,
): Promise<void> {
  if (isLegacyImportDone(markerKey)) return
  if (!rows || rows.length === 0) {
    markLegacyImportDone(markerKey)
    return
  }
  if ((await table.count()) > 0) {
    markLegacyImportDone(markerKey)
    return
  }
  await table.bulkPut(rows)
  markLegacyImportDone(markerKey)
}

let filesImport: Promise<void> | null = null
let aiImport: Promise<void> | null = null

/** 惰性导入旧文档库（结果记忆化，失败不抛）。Dexie 读写路径起点 await 它。 */
export function ensureLegacyImport(): Promise<void> {
  filesImport ??= importLegacyFilesDb().catch(() => {
    /* 迁移失败：新库照常工作，旧库数据保留；未落标记，下次启动重试 */
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
