// 本地多文档存储（Dexie / IndexedDB）
import Dexie, { type Table } from 'dexie'
import type { Presentation } from '../types/slides'

export interface PPTDocRecord {
  id: string
  name: string
  presentation: Presentation
  createdAt: number
  updatedAt: number
}

/** 未保存草稿（登录态/大文档同样落本地，刷新后优先恢复） */
export interface PPTDraftRecord {
  id: string
  name: string
  presentation: Presentation
  savedAt: number
}

export class PptxDb extends Dexie {
  documents!: Table<PPTDocRecord, string>
  drafts!: Table<PPTDraftRecord, string>

  constructor() {
    super('doculuma-pptx-files')
    this.version(1).stores({
      documents: 'id, updatedAt',
    })
    // v2：草稿表（与 documents 分离——草稿是"未保存"状态，保存成功后删除）
    this.version(2).stores({
      drafts: 'id',
    })
  }
}

export const pptxDb = new PptxDb()

/** 仅测试用：删除并重建数据库 */
export async function resetDb(): Promise<void> {
  await pptxDb.delete()
  await pptxDb.open()
}
