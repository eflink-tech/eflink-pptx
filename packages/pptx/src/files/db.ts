// 本地多文档管理（Dexie / IndexedDB）
import Dexie, { type Table } from 'dexie'
import type { Presentation } from '../types/slides'

export interface PPTDocRecord {
  id: string
  name: string
  presentation: Presentation
  createdAt: number
  updatedAt: number
}

export class PptxDb extends Dexie {
  documents!: Table<PPTDocRecord, string>

  constructor() {
    super('eflink-pptx-files')
    this.version(1).stores({
      documents: 'id, updatedAt',
    })
  }
}

export const pptxDb = new PptxDb()

/** 仅测试用：删除并重建数据库 */
export async function resetDb(): Promise<void> {
  await pptxDb.delete()
  await pptxDb.open()
}
