// 持久化：Dexie 多文档 + localStorage 崩溃恢复镜像 + 自动保存
import type { Presentation } from '../../types/slides'
import { createPresentation } from '../../types/slides'
import { pptxDb, type PPTDocRecord } from '../../files/db'
import { genId } from '../utils/id'

const LAST_DOC_KEY = 'eflink-pptx-last-doc'
const MIRROR_KEY = 'eflink-pptx-mirror'

export interface LoadedDoc {
  id: string
  name: string
  presentation: Presentation
}

/** localStorage 镜像（毫秒级写入，Dexie 节流写） */
function writeMirror(docId: string | undefined, name: string, presentation: Presentation): void {
  // docId 为 undefined 时不写入（JSON.stringify 会静默丢弃 undefined 字段，导致恢复时丢失）
  if (!docId) return
  try {
    localStorage.setItem(MIRROR_KEY, JSON.stringify({ id: docId, name, presentation }))
    localStorage.setItem(LAST_DOC_KEY, docId)
  } catch {
    /* 存储满等异常忽略 */
  }
}

/** 启动时载入文档：优先 localStorage 镜像（最近状态），否则 Dexie 该文档，否则新建 */
export async function loadStartupDoc(): Promise<LoadedDoc> {
  try {
    const mirror = localStorage.getItem(MIRROR_KEY)
    if (mirror) {
      const parsed = JSON.parse(mirror) as LoadedDoc
      if (parsed?.presentation?.slides?.length) {
        // mirror 可能缺少 id（旧版本 JSON.stringify 丢弃了 undefined），从 LAST_DOC_KEY 兜底
        if (!parsed.id || parsed.id === 'undefined') {
          const lastId = localStorage.getItem(LAST_DOC_KEY)
          if (lastId && lastId !== 'undefined') parsed.id = lastId
        }
        // 如果仍然没有有效 id，说明旧数据 id 丢失，但 presentation 数据完好
        // → 创建新 id 并回写 Dexie，保留用户数据
        if (!parsed.id || parsed.id === 'undefined') {
          parsed.id = genId('doc-')
          localStorage.setItem(MIRROR_KEY, JSON.stringify(parsed))
          localStorage.setItem(LAST_DOC_KEY, parsed.id)
          // 回写 Dexie
          const now = Date.now()
          try {
            await pptxDb.documents.put({ id: parsed.id, name: parsed.name, presentation: parsed.presentation, createdAt: now, updatedAt: now })
          } catch { /* Dexie 不可用时镜像仍生效 */ }
        }
        return parsed
      }
    }
  } catch { /* 忽略坏数据 */ }

  try {
    const lastId = localStorage.getItem(LAST_DOC_KEY)
    if (lastId) {
      const rec = await pptxDb.documents.get(lastId)
      if (rec) return { id: rec.id, name: rec.name, presentation: rec.presentation }
    }
  } catch { /* Dexie 不可用（隐私模式等） */ }

  return { id: genId('doc-'), name: '未命名演示文稿', presentation: createPresentation(genId('slide-')) }
}

/** 保存到 Dexie（不存在则创建），并同步镜像 */
export async function saveDoc(docId: string | undefined, name: string, presentation: Presentation): Promise<void> {
  if (!docId) return // docId 无效时跳过，防止写入损坏数据
  writeMirror(docId, name, presentation)
  const now = Date.now()
  try {
    const rec: PPTDocRecord = { id: docId, name, presentation, createdAt: now, updatedAt: now }
    const existing = await pptxDb.documents.get(docId)
    if (existing) rec.createdAt = existing.createdAt
    await pptxDb.documents.put(rec)
  } catch { /* Dexie 不可用时镜像仍生效 */ }
}

/** 自动保存节流（3s 防抖） */
let saveTimer: number | undefined
let pending: (() => void) | null = null

export function scheduleAutosave(fn: () => void): void {
  pending = fn
  if (saveTimer !== undefined) return
  saveTimer = window.setTimeout(() => {
    saveTimer = undefined
    const task = pending
    pending = null
    task?.()
  }, 3000)
}

/** 文档列表（按更新时间倒序） */
export async function listDocs(): Promise<Array<Pick<PPTDocRecord, 'id' | 'name' | 'updatedAt'>>> {
  try {
    const rows = await pptxDb.documents.orderBy('updatedAt').reverse().toArray()
    return rows.map((r) => ({ id: r.id, name: r.name, updatedAt: r.updatedAt }))
  } catch {
    return []
  }
}

export async function createDoc(name: string): Promise<LoadedDoc> {
  return { id: genId('doc-'), name, presentation: createPresentation(genId('slide-')) }
}

export async function deleteDoc(docId: string): Promise<void> {
  try { await pptxDb.documents.delete(docId) } catch { /* 忽略 */ }
  if (localStorage.getItem(LAST_DOC_KEY) === docId) {
    localStorage.removeItem(LAST_DOC_KEY)
    localStorage.removeItem(MIRROR_KEY)
  }
}

export async function duplicateDoc(docId: string): Promise<LoadedDoc | null> {
  try {
    const rec = await pptxDb.documents.get(docId)
    if (!rec) return null
    const copy: LoadedDoc = {
      id: genId('doc-'),
      name: `${rec.name} 副本`,
      presentation: structuredClone(rec.presentation),
    }
    await saveDoc(copy.id, copy.name, copy.presentation)
    return copy
  } catch {
    return null
  }
}

export async function renameDoc(docId: string, name: string): Promise<void> {
  try {
    const rec = await pptxDb.documents.get(docId)
    if (rec) await pptxDb.documents.put({ ...rec, name, updatedAt: Date.now() })
  } catch { /* 忽略 */ }
}
