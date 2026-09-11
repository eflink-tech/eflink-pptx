// 持久化：Dexie 多文档 + localStorage 崩溃恢复镜像（本地草稿） + 手动云端保存（⌘S/Ctrl+S）
// 保存策略：内容变化立即写 localStorage 镜像（防崩溃丢稿，不清除 dirty）；
// 远端/云端保存仅由手动触发，成功后才置 dirty=false。
import type { Presentation } from '../../types/slides'
import { createPresentation } from '../../types/slides'
import { pptxDb, type PPTDocRecord } from '../../files/db'
import { genId } from '../utils/id'

/**
 * 可切换存储后端：默认 Dexie(IndexedDB)；宿主注入后所有落库走宿主实现（如后端 API）。
 * 注入后 localStorage 镜像自动停用，避免本地旧数据遮蔽远端数据。
 */
export interface PptxStorageBackend {
  /** 不存在则创建，存在则整体覆盖 */
  put(rec: PPTDocRecord): Promise<void>
  get(id: string): Promise<PPTDocRecord | undefined>
  remove(id: string): Promise<void>
  /** 按更新时间倒序的文档元信息（不含正文） */
  list(): Promise<Array<Pick<PPTDocRecord, 'id' | 'name' | 'updatedAt'>>>
}

let backendOverride: PptxStorageBackend | null = null

/** 注册自定义存储后端（宿主在挂载编辑器前调用） */
export function setPptxStorageBackend(backend: PptxStorageBackend | null): void {
  backendOverride = backend
}

function hasBackend(): boolean {
  return backendOverride !== null
}

async function dbPut(rec: PPTDocRecord): Promise<void> {
  if (backendOverride) return backendOverride.put(rec)
  await pptxDb.documents.put(rec)
}

async function dbGet(id: string): Promise<PPTDocRecord | undefined> {
  if (backendOverride) return backendOverride.get(id)
  return pptxDb.documents.get(id)
}

async function dbRemove(id: string): Promise<void> {
  if (backendOverride) return backendOverride.remove(id)
  await pptxDb.documents.delete(id)
}

async function dbList(): Promise<Array<Pick<PPTDocRecord, 'id' | 'name' | 'updatedAt'>>> {
  if (backendOverride) return backendOverride.list()
  return pptxDb.documents.orderBy('updatedAt').reverse().toArray()
}

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
    // 自定义后端模式下仅维护"当前文档 id"指针，停用内容镜像（远端是单一数据源）
    localStorage.setItem(LAST_DOC_KEY, docId)
    if (hasBackend()) return
    localStorage.setItem(MIRROR_KEY, JSON.stringify({ id: docId, name, presentation }))
  } catch {
    /* 存储满等异常忽略 */
  }
}

/** 启动时载入文档。加载优先级：
 * - 登录态（自定义后端）：远端优先 —— 仅当远端加载失败或远端无此文档时，才用 localStorage 镜像恢复本地草稿，
 *   严禁用旧镜像覆盖远端已保存内容（幽灵修改防护）。
 * - 本地模式（无后端）：优先镜像（最近状态），否则 Dexie 该文档，否则新建。
 */
export async function loadStartupDoc(bootDocId?: string): Promise<LoadedDoc> {
  // 指定启动文档（如分享查看页的只读快照）：加载失败不回退访客本地 last-doc，避免串文档
  if (bootDocId) {
    try {
      const rec = await dbGet(bootDocId)
      if (rec) return { id: rec.id, name: rec.name, presentation: rec.presentation }
    } catch { /* 忽略，走新建兜底 */ }
    return { id: genId('doc-'), name: '未命名演示文稿', presentation: createPresentation(genId('slide-')) }
  }

  // 登录态（自定义后端）：远端优先，镜像仅作草稿兜底
  if (hasBackend()) {
    let lastId: string | null = null
    try { lastId = localStorage.getItem(LAST_DOC_KEY) } catch { /* 忽略 */ }
    if (lastId) {
      try {
        const rec = await dbGet(lastId)
        // 远端有此文档：以远端为准（镜像里的未保存草稿不覆盖远端内容）
        if (rec) return { id: rec.id, name: rec.name, presentation: rec.presentation }
      } catch { /* 远端加载失败 → 回退镜像草稿 */ }
    }
    // 远端无此文档或加载失败：镜像作为本地草稿恢复（无有效 id 的坏镜像直接丢弃，不回写远端）
    try {
      const mirror = localStorage.getItem(MIRROR_KEY)
      if (mirror) {
        const parsed = JSON.parse(mirror) as LoadedDoc
        if (parsed?.presentation?.slides?.length) {
          if (!parsed.id || parsed.id === 'undefined') {
            if (lastId && lastId !== 'undefined') parsed.id = lastId
          }
          if (parsed.id && parsed.id !== 'undefined') return parsed
        }
      }
    } catch { /* 忽略坏数据 */ }
    return { id: genId('doc-'), name: '未命名演示文稿', presentation: createPresentation(genId('slide-')) }
  }

  // 本地模式：优先镜像（最近状态），否则 Dexie 该文档，否则新建
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
            await dbPut({ id: parsed.id, name: parsed.name, presentation: parsed.presentation, createdAt: now, updatedAt: now })
          } catch { /* Dexie 不可用时镜像仍生效 */ }
        }
        return parsed
      }
    }
  } catch { /* 忽略坏数据 */ }

  try {
    const lastId = localStorage.getItem(LAST_DOC_KEY)
    if (lastId) {
      const rec = await dbGet(lastId)
      if (rec) return { id: rec.id, name: rec.name, presentation: rec.presentation }
    }
  } catch { /* Dexie 不可用（隐私模式等） */ }

  return { id: genId('doc-'), name: '未命名演示文稿', presentation: createPresentation(genId('slide-')) }
}

/** 云端保存入口（不存在则创建），并同步镜像。仅由手动保存触发（⌘S/Ctrl+S、保存按钮、分享前、window bridge） */
export async function saveDoc(docId: string | undefined, name: string, presentation: Presentation): Promise<void> {
  if (!docId) return // docId 无效时跳过，防止写入损坏数据
  writeMirror(docId, name, presentation)
  const now = Date.now()
  try {
    const rec: PPTDocRecord = { id: docId, name, presentation, createdAt: now, updatedAt: now }
    const existing = await dbGet(docId)
    if (existing) rec.createdAt = existing.createdAt
    await dbPut(rec)
  } catch (err) {
    // 本地模式：镜像已写入，容忍 Dexie 失败；登录态（后端）：向上抛出，
    // 让调用方提示"保存失败"，dirty 保持 true（严禁把失败当成已保存）
    if (hasBackend()) throw err
  }
}

/** 文档列表（按更新时间倒序） */
export async function listDocs(): Promise<Array<Pick<PPTDocRecord, 'id' | 'name' | 'updatedAt'>>> {
  try {
    const rows = await dbList()
    return rows.map((r) => ({ id: r.id, name: r.name, updatedAt: r.updatedAt }))
  } catch {
    return []
  }
}

export async function createDoc(name: string): Promise<LoadedDoc> {
  return { id: genId('doc-'), name, presentation: createPresentation(genId('slide-')) }
}

export async function deleteDoc(docId: string): Promise<void> {
  try { await dbRemove(docId) } catch { /* 忽略 */ }
  if (localStorage.getItem(LAST_DOC_KEY) === docId) {
    localStorage.removeItem(LAST_DOC_KEY)
    localStorage.removeItem(MIRROR_KEY)
  }
}

export async function duplicateDoc(docId: string): Promise<LoadedDoc | null> {
  try {
    const rec = await dbGet(docId)
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

/** 重命名：仅更新文档记录 name 字段（正文不动，不影响内容 dirty）。
 * 记录尚不存在（新文档未保存过）且传入 presentation 时，退化为整档落库，避免改名刷新后丢失。
 * 本地模式尽力而为；登录态（后端）失败时向上抛出，由调用方提示"重命名失败"（名字保持原值）。 */
export async function renameDoc(docId: string, name: string, presentation?: Presentation): Promise<void> {
  try {
    const rec = await dbGet(docId)
    if (rec) {
      await dbPut({ ...rec, name, updatedAt: Date.now() })
      return
    }
    if (presentation) {
      const now = Date.now()
      await dbPut({ id: docId, name, presentation, createdAt: now, updatedAt: now })
    }
  } catch (err) {
    if (hasBackend()) throw err
  }
}
