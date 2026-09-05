// 编辑器主状态：演示文稿数据 + 选择 + 历史记录 + 剪贴板
import { create } from 'zustand'
import { immer } from 'zustand/middleware/immer'
import type {
  PPTElement, Presentation, Slide, Theme,
} from '../types/slides'
import { createPresentation, createSlide } from '../types/slides'
import { genId } from '../core/utils/id'
import { unionBounds } from '../core/utils/geometry'
import { normalizePresentation } from '../core/chart/migrate'

const HISTORY_LIMIT = 50

export interface ClipboardData {
  slideId: string
  elements: PPTElement[]
}

interface EditorState {
  /** 当前文档 id（本地文件管理） */
  docId: string
  docName: string
  presentation: Presentation
  /** 当前页索引 */
  slideIndex: number
  /** 当前页选中元素 id */
  selectedIds: string[]
  /** 富文本编辑中的元素 id */
  editingId: string | null
  /** 表格单元格就地编辑：格式 "tableElementId:r:c"，null 表示非编辑态 */
  editingCellId: string | null
  /** 历史快照栈（undo 用）与未来栈（redo 用） */
  history: Presentation[]
  future: Presentation[]
  /** 复制/剪切暂存（元素级） */
  clipboard: ClipboardData | null
  /** 复制的幻灯片（页面级剪贴板，缩略图 Ctrl+C/V 用） */
  slideClipboard: Slide | null
  /** 脏标记（自动保存用） */
  dirty: boolean

  /* 文档级 */
  loadDocument: (doc: { id: string; name: string; presentation: Presentation }) => void
  renameDocument: (name: string) => void
  replacePresentation: (p: Presentation) => void

  /* 页面级 */
  addSlide: (atIndex?: number, slide?: Slide) => string
  deleteSlides: (indexes: number[]) => void
  copySlide: (index: number) => void
  moveSlide: (from: number, to: number) => void
  updateSlide: (index: number, partial: Partial<Slide>) => void
  setSlideBackground: (partial: Partial<Slide['background']>) => void
  setSlideTransition: (partial: Partial<NonNullable<Slide['transition']>>) => void
  setSlideNote: (note: string) => void
  /** 调整画布比例（viewportRatio = 宽/高），元素位置保持不动 */
  setViewportRatio: (ratio: number) => void

  /* 主题 */
  updateTheme: (partial: Partial<Theme>) => void

  /* 元素级 */
  addElement: (el: PPTElement, opts?: { select?: boolean }) => void
  updateElements: (ids: string[], updater: (el: PPTElement) => void, opts?: { history?: boolean }) => void
  removeElements: (ids: string[]) => void
  duplicateElements: (ids: string[]) => string[]
  setElementLevel: (ids: string[], level: 'up' | 'down' | 'top' | 'bottom') => void
  setGroup: (ids: string[], groupId?: string) => void

  /* 选择 */
  setSelected: (ids: string[]) => void
  toggleSelected: (id: string) => void
  selectAll: () => void
  setEditingId: (id: string | null) => void
  setEditingCellId: (id: string | null) => void
  /** 更新表格单元格文本（id 格式 "tableId:r:c"） */
  updateTableCellText: (cellId: string, text: string) => void

  /* 剪贴板 */
  copySelection: () => void
  cutSelection: () => void
  pasteClipboard: (offset?: number) => void
  /** 复制当前整页幻灯片到页面剪贴板 */
  copyCurrentSlide: () => void
  /** 粘贴页面剪贴板中的幻灯片（插入到当前页之后） */
  pasteSlide: () => void

  /* 历史 */
  pushHistory: () => void
  undo: () => void
  redo: () => void
  markSaved: () => void
}

/** 深拷贝元素（粘贴/复制用） */
function cloneElements(elements: PPTElement[], newIds = true): PPTElement[] {
  const cloned = structuredClone(elements)
  if (!newIds) return cloned
  const idMap = new Map<string, string>()
  for (const el of cloned) idMap.set(el.id, genId())
  for (const el of cloned) {
    el.id = idMap.get(el.id) ?? el.id
    if (el.groupId) {
      if (!idMap.has(el.groupId)) idMap.set(el.groupId, genId('g-'))
      el.groupId = idMap.get(el.groupId)
    }
  }
  return cloned
}

export const useEditorStore = create<EditorState>()(
  immer((set, get) => ({
    docId: 'default',
    docName: '未命名演示文稿',
    presentation: createPresentation('slide-0'),
    slideIndex: 0,
    selectedIds: [],
    editingId: null,
    editingCellId: null,
    history: [],
    future: [],
    clipboard: null,
    slideClipboard: null,
    dirty: false,

    loadDocument: (doc) =>
      set((s) => {
        s.docId = doc.id
        s.docName = doc.name
        // 规范化图表元素：旧文档的历史类型/开关字段迁移为 ECharts 版结构
        s.presentation = normalizePresentation(doc.presentation)
        s.slideIndex = 0
        s.selectedIds = []
        s.editingId = null
        s.editingCellId = null
        s.history = []
        s.future = []
        s.dirty = false
      }),

    renameDocument: (name) =>
      set((s) => {
        s.docName = name
        s.dirty = true
      }),

    replacePresentation: (p) =>
      set((s) => {
        s.presentation = normalizePresentation(p)
        s.slideIndex = 0
        s.selectedIds = []
        s.editingId = null
        s.editingCellId = null
        s.dirty = true
      }),

    addSlide: (atIndex, slide) => {
      const id = slide?.id ?? genId('slide-')
      set((s) => {
        const idx = atIndex ?? s.slideIndex + 1
        const next = slide ?? createSlide(id)
        s.presentation.slides.splice(idx, 0, next)
        s.slideIndex = idx
        s.selectedIds = []
        s.dirty = true
      })
      return id
    },

    deleteSlides: (indexes) =>
      set((s) => {
        if (s.presentation.slides.length <= indexes.length) return
        const sorted = [...indexes].sort((a, b) => b - a)
        for (const i of sorted) s.presentation.slides.splice(i, 1)
        s.slideIndex = Math.min(s.slideIndex, s.presentation.slides.length - 1)
        s.selectedIds = []
        s.dirty = true
      }),

    copySlide: (index) => {
      // 快照必须在 set 外部获取：immer draft 是 Proxy，structuredClone 会抛 DataCloneError
      const src = structuredClone(get().presentation.slides[index])
      src.id = genId('slide-')
      for (const el of src.elements) el.id = genId()
      set((s) => {
        s.presentation.slides.splice(index + 1, 0, src)
        s.slideIndex = index + 1
        s.selectedIds = []
        s.dirty = true
      })
    },

    moveSlide: (from, to) =>
      set((s) => {
        const [slide] = s.presentation.slides.splice(from, 1)
        if (!slide) return
        s.presentation.slides.splice(to, 0, slide)
        s.slideIndex = to
        s.dirty = true
      }),

    updateSlide: (index, partial) =>
      set((s) => {
        const slide = s.presentation.slides[index]
        if (!slide) return
        Object.assign(slide, partial)
        s.dirty = true
      }),

    setSlideBackground: (partial) =>
      set((s) => {
        const slide = s.presentation.slides[s.slideIndex]
        slide.background = { ...(slide.background ?? { type: 'solid', color: '#ffffff' }), ...partial } as Slide['background']
        s.dirty = true
      }),

    setSlideTransition: (partial) =>
      set((s) => {
        const slide = s.presentation.slides[s.slideIndex]
        slide.transition = { ...(slide.transition ?? { preset: 'fade', duration: 500, easing: 'ease' }), ...partial }
        s.dirty = true
      }),

    setSlideNote: (note) =>
      set((s) => {
        s.presentation.slides[s.slideIndex].note = note
        s.dirty = true
      }),

    setViewportRatio: (ratio) =>
      set((s) => {
        if (ratio > 0 && ratio !== s.presentation.viewportRatio) {
          s.presentation.viewportRatio = ratio
          s.dirty = true
        }
      }),

    updateTheme: (partial) =>
      set((s) => {
        Object.assign(s.presentation.theme, partial)
        s.dirty = true
      }),

    addElement: (el, opts) =>
      set((s) => {
        s.presentation.slides[s.slideIndex].elements.push(el)
        if (opts?.select !== false) {
          s.selectedIds = [el.id]
          s.editingId = null
          s.editingCellId = null
        }
        s.dirty = true
      }),

    updateElements: (ids, updater, opts) =>
      set((s) => {
        const elements = s.presentation.slides[s.slideIndex].elements
        for (const el of elements) {
          if (ids.includes(el.id)) updater(el)
        }
        if (opts?.history !== false) s.dirty = true
      }),

    removeElements: (ids) =>
      set((s) => {
        const slide = s.presentation.slides[s.slideIndex]
        slide.elements = slide.elements.filter((el) => !ids.includes(el.id))
        s.selectedIds = []
        s.editingId = null
        s.editingCellId = null
        s.dirty = true
      }),

    duplicateElements: (ids) => {
      const state = get()
      const slide = state.presentation.slides[state.slideIndex]
      const picked = slide.elements.filter((el) => ids.includes(el.id))
      const cloned = cloneElements(picked)
      for (const el of cloned) {
        el.x += 24
        el.y += 24
      }
      set((s) => {
        s.presentation.slides[s.slideIndex].elements.push(...cloned)
        s.selectedIds = cloned.map((el) => el.id)
        s.dirty = true
      })
      return cloned.map((el) => el.id)
    },

    setElementLevel: (ids, level) =>
      set((s) => {
        const slide = s.presentation.slides[s.slideIndex]
        const idSet = new Set(ids)
        const picked = slide.elements.filter((el) => idSet.has(el.id))
        if (!picked.length) return

        if (level === 'top') {
          slide.elements = [...slide.elements.filter((el) => !idSet.has(el.id)), ...picked]
        } else if (level === 'bottom') {
          slide.elements = [...picked, ...slide.elements.filter((el) => !idSet.has(el.id))]
        } else if (level === 'up') {
          // 数组末尾更高；从后往前交换，避免同组互相卡住
          for (let i = slide.elements.length - 2; i >= 0; i--) {
            if (idSet.has(slide.elements[i].id) && !idSet.has(slide.elements[i + 1].id)) {
              const tmp = slide.elements[i]
              slide.elements[i] = slide.elements[i + 1]
              slide.elements[i + 1] = tmp
            }
          }
        } else {
          // down：从前往后交换
          for (let i = 1; i < slide.elements.length; i++) {
            if (idSet.has(slide.elements[i].id) && !idSet.has(slide.elements[i - 1].id)) {
              const tmp = slide.elements[i]
              slide.elements[i] = slide.elements[i - 1]
              slide.elements[i - 1] = tmp
            }
          }
        }
        s.dirty = true
      }),

    setGroup: (ids, groupId) =>
      set((s) => {
        const elements = s.presentation.slides[s.slideIndex].elements
        for (const el of elements) {
          if (ids.includes(el.id)) el.groupId = groupId
        }
        s.dirty = true
      }),

    setSelected: (ids) =>
      set((s) => {
        s.selectedIds = ids
        s.editingId = null
        s.editingCellId = null
      }),

    toggleSelected: (id) =>
      set((s) => {
        s.selectedIds = s.selectedIds.includes(id)
          ? s.selectedIds.filter((v) => v !== id)
          : [...s.selectedIds, id]
        s.editingId = null
      }),

    selectAll: () =>
      set((s) => {
        s.selectedIds = s.presentation.slides[s.slideIndex].elements
          .filter((el) => !el.lock)
          .map((el) => el.id)
      }),

    setEditingId: (id) =>
      set((s) => {
        s.editingId = id
        s.editingCellId = null
        if (id) s.selectedIds = [id]
      }),

    setEditingCellId: (id) =>
      set((s) => {
        s.editingCellId = id
      }),

    updateTableCellText: (cellId, text) => {
      const parts = cellId.split(':')
      if (parts.length !== 3) return
      const tableId = parts[0]
      const r = Number(parts[1])
      const c = Number(parts[2])
      if (Number.isNaN(r) || Number.isNaN(c)) return
      // 内容有变化才推历史，避免失焦即产生空撤销步骤
      const current = get().presentation.slides[get().slideIndex].elements.find((el) => el.id === tableId)
      const prevText = current && current.type === 'table' ? current.cells[r]?.[c]?.text : undefined
      if (prevText === text) return
      get().pushHistory()
      set((s) => {
        const elements = s.presentation.slides[s.slideIndex].elements
        const table = elements.find((el) => el.id === tableId)
        if (!table || table.type !== 'table') return
        const row = table.cells[r]
        if (!row) return
        const cell = row[c]
        if (!cell) return
        cell.text = text
        s.dirty = true
      })
    },

    copySelection: () => {
      const state = get()
      const slide = state.presentation.slides[state.slideIndex]
      const elements = slide.elements.filter((el) => state.selectedIds.includes(el.id))
      if (elements.length) set((s) => { s.clipboard = { slideId: slide.id, elements: structuredClone(elements) } })
    },

    cutSelection: () => {
      get().copySelection()
      get().removeElements(get().selectedIds)
    },

    pasteClipboard: (offset = 0) => {
      const clip = get().clipboard
      if (!clip) return
      const cloned = cloneElements(clip.elements)
      for (const el of cloned) {
        el.x += 24 + offset
        el.y += 24 + offset
      }
      set((s) => {
        s.presentation.slides[s.slideIndex].elements.push(...cloned)
        s.selectedIds = cloned.map((el) => el.id)
        s.dirty = true
      })
    },

    copyCurrentSlide: () => {
      const state = get()
      const slide = state.presentation.slides[state.slideIndex]
      if (!slide) return
      // 必须在 set 外取快照：immer draft 是 Proxy，structuredClone 会报 DataCloneError
      const snapshot = structuredClone(slide)
      set((s) => { s.slideClipboard = snapshot })
    },

    pasteSlide: () => {
      const clip = get().slideClipboard
      if (!clip) return
      const cloned = structuredClone(clip)
      cloned.id = genId('slide-')
      for (const el of cloned.elements) el.id = genId()
      set((s) => {
        s.presentation.slides.splice(s.slideIndex + 1, 0, cloned)
        s.slideIndex = s.slideIndex + 1
        s.selectedIds = []
        s.dirty = true
      })
    },

    pushHistory: () => {
      // 注意：pushHistory 在「变更前」调用，保存当前状态供 undo；
      // 快照在 set 外部获取（draft Proxy 无法 structuredClone）
      const snapshot = structuredClone(get().presentation)
      set((s) => {
        s.history.push(snapshot)
        if (s.history.length > HISTORY_LIMIT) s.history.shift()
        s.future = []
        s.dirty = true
      })
    },

    undo: () => {
      const state = get()
      if (!state.history.length) return
      const prev = state.history[state.history.length - 1]
      const current = structuredClone(state.presentation)
      set((s) => {
        s.history.pop()
        s.future.push(current)
        s.presentation = prev
        s.slideIndex = Math.min(s.slideIndex, prev.slides.length - 1)
        s.selectedIds = []
        s.editingId = null
        s.dirty = true
      })
    },

    redo: () => {
      const state = get()
      if (!state.future.length) return
      const next = state.future[state.future.length - 1]
      const current = structuredClone(state.presentation)
      set((s) => {
        s.future.pop()
        s.history.push(current)
        s.presentation = next
        s.slideIndex = Math.min(s.slideIndex, next.slides.length - 1)
        s.selectedIds = []
        s.editingId = null
        s.dirty = true
      })
    },

    markSaved: () => set((s) => { s.dirty = false }),
  })),
)

/* ---------- 派生工具 ---------- */

export function useCurrentSlide(): Slide | undefined {
  return useEditorStore((s) => s.presentation.slides[s.slideIndex])
}

export function useSelectedElements(): PPTElement[] {
  return useEditorStore((s) => {
    const slide = s.presentation.slides[s.slideIndex]
    if (!slide) return []
    return slide.elements.filter((el) => s.selectedIds.includes(el.id))
  })
}

/** 多选元素的整体包围盒 */
export function getSelectionBounds(elements: PPTElement[], ids: string[]) {
  const picked = elements.filter((el) => ids.includes(el.id))
  if (!picked.length) return null
  return unionBounds(picked.map((el) => ({ x: el.x, y: el.y, w: el.w, h: el.h })))
}
