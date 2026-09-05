// UI 状态：面板/模态框/放映/画布视图
import { create } from 'zustand'

export type ModalName =
  | 'template' | 'theme' | 'export' | 'import' | 'findReplace'
  | 'hotkey' | 'fileManager' | 'aiSettings' | 'aiHistory' | 'about' | null

export type PlayerMode = 'off' | 'playing' | 'presenter'

export type RightPanelTab = 'style' | 'animation' | 'comment'

interface UIState {
  /** 左侧缩略图面板可见 */
  thumbnailsVisible: boolean
  /** AI 面板可见 */
  aiPanelVisible: boolean
  /** 右侧面板当前页签 */
  rightPanelTab: RightPanelTab
  /** 当前打开的模态框 */
  modal: ModalName
  /** 放映模式 */
  playerMode: PlayerMode
  /** 放映起始页（进入放映时的页索引） */
  playerStartIndex: number
  /** 画布缩放（1 = 100%；0 表示自动适配） */
  canvasScale: number
  /** 网格可见 */
  gridVisible: boolean
  /** 参考线可见 */
  guideLines: boolean
  /** 查找替换高亮（元素 id 列表） */
  findHitIds: string[]

  toggleThumbnails: () => void
  toggleAIPanel: () => void
  setRightPanelTab: (tab: RightPanelTab) => void
  openModal: (modal: ModalName) => void
  closeModal: () => void
  setPlayerMode: (mode: PlayerMode, startIndex?: number) => void
  setCanvasScale: (scale: number) => void
  toggleGrid: () => void
  toggleGuideLines: () => void
  setFindHits: (ids: string[]) => void
}

export const useUIStore = create<UIState>()((set) => ({
  thumbnailsVisible: true,
  aiPanelVisible: false,
  rightPanelTab: 'style',
  modal: null,
  playerMode: 'off',
  playerStartIndex: 0,
  canvasScale: 0,
  gridVisible: false,
  guideLines: true,
  findHitIds: [],

  toggleThumbnails: () => set((s) => ({ thumbnailsVisible: !s.thumbnailsVisible })),
  toggleAIPanel: () => set((s) => ({ aiPanelVisible: !s.aiPanelVisible })),
  setRightPanelTab: (tab) => set({ rightPanelTab: tab }),
  openModal: (modal) => set({ modal }),
  closeModal: () => set({ modal: null }),
  setPlayerMode: (mode, startIndex) =>
    set((s) => ({ playerMode: mode, playerStartIndex: startIndex ?? s.playerStartIndex })),
  setCanvasScale: (scale) => set({ canvasScale: scale }),
  toggleGrid: () => set((s) => ({ gridVisible: !s.gridVisible })),
  toggleGuideLines: () => set((s) => ({ guideLines: !s.guideLines })),
  setFindHits: (ids) => set({ findHitIds: ids }),
}))

/** 轻量消息提示（不引第三方库） */
export type ToastKind = 'success' | 'error' | 'info'
interface ToastItem { id: number; kind: ToastKind; text: string }
interface ToastState {
  toasts: ToastItem[]
  toast: (text: string, kind?: ToastKind) => void
  remove: (id: number) => void
}

let toastSeq = 1

export const useToastStore = create<ToastState>()((set, get) => ({
  toasts: [],
  toast: (text, kind = 'info') => {
    const id = toastSeq++
    set((s) => ({ toasts: [...s.toasts, { id, kind, text }] }))
    window.setTimeout(() => get().remove(id), 2600)
  },
  remove: (id) => set((s) => ({ toasts: s.toasts.filter((t) => t.id !== id) })),
}))
