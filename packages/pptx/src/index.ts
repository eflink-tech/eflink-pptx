/** @eflink-tech/pptx 对外导出面：编辑器组件 + store + 持久化辅助 + 数据模型 */
import './styles.css'

// 组件
export { PptxEditor } from './PptxEditor'
export { setPptxStorageBackend } from './core/editor/persistence'
export type { PptxStorageBackend } from './core/editor/persistence'

// 状态
export { useEditorStore } from './store/editorStore'
export { useUIStore } from './store/uiStore'

// 持久化（localStorage 镜像 + IndexedDB 多文档）
export {
  loadStartupDoc, saveDoc, scheduleAutosave, listDocs, createDoc, deleteDoc, duplicateDoc, renameDoc,
} from './core/editor/persistence'

// 元素工厂（程序化构建页面用）
export {
  createTextElement, createImageElement, createShapeElement, createLineElement,
  createTableElement, createChartElement, createVideoElement, createAudioElement,
  createFormulaElement, createElementByType,
} from './core/schema/factory'

// 数据模型
export { SLIDE_WIDTH, SLIDE_HEIGHT, PX_PER_INCH, createPresentation, createSlide } from './types/slides'
export type { Presentation, Slide, PPTElement, Theme } from './types/slides'
