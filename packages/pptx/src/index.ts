/** @eflink-tech/pptx 对外导出面：编辑器组件 + store + 持久化辅助 + 数据模型 */
// 顶栏/菜单栏样式来自 office-kit：宿主只引入 dist/styles.css，必须在此显式引入才会打进产物。
// office-kit 的 JS 不在构建外部化清单里（随包打进 dist/index.js），宿主无需直接安装它。
import '@eflink-tech/office-kit/styles.css'
import './styles.css'
import { migrateLegacyLocalStorage } from './storage/legacy'

// 品牌迁移（efLink → DocuLuma）：模块加载即把旧命名空间的 localStorage 键一次性改名，
// 早于任何编辑器挂载/读取路径；IndexedDB 旧库的导入是惰性的（见 storage/legacy.ts）。
migrateLegacyLocalStorage()

// 组件
export { PptxEditor } from './PptxEditor'
export type { PptxEditorProps, PptxEditorBrand } from './PptxEditor'
export type { OfficeLocale as PptxLocale } from '@eflink-tech/office-kit'

// 品牌切换：'doculuma' | 'eflink' 内置品牌，或传对象自定义
export { builtinEditorBrand, resolveEditorBrandId, EDITOR_BRAND_IDS } from './brand'
export type { EditorBrandId } from './brand'
export { setPptxStorageBackend } from './core/editor/persistence'
export { setEditorBackHref } from './core/editor/chrome'
export type { PptxStorageBackend } from './core/editor/persistence'

// 分享桥接：宿主注入分享实现（编辑器分享按钮调用）
export { setPptxShareHandler } from './core/share/shareBridge'
export type { PptxShareHandler, PptxShareResult } from './core/share/shareBridge'

// 状态
export { useEditorStore } from './store/editorStore'
export { useUIStore } from './store/uiStore'

// 持久化（localStorage 镜像草稿 + IndexedDB/后端多文档；云端保存为手动触发）
export {
  loadStartupDoc, saveDoc, listDocs, createDoc, deleteDoc, duplicateDoc, renameDoc,
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
