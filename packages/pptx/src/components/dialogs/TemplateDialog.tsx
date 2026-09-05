// 模板库对话框
import { Modal } from './ModalHost'
import { PPT_TEMPLATES } from '../../core/schema/templates'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { genId } from '../../core/utils/id'
import { SlideRenderer } from '../canvas/SlideRenderer'

export function TemplateDialog() {
  const closeModal = useUIStore.getState().closeModal
  const apply = (tpl: (typeof PPT_TEMPLATES)[number]) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    const slides = tpl.build()
    store.replacePresentation({
      slides,
      theme: { ...structuredClone(store.presentation.theme), colors: tpl.colors },
      width: store.presentation.width,
      viewportRatio: store.presentation.viewportRatio,
    })
    useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
    useToastStore.getState().toast(`已应用模板「${tpl.name}」`, 'success')
    closeModal()
  }

  const insertPage = (tpl: (typeof PPT_TEMPLATES)[number]) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    const slides = tpl.build()
    const pageIndex = store.slideIndex
    store.addSlide(pageIndex + 1, { ...slides[1], id: genId('slide-') })
    useToastStore.getState().toast('已插入内容页', 'success')
    closeModal()
  }

  // 缩略图尺寸：1280×720 按比例缩小
  const THUMB_W = 224
  const THUMB_H = Math.round(THUMB_W * (720 / 1280))
  const thumbScale = THUMB_W / 1280

  return (
    <Modal title="模板库" width={760}>
      <div className="grid grid-cols-3 gap-4">
        {PPT_TEMPLATES.map((tpl) => (
          <div key={tpl.id} className="rounded-lg border border-gray-200 p-3">
            <div className="overflow-hidden rounded" style={{ width: THUMB_W, height: THUMB_H }}>
              <div style={{ width: 1280, height: 720, transform: `scale(${thumbScale})`, transformOrigin: 'top left' }}>
                <SlideRenderer slide={tpl.build()[0]} width={1280} height={720} staticMode />
              </div>
            </div>
            <div className="mt-2 text-xs font-medium text-gray-700">{tpl.name}</div>
            <div className="mt-2 flex gap-1.5">
              <button
                className="flex-1 rounded bg-[#d14424] py-1.5 text-xs text-white hover:bg-[#b93a1d]"
                onClick={() => apply(tpl)}
              >
                应用整套
              </button>
              <button
                className="flex-1 rounded border border-gray-200 py-1.5 text-xs text-gray-600 hover:bg-gray-50"
                onClick={() => insertPage(tpl)}
              >
                插入内容页
              </button>
            </div>
          </div>
        ))}
      </div>
    </Modal>
  )
}
