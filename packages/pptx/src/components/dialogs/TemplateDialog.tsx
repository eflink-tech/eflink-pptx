// 模板库对话框
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'
import { PPT_TEMPLATES } from '../../core/schema/templates'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { genId } from '../../core/utils/id'
import { SlideRenderer } from '../canvas/SlideRenderer'

export const TEMPLATE_NAME_KEYS: Record<string, string> = {
  'tpl-brand': 'pptx.template.name.brand',
  'tpl-tech': 'pptx.template.name.tech',
  'tpl-fresh': 'pptx.template.name.fresh',
  'tpl-warm': 'pptx.template.name.warm',
  'tpl-blue': 'pptx.template.name.blue',
}

export function TemplateDialog() {
  const { t } = useI18n()
  const closeModal = useUIStore.getState().closeModal
  const apply = (tpl: (typeof PPT_TEMPLATES)[number]) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    const slides = tpl.build(t)
    store.replacePresentation({
      slides,
      theme: { ...structuredClone(store.presentation.theme), colors: tpl.colors },
      width: store.presentation.width,
      viewportRatio: store.presentation.viewportRatio,
    })
    useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
    useToastStore.getState().toast(t('pptx.template.applied', { name: TEMPLATE_NAME_KEYS[tpl.id] ? t(TEMPLATE_NAME_KEYS[tpl.id]) : tpl.name }), 'success')
    closeModal()
  }

  const insertPage = (tpl: (typeof PPT_TEMPLATES)[number]) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    const slides = tpl.build(t)
    const pageIndex = store.slideIndex
    store.addSlide(pageIndex + 1, { ...slides[1], id: genId('slide-') })
    useToastStore.getState().toast(t('pptx.template.pageInserted'), 'success')
    closeModal()
  }

  // 缩略图尺寸：1280×720 按比例缩小
  const THUMB_W = 224
  const THUMB_H = Math.round(THUMB_W * (720 / 1280))
  const thumbScale = THUMB_W / 1280

  return (
    <Modal title={t('pptx.template.title')} width={760}>
      <div className="grid grid-cols-3 gap-4">
        {PPT_TEMPLATES.map((tpl) => (
          <div key={tpl.id} className="rounded-lg border border-gray-200 p-3">
            <div className="overflow-hidden rounded" style={{ width: THUMB_W, height: THUMB_H }}>
              <div style={{ width: 1280, height: 720, transform: `scale(${thumbScale})`, transformOrigin: 'top left' }}>
                <SlideRenderer slide={tpl.build(t)[0]} width={1280} height={720} staticMode />
              </div>
            </div>
            <div className="mt-2 text-xs font-medium text-gray-700">{TEMPLATE_NAME_KEYS[tpl.id] ? t(TEMPLATE_NAME_KEYS[tpl.id]) : tpl.name}</div>
            <div className="mt-2 flex gap-1.5">
              <button
                className="flex-1 rounded bg-[#d14424] py-1.5 text-xs text-white hover:bg-[#b93a1d]"
                onClick={() => apply(tpl)}
              >
                {t('pptx.template.apply')}
              </button>
              <button
                className="flex-1 rounded border border-gray-200 py-1.5 text-xs text-gray-600 hover:bg-gray-50"
                onClick={() => insertPage(tpl)}
              >
                {t('pptx.template.insertPage')}
              </button>
            </div>
          </div>
        ))}
      </div>
    </Modal>
  )
}
