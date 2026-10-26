// 主题配色对话框：一键换色板（应用到图表默认色/背景可选）
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useToastStore } from '../../store/uiStore'
import { PPT_TEMPLATES } from '../../core/schema/templates'
import { TEMPLATE_NAME_KEYS } from './TemplateDialog'

const PALETTE_PRESETS = PPT_TEMPLATES.map((tpl) => ({ id: tpl.id, name: tpl.name, colors: tpl.colors, bg: tpl.background.type === 'solid' ? tpl.background.color ?? '#ffffff' : '#ffffff' }))

export function ThemeDialog() {
  const { t } = useI18n()
  const apply = (colors: string[], bg: string, applyBg: boolean) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    store.updateTheme({ colors })
    if (applyBg) {
      const pres = structuredClone(useEditorStore.getState().presentation)
      for (const slide of pres.slides) {
        if (slide.background?.type === 'solid') slide.background.color = bg
      }
      store.replacePresentation(pres)
    }
    useToastStore.getState().toast(t('pptx.theme.updated'), 'success')
  }

  return (
    <Modal title={t('pptx.theme.title')}>
      <div className="space-y-3">
        {PALETTE_PRESETS.map((p) => (
          <div key={p.id} className="flex items-center justify-between rounded-lg border border-gray-200 px-3 py-2.5">
            <div>
              <div className="text-xs font-medium text-gray-700">{TEMPLATE_NAME_KEYS[p.id] ? t(TEMPLATE_NAME_KEYS[p.id]) : p.name}</div>
              <div className="mt-1.5 flex gap-1">
                {p.colors.map((c, i) => <span key={i} className="size-5 rounded" style={{ background: c }} />)}
              </div>
            </div>
            <div className="flex gap-1.5">
              <button
                className="rounded bg-[#d14424] px-3 py-1 text-xs text-white hover:bg-[#b93a1d]"
                onClick={() => apply(p.colors, p.bg, false)}
              >
                {t('pptx.theme.apply')}
              </button>
              <button
                className="rounded border border-gray-200 px-3 py-1 text-xs text-gray-600 hover:bg-gray-50"
                onClick={() => apply(p.colors, p.bg, true)}
                title={t('pptx.theme.applyBgHint')}
              >
                {t('pptx.theme.applyBg')}
              </button>
            </div>
          </div>
        ))}
      </div>
    </Modal>
  )
}
