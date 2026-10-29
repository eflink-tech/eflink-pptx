// 导入对话框：PPTX（先预览再导入）/ JSON
import { useEffect, useRef, useState } from 'react'
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { parseJSONFile } from '../../core/export/json'
import { importPPTXDetailed } from '../../core/import'
import { previewPPTXDetailed } from '../../core/preview'
import type { ImportReport } from '../../types/slides'

/** 兼容性报告 skip 种类 → 文案 key */
const SKIP_LABEL_KEYS: Record<string, string> = {
  missingImage: 'pptx.import.skip.missingImage',
  missingChart: 'pptx.import.skip.missingChart',
  unknownChart: 'pptx.import.skip.unknownChart',
  elementParseFailed: 'pptx.import.skip.elementParseFailed',
  slideParseFailed: 'pptx.import.skip.slideParseFailed',
  smartartFallback: 'pptx.import.skip.smartartFallback',
  smartartDrawing: 'pptx.import.skip.smartartDrawing',
  smartartUnsupported: 'pptx.import.skip.smartartUnsupported',
  groupParseFailed: 'pptx.import.skip.groupParseFailed',
  groupRotation: 'pptx.import.skip.groupRotation',
  missingPlaceholder: 'pptx.import.skip.missingPlaceholder',
  graphicFrameUnknown: 'pptx.import.skip.graphicFrameUnknown',
  alternateContentUnsupported: 'pptx.import.skip.alternateContentUnsupported',
  model3dFallback: 'pptx.import.skip.model3dFallback',
  oleObjectUnsupported: 'pptx.import.skip.oleObjectUnsupported',
  olePreviewOnly: 'pptx.import.skip.olePreviewOnly',
  contentPartUnsupported: 'pptx.import.skip.contentPartUnsupported',
  unknownElement: 'pptx.import.skip.unknownElement',
  animation: 'pptx.import.skip.animation',
  hiddenSlide: 'pptx.import.skip.hiddenSlide',
  notesUnsupported: 'pptx.import.skip.notesUnsupported',
  transitionUnsupported: 'pptx.import.skip.transitionUnsupported',
  section: 'pptx.import.skip.section',
  patternFill: 'pptx.import.skip.patternFill',
  unsupportedFill: 'pptx.import.skip.unsupportedFill',
  unsupportedMedia: 'pptx.import.skip.unsupportedMedia',
  embeddedFont: 'pptx.import.skip.embeddedFont',
  textEffects: 'pptx.import.skip.textEffects',
  tableRichText: 'pptx.import.skip.tableRichText',
  tableBorders: 'pptx.import.skip.tableBorders',
  tableStyleBuiltin: 'pptx.import.skip.tableStyleBuiltin',
  imageClipGeom: 'pptx.import.skip.imageClipGeom',
  commentsUnsupported: 'pptx.import.skip.commentsUnsupported',
  chartCombo: 'pptx.import.skip.chartCombo',
  chartStyle: 'pptx.import.skip.chartStyle',
  chartBackground: 'pptx.import.skip.chartBackground',
  unsupportedEffect: 'pptx.import.skip.unsupportedEffect',
  groupFlip: 'pptx.import.skip.groupFlip',
}

/** 兼容性报告 skip 项 → 「文案 ×N」列表（applyImport 与预览视图共用） */
function formatSkipped(report: ImportReport, t: (key: string) => string): string[] {
  return Object.entries(report.skipped).map(([k, v]) => `${SKIP_LABEL_KEYS[k] ? t(SKIP_LABEL_KEYS[k]) : k} ×${v}`)
}

/** 预览状态：文件 + SVG 页面 + 兼容性报告 */
interface PreviewState {
  file: File
  pages: SVGSVGElement[]
  report: ImportReport
}

export function ImportDialog() {
  const { t } = useI18n()
  const toast = useToastStore.getState().toast
  const closeModal = useUIStore.getState().closeModal
  const [preview, setPreview] = useState<PreviewState | null>(null)
  const [busy, setBusy] = useState(false)
  const gridRef = useRef<HTMLDivElement>(null)

  // 预览变化时把 SVG DOM（非 React 元素）直接挂载到网格容器
  useEffect(() => {
    const grid = gridRef.current
    if (!grid) return
    grid.innerHTML = ''
    if (!preview) return
    preview.pages.forEach((page, i) => {
      const cell = document.createElement('div')
      cell.className = 'relative overflow-hidden rounded-lg border border-gray-200 bg-white'
      page.style.width = '100%'
      page.style.height = 'auto'
      page.style.display = 'block'
      cell.appendChild(page)
      const badge = document.createElement('div')
      badge.className = 'absolute right-1 top-1 rounded bg-black/60 px-1 text-[10px] text-white'
      badge.textContent = String(i + 1)
      cell.appendChild(badge)
      grid.appendChild(cell)
    })
  }, [preview])

  /** 确认导入：解析并替换当前演示文稿（复用现有导入逻辑） */
  const applyImport = async (file: File) => {
    toast(t('pptx.import.importing'))
    const { presentation: pres, report } = await importPPTXDetailed(file)
    useEditorStore.getState().pushHistory()
    useEditorStore.getState().replacePresentation(pres)
    const items = formatSkipped(report, t)
    if (items.length) {
      toast(t('pptx.import.pptxPartial', { count: pres.slides.length, items: items.join(t('pptx.listSeparator')) }), 'success')
    } else {
      toast(t('pptx.import.pptxDone', { count: pres.slides.length }), 'success')
    }
    useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
    closeModal()
  }

  const handleError = (error: unknown) => {
    toast(t('pptx.import.failed', { message: error instanceof Error ? error.message : t('pptx.export.unknownError') }), 'error')
  }

  const handleFile = async (file: File) => {
    try {
      if (file.name.toLowerCase().endsWith('.json') || file.name.toLowerCase().endsWith('.efppt.json')) {
        const text = await file.text()
        const pres = parseJSONFile(text)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        toast(t('pptx.import.jsonDone', { count: pres.slides.length }), 'success')
        useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
        closeModal()
      } else if (file.name.toLowerCase().endsWith('.pptx')) {
        // 先生成高保真预览，用户确认后再真正导入
        toast(t('pptx.import.previewing'))
        setBusy(true)
        try {
          const { pages, report } = await previewPPTXDetailed(file)
          setPreview({ file, pages, report })
        } finally {
          setBusy(false)
        }
      } else {
        toast(t('pptx.import.chooseFile'), 'error')
      }
    } catch (error) {
      handleError(error)
    }
  }

  // 预览视图：SVG 缩略网格 + 降级提示 + 返回/确认
  if (preview) {
    const items = formatSkipped(preview.report, t)
    return (
      <Modal title={t('pptx.import.previewTitle')} width={760}>
        <div ref={gridRef} className="grid max-h-[60vh] grid-cols-3 gap-3 overflow-auto" />
        {items.length > 0 && (
          <div className="mt-2 text-xs text-gray-500">{t('pptx.import.degraded', { items: items.join(t('pptx.listSeparator')) })}</div>
        )}
        <div className="mt-4 flex justify-end gap-2">
          <button
            className="rounded-lg border border-gray-300 px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 disabled:opacity-50"
            disabled={busy}
            onClick={() => setPreview(null)}
          >
            {t('pptx.import.back')}
          </button>
          <button
            className="rounded-lg bg-[#d14424] px-4 py-2 text-sm text-white hover:bg-[#b93a1d] disabled:opacity-50"
            disabled={busy}
            onClick={() => {
              // busy 期间按钮已禁用；原地导入，失败保留预览供重试
              setBusy(true)
              void applyImport(preview.file).catch((error) => {
                setBusy(false)
                handleError(error)
              })
            }}
          >
            {t('pptx.import.confirm')}
          </button>
        </div>
      </Modal>
    )
  }

  return (
    <Modal title={t('pptx.import.title')}>
      <label className="block cursor-pointer rounded-xl border-2 border-dashed border-gray-300 px-6 py-10 text-center hover:border-[#d14424]">
        <input
          type="file"
          accept=".pptx,.json"
          className="hidden disabled:opacity-50"
          disabled={busy}
          onChange={(e) => {
            const file = e.target.files?.[0]
            if (file) void handleFile(file)
            e.target.value = ''
          }}
        />
        <div className="text-sm text-gray-600">{t('pptx.import.choose')}</div>
        <div className="mt-1 text-xs text-gray-400">{t('pptx.import.support')}</div>
        <div className="mt-2 text-[11px] text-gray-400">{t('pptx.import.previewNote')}</div>
        <div className="mt-2 text-[11px] text-gray-400">{t('pptx.import.limitNote')}</div>
      </label>
    </Modal>
  )
}
