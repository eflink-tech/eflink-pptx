// 导入对话框：PPTX（先预览再导入）/ JSON
import { useEffect, useRef, useState } from 'react'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { parseJSONFile } from '../../core/export/json'
import { importPPTXDetailed } from '../../core/import'
import { previewPPTXDetailed } from '../../core/preview'
import type { ImportReport } from '../../types/slides'

/** 兼容性报告 skip 种类 → 用户可读文案 */
const SKIP_LABELS: Record<string, string> = {
  missingImage: '缺失图片',
  missingChart: '缺失图表数据',
  unknownChart: '未识别图表',
  elementParseFailed: '无法解析的元素',
  slideParseFailed: '无法解析的页面',
  smartartFallback: 'SmartArt 已转图片',
  groupParseFailed: '组合解析失败',
  groupRotation: '组合旋转未还原',
  missingPlaceholder: '占位符缺失',
  graphicFrameUnknown: '未识别的元素容器',
}

/** 兼容性报告 skip 项 → 「文案 ×N」列表（applyImport 与预览视图共用） */
function formatSkipped(report: ImportReport): string[] {
  return Object.entries(report.skipped).map(([k, v]) => `${SKIP_LABELS[k] ?? k} ×${v}`)
}

/** 预览状态：文件 + SVG 页面 + 兼容性报告 */
interface PreviewState {
  file: File
  pages: SVGSVGElement[]
  report: ImportReport
}

export function ImportDialog() {
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
    toast('正在导入 PPTX…')
    const { presentation: pres, report } = await importPPTXDetailed(file)
    useEditorStore.getState().pushHistory()
    useEditorStore.getState().replacePresentation(pres)
    const items = formatSkipped(report)
    if (items.length) {
      toast(`已导入 PPTX（${pres.slides.length} 页）；部分内容未完整还原：${items.join('、')}`, 'success')
    } else {
      toast(`已导入 PPTX（${pres.slides.length} 页）`, 'success')
    }
    useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
    closeModal()
  }

  const handleError = (error: unknown) => {
    toast(`导入失败：${error instanceof Error ? error.message : '未知错误'}`, 'error')
  }

  const handleFile = async (file: File) => {
    try {
      if (file.name.toLowerCase().endsWith('.json') || file.name.toLowerCase().endsWith('.efppt.json')) {
        const text = await file.text()
        const pres = parseJSONFile(text)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        toast(`已导入 JSON（${pres.slides.length} 页）`, 'success')
        useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
        closeModal()
      } else if (file.name.toLowerCase().endsWith('.pptx')) {
        // 先生成高保真预览，用户确认后再真正导入
        toast('正在生成预览…')
        setBusy(true)
        try {
          const { pages, report } = await previewPPTXDetailed(file)
          setPreview({ file, pages, report })
        } finally {
          setBusy(false)
        }
      } else {
        toast('请选择 .pptx 或 .json 文件', 'error')
      }
    } catch (error) {
      handleError(error)
    }
  }

  // 预览视图：SVG 缩略网格 + 降级提示 + 返回/确认
  if (preview) {
    const items = formatSkipped(preview.report)
    return (
      <Modal title="导入预览" width={760}>
        <div ref={gridRef} className="grid max-h-[60vh] grid-cols-3 gap-3 overflow-auto" />
        {items.length > 0 && (
          <div className="mt-2 text-xs text-gray-500">部分内容预览/还原存在降级：{items.join('、')}</div>
        )}
        <div className="mt-4 flex justify-end gap-2">
          <button
            className="rounded-lg border border-gray-300 px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 disabled:opacity-50"
            disabled={busy}
            onClick={() => setPreview(null)}
          >
            返回
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
            确认导入（可编辑）
          </button>
        </div>
      </Modal>
    )
  }

  return (
    <Modal title="导入">
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
        <div className="text-sm text-gray-600">点击选择文件</div>
        <div className="mt-1 text-xs text-gray-400">支持 .pptx（PowerPoint 演示文稿）与 .json（本工具工程文件）</div>
        <div className="mt-2 text-[11px] text-gray-400">PPTX 将先展示高保真预览，确认后再导入为可编辑内容</div>
        <div className="mt-2 text-[11px] text-gray-400">PPTX 导入尽力还原文本、形状、图片、表格与线条；复杂图表与艺术效果可能有差异</div>
      </label>
    </Modal>
  )
}
