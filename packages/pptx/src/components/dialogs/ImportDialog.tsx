// 导入对话框：PPTX / JSON
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { parseJSONFile } from '../../core/export/json'
import { importPPTXDetailed } from '../../core/import'

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
}

export function ImportDialog() {
  const toast = useToastStore.getState().toast
  const closeModal = useUIStore.getState().closeModal

  const handleFile = async (file: File) => {
    try {
      if (file.name.toLowerCase().endsWith('.json') || file.name.toLowerCase().endsWith('.efppt.json')) {
        const text = await file.text()
        const pres = parseJSONFile(text)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        toast(`已导入 JSON（${pres.slides.length} 页）`, 'success')
      } else if (file.name.toLowerCase().endsWith('.pptx')) {
        toast('正在解析 PPTX…')
        const { presentation: pres, report } = await importPPTXDetailed(file)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        const items = Object.entries(report.skipped).map(([k, v]) => `${SKIP_LABELS[k] ?? k} ×${v}`)
        if (items.length) {
          toast(`已导入 PPTX（${pres.slides.length} 页）；部分内容未完整还原：${items.join('、')}`, 'success')
        } else {
          toast(`已导入 PPTX（${pres.slides.length} 页）`, 'success')
        }
      } else {
        toast('请选择 .pptx 或 .json 文件', 'error')
        return
      }
      useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
      closeModal()
    } catch (error) {
      toast(`导入失败：${error instanceof Error ? error.message : '未知错误'}`, 'error')
    }
  }

  return (
    <Modal title="导入">
      <label className="block cursor-pointer rounded-xl border-2 border-dashed border-gray-300 px-6 py-10 text-center hover:border-[#d14424]">
        <input
          type="file"
          accept=".pptx,.json"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0]
            if (file) void handleFile(file)
            e.target.value = ''
          }}
        />
        <div className="text-sm text-gray-600">点击选择文件</div>
        <div className="mt-1 text-xs text-gray-400">支持 .pptx（PowerPoint 演示文稿）与 .json（本工具工程文件）</div>
        <div className="mt-2 text-[11px] text-gray-400">PPTX 导入尽力还原文本、形状、图片、表格与线条；复杂图表与艺术效果可能有差异</div>
      </label>
    </Modal>
  )
}
