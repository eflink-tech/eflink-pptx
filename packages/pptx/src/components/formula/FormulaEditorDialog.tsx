// 公式编辑器弹窗：650×550 弹窗，内嵌 MathLive 编辑器
import { useState } from 'react'
import { MathFieldEditor } from './MathFieldEditor'

interface FormulaEditorDialogProps {
  /** 当前 LaTeX 值 */
  value: string
  /** 确认回调 */
  onConfirm: (latex: string) => void
  /** 关闭回调 */
  onClose: () => void
}

/** MathLive 导出时会产生 \placeholder 占位符，KaTeX 无法渲染，需清理 */
function cleanLatex(latex: string): string {
  return latex
    .replace(/\\placeholder/g, '')       // 移除占位符
    .replace(/\^\{\}/g, '')               // 清理空上标 ^{}
    .replace(/_\{\}/g, '')                // 清理空下标 _{}
    .replace(/\{(\s*)\}/g, '')            // 清理空分组 {}
    .replace(/\s+/g, ' ')                 // 合并多余空格
    .trim()
}

export function FormulaEditorDialog({ value, onConfirm, onClose }: FormulaEditorDialogProps) {
  const [latex, setLatex] = useState(value)

  return (
    <div className="fixed inset-0 z-[150] flex items-center justify-center bg-black/40" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose() }}>
      <div className="modal-pop flex flex-col overflow-hidden rounded-xl bg-white shadow-2xl" style={{ width: 650, height: 550 }}>
        {/* 标题栏 */}
        <div className="flex items-center justify-between border-b border-gray-100 px-5 py-3">
          <span className="text-sm font-semibold text-gray-700">编辑公式</span>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100" onClick={onClose}>✕</button>
        </div>

        {/* 编辑区 */}
        <div className="flex-1 overflow-y-auto p-5">
          <MathFieldEditor
            value={latex}
            onChange={setLatex}
            className="formula-editor-large"
          />
        </div>

        {/* 底部按钮 */}
        <div className="flex items-center justify-end gap-2 border-t border-gray-100 px-5 py-3">
          <button className="rounded border border-gray-200 px-4 py-1.5 text-sm text-gray-600 hover:bg-gray-50" onClick={onClose}>
            取消
          </button>
          <button className="rounded bg-[#d14424] px-4 py-1.5 text-sm text-white hover:bg-[#b93a1d]" onClick={() => onConfirm(cleanLatex(latex))}>
            确认
          </button>
        </div>
      </div>
    </div>
  )
}
