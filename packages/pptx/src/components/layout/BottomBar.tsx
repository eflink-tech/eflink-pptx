// 底部状态栏：文档名（点击改名）/ 保存状态 / 页码导航 / 网格 / 缩放
import { useRef, useState } from 'react'
import { ChevronLeft, ChevronRight, ZoomIn, ZoomOut, Pencil } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { renameDoc } from '../../core/editor/persistence'

export function BottomBar() {
  const slideCount = useEditorStore((s) => s.presentation.slides.length)
  const slideIndex = useEditorStore((s) => s.slideIndex)
  const docName = useEditorStore((s) => s.docName)
  const dirty = useEditorStore((s) => s.dirty)
  const canvasScale = useUIStore((s) => s.canvasScale)
  const gridVisible = useUIStore((s) => s.gridVisible)

  // 文档名行内编辑：Esc 先置跳过标记再 blur，避免 onBlur 误提交（Enter/失焦提交，空值不提交）
  const [nameEditing, setNameEditing] = useState(false)
  const skipCommitRef = useRef(false)

  // 改名提交：独立保存（仅更新文档记录 name，不影响内容 dirty）；失败 toast 且名字保持原值（回滚）
  const commitRename = async (raw: string) => {
    const name = raw.trim()
    const s = useEditorStore.getState()
    if (!name || name === s.docName) return
    try {
      await renameDoc(s.docId, name, s.presentation)
      useEditorStore.getState().renameDocument(name)
    } catch {
      useToastStore.getState().toast('重命名失败，请重试', 'error')
    }
  }

  const goto = (i: number) => {
    const clamped = Math.min(Math.max(0, i), slideCount - 1)
    useEditorStore.setState({ slideIndex: clamped, selectedIds: [], editingId: null })
  }

  const setScale = (next: number) => {
    useUIStore.getState().setCanvasScale(next <= 0 ? 0 : Math.min(4, Math.max(0.1, next)))
  }

  // canvasScale 0 = 自适应
  const percent = canvasScale > 0 ? Math.round(canvasScale * 100) : null

  return (
    <div className="flex h-7 shrink-0 select-none items-center justify-between border-t border-[#e0e0e0] bg-[#f8f8f8] px-3 text-[11px] text-[#888]">
      {/* 左侧：文档名（点击改名）+ 保存状态 + 页码导航 */}
      <div className="flex min-w-0 items-center gap-3">
        <div className="flex min-w-0 items-center" data-testid="doc-name">
          {nameEditing
            ? (
                <input
                  autoFocus
                  defaultValue={docName}
                  onFocus={(e) => e.currentTarget.select()}
                  className="w-36 rounded border border-[#e0e0e0] bg-white px-1 py-0.5 text-[11px] text-[#333] outline-none focus:border-[#d14424]"
                  onBlur={(e) => {
                    setNameEditing(false)
                    if (skipCommitRef.current) { skipCommitRef.current = false; return }
                    void commitRename(e.target.value)
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                    else if (e.key === 'Escape') { skipCommitRef.current = true; (e.target as HTMLInputElement).blur() }
                  }}
                />
              )
            : (
                <button
                  className="flex min-w-0 items-center gap-1 rounded px-1.5 py-0.5 text-left text-[#666] transition-colors hover:bg-[#e8e8e8]"
                  title="点击重命名"
                  onClick={() => setNameEditing(true)}
                >
                  <Pencil size={11} className="shrink-0 text-[#999]" />
                  <span className="max-w-[200px] truncate">{docName}</span>
                </button>
              )}
        </div>

        {/* 保存状态指示：dirty = 红色未保存；干净 = 灰色已保存（手动 ⌘S/Ctrl+S 云端保存成功后更新） */}
        <span className="shrink-0" data-testid="save-state" title={dirty ? '有未保存的修改，按 Ctrl+S 保存' : '所有修改已保存'}>
          {dirty
            ? <span className="text-[#e02e2e]">● 未保存</span>
            : <span className="text-gray-400">✓ 已保存</span>}
        </span>

        <div className="h-3.5 w-px shrink-0 bg-[#e0e0e0]" />

        <div className="flex shrink-0 items-center gap-1">
          <button className="rounded p-1 hover:bg-[#e8e8e8] disabled:opacity-40" disabled={slideIndex === 0} onClick={() => goto(slideIndex - 1)}>
            <ChevronLeft size={14} />
          </button>
          <span data-testid="slide-nav">{slideIndex + 1} / {slideCount}</span>
          <button className="rounded p-1 hover:bg-[#e8e8e8] disabled:opacity-40" disabled={slideIndex === slideCount - 1} onClick={() => goto(slideIndex + 1)}>
            <ChevronRight size={14} />
          </button>
        </div>
      </div>

      <div className="flex shrink-0 items-center gap-2">
        <button className={`rounded px-2 py-0.5 ${gridVisible ? 'bg-[#fbeae5] text-[#d14424]' : 'hover:bg-[#e8e8e8]'}`} onClick={() => useUIStore.getState().toggleGrid()}>
          网格
        </button>
        <div className="flex items-center gap-1">
          <button className="rounded p-1 hover:bg-[#e8e8e8]" onClick={() => setScale((canvasScale > 0 ? canvasScale : 1) - 0.1)}>
            <ZoomOut size={14} />
          </button>
          <button
            className="w-14 rounded px-1 py-0.5 text-center hover:bg-[#e8e8e8]"
            title="点击恢复自适应"
            onClick={() => setScale(0)}
            data-testid="zoom-level"
          >
            {percent ? `${percent}%` : '自适应'}
          </button>
          <button className="rounded p-1 hover:bg-[#e8e8e8]" onClick={() => setScale((canvasScale > 0 ? canvasScale : 1) + 0.1)}>
            <ZoomIn size={14} />
          </button>
        </div>
      </div>
    </div>
  )
}
