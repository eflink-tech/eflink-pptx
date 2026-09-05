// 底部状态栏：页码导航 / 缩放 / 网格
import { ChevronLeft, ChevronRight, ZoomIn, ZoomOut } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'

export function BottomBar() {
  const slideCount = useEditorStore((s) => s.presentation.slides.length)
  const slideIndex = useEditorStore((s) => s.slideIndex)
  const canvasScale = useUIStore((s) => s.canvasScale)
  const gridVisible = useUIStore((s) => s.gridVisible)

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
    <div className="flex h-8 shrink-0 items-center justify-between border-t border-gray-200 bg-white px-3 text-xs text-gray-500">
      <div className="flex items-center gap-1">
        <button className="rounded p-1 hover:bg-gray-100 disabled:opacity-40" disabled={slideIndex === 0} onClick={() => goto(slideIndex - 1)}>
          <ChevronLeft size={14} />
        </button>
        <span data-testid="slide-nav">{slideIndex + 1} / {slideCount}</span>
        <button className="rounded p-1 hover:bg-gray-100 disabled:opacity-40" disabled={slideIndex === slideCount - 1} onClick={() => goto(slideIndex + 1)}>
          <ChevronRight size={14} />
        </button>
      </div>

      <div className="flex items-center gap-2">
        <button className={`rounded px-2 py-0.5 ${gridVisible ? 'bg-[#fbeae5] text-[#d14424]' : 'hover:bg-gray-100'}`} onClick={() => useUIStore.getState().toggleGrid()}>
          网格
        </button>
        <div className="flex items-center gap-1">
          <button className="rounded p-1 hover:bg-gray-100" onClick={() => setScale((canvasScale > 0 ? canvasScale : 1) - 0.1)}>
            <ZoomOut size={14} />
          </button>
          <button
            className="w-14 rounded px-1 py-0.5 text-center hover:bg-gray-100"
            title="点击恢复自适应"
            onClick={() => setScale(0)}
            data-testid="zoom-level"
          >
            {percent ? `${percent}%` : '自适应'}
          </button>
          <button className="rounded p-1 hover:bg-gray-100" onClick={() => setScale((canvasScale > 0 ? canvasScale : 1) + 0.1)}>
            <ZoomIn size={14} />
          </button>
        </div>
      </div>
    </div>
  )
}
