// 左侧幻灯片缩略图面板：选择/排序/右键菜单
import { useState } from 'react'
import { useEditorStore } from '../../store/editorStore'
import { SlideRenderer } from '../canvas/SlideRenderer'

const THUMB_WIDTH = 168

export function ThumbnailPanel() {
  const presentation = useEditorStore((s) => s.presentation)
  const slideIndex = useEditorStore((s) => s.slideIndex)
  const [menu, setMenu] = useState<{ index: number; x: number; y: number } | null>(null)
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  /** 当前拖拽悬停的目标索引（用于显示插入指示线） */
  const [dropIndex, setDropIndex] = useState<number | null>(null)

  const ratio = presentation.viewportRatio
  const thumbH = Math.round(THUMB_WIDTH / ratio)

  /**
   * 计算放置下后的目标索引。
   * 由于 moveSlide 先 remove 再 insert，从前往后拖时需补偿：
   * - 向后拖（dragIndex < target）：insert 到 target 位置即最终位置
   * - 向前拖（dragIndex > target）：insert 到 target 位置即最终位置
   * 所以统一用 target 作为 insertIndex 即可（moveSlide 内部 splice 顺序已处理）
   */
  const handleDrop = (targetIndex: number) => {
    if (dragIndex === null || dragIndex === targetIndex) return
    const store = useEditorStore.getState()
    store.pushHistory()
    store.moveSlide(dragIndex, targetIndex)
    setDragIndex(null)
    setDropIndex(null)
  }

  /** 判断某个缩略图上方是否应显示插入指示线 */
  const showTopIndicator = (index: number) =>
    dropIndex !== null && dragIndex !== null && dropIndex === index && dragIndex > index

  /** 判断某个缩略图下方是否应显示插入指示线 */
  const showBottomIndicator = (index: number) =>
    dropIndex !== null && dragIndex !== null && dropIndex === index && dragIndex < index

  return (
    <div className="flex h-full w-[200px] shrink-0 flex-col border-r border-gray-200 bg-white" data-testid="thumbnail-panel">
      <div className="flex items-center justify-between px-3 py-2">
        <span className="text-xs text-gray-500">幻灯片 {presentation.slides.length}</span>
        <button
          className="rounded p-1 text-gray-500 hover:bg-gray-100"
          title="新建幻灯片"
          data-testid="add-slide"
          onClick={() => {
            const store = useEditorStore.getState()
            store.pushHistory()
            store.addSlide(store.slideIndex + 1)
          }}
        >
          ＋
        </button>
      </div>
      <div className="flex-1 space-y-2 overflow-y-auto pb-4">
        {presentation.slides.map((slide, index) => (
          <div key={slide.id} className="relative flex items-center gap-1.5">
            {/* 拖拽插入指示线（上） */}
            {showTopIndicator(index) && (
              <div className="absolute -top-1.5 left-7 right-0 z-10 h-0.5 rounded-full bg-[#d14424]" />
            )}
            <span className="w-4 shrink-0 text-right text-[10px] text-gray-400">{index + 1}</span>
            <div
              className={`group relative flex-1 cursor-pointer rounded-md border-2 p-0.5 transition-colors ${
                index === slideIndex ? 'border-[#d14424]' : 'border-transparent hover:border-gray-300'
              } ${dragIndex === index ? 'opacity-40' : ''}`}
              draggable
              onDragStart={() => setDragIndex(index)}
              onDragOver={(e) => {
                e.preventDefault()
                setDropIndex(index)
              }}
              onDragLeave={() => setDropIndex(null)}
              onDrop={(e) => {
                e.preventDefault()
                handleDrop(index)
              }}
              onDragEnd={() => {
                setDragIndex(null)
                setDropIndex(null)
              }}
              onClick={() => useEditorStore.setState({ slideIndex: index, selectedIds: [], editingId: null })}
              onContextMenu={(e) => {
                e.preventDefault()
                setMenu({ index, x: e.clientX, y: e.clientY })
              }}
              data-testid={`thumb-${index}`}
            >
              <div className="overflow-hidden rounded bg-white shadow-sm">
                <div style={{ width: THUMB_WIDTH, height: thumbH }}>
                  <div style={{ transform: `scale(${THUMB_WIDTH / presentation.width})`, transformOrigin: 'top left' }}>
                    <SlideRenderer slide={slide} width={presentation.width} height={Math.round(presentation.width / ratio)} staticMode />
                  </div>
                </div>
              </div>
            </div>
            {/* 拖拽插入指示线（下） */}
            {showBottomIndicator(index) && (
              <div className="absolute -bottom-1.5 left-7 right-0 z-10 h-0.5 rounded-full bg-[#d14424]" />
            )}
          </div>
        ))}
      </div>

      {menu && (
        <>
          <div className="fixed inset-0 z-40" onClick={() => setMenu(null)} onContextMenu={(e) => { e.preventDefault(); setMenu(null) }} />
          <div
            className="fixed z-50 w-40 rounded-md border border-gray-200 bg-white py-1 text-sm shadow-lg"
            style={{ left: menu.x, top: menu.y }}
          >
            {[
              { label: '在此之后新建', action: () => { const s = useEditorStore.getState(); s.pushHistory(); s.addSlide(menu.index + 1) } },
              { label: '复制本页', action: () => { const s = useEditorStore.getState(); s.pushHistory(); s.copySlide(menu.index) } },
              { label: '上移', action: () => { if (menu.index > 0) { const s = useEditorStore.getState(); s.pushHistory(); s.moveSlide(menu.index, menu.index - 1) } } },
              { label: '下移', action: () => { const s = useEditorStore.getState(); if (menu.index < s.presentation.slides.length - 1) { s.pushHistory(); s.moveSlide(menu.index, menu.index + 1) } } },
              { label: '删除本页', action: () => { const s = useEditorStore.getState(); s.pushHistory(); s.deleteSlides([menu.index]) }, danger: true },
            ].map((item) => (
              <button
                key={item.label}
                className={`block w-full px-3 py-1.5 text-left hover:bg-gray-100 ${'danger' in item && item.danger ? 'text-red-500' : ''}`}
                onClick={() => { item.action(); setMenu(null) }}
              >
                {item.label}
              </button>
            ))}
          </div>
        </>
      )}
    </div>
  )
}
