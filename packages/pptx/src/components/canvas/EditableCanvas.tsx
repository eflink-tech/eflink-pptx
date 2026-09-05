// 编辑画布：自适应缩放 + 选择/拖拽/缩放/旋转/框选/吸附
import { useCallback, useEffect, useRef, useState } from 'react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import type { LineElement } from '../../types/slides'
import { applyResize, unionBounds, type HandleDir } from '../../core/utils/geometry'
import { calcSnap, collectSnapTargets, type GuideLine } from '../../core/utils/align'
import { scaleLinePoints } from '../../core/utils/interact'
import {
  collectAllAnchors, lineLocalToWorld, resolveLineEndpoint, setLineEndpointWorld,
  setLineMidPointWorld, snapToAnchor, refreshAttachedLines,
} from '../../core/editor/lineLinker'

import { SlideRenderer } from './SlideRenderer'
import { TextFormatToolbar } from '../richtext/TextFormatToolbar'
import { CanvasContextMenu, type ContextMenuState } from './CanvasContextMenu'
import { createTextElement } from '../../core/schema/factory'
import { insertImageFile, isImageFile } from '../../core/editor/media'

interface OrigElement { id: string; x: number; y: number; w: number; h: number; rotate: number }

type Interaction =
  | { kind: 'move'; startX: number; startY: number; origs: OrigElement[] }
  | { kind: 'resize'; handle: HandleDir; startX: number; startY: number; origs: OrigElement[]; origBounds: { x: number; y: number; w: number; h: number } }
  | { kind: 'rotate'; startAngle: number; origRotate: number; ids: string[] }
  | { kind: 'linePoint'; id: string; index: number; origWorld: [number, number]; startX: number; startY: number }
  | { kind: 'lineEndpoint'; id: string; endpoint: 'start' | 'end'; startX: number; startY: number }
  | { kind: 'box'; startX: number; startY: number; current: { x: number; y: number } }
  | null

/** 鼠标事件转画布坐标 */
function toCanvasPoint(e: MouseEvent | React.MouseEvent, rect: DOMRect, scale: number): { x: number; y: number } {
  return { x: (e.clientX - rect.left) / scale, y: (e.clientY - rect.top) / scale }
}

export function EditableCanvas() {
  const containerRef = useRef<HTMLDivElement>(null)
  const slideWrapRef = useRef<HTMLDivElement>(null)
  const interactionRef = useRef<Interaction>(null)
  const [guides, setGuides] = useState<GuideLine[]>([])
  const [box, setBox] = useState<{ x: number; y: number; w: number; h: number } | null>(null)
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [lineAnchorHint, setLineAnchorHint] = useState(false)

  const presentation = useEditorStore((s) => s.presentation)
  const slideIndex = useEditorStore((s) => s.slideIndex)
  const selectedIds = useEditorStore((s) => s.selectedIds)
  const editingId = useEditorStore((s) => s.editingId)
  const editingCellId = useEditorStore((s) => s.editingCellId)
  const canvasScaleCfg = useUIStore((s) => s.canvasScale)
  const gridVisible = useUIStore((s) => s.gridVisible)
  const findHitIds = useUIStore((s) => s.findHitIds)

  const slide = presentation.slides[slideIndex]
  const slideW = presentation.width
  const slideH = Math.round(presentation.width / presentation.viewportRatio)

  /* 自适应缩放 */
  const [fitScale, setFitScale] = useState(1)
  useEffect(() => {
    const container = containerRef.current
    if (!container) return
    const observer = new ResizeObserver(() => {
      const pad = 80
      const w = container.clientWidth - pad
      const h = container.clientHeight - pad
      if (w <= 0 || h <= 0) return
      setFitScale(Math.min(w / slideW, h / slideH))
    })
    observer.observe(container)
    return () => observer.disconnect()
  }, [slideW, slideH])

  const scale = canvasScaleCfg > 0 ? canvasScaleCfg : fitScale

  /* ---------- 交互 ---------- */

  const getCanvasRect = useCallback(() => {
    const el = slideWrapRef.current
    return el ? el.getBoundingClientRect() : null
  }, [])

  /** 命中检测（含组）：返回应选中的 id 集合 */
  const hitElementId = useCallback((target: EventTarget | null): string | null => {
    const node = (target as HTMLElement | null)?.closest?.('.pptx-element') as HTMLElement | null
    if (!node) return null
    const id = node.dataset.elId ?? null
    return id
  }, [])

  const startMove = useCallback((e: React.MouseEvent, ids: string[]) => {
    const store = useEditorStore.getState()
    const slideNow = store.presentation.slides[store.slideIndex]
    const origs = slideNow.elements
      .filter((el) => ids.includes(el.id))
      .map((el) => ({ id: el.id, x: el.x, y: el.y, w: el.w, h: el.h, rotate: el.rotate ?? 0 }))
    if (!origs.length) return
    store.pushHistory()
    document.body.style.cursor = 'move'
    interactionRef.current = { kind: 'move', startX: e.clientX, startY: e.clientY, origs }
  }, [])

  const onMouseDown = useCallback((e: React.MouseEvent) => {
    if (e.button !== 0) return
    const store = useEditorStore.getState()
    // 正在富文本编辑时：点在编辑区内交给 TipTap；点外部则退出编辑
    if (store.editingId) {
      const node = (e.target as HTMLElement).closest('.pptx-richtext-editor')
      if (node) return
      store.setEditingId(null)
    }
    // 正在表格单元格编辑时：点在编辑区内交给 contentEditable；点外部则先保存再退出编辑
    if (store.editingCellId) {
      const editorNode = (e.target as HTMLElement).closest('.pptx-table-cell-editor')
      if (editorNode) return
      // mousedown 先于 blur 触发，需要在此手动保存内容
      const editor = document.querySelector('.pptx-table-cell-editor')
      if (editor) {
        const text = editor.textContent ?? ''
        store.updateTableCellText(store.editingCellId, text)
      }
      store.setEditingCellId(null)
    }

    const hitId = hitElementId(e.target)
    if (!hitId) {
      // 空白处：开始框选
      const rect = getCanvasRect()
      if (!rect) return
      const p = toCanvasPoint(e, rect, scale)
      if (!e.shiftKey) store.setSelected([])
      interactionRef.current = { kind: 'box', startX: p.x, startY: p.y, current: p }
      setBox({ x: p.x, y: p.y, w: 0, h: 0 })
      e.preventDefault()
      return
    }

    const slideNow = store.presentation.slides[store.slideIndex]
    const hitEl = slideNow.elements.find((el) => el.id === hitId)
    if (!hitEl) return
    if (hitEl.lock) return

    let ids: string[]
    if (e.shiftKey) {
      store.toggleSelected(hitEl.groupId ? slideNow.elements.filter((el) => el.groupId === hitEl.groupId).map((el) => el.id)[0] : hitId)
      return
    }
    if (store.selectedIds.includes(hitId)) {
      ids = store.selectedIds
    } else if (hitEl.groupId) {
      ids = slideNow.elements.filter((el) => el.groupId === hitEl.groupId).map((el) => el.id)
      store.setSelected(ids)
    } else {
      ids = [hitId]
      store.setSelected(ids)
    }
    // 阻止原生文字选择/拖拽（否则拖动期间 mousemove 停止派发，元素无法移动）
    e.preventDefault()
    startMove(e, ids)
  }, [getCanvasRect, hitElementId, scale, startMove])

  /* 缩放/旋转柄 */
  const startResize = useCallback((e: React.MouseEvent, handle: HandleDir) => {
    e.stopPropagation()
    e.preventDefault()
    const store = useEditorStore.getState()
    if (store.editingId) store.setEditingId(null)
    const slideNow = store.presentation.slides[store.slideIndex]
    const ids = store.selectedIds
    const origs = slideNow.elements.filter((el) => ids.includes(el.id)).map((el) => ({ id: el.id, x: el.x, y: el.y, w: el.w, h: el.h, rotate: el.rotate ?? 0 }))
    if (!origs.length) return
    const bounds = ids.length === 1
      ? { x: origs[0].x, y: origs[0].y, w: origs[0].w, h: origs[0].h }
      : unionBounds(origs)
    store.pushHistory()
    document.body.style.cursor = 'move'
    interactionRef.current = { kind: 'resize', handle, startX: e.clientX, startY: e.clientY, origs, origBounds: bounds }
  }, [])

  const startRotate = useCallback((e: React.MouseEvent) => {
    e.stopPropagation()
    e.preventDefault()
    const store = useEditorStore.getState()
    if (store.editingId) store.setEditingId(null)
    const slideNow = store.presentation.slides[store.slideIndex]
    const ids = store.selectedIds
    const first = slideNow.elements.find((el) => ids.includes(el.id))
    if (!first) return
    const rect = getCanvasRect()
    if (!rect) return
    const center = { x: first.x + first.w / 2, y: first.y + first.h / 2 }
    const p = toCanvasPoint(e, rect, scale)
    store.pushHistory()
    document.body.style.cursor = 'grabbing'
    interactionRef.current = {
      kind: 'rotate',
      startAngle: Math.atan2(p.y - center.y, p.x - center.x) * 180 / Math.PI,
      origRotate: first.rotate ?? 0,
      ids,
    }
  }, [getCanvasRect, scale])

  /** 拖拽线条控制点 */
  const startLinePointDrag = useCallback((e: React.MouseEvent, el: LineElement, index: number) => {
    e.stopPropagation()
    if (el.lock) return
    const orig = el.points?.[index]
    if (!orig) return
    useEditorStore.getState().pushHistory()
    setLineAnchorHint(true)
    interactionRef.current = {
      kind: 'linePoint', id: el.id, index,
      origWorld: lineLocalToWorld(el, orig),
      startX: e.clientX, startY: e.clientY,
    }
  }, [])

  const startLineEndpointDrag = useCallback((e: React.MouseEvent, el: LineElement, endpoint: 'start' | 'end') => {
    e.stopPropagation()
    e.preventDefault()
    if (el.lock) return
    useEditorStore.getState().pushHistory()
    setLineAnchorHint(true)
    interactionRef.current = { kind: 'lineEndpoint', id: el.id, endpoint, startX: e.clientX, startY: e.clientY }
  }, [])

  /* 粘贴图片文件 → 插入画布（编辑富文本时不拦截） */
  useEffect(() => {
    const onPaste = (e: ClipboardEvent) => {
      if (useEditorStore.getState().editingId) return
      const target = e.target as HTMLElement | null
      if (target?.closest?.('input, textarea, [contenteditable="true"]')) return
      const items = Array.from(e.clipboardData?.items ?? [])
      const imageItem = items.find((it) => it.kind === 'file' && it.type.startsWith('image/'))
      const file = imageItem?.getAsFile()
      if (file) {
        e.preventDefault()
        void insertImageFile(file)
      }
    }
    window.addEventListener('paste', onPaste)
    return () => window.removeEventListener('paste', onPaste)
  }, [])

  /* 全局 mousemove / mouseup */
  useEffect(() => {
    const onMouseMove = (e: MouseEvent) => {
      const it = interactionRef.current
      if (!it) return
      const store = useEditorStore.getState()
      const slideNow = store.presentation.slides[store.slideIndex]
      const rect = getCanvasRect()
      if (!rect) return

      if (it.kind === 'move') {
        let dx = (e.clientX - it.startX) / scale
        let dy = (e.clientY - it.startY) / scale
        let snapGuides: GuideLine[] = []
        if (!e.altKey) {
          const origBounds = unionBounds(it.origs)
          const targets = collectSnapTargets(slideNow.elements, it.origs.map((o) => o.id), slideW, slideH)
          const snap = calcSnap({ x: origBounds.x + dx, y: origBounds.y + dy, w: origBounds.w, h: origBounds.h }, targets)
          if (snap.guides.length) {
            dx += snap.dx
            dy += snap.dy
            snapGuides = snap.guides
          }
        }
        setGuides(snapGuides)
        const movedIds = it.origs.map((o) => o.id)
        store.updateElements(movedIds, (el) => {
          const orig = it.origs.find((o) => o.id === el.id)
          if (!orig) return
          el.x = Math.round(orig.x + dx)
          el.y = Math.round(orig.y + dy)
          if (el.type === 'line') {
            el.startAttach = undefined
            el.endAttach = undefined
          }
        }, { history: false })
        refreshAttachedLines(movedIds)
      } else if (it.kind === 'resize') {
        const dx = (e.clientX - it.startX) / scale
        const dy = (e.clientY - it.startY) / scale
        const keepRatio = e.shiftKey
        if (it.origs.length === 1) {
          const orig = it.origs[0]
          const el = slideNow.elements.find((el) => el.id === orig.id)
          if (!el) return
          const nb = applyResize({ x: orig.x, y: orig.y, w: orig.w, h: orig.h }, orig.rotate, it.handle, dx, dy, keepRatio)
          store.updateElements([orig.id], (target) => {
            if (target.type === 'line') scaleLinePoints(target as LineElement, orig.w, orig.h, nb.w, nb.h)
            target.x = nb.x
            target.y = nb.y
            target.w = Math.round(nb.w)
            target.h = Math.round(nb.h)
          }, { history: false })
          refreshAttachedLines([orig.id])
        } else {
          const nb = applyResize(it.origBounds, 0, it.handle, dx, dy, keepRatio)
          const kx = it.origBounds.w === 0 ? 1 : nb.w / it.origBounds.w
          const ky = it.origBounds.h === 0 ? 1 : nb.h / it.origBounds.h
          const resizedIds = it.origs.map((o) => o.id)
          store.updateElements(resizedIds, (target) => {
            const orig = it.origs.find((o) => o.id === target.id)
            if (!orig) return
            const nw = orig.w * kx
            const nh = orig.h * ky
            if (target.type === 'line') scaleLinePoints(target as LineElement, orig.w, orig.h, nw, nh)
            target.x = nb.x + (orig.x - it.origBounds.x) * kx
            target.y = nb.y + (orig.y - it.origBounds.y) * ky
            target.w = nw
            target.h = nh
          }, { history: false })
          refreshAttachedLines(resizedIds)
        }
      } else if (it.kind === 'rotate') {
        const first = slideNow.elements.find((el) => it.ids.includes(el.id))
        if (!first) return
        const center = { x: first.x + first.w / 2, y: first.y + first.h / 2 }
        const p = toCanvasPoint(e, rect, scale)
        let angle = it.origRotate + (Math.atan2(p.y - center.y, p.x - center.x) * 180 / Math.PI - it.startAngle)
        angle = ((angle % 360) + 360) % 360
        if (!e.altKey) {
          const snap15 = Math.round(angle / 15) * 15
          if (Math.abs(angle - snap15) < 4) angle = snap15
        }
        store.updateElements(it.ids, (el) => { el.rotate = Math.round(angle) }, { history: false })
        refreshAttachedLines(it.ids)
      } else if (it.kind === 'linePoint') {
        const p = toCanvasPoint(e, rect, scale)
        const wx = p.x
        const wy = p.y
        store.updateElements([it.id], (target) => {
          if (target.type !== 'line') return
          setLineMidPointWorld(target, it.index, [wx, wy], store.presentation.slides[store.slideIndex].elements)
        }, { history: false })
      } else if (it.kind === 'lineEndpoint') {
        const p = toCanvasPoint(e, rect, scale)
        const snap = snapToAnchor(p.x, p.y, slideNow.elements, [it.id])
        store.updateElements([it.id], (target) => {
          if (target.type !== 'line') return
          setLineEndpointWorld(target, it.endpoint, [snap.x, snap.y], snap.attach, store.presentation.slides[store.slideIndex].elements)
        }, { history: false })
      } else if (it.kind === 'box') {
        const p = toCanvasPoint(e, rect, scale)
        const b = {
          x: Math.min(it.startX, p.x), y: Math.min(it.startY, p.y),
          w: Math.abs(p.x - it.startX), h: Math.abs(p.y - it.startY),
        }
        setBox(b)
        const hits = slideNow.elements
          .filter((el) => !el.lock && el.x < b.x + b.w && el.x + el.w > b.x && el.y < b.y + b.h && el.y + el.h > b.y)
          .map((el) => el.id)
        store.setSelected(hits)
      }
    }

    const onMouseUp = () => {
      const it = interactionRef.current
      if (it?.kind === 'box') setBox(null)
      if (it?.kind === 'lineEndpoint' || it?.kind === 'linePoint') setLineAnchorHint(false)
      if (it?.kind === 'move' || it?.kind === 'resize' || it?.kind === 'rotate' || it?.kind === 'linePoint' || it?.kind === 'lineEndpoint') {
        document.body.style.cursor = ''
      }
      interactionRef.current = null
      setGuides([])
    }

    window.addEventListener('mousemove', onMouseMove)
    window.addEventListener('mouseup', onMouseUp)
    return () => {
      window.removeEventListener('mousemove', onMouseMove)
      window.removeEventListener('mouseup', onMouseUp)
    }
  }, [getCanvasRect, scale, slideH, slideW])

  if (!slide) return null

  const selectedEls = slide.elements.filter((el) => selectedIds.includes(el.id))
  const single = selectedEls.length === 1 ? selectedEls[0] : null

  return (
    <div
      ref={containerRef}
      className="relative flex size-full items-center justify-center overflow-hidden bg-[#f0f1f3]"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault()
        const files = Array.from(e.dataTransfer.files).filter(isImageFile)
        for (const file of files) void insertImageFile(file)
      }}
      onWheel={(e) => {
        if (!e.ctrlKey && !e.metaKey) return
        e.preventDefault()
        const cur = useUIStore.getState().canvasScale
        const base = cur > 0 ? cur : fitScale
        const next = Math.min(4, Math.max(0.1, base * (e.deltaY < 0 ? 1.1 : 0.9)))
        useUIStore.getState().setCanvasScale(next)
      }}
    >
      <div
        ref={slideWrapRef}
        className={`relative shadow-[0_4px_24px_rgba(0,0,0,0.14)] ${editingId ? '' : 'select-none'}`}
        data-testid="canvas-slide"
        style={{
          width: slideW,
          height: slideH,
          transform: `scale(${scale})`,
          transformOrigin: 'center center',
          backgroundImage: gridVisible
            ? 'linear-gradient(#e6e6e680 1px, transparent 1px), linear-gradient(90deg, #e6e6e680 1px, transparent 1px)'
            : undefined,
          backgroundSize: gridVisible ? '40px 40px' : undefined,
        }}
        onMouseDown={onMouseDown}
        onContextMenu={(e) => {
          e.preventDefault()
          if (useEditorStore.getState().editingId) useEditorStore.getState().setEditingId(null)
          setContextMenu({ x: e.clientX, y: e.clientY, elementId: hitElementId(e.target) })
        }}
        onDoubleClick={(e) => {
          const id = hitElementId(e.target)
          const store = useEditorStore.getState()
          interactionRef.current = null
          document.body.style.cursor = ''
          if (!id) {
            // 双击空白处：在该位置创建文本框并进入编辑
            const rect = getCanvasRect()
            if (!rect) return
            const p = toCanvasPoint(e, rect, scale)
            store.pushHistory()
            const el = createTextElement(Math.round(p.x - 100), Math.round(p.y - 20), 200, 40, '')
            store.addElement(el)
            store.setEditingId(el.id)
            return
          }
          const el = store.presentation.slides[store.slideIndex].elements.find((el) => el.id === id)
          if (!el) return
          if (el.type === 'text' || el.type === 'shape') {
            store.setEditingId(id)
          } else if (el.type === 'table') {
            // 双击表格：计算点击位置对应的单元格，进入就地编辑
            const tableEl = el as import('../../types/slides').TableElement
            const elNode = (e.target as HTMLElement).closest('[data-el-id]') as HTMLElement | null
            if (!elNode) return
            const rect = elNode.getBoundingClientRect()
            const clickX = e.clientX - rect.left
            const clickY = e.clientY - rect.top
            const w = rect.width
            const h = rect.height
            // 计算列索引
            let colX = 0
            let col = 0
            for (let i = 0; i < tableEl.colSizes.length; i++) {
              const colW = tableEl.colSizes[i] * w
              if (clickX < colX + colW) { col = i; break }
              colX += colW
            }
            // 计算行索引
            let rowY = 0
            let row = 0
            for (let i = 0; i < tableEl.rowSizes.length; i++) {
              const rowH = tableEl.rowSizes[i] * h
              if (clickY < rowY + rowH) { row = i; break }
              rowY += rowH
            }
            store.setEditingCellId(`${id}:${row}:${col}`)
          }
        }}
      >
        <SlideRenderer
          slide={slide}
          width={slideW}
          height={slideH}
          editingId={editingId}
          onEditChange={(id, html) => {
            useEditorStore.getState().updateElements([id], (target) => {
              if (target.type === 'text') target.content = html
              else if (target.type === 'shape') target.text = html.replace(/<[^>]+>/g, '')
            }, { history: false })
          }}
          editingCellId={editingCellId}
          onCellClick={(cellId) => {
            const store = useEditorStore.getState()
            // 单击单元格：选中表格，并进入编辑
            const parts = cellId.split(':')
            if (parts.length === 3) {
              const tableId = parts[0]
              store.setSelected([tableId])
              store.setEditingCellId(cellId)
            }
          }}
          onCellBlur={(cellId, text) => {
            useEditorStore.getState().updateTableCellText(cellId, text)
            useEditorStore.getState().setEditingCellId(null)
          }}
          onElementDoubleClick={(el) => {
            const store = useEditorStore.getState()
            interactionRef.current = null
            document.body.style.cursor = ''
            if (el.type === 'text' || el.type === 'shape') {
              store.setEditingId(el.id)
            } else if (el.type === 'table') {
              store.setSelected([el.id])
            }
          }}
        />


        {/* 吸附参考线（元素间对齐用虚线） */}
        {guides.map((g, i) => (
          <div
            key={i}
            className="pointer-events-none absolute"
            style={g.orient === 'v'
              ? { left: g.pos, top: g.from, width: 0, height: Math.max(0, g.to - g.from), borderLeft: '1px dashed #d14424' }
              : { left: g.from, top: g.pos, width: Math.max(0, g.to - g.from), height: 0, borderTop: '1px dashed #d14424' }}
          />
        ))}

        {/* 框选框 */}
        {box && (
          <div
            className="pointer-events-none absolute border border-[#d14424] bg-[#d1442411]"
            style={{ left: box.x, top: box.y, width: box.w, height: box.h }}
          />
        )}

        {/* 选中态描边：选中=实线框，编辑=虚线框；线条按端点范围显示 */}
        {single && (() => {
          if (single.type === 'line') {
            const ws = resolveLineEndpoint(single, 'start', slide.elements)
            const we = resolveLineEndpoint(single, 'end', slide.elements)
            const pad = 4
            const lx = Math.min(ws[0], we[0]) - pad
            const ly = Math.min(ws[1], we[1]) - pad
            const lw = Math.max(Math.abs(we[0] - ws[0]), 8) + pad * 2
            const lh = Math.max(Math.abs(we[1] - ws[1]), 8) + pad * 2
            return (
              <div
                className={`pointer-events-none absolute border border-[#d14424] ${single.id === editingId ? 'border-dashed' : 'border-solid'}`}
                style={{ left: lx, top: ly, width: lw, height: lh }}
              />
            )
          }
          return (
            <div
              className={`pointer-events-none absolute border border-[#d14424] ${single.id === editingId ? 'border-dashed' : 'border-solid'}`}
              style={{
                left: single.x, top: single.y, width: single.w, height: single.h,
                transform: single.rotate ? `rotate(${single.rotate}deg)` : undefined,
              }}
            />
          )
        })()}
        {selectedEls.length > 1 && selectedEls.map((el) => (
          <div
            key={`sel-${el.id}`}
            className="pointer-events-none absolute border border-[#d1442488]"
            style={{ left: el.x, top: el.y, width: el.w, height: el.h, transform: el.rotate ? `rotate(${el.rotate}deg)` : undefined }}
          />
        ))}

        {/* 操作柄：选中/编辑态均保留；线条使用专用端点柄 */}
        {single && single.type !== 'line' && (
          <div
            className="pointer-events-none absolute"
            style={{
              left: single.x,
              top: single.y,
              width: single.w,
              height: single.h,
              transform: single.rotate ? `rotate(${single.rotate}deg)` : undefined,
            }}
          >
            {!single.lock && (
              <>
                {(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as HandleDir[]).map((h) => {
                  const pos: React.CSSProperties = {
                    nw: { left: -5, top: -5, cursor: 'nwse-resize' },
                    n: { left: '50%', marginLeft: -5, top: -5, cursor: 'ns-resize' },
                    ne: { right: -5, top: -5, cursor: 'nesw-resize' },
                    e: { right: -5, top: '50%', marginTop: -5, cursor: 'ew-resize' },
                    se: { right: -5, bottom: -5, cursor: 'nwse-resize' },
                    s: { left: '50%', marginLeft: -5, bottom: -5, cursor: 'ns-resize' },
                    sw: { left: -5, bottom: -5, cursor: 'nesw-resize' },
                    w: { left: -5, top: '50%', marginTop: -5, cursor: 'ew-resize' },
                  }[h]
                  return (
                    <div
                      key={h}
                      className="element-handle pointer-events-auto z-20"
                      style={{ ...pos }}
                      onMouseDown={(e) => startResize(e, h)}
                    />
                  )
                })}
                {/* 旋转柄 */}
                <div
                  className="pointer-events-auto absolute left-1/2 z-20 flex h-5 w-5 cursor-grab items-center justify-center"
                  style={{ top: -34, marginLeft: -10 }}
                  onMouseDown={startRotate}
                  title="旋转"
                >
                  <div className="size-4 rounded-full border border-[#d14424] bg-white" />
                </div>
                <div className="absolute left-1/2 top-[-30px] h-[26px] w-px bg-[#d14424aa]" style={{ marginLeft: -0.5 }} />
              </>
            )}
          </div>
        )}

        {/* 线条端点 / 控制点 + 图形连接锚点 */}
        {single && single.type === 'line' && single.id !== editingId && (() => {
          const lineEl = single
          const ws = resolveLineEndpoint(lineEl, 'start', slide.elements)
          const we = resolveLineEndpoint(lineEl, 'end', slide.elements)
          return (
            <>
              {([
                ['start', ws, lineEl.startAttach] as const,
                ['end', we, lineEl.endAttach] as const,
              ]).map(([ep, pt, attach]) => (
                <div
                  key={`lep-${ep}`}
                  className={`absolute z-20 size-3 cursor-crosshair rounded-full border-2 bg-white ${attach ? 'border-[#22c55e]' : 'border-[#d14424]'}`}
                  style={{ left: pt[0] - 6, top: pt[1] - 6 }}
                  title={attach ? '已连接图形（拖拽可改挂点）' : '拖拽端点连接图形'}
                  onMouseDown={(e) => startLineEndpointDrag(e, lineEl, ep)}
                />
              ))}
              {(lineEl.points ?? []).map((p, i) => {
                const wp = lineLocalToWorld(lineEl, p)
                return (
                  <div
                    key={`lp-${i}`}
                    className="absolute z-10 size-2.5 cursor-move rounded-full border border-[#d14424] bg-white"
                    style={{ left: wp[0] - 5, top: wp[1] - 5 }}
                    title="拖拽调整控制点"
                    onMouseDown={(e) => startLinePointDrag(e, lineEl, i)}
                  />
                )
              })}
            </>
          )
        })()}

        {(lineAnchorHint || (single?.type === 'line')) && collectAllAnchors(slide.elements, single?.type === 'line' ? single.id : undefined).map((a) => (
          <div
            key={`anchor-${a.elementId}-${a.anchor}`}
            className={`pointer-events-none absolute z-[15] size-2 rounded-full border ${lineAnchorHint ? 'border-[#d14424] bg-[#fbeae5]' : 'border-[#d1442488] bg-white'}`}
            style={{ left: a.x - 4, top: a.y - 4 }}
          />
        ))}

        {/* 查找命中高亮 */}
        {slide.elements.map((el) => {
          if (!findHitIds.includes(el.id)) return null
          return (
            <div
              key={`hit-${el.id}`}
              className="pointer-events-none absolute border-2 border-[#f5c56b] bg-[#f5c56b22]"
              style={{ left: el.x, top: el.y, width: el.w, height: el.h }}
            />
          )
        })}
      </div>

      {/* 右键菜单必须在 scale 容器外：祖先 transform 会使 fixed 相对画布而非视口定位 */}
      {contextMenu && <CanvasContextMenu menu={contextMenu} onClose={() => setContextMenu(null)} />}

      {/* 富文本格式工具条（不随画布缩放） */}
      {editingId && <TextFormatToolbar />}

      {/* 缩放比例显示 */}
      <div className="pointer-events-none absolute bottom-2 right-3 text-xs text-gray-400">
        {Math.round(scale * 100)}%
      </div>
    </div>
  )
}
