// 元素静态渲染器：编辑画布 / 缩略图 / 放映 三态复用
// 元素按画布原始 px 尺寸渲染，整体缩放交给外层 SlideView 的 transform
import { memo, useEffect, useMemo, useRef, useState } from 'react'
import { renderFormulaMarkup } from '../../core/render/formula'
import type {
  AudioElement, FormulaElement, ImageElement, LineElement,
  PPTElement, ShapeElement, TableElement, TextElement, VideoElement,
} from '../../types/slides'
import { getShapePath } from '../../core/render/shape'
import { ChartRenderer } from './ChartRenderer'
import { RichTextEditor } from '../richtext/RichTextEditor'

/** 形状渐变填充的 SVG defs */
function ShapeGradient({ id, gradient }: { id: string; gradient: { type: string; colors: Array<{ pos: number; color: string }>; rotate?: number } }) {
  if (gradient.type === 'radial') {
    return (
      <radialGradient id={id}>
        {gradient.colors.map((c, i) => <stop key={i} offset={`${c.pos}%`} stopColor={c.color} />)}
      </radialGradient>
    )
  }
  const rad = ((gradient.rotate ?? 90) * Math.PI) / 180
  const x1 = 50 - Math.cos(rad) * 50
  const y1 = 50 - Math.sin(rad) * 50
  const x2 = 50 + Math.cos(rad) * 50
  const y2 = 50 + Math.sin(rad) * 50
  return (
    <linearGradient id={id} x1={`${x1}%`} y1={`${y1}%`} x2={`${x2}%`} y2={`${y2}%`}>
      {gradient.colors.map((c, i) => <stop key={i} offset={`${c.pos}%`} stopColor={c.color} />)}
    </linearGradient>
  )
}

function TextItem({
  el,
  editing,
  onChange,
}: {
  el: TextElement
  editing?: boolean
  onChange?: (html: string) => void
}) {
  const boxRef = useRef<HTMLDivElement>(null)
  const [shrink, setShrink] = useState(1)

  // 自动收缩：内容超高时按比例整体缩放（缩略图/放映态一致）；编辑中不收缩以免跳动
  useEffect(() => {
    if (editing || !el.autoSize) {
      setShrink(1)
      return
    }
    const node = boxRef.current
    if (!node) return
    const inner = node.scrollHeight
    const outer = node.clientHeight
    if (outer > 0 && inner > outer) setShrink(Math.max(0.3, outer / inner))
    else setShrink(1)
  }, [editing, el.autoSize, el.content, el.w, el.h, el.lineHeight, el.padding])

  const style: React.CSSProperties = {
    lineHeight: el.lineHeight ?? 1.5,
    letterSpacing: el.charSpace ? `${el.charSpace}px` : undefined,
    padding: el.padding ?? 8,
    WebkitTextStroke: el.textStroke ? `${el.textStroke.width}px ${el.textStroke.color}` : undefined,
    textShadow: el.shadow ? `${el.shadow.h}px ${el.shadow.v}px ${el.shadow.blur}px ${el.shadow.color}` : undefined,
    writingMode: el.vertical ? 'vertical-rl' : undefined,
    background: typeof el.fill === 'string' ? el.fill : undefined,
    border: el.outline ? `${el.outline.width}px ${el.outline.style} ${el.outline.color}` : undefined,
  }
  return (
    <div
      ref={boxRef}
      className={`size-full overflow-hidden break-words ${editing ? 'pptx-richtext-editor' : ''}`}
      style={style}
    >
      {editing && onChange ? (
        <RichTextEditor elementId={el.id} initialHTML={el.content} onChange={onChange} />
      ) : (
        <div
          className="pptx-text-html"
          style={{
            transform: el.autoSize && shrink < 1 ? `scale(${shrink})` : undefined,
            transformOrigin: 'top left',
            width: '100%',
          }}
          dangerouslySetInnerHTML={{ __html: el.content }}
        />
      )}
    </div>
  )
}

function ImageItem({ el }: { el: ImageElement }) {
  const shadow = el.shadow ? `${el.shadow.h}px ${el.shadow.v}px ${el.shadow.blur}px ${el.shadow.color}` : undefined
  // 裁剪：clip 为源图比例区域 [0,1]，内层 img 放大到裁剪区恰好填满元素框
  if (el.clip && !el.clipShape) {
    const { x, y, w: cw, h: ch } = el.clip
    return (
      <div
        className="relative size-full overflow-hidden"
        style={{ borderRadius: el.radius, boxShadow: shadow, border: el.outline ? `${el.outline.width}px ${el.outline.style} ${el.outline.color}` : undefined }}
      >
        <img
          src={el.src}
          draggable={false}
          className="absolute select-none object-fill"
          style={{
            width: `${100 / cw}%`,
            height: `${100 / ch}%`,
            left: `${(-x * 100) / cw}%`,
            top: `${(-y * 100) / ch}%`,
            transform: `scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})`,
          }}
          alt=""
        />
      </div>
    )
  }
  if (el.clipShape) {
    const meta = getShapePath(el.clipShape)
    return (
      <div className="relative size-full" style={{ filter: shadow ? `drop-shadow(${shadow})` : undefined }}>
        <svg viewBox="0 0 100 100" preserveAspectRatio="none" className="size-full" style={{ transform: `scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})` }}>
          <defs>
            <clipPath id={`clip-${el.id}`} clipPathUnits="objectBoundingBox">
              {/* objectBoundingBox 需要 0-1 坐标：嵌套 svg 不支持，改用 userSpaceOnUse 于 0-100 viewBox */}
            </clipPath>
            <clipPath id={`clipu-${el.id}`}>
              <path d={meta.path} clipRule={meta.evenodd ? 'evenodd' : undefined} fillRule={meta.evenodd ? 'evenodd' : undefined} />
            </clipPath>
          </defs>
          <image
            href={el.src}
            width="100"
            height="100"
            preserveAspectRatio="xMidYMid slice"
            clipPath={`url(#clipu-${el.id})`}
          />
        </svg>
      </div>
    )
  }
  return (
    <div className="size-full overflow-hidden" style={{ borderRadius: el.radius, boxShadow: shadow, border: el.outline ? `${el.outline.width}px ${el.outline.style} ${el.outline.color}` : undefined }}>
      <img
        src={el.src}
        draggable={false}
        className="size-full select-none object-fill"
        style={{ transform: `scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})` }}
        alt=""
      />
    </div>
  )
}

function shapeTextToHTML(text: string | undefined): string {
  if (!text) return '<p></p>'
  if (/<[a-z][\s\S]*>/i.test(text)) return text
  return `<p>${text}</p>`
}

function ShapeItem({
  el,
  editing,
  onChange,
}: {
  el: ShapeElement
  editing?: boolean
  onChange?: (html: string) => void
}) {
  const meta = getShapePath(el.shapeKey)
  const gradientId = `sg-${el.id}`
  const isGradient = typeof el.fill === 'object' && el.fill !== null
  const fill = isGradient ? `url(#${gradientId})` : typeof el.fill === 'string' ? el.fill : 'none'
  const strokeScale = el.outline && el.outline.width > 0 ? (100 / Math.sqrt(el.w * el.h || 1)) : 0
  const align = el.align ?? 'center'
  const valign = el.valign ?? 'middle'
  const textLayerStyle: React.CSSProperties = {
    justifyContent: align === 'left' ? 'flex-start' : align === 'right' ? 'flex-end' : 'center',
    alignItems: valign === 'top' ? 'flex-start' : valign === 'bottom' ? 'flex-end' : 'center',
    color: el.defaultColor ?? '#fff',
    fontSize: el.fontSize ?? 18,
    fontWeight: el.fontWeight ?? 'normal',
    lineHeight: el.lineHeight ?? 1.2,
    fontFamily: el.defaultFontName,
  }
  return (
    <div
      className="relative size-full"
      style={{ filter: el.shadow ? `drop-shadow(${el.shadow.h}px ${el.shadow.v}px ${el.shadow.blur}px ${el.shadow.color})` : undefined }}
    >
      <svg
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        className="size-full"
        style={{ transform: `scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})`, overflow: 'visible' }}
      >
        <defs>
          {isGradient && <ShapeGradient id={gradientId} gradient={el.fill as never} />}
        </defs>
        <path
          d={meta.path}
          fill={fill}
          fillRule={meta.evenodd ? 'evenodd' : undefined}
          stroke={el.outline && el.outline.width > 0 ? el.outline.color : undefined}
          strokeWidth={strokeScale * (el.outline?.width ?? 0)}
          strokeDasharray={el.outline?.style === 'dashed' ? '8 4' : el.outline?.style === 'dotted' ? '2 3' : undefined}
          strokeLinejoin="round"
        />
      </svg>
      {(el.text || editing) && (
        <div
          className={`absolute inset-0 flex overflow-hidden whitespace-pre-wrap break-words px-1 ${editing ? 'pptx-richtext-editor pptx-shape-text-editor cursor-text' : 'select-none'}`}
          style={textLayerStyle}
          data-align={align}
          data-valign={valign}
        >
          {editing && onChange ? (
            <RichTextEditor elementId={el.id} initialHTML={shapeTextToHTML(el.text)} onChange={onChange} />
          ) : (
            el.text
          )}
        </div>
      )}
    </div>
  )
}

function arrowMarkerPath(kind: string): string {
  switch (kind) {
    case 'arrow': return 'M2 2L12 7L2 12'
    case 'triangle': return 'M2 2L12 7L2 12Z'
    default: return ''
  }
}

function buildLinePath(el: LineElement): string {
  const all = [el.start, ...(el.points ?? []), el.end]
  if (el.lineType === 'curve' && all.length === 4) {
    const [s, c1, c2, e] = all
    return `M${s[0]} ${s[1]} C${c1[0]} ${c1[1]} ${c2[0]} ${c2[1]} ${e[0]} ${e[1]}`
  }
  if (el.lineType === 'curve' && all.length === 3) {
    const [s, c, e] = all
    return `M${s[0]} ${s[1]} Q${c[0]} ${c[1]} ${e[0]} ${e[1]}`
  }
  return all.map((p, i) => `${i ? 'L' : 'M'}${p[0]} ${p[1]}`).join(' ')
}

function LineItem({ el }: { el: LineElement }) {
  const w = Math.max(el.w, 1)
  const h = Math.max(el.h, 1)
  const d = buildLinePath(el)
  const dash = el.lineStyle === 'dashed' ? '12 8' : el.lineStyle === 'dotted' ? '2 6' : undefined
  const uid = `ln-${el.id}`
  return (
    <svg width={w} height={h} style={{ overflow: 'visible', display: 'block' }}>
      <defs>
        {[el.startArrow, el.endArrow].map((arrow, idx) => arrow && arrow !== 'none' && arrow !== 'dot'
          ? (
              <marker
                key={idx}
                id={`${uid}-${idx}`}
                markerWidth="12"
                markerHeight="12"
                refX="7"
                refY="7"
                orient="auto"
                markerUnits="userSpaceOnUse"
              >
                <path d={arrowMarkerPath(arrow)} fill="none" stroke={el.color} strokeWidth={el.lineWidth} strokeLinejoin="miter" />
              </marker>
            )
          : null)}
      </defs>
      <path
        d={d}
        fill="none"
        stroke={el.color}
        strokeWidth={el.lineWidth}
        strokeDasharray={dash}
        strokeLinecap="round"
        strokeLinejoin="round"
        markerStart={el.startArrow && el.startArrow !== 'none' && el.startArrow !== 'dot' ? `url(#${uid}-0)` : undefined}
        markerEnd={el.endArrow && el.endArrow !== 'none' && el.endArrow !== 'dot' ? `url(#${uid}-1)` : undefined}
      />
      {el.startArrow === 'dot' && <circle cx={el.start[0]} cy={el.start[1]} r={el.lineWidth + 2} fill={el.color} />}
      {el.endArrow === 'dot' && <circle cx={el.end[0]} cy={el.end[1]} r={el.lineWidth + 2} fill={el.color} />}
    </svg>
  )
}

function TableItem({
  el,
  editingCellId,
  onCellClick,
  onCellBlur,
}: {
  el: TableElement
  editingCellId?: string | null
  onCellClick?: (cellId: string) => void
  onCellBlur?: (cellId: string, text: string) => void
}) {
  return (
    <table className="h-full w-full table-fixed border-collapse" style={{ border: `${el.outline?.width ?? 1}px ${el.outline?.style ?? 'solid'} ${el.outline?.color ?? '#ccc'}` }}>
      <colgroup>
        {el.colSizes.map((cw, i) => <col key={i} style={{ width: `${cw * 100}%` }} />)}
      </colgroup>
      <tbody>
        {el.cells.map((row, r) => {
          if (!row) return null
          return (
            <tr key={r} style={{ height: `${(el.rowSizes[r] ?? 1 / el.rowSizes.length) * 100}%` }}>
              {row.map((cell, c) => {
                if (!cell) return null
                const st = cell.style
                const cellId = `${el.id}:${r}:${c}`
                const isEditing = editingCellId === cellId
                return (
                  <td
                    key={c}
                    colSpan={cell.colspan && cell.colspan > 1 ? cell.colspan : undefined}
                    rowSpan={cell.rowspan && cell.rowspan > 1 ? cell.rowspan : undefined}
                    className="relative border border-[#e0e0e0] px-2 align-middle"
                    style={{
                      color: st?.color, background: st?.backcolor,
                      fontSize: st?.fontsize, fontFamily: st?.fontface,
                      fontWeight: st?.bold ? 700 : undefined,
                      fontStyle: st?.italic ? 'italic' : undefined,
                      textDecoration: st?.underline ? 'underline' : undefined,
                      textAlign: st?.align, verticalAlign: st?.valign,
                      cursor: isEditing ? 'text' : 'default',
                    }}
                    onClick={(e) => {
                      e.stopPropagation()
                      onCellClick?.(cellId)
                    }}
                  >
                    {isEditing ? (
                      <div
                        className="pptx-table-cell-editor w-full h-full outline-none"
                        contentEditable
                        suppressContentEditableWarning
                        autoFocus
                        style={{
                          color: st?.color ?? '#333',
                          background: 'transparent',
                          fontSize: st?.fontsize ?? 14,
                          fontFamily: st?.fontface,
                          fontWeight: st?.bold ? 700 : undefined,
                          fontStyle: st?.italic ? 'italic' : undefined,
                          textDecoration: st?.underline ? 'underline' : undefined,
                          textAlign: st?.align ?? 'left',
                        }}
                        onBlur={(e) => {
                          onCellBlur?.(cellId, e.currentTarget.textContent ?? '')
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Escape') {
                            e.currentTarget.blur()
                          }
                        }}
                      >
                        {cell.text}
                      </div>
                    ) : (
                      cell.text
                    )}
                  </td>
                )
              })}
            </tr>
          )
        })}
      </tbody>
    </table>
  )
}

function VideoItem({ el }: { el: VideoElement }) {
  return (
    <video
      src={el.src}
      poster={el.poster}
      loop={el.loop}
      autoPlay={el.autoPlay}
      controls
      className="size-full bg-black object-contain"
    />
  )
}

function AudioItem({ el }: { el: AudioElement }) {
  return (
    <div className="flex size-full items-center justify-center rounded-full bg-[#f1f3f5] shadow-md">
      <audio src={el.src} loop={el.loop} autoPlay={el.autoPlay} controls className="w-[85%]" />
    </div>
  )
}

function FormulaItem({ el }: { el: FormulaElement }) {
  const html = useMemo(() => renderFormulaMarkup(el.latex), [el.latex])
  return (
    <div
      className="pptx-formula-display flex size-full items-center justify-center overflow-hidden"
      style={{ color: el.color, fontSize: el.fontSize }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  )
}

export const ElementRenderer = memo(function ElementRenderer({
  el,
  editing,
  onEditChange,
  editingCellId,
  onCellClick,
  onCellBlur,
  onDoubleClick,
}: {
  el: PPTElement
  /** 文本框 / 形状内文本就地编辑（与静态渲染共用同一容器，保证所见即所得） */
  editing?: boolean
  onEditChange?: (html: string) => void
  /** 表格单元格就地编辑 */
  editingCellId?: string | null
  onCellClick?: (cellId: string) => void
  onCellBlur?: (cellId: string, text: string) => void
  /** 双击元素回调 */
  onDoubleClick?: (el: PPTElement) => void
}) {
  const base: React.CSSProperties = {
    position: 'absolute',
    left: el.x,
    top: el.y,
    width: el.w,
    height: el.h,
    transform: el.rotate ? `rotate(${el.rotate}deg)` : undefined,
    opacity: el.opacity ?? 1,
  }
  return (
    <div
      className={`pptx-element ${editing ? 'cursor-text' : el.lock ? 'cursor-default' : 'cursor-move'}`}
      style={base}
      data-el-id={el.id}
      onDoubleClick={(e) => {
        e.stopPropagation()
        onDoubleClick?.(el)
      }}
    >
      {el.type === 'text' && <TextItem el={el} editing={editing} onChange={onEditChange} />}
      {el.type === 'image' && <ImageItem el={el} />}
      {el.type === 'shape' && <ShapeItem el={el} editing={editing} onChange={onEditChange} />}
      {el.type === 'line' && <LineItem el={el} />}
      {el.type === 'table' && <TableItem el={el} editingCellId={editingCellId} onCellClick={onCellClick} onCellBlur={onCellBlur} />}
      {el.type === 'chart' && <ChartRenderer el={el} />}
      {el.type === 'video' && <VideoItem el={el} />}
      {el.type === 'audio' && <AudioItem el={el} />}
      {el.type === 'formula' && <FormulaItem el={el} />}
    </div>
  )
})
