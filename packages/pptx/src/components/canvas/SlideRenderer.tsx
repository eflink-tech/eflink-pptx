// 幻灯片渲染：背景 + 元素集合（编辑/缩略图/放映共用）
// scale<=0 时表示由外层容器自适应（此组件不做 transform）
import { memo } from 'react'
import type { Background, Slide } from '../../types/slides'
import { ElementRenderer } from './ElementRenderer'

function backgroundStyle(bg: Background | undefined): React.CSSProperties {
  if (!bg) return { background: '#fff' }
  if (bg.type === 'solid') return { background: bg.color ?? '#fff' }
  if (bg.type === 'gradient' && bg.gradient) {
    const g = bg.gradient
    const stops = g.colors.map((c) => `${c.color} ${c.pos}%`).join(', ')
    if (g.type === 'radial') return { background: `radial-gradient(circle, ${stops})` }
    return { background: `linear-gradient(${g.rotate ?? 90}deg, ${stops})` }
  }
  if (bg.type === 'image' && bg.image) {
    if (bg.image.size === 'repeat') {
      return { backgroundImage: `url(${bg.image.src})`, backgroundRepeat: 'repeat', backgroundSize: 'auto' }
    }
    return {
      backgroundImage: `url(${bg.image.src})`,
      backgroundSize: bg.image.size === 'contain' ? 'contain' : 'cover',
      backgroundPosition: 'center',
      backgroundRepeat: 'no-repeat',
    }
  }
  return { background: '#fff' }
}

interface Props {
  slide: Slide
  width: number
  height: number
  /** 静态模式（缩略图/放映）：隐藏交互痕迹 */
  staticMode?: boolean
  /** 正在就地编辑的元素 id（仅编辑画布） */
  editingId?: string | null
  /** 文本就地编辑内容回调 */
  onEditChange?: (elementId: string, html: string) => void
  /** 正在就地编辑的表格单元格 id（格式 "tableId:r:c"） */
  editingCellId?: string | null
  /** 单元格点击回调 */
  onCellClick?: (cellId: string) => void
  /** 单元格失焦回调 */
  onCellBlur?: (cellId: string, text: string) => void
  /** 双击元素回调 */
  onElementDoubleClick?: (el: import('../../types/slides').PPTElement) => void
  className?: string
}

export const SlideRenderer = memo(function SlideRenderer({
  slide, width, height, staticMode, editingId, onEditChange, editingCellId, onCellClick, onCellBlur, onElementDoubleClick, className,
}: Props) {
  return (
    <div
      className={`relative overflow-hidden ${staticMode ? 'pointer-events-none' : ''} ${className ?? ''}`}
      style={{ width, height, ...backgroundStyle(slide.background) }}
    >
      {slide.elements.map((el) => (
        <ElementRenderer
          key={el.id}
          el={el}
          editing={editingId === el.id && (el.type === 'text' || el.type === 'shape')}
          onEditChange={
            editingId === el.id && onEditChange
              ? (html) => onEditChange(el.id, html)
              : undefined
          }
          editingCellId={editingCellId}
          onCellClick={onCellClick}
          onCellBlur={onCellBlur}
          onDoubleClick={onElementDoubleClick}
        />
      ))}
    </div>
  )
})
