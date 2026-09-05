/** 几何工具：旋转、包围盒、对齐吸附 */

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

export interface Point {
  x: number
  y: number
}

/** 将元素本地坐标点（相对未旋转包围盒左上角）旋转到画布坐标 */
export function rotatePoint(cx: number, cy: number, x: number, y: number, deg: number): Point {
  const rad = (deg * Math.PI) / 180
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  const dx = x - cx
  const dy = y - cy
  return {
    x: cx + dx * cos - dy * sin,
    y: cy + dx * sin + dy * cos,
  }
}

/** 旋转矩形的正交包围盒（画布坐标） */
export function rotatedRectBounds(rect: Rect, rotate = 0): Rect {
  if (!rotate) return { ...rect }
  const cx = rect.x + rect.w / 2
  const cy = rect.y + rect.h / 2
  const corners = [
    rotatePoint(cx, cy, rect.x, rect.y, rotate),
    rotatePoint(cx, cy, rect.x + rect.w, rect.y, rotate),
    rotatePoint(cx, cy, rect.x, rect.y + rect.h, rotate),
    rotatePoint(cx, cy, rect.x + rect.w, rect.y + rect.h, rotate),
  ]
  const xs = corners.map((p) => p.x)
  const ys = corners.map((p) => p.y)
  const minX = Math.min(...xs)
  const minY = Math.min(...ys)
  return { x: minX, y: minY, w: Math.max(...xs) - minX, h: Math.max(...ys) - minY }
}

/** 多个矩形的整体包围盒 */
export function unionBounds(rects: Rect[]): Rect {
  const minX = Math.min(...rects.map((r) => r.x))
  const minY = Math.min(...rects.map((r) => r.y))
  const maxX = Math.max(...rects.map((r) => r.x + r.w))
  const maxY = Math.max(...rects.map((r) => r.y + r.h))
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY }
}

export const deg2rad = (deg: number) => (deg * Math.PI) / 180

/* ---------- 缩放（支持旋转） ---------- */

export type HandleDir = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

const HANDLE_S: Record<HandleDir, [number, number]> = {
  nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0],
  se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0],
}

const MIN_SIZE = 10

/**
 * 计算缩放后的新矩形（x/y/w/h 均为画布坐标，rect 可带旋转）。
 * dx/dy 为画布坐标下的鼠标位移；keepRatio 时按宽高比缩放。
 */
export function applyResize(
  rect: Rect,
  rotate: number,
  handle: HandleDir,
  dx: number,
  dy: number,
  keepRatio = false,
): Rect {
  const [sx, sy] = HANDLE_S[handle]
  const rad = deg2rad(rotate)
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)

  // 画布位移 → 元素本地位移（逆旋转）
  const ldx = dx * cos + dy * sin
  const ldy = -dx * sin + dy * cos

  let w2 = rect.w
  let h2 = rect.h
  if (sx !== 0) w2 = Math.max(MIN_SIZE, rect.w + ldx * sx)
  if (sy !== 0) h2 = Math.max(MIN_SIZE, rect.h + ldy * sy)

  if (keepRatio && (sx !== 0 || sy !== 0)) {
    const ratio = rect.w / rect.h
    if (sx !== 0 && sy !== 0) {
      h2 = Math.max(MIN_SIZE, w2 / ratio)
      w2 = h2 * ratio
    } else if (sx !== 0) {
      h2 = Math.max(MIN_SIZE, w2 / ratio)
    } else {
      w2 = Math.max(MIN_SIZE, h2 * ratio)
    }
  }

  // 固定边/角（移动边的反方向）在旧尺寸下的本地坐标
  const fx = -sx
  const fy = -sy
  const cx = rect.x + rect.w / 2
  const cy = rect.y + rect.h / 2
  // 固定点画布坐标（旧尺寸）
  const pfx = cx + (fx * rect.w / 2) * cos - (fy * rect.h / 2) * sin
  const pfy = cy + (fx * rect.w / 2) * sin + (fy * rect.h / 2) * cos
  // 新中心 = 固定点 − 旋转后的新固定偏移
  const ncx = pfx - (fx * w2 / 2) * cos + (fy * h2 / 2) * sin
  const ncy = pfy - (fx * w2 / 2) * sin - (fy * h2 / 2) * cos

  return { x: ncx - w2 / 2, y: ncy - h2 / 2, w: w2, h: h2 }
}
