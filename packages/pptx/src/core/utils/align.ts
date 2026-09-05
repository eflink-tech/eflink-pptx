// 对齐 / 分布 / 吸附参考线（纯函数）
import type { PPTElement } from '../../types/slides'
import type { Rect } from './geometry'

export type AlignMode =
  | 'left' | 'hcenter' | 'right'
  | 'top' | 'vcenter' | 'bottom'

export type DistributeMode = 'horizontal' | 'vertical'

export interface GuideLine {
  /** 'v' 垂直线（x 定位） | 'h' 水平线（y 定位） */
  orient: 'v' | 'h'
  pos: number
  /** 参考线起止（画布坐标） */
  from: number
  to: number
}

export interface SnapResult {
  dx: number
  dy: number
  guides: GuideLine[]
}

/** 收集所有吸附目标边（元素 + 画布中线/边线） */
export function collectSnapTargets(
  elements: PPTElement[],
  excludeIds: string[],
  slideW: number,
  slideH: number,
): Array<{ x: number[]; y: number[]; rect: Rect }> {
  const targets: Array<{ x: number[]; y: number[]; rect: Rect }> = []
  for (const el of elements) {
    if (excludeIds.includes(el.id)) continue
    const rect = { x: el.x, y: el.y, w: el.w, h: el.h }
    targets.push({
      x: [rect.x, rect.x + rect.w / 2, rect.x + rect.w],
      y: [rect.y, rect.y + rect.h / 2, rect.y + rect.h],
      rect,
    })
  }
  targets.push({
    x: [0, slideW / 2, slideW],
    y: [0, slideH / 2, slideH],
    rect: { x: 0, y: 0, w: slideW, h: slideH },
  })
  return targets
}

/** 拖拽吸附：bounds 为拖拽前包围盒，返回偏移修正与参考线 */
export function calcSnap(
  bounds: Rect,
  targets: Array<{ x: number[]; y: number[]; rect: Rect }>,
  threshold = 5,
): SnapResult {
  const selfX = [bounds.x, bounds.x + bounds.w / 2, bounds.x + bounds.w]
  const selfY = [bounds.y, bounds.y + bounds.h / 2, bounds.y + bounds.h]

  let bestDx = 0
  let bestDy = 0
  let bestDistX = threshold + 1
  let bestDistY = threshold + 1
  let guideV: GuideLine | null = null
  let guideH: GuideLine | null = null

  for (const t of targets) {
    for (const sx of selfX) {
      for (const tx of t.x) {
        const d = tx - sx
        if (Math.abs(d) <= threshold && Math.abs(d) < Math.abs(bestDistX)) {
          bestDistX = d
          bestDx = d
          guideV = { orient: 'v', pos: tx, from: Math.min(bounds.y, t.rect.y) - 12, to: Math.max(bounds.y + bounds.h, t.rect.y + t.rect.h) + 12 }
        }
      }
    }
    for (const sy of selfY) {
      for (const ty of t.y) {
        const d = ty - sy
        if (Math.abs(d) <= threshold && Math.abs(d) < Math.abs(bestDistY)) {
          bestDistY = d
          bestDy = d
          guideH = { orient: 'h', pos: ty, from: Math.min(bounds.x, t.rect.x) - 12, to: Math.max(bounds.x + bounds.w, t.rect.x + t.rect.w) + 12 }
        }
      }
    }
  }

  const guides: GuideLine[] = []
  if (guideV && Math.abs(bestDistX) <= threshold) guides.push(guideV)
  if (guideH && Math.abs(bestDistY) <= threshold) guides.push(guideH)
  return { dx: bestDx, dy: bestDy, guides }
}

/** 对齐选中元素（参照整体包围盒；单选时参照画布） */
export function alignElements(
  elements: PPTElement[],
  ids: string[],
  mode: AlignMode,
  slideW: number,
  slideH: number,
): void {
  const picked = elements.filter((el) => ids.includes(el.id))
  if (!picked.length) return
  const multi = picked.length > 1
  const minX = multi ? Math.min(...picked.map((e) => e.x)) : 0
  const maxX = multi ? Math.max(...picked.map((e) => e.x + e.w)) : slideW
  const minY = multi ? Math.min(...picked.map((e) => e.y)) : 0
  const maxY = multi ? Math.max(...picked.map((e) => e.y + e.h)) : slideH

  for (const el of picked) {
    switch (mode) {
      case 'left': el.x = minX; break
      case 'hcenter': el.x = minX + (maxX - minX) / 2 - el.w / 2; break
      case 'right': el.x = maxX - el.w; break
      case 'top': el.y = minY; break
      case 'vcenter': el.y = minY + (maxY - minY) / 2 - el.h / 2; break
      case 'bottom': el.y = maxY - el.h; break
    }
  }
}

/** 等距分布（3 个及以上元素生效） */
export function distributeElements(
  elements: PPTElement[],
  ids: string[],
  mode: DistributeMode,
): void {
  const picked = elements.filter((el) => ids.includes(el.id))
  if (picked.length < 3) return
  const isH = mode === 'horizontal'
  const sorted = [...picked].sort((a, b) => (isH ? a.x - b.x : a.y - b.y))
  const first = sorted[0]
  const last = sorted[sorted.length - 1]
  const totalSize = sorted.reduce((acc, el) => acc + (isH ? el.w : el.h), 0)
  const span = isH
    ? (last.x + last.w) - first.x
    : (last.y + last.h) - first.y
  const gap = (span - totalSize) / (sorted.length - 1)
  let cursor = isH ? first.x : first.y
  for (const el of sorted) {
    if (isH) {
      el.x = cursor
      cursor += el.w + gap
    } else {
      el.y = cursor
      cursor += el.h + gap
    }
  }
}
