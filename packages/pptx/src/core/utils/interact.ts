// 交互辅助：线段端点缩放、多选元素按包围盒映射
import type { LineElement, PPTElement } from '../../types/slides'
import type { Rect } from './geometry'

/** 线条元素尺寸变化时，按比例缩放端点/控制点 */
export function scaleLinePoints(el: LineElement, ow: number, oh: number, nw: number, nh: number): void {
  const kx = ow === 0 ? 1 : nw / ow
  const ky = oh === 0 ? 1 : nh / oh
  el.start = [el.start[0] * kx, el.start[1] * ky]
  el.end = [el.end[0] * kx, el.end[1] * ky]
  if (el.points) el.points = el.points.map((p) => [p[0] * kx, p[1] * ky] as [number, number])
}

/** 多选缩放：把每个元素按旧包围盒→新包围盒线性映射 */
export function mapElementsToBounds(elements: PPTElement[], ob: Rect, nb: Rect): void {
  const kx = ob.w === 0 ? 1 : nb.w / ob.w
  const ky = ob.h === 0 ? 1 : nb.h / ob.h
  for (const el of elements) {
    el.x = nb.x + (el.x - ob.x) * kx
    el.y = nb.y + (el.y - ob.y) * ky
    const nw = el.w * kx
    const nh = el.h * ky
    if (el.type === 'line') scaleLinePoints(el, el.w, el.h, nw, nh)
    el.w = nw
    el.h = nh
  }
}
