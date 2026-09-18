// src/core/import/elements/frame-geom.ts
/** p:pic / p:graphicFrame 通用几何：xfrm off/ext → 目标画布 px 矩形 */
import { attr, directChild, firstDescendant } from '../xml'
import { mapX, mapY } from '../context'
import type { GroupXform, ParseContext } from '../context'

export interface FrameGeom {
  x: number
  y: number
  w: number
  h: number
}

/** 解析节点 xfrm（graphicFrame 用 p:xfrm，pic 用 a:xfrm）并折算组合变换 + 画布缩放（EMU→px 96dpi） */
export function parseFrameGeom(node: Element, xf: GroupXform, ctx: ParseContext): FrameGeom | null {
  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!xfrm || !off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)
  return {
    x: Math.round(mapX(xf, ex) / 9525 * ctx.scale.x),
    y: Math.round(mapY(xf, ey) / 9525 * ctx.scale.y),
    w: Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x)),
    h: Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y)),
  }
}
