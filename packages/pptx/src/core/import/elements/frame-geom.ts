// src/core/import/elements/frame-geom.ts
/** p:pic / p:graphicFrame 通用几何：xfrm off/ext → 目标画布 px 矩形 */
import { attr, directChild, firstDescendant } from '../xml'
import { mapPoint, xformScale } from '../context'
import type { GroupXform, ParseContext } from '../context'

export interface FrameGeom {
  x: number
  y: number
  w: number
  h: number
}

/** 元素自身的 rot/flip 与组合变换合成（模型只支持轴对齐盒 + 绕自身中心旋转） */
export interface FrameRot {
  /** 累计旋转角（度） */
  rot: number
  flipH: boolean
  flipV: boolean
}

export function frameRotOf(xfrm: Element | null, xf: Readonly<GroupXform>): FrameRot {
  const ownRot = xfrm ? parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000 : 0
  const ownFlipH = xfrm ? attr(xfrm, 'flipH') === '1' : false
  const ownFlipV = xfrm ? attr(xfrm, 'flipV') === '1' : false
  return {
    rot: ownRot + xf.rot,
    flipH: ownFlipH !== xf.flipH,
    flipV: ownFlipV !== xf.flipV,
  }
}

/** EMU 子空间矩形 → 画布 px：中心走完整仿射、宽高走轴向缩放（旋转后仍保持包围盒中心一致） */
export function mapEmuBox(
  xf: Readonly<GroupXform>,
  box: { x: number; y: number; w: number; h: number },
  scale: { x: number; y: number },
): FrameGeom {
  const [mx, my] = mapPoint(xf, box.x + box.w / 2, box.y + box.h / 2)
  const [sx, sy] = xformScale(xf)
  const w = box.w * sx
  const h = box.h * sy
  return {
    x: Math.round((mx - w / 2) / 9525 * scale.x),
    y: Math.round((my - h / 2) / 9525 * scale.y),
    w: Math.max(1, Math.round(w / 9525 * scale.x)),
    h: Math.max(1, Math.round(h / 9525 * scale.y)),
  }
}

/** 解析节点 xfrm（graphicFrame 用 p:xfrm，pic 用 a:xfrm）并折算组合变换 + 画布缩放（EMU→px 96dpi） */
export function parseFrameGeom(node: Element, xf: GroupXform, ctx: ParseContext): FrameGeom | null {
  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!xfrm || !off || !ext) return null
  return mapEmuBox(xf, {
    x: parseInt(attr(off, 'x') ?? '0', 10),
    y: parseInt(attr(off, 'y') ?? '0', 10),
    w: parseInt(attr(ext, 'cx') ?? '0', 10),
    h: parseInt(attr(ext, 'cy') ?? '0', 10),
  }, ctx.scale)
}
