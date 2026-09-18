// src/core/import/elements/group.ts
/** p:grpSp 递归展开：组合变换（off/ext vs chOff/chExt）折算到子元素 */
import { attr, directChild, firstDescendant } from '../xml'
import type { GroupXform } from '../context'
import type { PPTElement } from '../../../types/slides'
import { parseSpTreeNode } from './index'
import type { ParseContext } from '../context'
import type { PptxPackage } from '../package'

/** 由 grpSp 的 xfrm 计算子空间 → 当前空间的复合变换 */
export function childXform(grpSp: Element, xf: GroupXform): GroupXform {
  const xfrm = firstDescendant(grpSp, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const chOff = xfrm ? directChild(xfrm, 'a:chOff') : null
  const chExt = xfrm ? directChild(xfrm, 'a:chExt') : null
  if (!off || !ext || !chOff || !chExt) return xf
  const offX = parseInt(attr(off, 'x') ?? '0', 10)
  const offY = parseInt(attr(off, 'y') ?? '0', 10)
  const chX = parseInt(attr(chOff, 'x') ?? '0', 10)
  const chY = parseInt(attr(chOff, 'y') ?? '0', 10)
  const cx = parseInt(attr(ext, 'cx') ?? '0', 10)
  const cy = parseInt(attr(ext, 'cy') ?? '0', 10)
  const chCx = parseInt(attr(chExt, 'cx') ?? '0', 10) || 1
  const chCy = parseInt(attr(chExt, 'cy') ?? '0', 10) || 1
  const sx = chCx ? cx / chCx : 1
  const sy = chCy ? cy / chCy : 1
  return {
    // 复合：outer( inner(v) ) = xf.ox + (offX + (v - chX) * s) * xf.sx
    ox: xf.ox + (offX - chX * sx) * xf.sx,
    oy: xf.oy + (offY - chY * sy) * xf.sy,
    sx: xf.sx * sx,
    sy: xf.sy * sy,
    // rot 保持六万分之一度原始值累加，由 parseShapeEl 统一 /60000
    rot: xf.rot + parseInt(attr(xfrm, 'rot') ?? '0', 10),
  }
}

/** 展开组合，返回打平后的元素列表（子元素失败由分发器逐个跳过） */
export async function parseGroupEl(
  grpSp: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  const cxf = childXform(grpSp, xf)
  const out: PPTElement[] = []
  // 组合子节点是 grpSp 的直接子级（p:sp/p:pic/p:grpSp/...），交由分发器处理
  for (const child of Array.from(grpSp.children)) {
    const parsed = await parseSpTreeNode(child, ctx, cxf, pkg)
    out.push(...parsed)
  }
  return out
}
