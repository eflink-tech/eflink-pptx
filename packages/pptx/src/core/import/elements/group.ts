// src/core/import/elements/group.ts
/** p:grpSp 递归展开：组合仿射（off/ext vs chOff/chExt + rot/flip）折算到子元素 */
import { attr, directChild } from '../xml'
import { composeXform } from '../context'
import type { GroupXform } from '../context'
import type { PPTElement } from '../../../types/slides'
import { parseSpTreeNode } from './index'
import type { ParseContext } from '../context'
import type { PptxPackage } from '../package'

/** 由 grpSp 的 xfrm 计算子空间 → 当前空间的复合仿射（含缩放/旋转/翻转） */
export function childXform(grpSp: Element, xf: GroupXform): GroupXform {
  // 只取本组合 grpSpPr 下的 xfrm：全局搜索会在缺少 grpSpPr 时抓子形状的 xfrm 造成二次偏移
  const grpSpPr = directChild(grpSp, 'p:grpSpPr')
  const xfrm = grpSpPr ? directChild(grpSpPr, 'a:xfrm') : null
  if (!xfrm) return xf
  const off = directChild(xfrm, 'a:off')
  const ext = directChild(xfrm, 'a:ext')
  if (!off || !ext) return xf
  const offX = parseInt(attr(off, 'x') ?? '0', 10)
  const offY = parseInt(attr(off, 'y') ?? '0', 10)
  const cx = parseInt(attr(ext, 'cx') ?? '0', 10)
  const cy = parseInt(attr(ext, 'cy') ?? '0', 10)
  // chOff/chExt 缺失视为与 off/ext 重合（不缩放、不平移）
  const chOff = directChild(xfrm, 'a:chOff')
  const chExt = directChild(xfrm, 'a:chExt')
  const chX = chOff ? parseInt(attr(chOff, 'x') ?? '0', 10) : offX
  const chY = chOff ? parseInt(attr(chOff, 'y') ?? '0', 10) : offY
  const chCx = chExt ? parseInt(attr(chExt, 'cx') ?? '0', 10) || 1 : cx || 1
  const chCy = chExt ? parseInt(attr(chExt, 'cy') ?? '0', 10) || 1 : cy || 1
  const sx = cx / chCx
  const sy = cy / chCy
  const flipH = attr(xfrm, 'flipH') === '1'
  const flipV = attr(xfrm, 'flipV') === '1'
  const deg = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000
  const theta = (deg * Math.PI) / 180
  const cos = Math.cos(theta)
  const sin = Math.sin(theta)

  // 局部变换 L = T(中心) · R(θ) · F · S · T(-子中心)：p' = 组合中心 + R·F·S·(p - 子中心)
  const kx = sx * (flipH ? -1 : 1)
  const ky = sy * (flipV ? -1 : 1)
  const l00 = cos * kx
  const l01 = -sin * ky
  const l10 = sin * kx
  const l11 = cos * ky
  const ccx = offX + cx / 2
  const ccy = offY + cy / 2
  const chcx = chX + chCx / 2
  const chcy = chY + chCy / 2
  const tx = ccx - (l00 * chcx + l01 * chcy)
  const ty = ccy - (l10 * chcx + l11 * chcy)

  // 复合 outer ∘ local
  return composeXform(xf, {
    a: l00, b: l10, c: l01, d: l11, e: tx, f: ty,
    rot: deg, flipH, flipV,
  })
}

/** 展开组合，返回打平后的元素列表（子元素失败由分发器逐个跳过） */
export async function parseGroupEl(
  grpSp: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  const cxf = childXform(grpSp, xf)
  const grpSpPr = directChild(grpSp, 'p:grpSpPr')
  const childCtx = groupChildContext(grpSpPr, ctx)
  const out: PPTElement[] = []
  // 组合子节点是 grpSp 的直接子级（p:sp/p:pic/p:grpSp/...），交由分发器处理
  for (const child of Array.from(grpSp.children)) {
    const parsed = await parseSpTreeNode(child, childCtx, cxf, pkg)
    out.push(...parsed)
  }
  return out
}

const FILL_LOCALS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:blipFill', 'a:grpFill']

/** 组合自身的填充：颜色/渐变/图片填充作为子元素的 a:grpFill 来源；
 *  box 为填充覆盖区域（子空间 EMU，chOff/chExt 叠加 a:fillRect 内缩） */
function groupFillOf(grpSpPr: Element): ParseContext['groupFill'] {
  let fill: Element | null = null
  for (const name of FILL_LOCALS) {
    fill = directChild(grpSpPr, name)
    if (fill) break
  }
  // noFill/grpFill：组合无可继承的实体填充（grpFill 继续沿用更外层组合）
  if (!fill || fill.localName === 'noFill' || fill.localName === 'grpFill') return undefined
  const xfrm = directChild(grpSpPr, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const chOff = xfrm ? directChild(xfrm, 'a:chOff') : null
  const chExt = xfrm ? directChild(xfrm, 'a:chExt') : null
  const baseX = parseInt(attr(chOff ?? off, 'x') ?? '0', 10)
  const baseY = parseInt(attr(chOff ?? off, 'y') ?? '0', 10)
  const baseW = parseInt(attr(chExt ?? ext, 'cx') ?? '0', 10) || 1
  const baseH = parseInt(attr(chExt ?? ext, 'cy') ?? '0', 10) || 1
  const stretch = directChild(fill, 'a:stretch')
  const fillRect = stretch ? directChild(stretch, 'a:fillRect') : null
  const frac = (name: string): number => (parseInt(attr(fillRect, name) ?? '0', 10) || 0) / 100000
  const l = frac('l')
  const t = frac('t')
  const r = frac('r')
  const b = frac('b')
  return {
    fill,
    box: {
      x: baseX + l * baseW,
      y: baseY + t * baseH,
      w: Math.max(1, baseW * (1 - l - r)),
      h: Math.max(1, baseH * (1 - t - b)),
    },
  }
}

/** 派生继承组合填充的子上下文（组合未声明时沿用外层，嵌套组合语义一致）。
 *  组合阴影不向子元素分摊：逐子元素叠加会把同一投影重复多次，视觉显著重于原图。 */
export function groupChildContext<T extends ParseContext>(grpSpPr: Element | null, ctx: T): T {
  if (!grpSpPr) return ctx
  const groupFill = groupFillOf(grpSpPr) ?? ctx.groupFill
  if (groupFill === ctx.groupFill) return ctx
  return { ...ctx, groupFill }
}
