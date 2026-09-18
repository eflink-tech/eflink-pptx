/** p:grpSp → <g>：SVG transform 完整表达 child→parent 映射与组合旋转（预览无需一期的 groupRotation 降级）。
 * 子元素按子空间坐标原样渲染，由组 g 的 transform 统一映射到父空间。 */
import { attr, directChild } from '../import/xml'
import { emu2pxF, svgEl, type PreviewCtx } from './svg'
import { renderSpTreeNode } from './dispatch'

export async function renderGroup(grpSp: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  // xfrm 只从本组合的 p:grpSpPr 取：全局子树搜索会静默抓到第一个子形状的 xfrm，造成子元素双重偏移
  const grpSpPr = directChild(grpSp, 'p:grpSpPr')
  const xfrm = grpSpPr ? directChild(grpSpPr, 'a:xfrm') : null
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const chOff = xfrm ? directChild(xfrm, 'a:chOff') : null
  const chExt = xfrm ? directChild(xfrm, 'a:chExt') : null
  const g = svgEl('g')
  if (off && ext && chOff && chExt) {
    const ox = emu2pxF(parseInt(attr(off, 'x') ?? '0', 10))
    const oy = emu2pxF(parseInt(attr(off, 'y') ?? '0', 10))
    const cx = emu2pxF(parseInt(attr(ext, 'cx') ?? '0', 10))
    const cy = emu2pxF(parseInt(attr(ext, 'cy') ?? '0', 10))
    const chx = emu2pxF(parseInt(attr(chOff, 'x') ?? '0', 10))
    const chy = emu2pxF(parseInt(attr(chOff, 'y') ?? '0', 10))
    const chcx = emu2pxF(parseInt(attr(chExt, 'cx') ?? '0', 10)) || 1 // 防除零：0 视为 1（不缩放）
    const chcy = emu2pxF(parseInt(attr(chExt, 'cy') ?? '0', 10)) || 1
    // SVG transform 串 "A B" 矩阵为 A×B（右侧先作用）；OOXML 语义 p' = R(绕父空间组合中心)·T(off)·S·T(-chOff)·p，
    // 因此 rotate 必须在串首（最后作用于已映射到父空间的坐标），置于串尾会错误地旋转子空间坐标
    const parts: string[] = []
    const rot = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000
    if (rot) parts.push(`rotate(${rot},${ox + cx / 2},${oy + cy / 2})`)
    parts.push(`translate(${ox},${oy})`, `scale(${cx / chcx},${cy / chcy})`, `translate(${-chx},${-chy})`)
    g.setAttribute('transform', parts.join(' '))
  }
  for (const child of Array.from(grpSp.children)) {
    g.append(...(await renderSpTreeNode(child, ctx)))
  }
  return g.children.length ? g : null
}
