/** p:grpSp → <g>：SVG transform 完整表达 child→parent 映射（缩放/旋转/翻转），子元素按子空间坐标原样渲染，
 * 由组 g 的 transform 统一映射到父空间。 */
import { attr, directChild } from '../import/xml'
import { emu2pxF, svgEl, type PreviewCtx } from './svg'
import { renderSpTreeNode } from './dispatch'

export async function renderGroup(grpSp: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  // xfrm 只从本组合的 p:grpSpPr 取：全局子树搜索会静默抓到第一个子形状的 xfrm，造成子元素双重偏移
  const grpSpPr = directChild(grpSp, 'p:grpSpPr')
  const xfrm = grpSpPr ? directChild(grpSpPr, 'a:xfrm') : null
  const g = svgEl('g')
  if (xfrm) {
    const off = directChild(xfrm, 'a:off')
    const ext = directChild(xfrm, 'a:ext')
    if (off && ext) {
      const ox = emu2pxF(parseInt(attr(off, 'x') ?? '0', 10))
      const oy = emu2pxF(parseInt(attr(off, 'y') ?? '0', 10))
      const cx = emu2pxF(parseInt(attr(ext, 'cx') ?? '0', 10))
      const cy = emu2pxF(parseInt(attr(ext, 'cy') ?? '0', 10))
      // chOff/chExt 缺失视为与 off/ext 重合（不缩放、不平移）
      const chOff = directChild(xfrm, 'a:chOff')
      const chExt = directChild(xfrm, 'a:chExt')
      const chx = chOff ? emu2pxF(parseInt(attr(chOff, 'x') ?? '0', 10)) : ox
      const chy = chOff ? emu2pxF(parseInt(attr(chOff, 'y') ?? '0', 10)) : oy
      const chcx = (chExt ? emu2pxF(parseInt(attr(chExt, 'cx') ?? '0', 10)) : cx) || 1
      const chcy = (chExt ? emu2pxF(parseInt(attr(chExt, 'cy') ?? '0', 10)) : cy) || 1
      const flipH = attr(xfrm, 'flipH') === '1'
      const flipV = attr(xfrm, 'flipV') === '1'
      const sx = (flipH ? -1 : 1) * (cx / chcx)
      const sy = (flipV ? -1 : 1) * (cy / chcy)
      // OOXML 语义 p' = R(绕组合中心)·T(off)·F·S·T(-chOff)·p；SVG 串 A B 为 A×B（右侧先作用），
      // rotate 必须置首（最后作用于父空间坐标）。镜像后需把缩放原点补到 x0/y0，保持组合中心不变
      const x0 = ox + (flipH ? cx : 0)
      const y0 = oy + (flipV ? cy : 0)
      const parts: string[] = []
      const rot = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000
      if (rot) parts.push(`rotate(${rot},${ox + cx / 2},${oy + cy / 2})`)
      parts.push(`translate(${x0},${y0})`, `scale(${sx},${sy})`, `translate(${-chx},${-chy})`)
      g.setAttribute('transform', parts.join(' '))
    }
  }
  for (const child of Array.from(grpSp.children)) {
    g.append(...(await renderSpTreeNode(child, ctx)))
  }
  return g.children.length ? g : null
}
