/** SVG 预览助手：命名空间元素构建、单位换算、元素几何 Box。
 * 预览坐标系 = 源画布 px（EMU/9525 浮点），viewBox 即源画布尺寸；因此 ctx.scale 恒为 {1,1}。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import type { ParseContext } from '../import/context'

export const SVG_NS = 'http://www.w3.org/2000/svg'

/** 预览上下文：结构兼容 ParseContext（可直接传给一期解析函数），额外持有 defs 与 id 生成器 */
export interface PreviewCtx extends ParseContext {
  /** 当前页 SVG 的 <defs>：渐变 / 滤镜 / 裁剪路径登记处 */
  defs: SVGElement
  /** 唯一 id 生成器（实现方保证跨页唯一，如 `${prefix}-p${pageIdx}-${n++}`） */
  uid: (prefix: string) => string
}

/** 构建 SVG 命名空间元素；值为 undefined/null 的属性跳过 */
export function svgEl(tag: string, attrs?: Record<string, string | number | undefined | null>): SVGElement {
  const el = document.createElementNS(SVG_NS, tag)
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null) continue
      el.setAttribute(k, String(v))
    }
  }
  return el
}

/** EMU → px（96dpi，浮点不取整：viewBox 用浮点保证布局精确） */
export function emu2pxF(emu: number): number {
  return emu / 9525
}

/** 字号/字距（1/100 pt）→ px */
export function sz2px(hundredthsPt: number): number {
  return hundredthsPt / 100 / 0.75
}

export interface Box {
  x: number
  y: number
  w: number
  h: number
  /** 度 */
  rot: number
  flipH: boolean
  flipV: boolean
}

/** 元素几何：a:xfrm（sp/pic 用）或 p:xfrm（graphicFrame 用）→ 源画布 px 浮点 Box。
 * 无 xfrm 时回退占位符位置表（EMU → px），与一期 parseShapeEl 的占位符继承语义一致。 */
export function geomOf(node: Element, ctx: PreviewCtx): Box | null {
  const xfrm = firstDescendant(node, 'a:xfrm') ?? firstDescendant(node, 'p:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (xfrm && off && ext) {
    return {
      x: emu2pxF(parseInt(attr(off, 'x') ?? '0', 10)),
      y: emu2pxF(parseInt(attr(off, 'y') ?? '0', 10)),
      w: emu2pxF(parseInt(attr(ext, 'cx') ?? '0', 10)),
      h: emu2pxF(parseInt(attr(ext, 'cy') ?? '0', 10)),
      rot: parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000,
      flipH: attr(xfrm, 'flipH') === '1',
      flipV: attr(xfrm, 'flipV') === '1',
    }
  }
  const ph = firstDescendant(firstDescendant(node, 'p:nvSpPr') ?? node, 'p:ph')
  if (ph) {
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    const pos = ctx.placeholders.get(key)
    if (pos) {
      return {
        x: emu2pxF(pos.x), y: emu2pxF(pos.y), w: emu2pxF(pos.w), h: emu2pxF(pos.h),
        rot: 0, flipH: false, flipV: false,
      }
    }
  }
  return null
}

/** 元素变换属性串：flip（绕自身包围盒翻转）+ rot（绕自身中心）。无变换返回 undefined。 */
export function boxTransform(box: Box): string | undefined {
  const parts: string[] = []
  if (box.flipH || box.flipV) {
    const fx = box.flipH ? -1 : 1
    const fy = box.flipV ? -1 : 1
    parts.push(
      `translate(${box.x + (fx < 0 ? box.w : 0)},${box.y + (fy < 0 ? box.h : 0)}) scale(${fx},${fy}) translate(${-box.x},${-box.y})`,
    )
  }
  if (box.rot) {
    parts.push(`rotate(${box.rot},${box.x + box.w / 2},${box.y + box.h / 2})`)
  }
  return parts.length ? parts.join(' ') : undefined
}
