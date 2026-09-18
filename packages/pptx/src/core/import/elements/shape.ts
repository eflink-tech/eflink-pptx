/** p:sp / p:cxnSp → shape / line / text 元素 */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { custGeomToPath, getShapeKey } from '../geometry'
import { txBodyToHTML, defRPrOf } from './text'
import { fontStackOf } from '../fonts'
import { genId } from '../../utils/id'
import { mapX, mapY, addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type {
  Gradient, LineElement, PPTElement, ShapeElement, ShadowEffect, TextElement,
} from '../../../types/slides'

/** EMU→px（96dpi） */
export function emu2px(emu: number): number {
  return emu / 9525
}

/** a:gradFill → 内部 Gradient（linear 取停站色，角度 1/60000 deg → rotate） */
export function parseGradient(gradFill: Element, ctx: ParseContext): Gradient | undefined {
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildren(gsLst, 'a:gs') : []
  if (!stops.length) return undefined
  const colors = stops.map((gs) => ({
    pos: (parseInt(attr(gs, 'pos') ?? '0', 10)) / 100000,
    color: resolveColor(gs, ctx.theme) ?? '#000000',
  }))
  const lin = directChild(gradFill, 'a:lin')
  const rotate = lin ? Math.round(parseInt(attr(lin, 'ang') ?? '0', 10) / 60000) : 0
  return { type: 'linear', colors, rotate }
}

/** a:outerShdw → 内部 ShadowEffect */
export function parseShadow(effectLst: Element | null, ctx: ParseContext): ShadowEffect | undefined {
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (!shdw) return undefined
  const blur = Math.round(emu2px(parseInt(attr(shdw, 'blurRad') ?? '0', 10)))
  const dist = emu2px(parseInt(attr(shdw, 'dist') ?? '0', 10))
  const dirRad = (parseInt(attr(shdw, 'dir') ?? '0', 10) / 60000) * (Math.PI / 180)
  return {
    h: Math.round(dist * Math.cos(dirRad)),
    v: Math.round(dist * Math.sin(dirRad)),
    blur,
    color: resolveColor(shdw, ctx.theme) ?? '#00000000',
  }
}

function dashOf(ln: Element | null): 'solid' | 'dashed' | 'dotted' {
  const dash = ln ? attr(directChild(ln, 'a:prstDash'), 'val') : null
  if (dash === 'dash' || dash === 'sysDash' || dash === 'lgDash') return 'dashed'
  if (dash === 'dot' || dash === 'sysDot' || dash === 'lgDot') return 'dotted'
  return 'solid'
}

function arrowOf(end: Element | null): LineElement['startArrow'] {
  const type = attr(end, 'type')
  if (type === 'triangle' || type === 'stealth') return 'triangle'
  if (type === 'arrow') return 'arrow'
  if (type === 'oval') return 'dot'
  return ''
}

function alignFromAlgn(algn: string | null): 'left' | 'center' | 'right' {
  return algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left'
}

/** p:sp / p:cxnSp 解析；xf 为组合变换（EMU 空间） */
export async function parseShapeEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement | null> {
  const xfrm = firstDescendant(node, 'a:xfrm')
  const ph = firstDescendant(firstDescendant(node, 'p:nvSpPr') ?? node, 'p:ph')
  let emuX: number, emuY: number, emuW: number, emuH: number
  if (xfrm) {
    const off = directChild(xfrm, 'a:off')
    const ext = directChild(xfrm, 'a:ext')
    if (!off || !ext) return null
    emuX = parseInt(attr(off, 'x') ?? '0', 10)
    emuY = parseInt(attr(off, 'y') ?? '0', 10)
    emuW = parseInt(attr(ext, 'cx') ?? '0', 10)
    emuH = parseInt(attr(ext, 'cy') ?? '0', 10)
  } else if (ph) {
    // 占位符位置继承（版式/母版）
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    const pos = ctx.placeholders.get(key)
    if (!pos) return null
    emuX = pos.x; emuY = pos.y; emuW = pos.w; emuH = pos.h
  } else {
    return null
  }
  const x = Math.round(mapX(xf, emuX) / 9525 * ctx.scale.x)
  const y = Math.round(mapY(xf, emuY) / 9525 * ctx.scale.y)
  const w = Math.max(1, Math.round((mapX(xf, emuX + emuW) - mapX(xf, emuX)) / 9525 * ctx.scale.x))
  const h = Math.max(1, Math.round((mapY(xf, emuY + emuH) - mapY(xf, emuY)) / 9525 * ctx.scale.y))
  const rot = (xfrm ? parseInt(attr(xfrm, 'rot') ?? '0', 10) : 0) / 60000 + (xf.rot ? xf.rot / 60000 : 0)
  const flipH = xfrm ? attr(xfrm, 'flipH') === '1' : false
  const flipV = xfrm ? attr(xfrm, 'flipV') === '1' : false

  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const txBody = firstDescendant(node, 'p:txBody') ?? firstDescendant(node, 'a:txBody')
  const prstGeom = firstDescendant(node, 'a:prstGeom')
  const prst = attr(prstGeom, 'prst') ?? ''
  const custGeom = spPr ? directChild(spPr, 'a:custGeom') : null
  const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
  const noFill = spPr ? directChild(spPr, 'a:noFill') : null
  const gradFill = spPr ? directChild(spPr, 'a:gradFill') : null
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const effectLst = spPr ? directChild(spPr, 'a:effectLst') : null
  const shadow = parseShadow(effectLst, ctx)

  // 连接线
  if (node.nodeName === 'p:cxnSp' || prst === 'line' || prst === 'straightConnector1') {
    const line: LineElement = {
      id: genId('l-'), type: 'line', x, y, w, h,
      start: [flipH ? w : 0, flipV ? h : 0],
      end: [flipH ? 0 : w, flipV ? 0 : h],
      lineType: 'straight',
      color: resolveColor(ln ? directChild(ln, 'a:solidFill') : solidFill, ctx.theme) ?? '#333333',
      lineWidth: Math.max(1, Math.round(emu2px(parseInt(attr(ln, 'w') ?? '12700', 10)))),
      lineStyle: dashOf(ln),
      startArrow: arrowOf(ln ? directChild(ln, 'a:headEnd') : null),
      endArrow: arrowOf(ln ? directChild(ln, 'a:tailEnd') : null),
      name: '线条',
    }
    return line
  }

  // 纯文本框：rect/textbox 且无填充无边框（noFill 视为无填充，与旧实现兼容；带 a:ln 的走形状分支保留 outline）
  if (node.nodeName === 'p:sp' && (prst === 'rect' || prst === 'textbox' || (!prst && !custGeom)) && txBody && !solidFill && !gradFill && !ln) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath)
    const text: TextElement = {
      id: genId('t-'), type: 'text', x, y, w, h,
      content: body.html,
      rotate: rot ? Math.round(rot) : undefined,
      defaultColor: '#333333', lineHeight: 1.5, padding: 8, name: '文本框',
    }
    if (body.autoSize) text.autoSize = true
    if (body.vertical) text.vertical = true
    if (shadow) text.shadow = shadow
    return text
  }

  // 形状
  const customPath = custGeom ? custGeomToPath(custGeom, w, h) : null
  let fill: string | Gradient | undefined
  if (gradFill) fill = parseGradient(gradFill, ctx)
  else if (noFill) fill = '#00000000'
  else if (solidFill) fill = resolveColor(solidFill, ctx.theme) ?? '#00000000'

  const outline = ln
    ? {
        color: resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? '#00000000',
        width: Math.max(1, Math.round(emu2px(parseInt(attr(ln, 'w') ?? '0', 10)))),
        style: dashOf(ln),
      }
    : undefined

  const shape: ShapeElement = {
    id: genId('s-'), type: 'shape', x, y, w, h,
    shapeKey: prst ? getShapeKey(prst) : 'rect',
    fill: fill ?? '#00000000',
    outline,
    flipH: flipH || undefined,
    flipV: flipV || undefined,
    rotate: rot ? Math.round(rot) : undefined,
    align: 'center', valign: 'middle', name: '形状',
  }
  if (customPath) shape.path = customPath
  if (shadow) shape.shadow = shadow

  // 形状内文本（首段样式近似）
  if (txBody) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath)
    // HTML → 纯文本：剥标签后须清理空段填充的 &nbsp; 并解码其余实体（escapeHTML 的 &amp;/&lt;/&gt; 等），
    // 否则画布按纯文本渲染会字面显示 "&nbsp;"；空行折叠 + trim 后保留段间换行（渲染层 whitespace-pre-wrap）
    const plain = decodeHTMLText(body.html
      .replace(/<li[^>]*>/g, '\n').replace(/<p[^>]*>/g, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, '')).replace(/\n{2,}/g, '\n').trim()
    if (plain) {
      shape.text = plain
      // 首 run 样式（a:r 优先，纯 a:fld 段落也可取）；显式属性缺失时回退 lstStyle defRPr 默认值
      const firstRun = firstDescendant(txBody, 'a:r') ?? firstDescendant(txBody, 'a:fld')
      const rPr = firstRun ? firstDescendant(firstRun, 'a:rPr') : null
      const defRPr = defRPrOf(directChild(txBody, 'a:lstStyle'), 0)
      shape.defaultColor = resolveColor(
        (rPr ? directChild(rPr, 'a:solidFill') : null) ?? (defRPr ? directChild(defRPr, 'a:solidFill') : null),
        ctx.theme,
      ) ?? '#FFFFFF'
      const sz = (rPr ? attr(rPr, 'sz') : null) ?? (defRPr ? attr(defRPr, 'sz') : null)
      if (sz) shape.fontSize = Math.round(parseInt(sz, 10) / 100 / 0.75)
      // 首 run 字体栈（latin + ea，含 +mj/+mn 主题引用解析）
      const latin = attr((rPr ? directChild(rPr, 'a:latin') : null) ?? (defRPr ? directChild(defRPr, 'a:latin') : null), 'typeface')
      const ea = attr((rPr ? directChild(rPr, 'a:ea') : null) ?? (defRPr ? directChild(defRPr, 'a:ea') : null), 'typeface')
      if (latin || ea) shape.defaultFontName = fontStackOf(latin, ea, ctx.theme)
      const firstP = firstDescendant(txBody, 'a:p')
      shape.align = alignFromAlgn(firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null)
      const bodyPr = directChild(txBody, 'a:bodyPr')
      const anchor = attr(bodyPr, 'anchor')
      shape.valign = anchor === 't' ? 'top' : anchor === 'b' ? 'bottom' : 'middle'
    }
  }
  return shape
}

/** 未识别的 shape 特性统一在此报告 */
export function skipShape(ctx: ParseContext, reason: string): void {
  addSkipped(ctx.report, reason)
}

/** 解码 HTML 实体（&amp;/&lt;/&gt;/&quot;/&#39; 等）：借助 text/html 解析器的文本语义 */
function decodeHTMLText(s: string): string {
  const doc = new DOMParser().parseFromString(s, 'text/html')
  return doc.documentElement.textContent ?? s
}
