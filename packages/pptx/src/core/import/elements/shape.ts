/** p:sp / p:cxnSp → shape / line / text 元素 */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { refColor, resolveColor, styleMatrixEntry } from '../styles'
import { resolveSchemeColor } from '../theme'
import type { PptxTheme } from '../theme'
import { custGeomToPath, presetGeomToPath, getShapeKey } from '../geometry'
import { txBodyToHTML, defRPrOf, lstAlgnOf } from './text'
import { fontStackOf } from '../fonts'
import { genId } from '../../utils/id'
import {
  addSkipped, resolvePlaceholderGeometry, inheritedPlaceholderText, inheritedPlaceholderShape,
  fallbackPlaceholderBox, parseElementLink,
} from '../context'
import { frameRotOf, mapEmuBox } from './frame-geom'
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
export function parseGradient(gradFill: Element, ctx: ParseContext, phColor?: Element | null): Gradient | undefined {
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildren(gsLst, 'a:gs') : []
  if (!stops.length) return undefined
  const colors = stops.map((gs) => ({
    pos: (parseInt(attr(gs, 'pos') ?? '0', 10)) / 100000,
    color: resolveColor(gs, ctx.theme, phColor) ?? '#000000',
  }))
  const lin = directChild(gradFill, 'a:lin')
  const rotate = lin ? Math.round(parseInt(attr(lin, 'ang') ?? '0', 10) / 60000) : 0
  return { type: 'linear', colors, rotate }
}

/** a:outerShdw → 内部 ShadowEffect */
export function parseShadow(effectLst: Element | null, ctx: ParseContext, phColor?: Element | null): ShadowEffect | undefined {
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (!shdw) return undefined
  const blur = Math.round(emu2px(parseInt(attr(shdw, 'blurRad') ?? '0', 10)))
  const dist = emu2px(parseInt(attr(shdw, 'dist') ?? '0', 10))
  const dirRad = (parseInt(attr(shdw, 'dir') ?? '0', 10) / 60000) * (Math.PI / 180)
  return {
    h: Math.round(dist * Math.cos(dirRad)),
    v: Math.round(dist * Math.sin(dirRad)),
    blur,
    color: resolveColor(shdw, ctx.theme, phColor) ?? '#00000000',
  }
}

/** 未支持的效果（发光/倒影/柔化/内阴影/棱台/3D 旋转等）计入报告；仅 outerShdw 可还原 */
export function reportUnsupportedEffects(effectLst: Element | null, ctx: ParseContext, kind = 'unsupportedEffect'): void {
  if (!effectLst) return
  for (const child of Array.from(effectLst.children)) {
    if (child.localName !== 'outerShdw') addSkipped(ctx.report, kind)
  }
}

/** 非纯色填充（图案/图片/组合）无法用当前模型表达，计入报告 */
export function reportUnsupportedFill(fill: Element | null, ctx: ParseContext): void {
  if (!fill) return
  if (fill.localName === 'pattFill') addSkipped(ctx.report, 'patternFill')
  else addSkipped(ctx.report, 'unsupportedFill')
}

/** p:style 的 effectRef → 矩阵 effectLst（idx=0 或矩阵缺失返回 null） */
export function effectFromRef(styleRef: Element | null, theme: PptxTheme): Element | null {
  const ref = styleRef ? directChild(styleRef, 'a:effectRef') : null
  if (!ref) return null
  const idx = parseInt(attr(ref, 'idx') ?? '0', 10)
  return styleMatrixEntry(theme, 'effect', idx)
}

/** p:style 的 fontRef 颜色（形状/文本框文字默认色） */
export function fontRefColor(styleRef: Element | null, theme: PptxTheme): string | undefined {
  const ref = styleRef ? directChild(styleRef, 'a:fontRef') : null
  return resolveColor(ref, theme) ?? undefined
}

/** p:style 的 fontRef idx → 主题字体引用（'+mj-lt' / '+mn-lt'），供 fontStackOf 解析 */
function fontRefFace(styleRef: Element | null): string | null {
  const ref = styleRef ? directChild(styleRef, 'a:fontRef') : null
  const idx = attr(ref, 'idx')
  if (idx === 'major') return '+mj-lt'
  if (idx === 'minor') return '+mn-lt'
  return null
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
  const ph = firstDescendant(firstDescendant(node, 'p:nvSpPr') ?? firstDescendant(node, 'p:nvCxnSpPr') ?? node, 'p:ph')
  const cNvPr = firstDescendant(node, 'p:cNvPr')
  const link = await parseElementLink(cNvPr, ctx)
  const altText = attr(cNvPr, 'descr')?.trim() || undefined
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
    // 占位符位置继承（版式/母版）：idx 优先、type 兜底；均缺失时用类型默认位置并计入报告
    const pos = resolvePlaceholderGeometry(ctx, ph)
    if (pos) {
      emuX = pos.x; emuY = pos.y; emuW = pos.w; emuH = pos.h
    } else if (ctx.slideSize) {
      addSkipped(ctx.report, 'missingPlaceholder')
      const box = fallbackPlaceholderBox(ph, ctx.slideSize)
      emuX = box.x; emuY = box.y; emuW = box.w; emuH = box.h
    } else {
      addSkipped(ctx.report, 'missingPlaceholder')
      return null
    }
  } else {
    return null
  }
  const { x, y, w, h } = mapEmuBox(xf, { x: emuX, y: emuY, w: emuW, h: emuH }, ctx.scale)
  const { rot, flipH, flipV } = frameRotOf(xfrm, xf)

  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const txBody = firstDescendant(node, 'p:txBody') ?? firstDescendant(node, 'a:txBody')
  // 占位符文字样式继承（版式/母版原型 lstStyle/bodyPr → 母版 txStyles）
  const inherited = ph ? inheritedPlaceholderText(ctx, ph) : undefined
  // 占位符外观继承（版式/母版原型 spPr 的几何/填充/边框/效果，p:style 引用）：
  // 幻灯片占位符常只写空 spPr 甚至完全省略，按属性粒度回退到版式原型
  const phShape = ph ? inheritedPlaceholderShape(ctx, ph) : undefined
  const protoSpPr = phShape?.spPr ?? null
  // 属性组继承：自身 spPr 声明了该组任一元素即整组以自身为准，否则整组回退版式原型
  const groupEl = (...names: string[]): Element | null => {
    for (const src of [spPr, protoSpPr]) {
      if (!src) continue
      for (const name of names) {
        const el = directChild(src, name)
        if (el) return el
      }
    }
    return null
  }
  const geomEl = groupEl('a:prstGeom', 'a:custGeom')
  const prstGeom = geomEl?.localName === 'prstGeom' ? geomEl : null
  const custGeom = geomEl?.localName === 'custGeom' ? geomEl : null
  const prst = attr(prstGeom, 'prst') ?? ''
  const fillEl = groupEl('a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:blipFill', 'a:grpFill')
  const solidFill = fillEl?.localName === 'solidFill' ? fillEl : null
  const noFill = fillEl?.localName === 'noFill' ? fillEl : null
  const gradFill = fillEl?.localName === 'gradFill' ? fillEl : null
  const ln = groupEl('a:ln')
  const styleRef = directChild(node, 'p:style') ?? phShape?.style ?? null
  const explicitEffectLst = groupEl('a:effectLst')
  const effectLst = explicitEffectLst ?? effectFromRef(styleRef, ctx.theme)
  const effectPh = explicitEffectLst ? null : refColor(styleRef ? directChild(styleRef, 'a:effectRef') : null)
  const shadow = parseShadow(effectLst, ctx, effectPh)
  reportUnsupportedEffects(effectLst, ctx)

  // 连接线
  if (node.localName === 'cxnSp' || prst === 'line' || prst === 'straightConnector1') {
    const cxnLnRef = styleRef ? directChild(styleRef, 'a:lnRef') : null
    const cxnLnIdx = parseInt(attr(cxnLnRef, 'idx') ?? '0', 10)
    const cxnLnMatrix = cxnLnIdx > 0 ? styleMatrixEntry(ctx.theme, 'line', cxnLnIdx) : null
    const cxnPh = refColor(cxnLnRef)
    const lineStyleEl = ln ?? cxnLnMatrix
    const color = ln
      ? resolveColor(directChild(ln, 'a:solidFill'), ctx.theme)
      : cxnLnMatrix
        ? resolveColor(directChild(cxnLnMatrix, 'a:solidFill'), ctx.theme, cxnPh)
        : null
    // 连接线线型：折线（bentConnector2/3/4/5）→ broken + 中间点；曲线 → curve + 控制点
    const start: [number, number] = [flipH ? w : 0, flipV ? h : 0]
    const end: [number, number] = [flipH ? 0 : w, flipV ? 0 : h]
    const midX = (start[0] + end[0]) / 2
    const midY = (start[1] + end[1]) / 2
    let lineType: LineElement['lineType'] = 'straight'
    let points: Array<[number, number]> | undefined
    const hasExtent = w >= 2 && h >= 2
    if (hasExtent && prst === 'bentConnector2') {
      lineType = 'broken'
      points = [[end[0], start[1]]]
    } else if (hasExtent && prst === 'bentConnector3') {
      lineType = 'broken'
      points = [[midX, start[1]], [midX, end[1]]]
    } else if (hasExtent && (prst === 'bentConnector4' || prst === 'bentConnector5')) {
      lineType = 'broken'
      points = [[midX, start[1]], [midX, midY], [end[0], midY]]
    } else if (hasExtent && prst === 'curvedConnector2') {
      lineType = 'curve'
      points = [[end[0], start[1]]]
    } else if (hasExtent && prst === 'curvedConnector3') {
      lineType = 'curve'
      points = [[midX, start[1]], [midX, end[1]]]
    } else if (hasExtent && (prst === 'curvedConnector4' || prst === 'curvedConnector5')) {
      lineType = 'curve'
      points = [[midX, start[1]], [midX, end[1]]]
    }
    const line: LineElement = {
      id: genId('l-'), type: 'line', x, y, w, h,
      start,
      end,
      points,
      lineType,
      color: color ?? (cxnLnRef ? resolveColor(cxnLnRef, ctx.theme) : null) ?? resolveColor(solidFill, ctx.theme) ?? '#333333',
      lineWidth: Math.max(1, Math.round(emu2px(parseInt(attr(lineStyleEl, 'w') ?? '12700', 10)))),
      lineStyle: dashOf(lineStyleEl),
      startArrow: arrowOf(ln ? directChild(ln, 'a:headEnd') : null),
      endArrow: arrowOf(ln ? directChild(ln, 'a:tailEnd') : null),
      name: altText ?? '线条',
    }
    if (rot) line.rotate = Math.round(rot)
    if (link) line.link = link
    return line
  }

  // 纯文本框：rect/textbox 且无填充无边框（noFill 视为无填充，与旧实现兼容；a:ln 内 a:noFill
  // 即 PowerPoint 文本框的标准「无边框线」声明，同样视为无边框——此类框需走文本分支保留
  // insets/spAutoFit/富文本，落入形状分支会被降级为纯文本且丢失垂直内边距）；带实体边框的走形状分支
  const lnNoFill = ln ? Boolean(directChild(ln, 'a:noFill')) : false
  if (node.localName === 'sp' && (prst === 'rect' || prst === 'textbox' || (!prst && !custGeom)) && txBody && !solidFill && !gradFill && (!ln || lnNoFill)) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath, inherited)
    const inheritedDef = defRPrOf(inherited?.lstStyle ?? null, 0)
    const inheritedColor = resolveColor(inheritedDef ? directChild(inheritedDef, 'a:solidFill') : null, ctx.theme)
    const text: TextElement = {
      id: genId('t-'), type: 'text', x, y, w, h,
      content: body.html,
      rotate: rot ? Math.round(rot) : undefined,
      // 行距取 PowerPoint 单倍行距近似 1.2（源文件通常不声明行距，硬编码 1.5 会放大内容需求导致裁剪；
      // 段级 lnSpc 已由内联 line-height 覆盖）；padding 取 bodyPr insets 折算值（缺省回退编辑器默认 8）
      defaultColor: inheritedColor ?? fontRefColor(styleRef, ctx.theme) ?? resolveSchemeColor('tx1', ctx.theme),
      lineHeight: 1.2, padding: body.padding ?? 8, name: altText ?? '文本框',
    }
    if (link) text.link = link
    if (body.autoSize) text.autoSize = true
    if (body.autoFitShape) text.autoFit = true
    if (body.vertical) text.vertical = true
    // bodyPr anchor：ctr/b 映射 valign（t/未声明为 OOXML 默认顶对齐，不写字段）
    if (body.anchor === 'ctr') text.valign = 'middle'
    else if (body.anchor === 'b') text.valign = 'bottom'
    if (shadow) text.shadow = shadow
    // 首 run 的文字描边 / 文字阴影（TextElement 级字段；rPr effectLst 优先于 spPr 效果）
    const firstRunEl = firstDescendant(txBody, 'a:r') ?? firstDescendant(txBody, 'a:fld')
    const rPr0 = firstRunEl ? firstDescendant(firstRunEl, 'a:rPr') : null
    const ln0 = rPr0 ? directChild(rPr0, 'a:ln') : null
    const strokeColor = ln0 ? resolveColor(directChild(ln0, 'a:solidFill'), ctx.theme) : null
    if (strokeColor) {
      text.textStroke = { color: strokeColor, width: Math.max(1, Math.round(emu2px(parseInt(attr(ln0, 'w') ?? '0', 10)))) }
    }
    const runShadow = parseShadow(rPr0 ? directChild(rPr0, 'a:effectLst') : null, ctx)
    if (runShadow) text.shadow = runShadow
    return text
  }

  // 形状
  // custGeom / 带调整参数的预设几何（如 snip1Rect）→ 精确 path；粗映射会丢失切角等视觉特征
  const customPath = custGeom
    ? custGeomToPath(custGeom, w, h)
    : presetGeomToPath(prstGeom, prst, w, h)
  // p:style 样式引用：多数设计稿形状不写显式 solidFill，靠 fillRef/lnRef 引用主题样式矩阵；
  // idx=0 表示显式无填充/无边框；矩阵内 phClr 由引用处颜色替换
  const fillRef = styleRef ? directChild(styleRef, 'a:fillRef') : null
  const lnRef = styleRef ? directChild(styleRef, 'a:lnRef') : null
  const fillIdx = parseInt(attr(fillRef, 'idx') ?? '0', 10)
  const lnIdx = parseInt(attr(lnRef, 'idx') ?? '0', 10)
  const fillPh = refColor(fillRef)
  const lnPh = refColor(lnRef)

  let fill: string | Gradient | undefined
  if (gradFill) fill = parseGradient(gradFill, ctx)
  else if (noFill) fill = '#00000000'
  else if (solidFill) fill = resolveColor(solidFill, ctx.theme) ?? '#00000000'
  else if (fillRef && fillIdx > 0) {
    const matrixFill = styleMatrixEntry(ctx.theme, 'fill', fillIdx)
    if (matrixFill?.localName === 'solidFill') fill = resolveColor(matrixFill, ctx.theme, fillPh) ?? '#00000000'
    else if (matrixFill?.localName === 'gradFill') fill = parseGradient(matrixFill, ctx, fillPh)
    else if (matrixFill?.localName === 'noFill') fill = '#00000000'
    else {
      if (matrixFill) reportUnsupportedFill(matrixFill, ctx)
      fill = resolveColor(fillRef, ctx.theme) ?? '#00000000'
    }
  } else if (fillIdx === 0 && fillRef) fill = '#00000000'

  const lnMatrix = ln ? null : styleMatrixEntry(ctx.theme, 'line', lnIdx)
  const outline = ln
    ? {
        color: resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? '#00000000',
        width: Math.max(1, Math.round(emu2px(parseInt(attr(ln, 'w') ?? '0', 10)))),
        style: dashOf(ln),
      }
    : lnRef && lnIdx > 0
      ? lnMatrix
        ? {
            color: resolveColor(directChild(lnMatrix, 'a:solidFill'), ctx.theme, lnPh) ?? '#00000000',
            width: Math.max(1, Math.round(emu2px(parseInt(attr(lnMatrix, 'w') ?? '12700', 10)))),
            style: dashOf(lnMatrix),
          }
        : {
            // 主题缺样式矩阵：按 idx 近似细/中/粗
            color: resolveColor(lnRef, ctx.theme) ?? '#00000000',
            width: Math.min(3, Math.max(1, lnIdx)),
            style: 'solid' as const,
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
    align: 'center', valign: 'middle', name: altText ?? '形状',
  }
  if (link) shape.link = link
  if (customPath) shape.path = customPath
  if (shadow) shape.shadow = shadow

  // 形状内文本（首段样式近似）
  if (txBody) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath, inherited)
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
      const defRPr = defRPrOf(directChild(txBody, 'a:lstStyle') ?? inherited?.lstStyle ?? null, 0)
      shape.defaultColor = resolveColor(
        (rPr ? directChild(rPr, 'a:solidFill') : null) ?? (defRPr ? directChild(defRPr, 'a:solidFill') : null),
        ctx.theme,
      ) ?? fontRefColor(styleRef, ctx.theme) ?? resolveSchemeColor('tx1', ctx.theme)
      // 字体引用：p:style fontRef idx 指向主题 major/minor（run/lstStyle 未声明字体时生效）
      if (!attr(rPr ? directChild(rPr, 'a:latin') : null, 'typeface') && !attr(defRPr ? directChild(defRPr, 'a:latin') : null, 'typeface')) {
        const face = fontRefFace(styleRef)
        if (face) shape.defaultFontName = fontStackOf(face, face, ctx.theme)
      }
      const sz = (rPr ? attr(rPr, 'sz') : null) ?? (defRPr ? attr(defRPr, 'sz') : null)
      if (sz) shape.fontSize = Math.round(parseInt(sz, 10) / 100 / 0.75)
      // 首 run 字体栈（latin + ea，含 +mj/+mn 主题引用解析）
      const latin = attr((rPr ? directChild(rPr, 'a:latin') : null) ?? (defRPr ? directChild(defRPr, 'a:latin') : null), 'typeface')
      const ea = attr((rPr ? directChild(rPr, 'a:ea') : null) ?? (defRPr ? directChild(defRPr, 'a:ea') : null), 'typeface')
      if (latin || ea) shape.defaultFontName = fontStackOf(latin, ea, ctx.theme)
      const firstP = firstDescendant(txBody, 'a:p')
      // 段落显式 pPr@algn 缺失时回退 lstStyle lvl1pPr@algn（叠加文本框常只有 lstStyle 声明居中）
      const firstAlgn = (firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null)
        ?? lstAlgnOf(directChild(txBody, 'a:lstStyle') ?? inherited?.lstStyle ?? null, 0)
      shape.align = alignFromAlgn(firstAlgn)
      const bodyPr = directChild(txBody, 'a:bodyPr') ?? inherited?.bodyPr ?? null
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
