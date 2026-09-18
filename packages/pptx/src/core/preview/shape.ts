/** p:sp / p:cxnSp → SVG：形状路径（嵌套 svg 0-100 视口）/ 线条；渐变、阴影、箭头登记到 defs。
 * 已知取舍：嵌套 svg 的 preserveAspectRatio="none" 会非均匀缩放描边（极端宽高比下描边粗细略有失真）。 */
import { attr, directChild, directChildren, firstDescendant } from '../import/xml'
import { resolveColor } from '../import/styles'
import { custGeomToPath, getShapeKey } from '../import/geometry'
import { getShapePath } from '../render/shape'
import { emu2pxF, svgEl, geomOf, boxTransform, type PreviewCtx } from './svg'
import { renderText } from './text'

/** a:gradFill → defs 登记 linearGradient（全部停站保留；a:lin@ang 从 3 点钟方向顺时针，SVG y 向下三角函数直接适用） */
export function registerLinearGradient(gradFill: Element, ctx: PreviewCtx): string {
  const id = ctx.uid('grad')
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildren(gsLst, 'a:gs') : []
  const lin = directChild(gradFill, 'a:lin')
  const ang = ((parseInt(attr(lin, 'ang') ?? '0', 10)) / 60000) * (Math.PI / 180)
  const grad = svgEl('linearGradient', {
    id,
    x1: 0.5 - Math.cos(ang) / 2,
    y1: 0.5 - Math.sin(ang) / 2,
    x2: 0.5 + Math.cos(ang) / 2,
    y2: 0.5 + Math.sin(ang) / 2,
    gradientUnits: 'objectBoundingBox',
  })
  for (const gs of stops) {
    grad.appendChild(svgEl('stop', {
      offset: parseInt(attr(gs, 'pos') ?? '0', 10) / 100000,
      // resolveColor 输出 #RRGGBB（或含 alpha 的 8 位形式），stop-color 均支持
      'stop-color': resolveColor(gs, ctx.theme) ?? '#000000',
    }))
  }
  ctx.defs.appendChild(grad)
  return `url(#${id})`
}

/** a:outerShdw → defs 登记 feDropShadow filter（stdDeviation = blur/2），无阴影返回 undefined */
export function registerShadowFilter(effectLst: Element | null, ctx: PreviewCtx): string | undefined {
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (!shdw) return undefined
  const id = ctx.uid('shadow')
  const blur = emu2pxF(parseInt(attr(shdw, 'blurRad') ?? '0', 10))
  const dist = emu2pxF(parseInt(attr(shdw, 'dist') ?? '0', 10))
  const dirRad = ((parseInt(attr(shdw, 'dir') ?? '0', 10)) / 60000) * (Math.PI / 180)
  const filter = svgEl('filter', { id, x: '-50%', y: '-50%', width: '200%', height: '200%' })
  filter.appendChild(svgEl('feDropShadow', {
    dx: dist * Math.cos(dirRad),
    dy: dist * Math.sin(dirRad),
    stdDeviation: blur / 2,
    'flood-color': resolveColor(shdw, ctx.theme) ?? '#000000',
  }))
  ctx.defs.appendChild(filter)
  return `url(#${id})`
}

/** a:ln → 描边属性；缺失或 a:noFill 返回 { 'stroke-width': 0 }（无边框） */
export function lnOf(ln: Element | null, ctx: PreviewCtx): Record<string, string | number> {
  if (!ln || directChild(ln, 'a:noFill')) return { 'stroke-width': 0 }
  const dash = attr(directChild(ln, 'a:prstDash'), 'val') ?? ''
  const dasharray = dash === 'dash' || dash === 'sysDash' || dash === 'lgDash'
    ? '12,6'
    : dash === 'dot' || dash === 'sysDot' || dash === 'lgDot' ? '2,4' : undefined
  // resolveColor 只对容器的直接颜色子节点求值，需显式取 a:ln 下的 a:solidFill
  const out: Record<string, string | number> = {
    stroke: resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? '#000000',
    'stroke-width': emu2pxF(parseInt(attr(ln, 'w') ?? '12700', 10)), // 缺省 12700 EMU = 1pt
  }
  if (dasharray) out['stroke-dasharray'] = dasharray
  return out
}

/** 端点箭头 → defs 登记 marker（triangle/stealth 实心三角、arrow 开放箭头、oval 圆点） */
function registerMarker(ctx: PreviewCtx, type: string, color: string): string {
  const id = ctx.uid('marker')
  let marker: SVGElement
  if (type === 'oval') {
    marker = svgEl('marker', { id, markerWidth: 8, markerHeight: 8, refX: 4, refY: 4, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('circle', { cx: 4, cy: 4, r: 3, fill: color }))
  } else if (type === 'arrow') {
    marker = svgEl('marker', { id, markerWidth: 10, markerHeight: 10, refX: 9, refY: 5, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('path', { d: 'M1,1 L9,5 L1,9', fill: 'none', stroke: color, 'stroke-width': 1.5 }))
  } else {
    marker = svgEl('marker', { id, markerWidth: 10, markerHeight: 10, refX: 9, refY: 5, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('path', { d: 'M0,0 L10,5 L0,10 Z', fill: color }))
  }
  ctx.defs.appendChild(marker)
  return `url(#${id})`
}

/** p:cxnSp / 直线 prst → <line>（flip 决定起止点，rot 上 transform；flip 不再进 transform 以免双重应用） */
export function renderLine(node: Element, ctx: PreviewCtx): SVGElement | null {
  const box = geomOf(node, ctx)
  if (!box) return null
  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
  // 颜色取 a:ln 下的 a:solidFill（resolveColor 只对直接颜色子节点求值），无边框时回退 spPr 填充色
  const color = resolveColor(ln ? directChild(ln, 'a:solidFill') : solidFill, ctx.theme) ?? '#333333'
  const line = svgEl('line', {
    x1: box.flipH ? box.x + box.w : box.x,
    y1: box.flipV ? box.y + box.h : box.y,
    x2: box.flipH ? box.x : box.x + box.w,
    y2: box.flipV ? box.y : box.y + box.h,
    stroke: color,
    'stroke-width': ln ? emu2pxF(parseInt(attr(ln, 'w') ?? '12700', 10)) : 1,
    transform: boxTransform({ ...box, flipH: false, flipV: false }), // flip 已折入起止点
  })
  // lnOf 的无边框降级（stroke-width:0）仅适用于存在 a:ln 的场景；线条缺 a:ln 时保留缺省 1pt 可见描边
  if (ln) {
    for (const [k, v] of Object.entries(lnOf(ln, ctx))) {
      if (k === 'stroke' || v === undefined) continue
      line.setAttribute(k, String(v))
    }
  }
  const head = attr(ln ? directChild(ln, 'a:headEnd') : null, 'type')
  const tail = attr(ln ? directChild(ln, 'a:tailEnd') : null, 'type')
  if (head) line.setAttribute('marker-start', registerMarker(ctx, head, color))
  if (tail) line.setAttribute('marker-end', registerMarker(ctx, tail, color))
  return line
}

/** p:sp → <g>：路径（custGeom 优先，否则预设形状库）+ 渐变/阴影 + 形状内文本 */
export async function renderShape(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const prst = attr(firstDescendant(node, 'a:prstGeom'), 'prst') ?? ''
  if (node.nodeName === 'p:cxnSp' || prst === 'line' || prst === 'straightConnector1') {
    return renderLine(node, ctx)
  }
  const box = geomOf(node, ctx)
  if (!box) return null
  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const custGeom = spPr ? directChild(spPr, 'a:custGeom') : null

  const g = svgEl('g', { transform: boxTransform(box) })
  // 路径：custGeomToPath 产出 0-100 视口空间（w/h 传 px 值，内部 ×9525 还原 EMU 归一化）；
  // 预设形状库同为 0-100（未知 key 回退 rect）
  const meta = custGeom ? null : getShapePath(prst ? getShapeKey(prst) : 'rect')
  const d = custGeom ? custGeomToPath(custGeom, box.w, box.h) : meta?.path ?? null
  if (d) {
    const noFill = spPr ? directChild(spPr, 'a:noFill') : null
    const gradFill = spPr ? directChild(spPr, 'a:gradFill') : null
    const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
    const ln = spPr ? directChild(spPr, 'a:ln') : null
    const effectLst = spPr ? directChild(spPr, 'a:effectLst') : null
    const nested = svgEl('svg', {
      x: box.x, y: box.y, width: box.w, height: box.h,
      viewBox: '0 0 100 100', preserveAspectRatio: 'none', overflow: 'visible',
    })
    nested.appendChild(svgEl('path', {
      d,
      fill: gradFill
        ? registerLinearGradient(gradFill, ctx)
        : noFill ? 'none'
          : resolveColor(solidFill, ctx.theme) ?? 'none',
      filter: registerShadowFilter(effectLst, ctx),
      ...lnOf(ln, ctx),
      'fill-rule': meta?.evenodd ? 'evenodd' : undefined,
    }))
    g.appendChild(nested)
  }

  const txBody = firstDescendant(node, 'p:txBody') ?? firstDescendant(node, 'a:txBody')
  if (txBody) {
    const text = await renderText(txBody, ctx, box)
    if (text) g.appendChild(text)
  }
  return g.children.length ? g : null
}
