/** 单页渲染：背景回退链（slide→layout→master→lt1 兜底）+ 版式/母版装饰与占位符位置表 + slide 自身元素 → <svg>。
 * slide 部件缺失返回 null（由调用方计入 slideParseFailed）；背景/装饰/元素逐项降级，不中断整页。 */
import { attr, firstDescendant, parseXML } from '../import/xml'
import { collectPlaceholders, findAncestry, parseBackgroundFill } from '../import/master'
import { parseThemeForMaster } from '../import/theme'
import type { PptxPackage } from '../import/package'
import type { ImportReport } from '../../types/slides'
import { SVG_NS, emu2pxF, svgEl, type PreviewCtx } from './svg'
import { registerLinearGradient } from './shape'
import { renderSpTreeNode } from './dispatch'

/** 参与 spTree 渲染的节点类型（与一期 parseSlideAncestry 的装饰过滤口径一致） */
const ELEMENT_NODES = ['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame', 'p:grpSp']

/** 部件 spTree 定位：p:cSld > p:spTree，cSld 缺失时容错到根下找 */
function spTreeOf(doc: Document): Element | null {
  return firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')
}

/** 单部件背景解析并追加全幅背景节点；解析/映射失败返回 false 由调用方继续回退下一级 */
async function appendBgPart(
  svg: SVGSVGElement,
  pkg: PptxPackage,
  partPath: string,
  ctx: PreviewCtx,
  w: number,
  h: number,
): Promise<boolean> {
  const xml = await pkg.text(partPath)
  if (!xml) return false
  const doc = parseXML(xml)
  const bg = firstDescendant(doc.documentElement, 'p:bg')
  const fill = parseBackgroundFill(bg, ctx.theme)
  if (!bg || !fill) return false
  if (fill.type === 'solid') {
    svg.appendChild(svgEl('rect', { x: 0, y: 0, width: w, height: h, fill: fill.color }))
    return true
  }
  if (fill.type === 'gradient') {
    // 复用 shape.ts 的 linearGradient 登记（a:lin@ang 口径一致）；取原始 a:gradFill 元素保留全部停站
    const gradFill = firstDescendant(bg, 'a:gradFill')
    if (!gradFill) return false
    svg.appendChild(svgEl('rect', { x: 0, y: 0, width: w, height: h, fill: registerLinearGradient(gradFill, ctx) }))
    return true
  }
  // 图片背景：src 解析失败继续回退下一级（避免悬空 image），与一期 parseSlideAncestry 语义一致
  const embedId = attr(firstDescendant(bg, 'a:blip'), 'r:embed')
  const target = embedId ? await pkg.relTarget(partPath, embedId) : null
  const src = target ? await pkg.mediaDataUrl(target) : undefined
  if (!src) return false
  svg.appendChild(svgEl('image', {
    x: 0, y: 0, width: w, height: h, href: src,
    preserveAspectRatio: 'xMidYMid slice', // cover 语义
  }))
  return true
}

export async function renderSlide(
  pkg: PptxPackage,
  slidePath: string,
  srcW: number,
  srcH: number,
  pageIdx: number,
  report: ImportReport,
): Promise<SVGSVGElement | null> {
  const xml = await pkg.text(slidePath)
  if (!xml) return null
  const slideDoc = parseXML(xml)
  const { layoutPath, masterPath } = await findAncestry(pkg, slidePath)
  const theme = await parseThemeForMaster(pkg, masterPath)

  // SVG 根：viewBox = 源画布尺寸（EMU→px 浮点），预览坐标 = 源画布 px
  const w = emu2pxF(srcW)
  const h = emu2pxF(srcH)
  const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement
  svg.setAttribute('xmlns', SVG_NS)
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`)
  svg.setAttribute('width', String(w))
  svg.setAttribute('height', String(h))
  const defs = svgEl('defs')
  svg.appendChild(defs)

  // 占位符位置表：版式/母版合入（不覆盖），key = idx ?? type（EMU 空间，与 geomOf 回退口径一致）
  const placeholders = new Map<string, { x: number; y: number; w: number; h: number }>()
  let n = 0
  const makeCtx = (partPath: string): PreviewCtx => ({
    pkg,
    partPath,
    theme,
    report,
    scale: { x: 1, y: 1 },
    placeholders,
    defs,
    uid: (prefix: string) => `${prefix}-p${pageIdx}-${n++}`,
  })

  // 背景回退链：slide → layout → master，最终 lt1 兜底
  let hasBg = false
  for (const partPath of [slidePath, layoutPath, masterPath]) {
    if (!partPath) continue
    if (await appendBgPart(svg, pkg, partPath, makeCtx(partPath), w, h)) {
      hasBg = true
      break
    }
  }
  if (!hasBg) {
    svg.appendChild(svgEl('rect', { x: 0, y: 0, width: w, height: h, fill: theme.schemeColors.lt1 ?? '#FFFFFF' }))
  }

  // 版式/母版层：占位符位置合入 + 非占位符装饰渲染
  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath) continue
    const partXml = await pkg.text(partPath)
    if (!partXml) continue
    const spTree = spTreeOf(parseXML(partXml))
    if (!spTree) continue
    for (const [k, v] of collectPlaceholders(spTree)) {
      if (!placeholders.has(k)) placeholders.set(k, v)
    }
    const partCtx = makeCtx(partPath)
    for (const child of Array.from(spTree.children)) {
      if (!ELEMENT_NODES.includes(child.nodeName)) continue
      if (firstDescendant(child, 'p:ph')) continue // 占位符原型不直接渲染
      svg.append(...(await renderSpTreeNode(child, partCtx)))
    }
  }

  // slide 层：slide 占位符（常无 xfrm）由 geomOf 经 placeholders 表回退位置
  const slideTree = spTreeOf(slideDoc)
  if (slideTree) {
    const slideCtx = makeCtx(slidePath)
    for (const child of Array.from(slideTree.children)) {
      if (!ELEMENT_NODES.includes(child.nodeName)) continue
      svg.append(...(await renderSpTreeNode(child, slideCtx)))
    }
  }
  return svg
}
