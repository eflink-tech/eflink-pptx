// PPTX 导入：解压 OOXML → 内部 schema（尽力还原文本/形状/图片/表格/线条）
import JSZip from 'jszip'
import type {
  PPTElement, Presentation, ShapeElement, Slide, TableCell, TableElement,
} from '../../types/slides'
import { createDefaultTheme } from '../../types/slides'
import { genId } from '../utils/id'

/** EMU → px（1px = 9525 EMU @96dpi） */
const EMU2PX = (emu: number) => Math.round(emu / 9525)

/** pptx prstGeom → 内部形状 key */
const PRST_MAP: Record<string, string> = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle',
  rtTriangle: 'rtTriangle', parallelogram: 'parallelogram', trapezoid: 'trapezoid',
  diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', heptagon: 'heptagon',
  octagon: 'octagon', donut: 'donut', pie: 'pie', chord: 'chord', arc: 'arc',
  frame: 'frame', halfFrame: 'halfFrame', can: 'can', cube: 'cube', plaque: 'plaque',
  cross: 'cross', plus: 'plus',
  rightArrow: 'arrowRight', leftArrow: 'arrowLeft', upArrow: 'arrowUp', downArrow: 'arrowDown',
  leftRightArrow: 'arrowLeftRight', upDownArrow: 'arrowUpDown', quadArrow: 'arrowQuad',
  bentArrow: 'bentArrow', curvedRightArrow: 'curvedRightArrow', chevron: 'chevron', homePlate: 'homePlate',
  star4: 'star4', star5: 'star5', star6: 'star6', star8: 'star8', star16: 'star16',
  heart: 'heart', lightningBolt: 'lightning', sun: 'sun', moon: 'moon', cloud: 'cloud',
  smileyFace: 'smile',
  flowChartProcess: 'flowProcess', flowChartDecision: 'flowDecision', flowChartData: 'flowData',
  flowChartDocument: 'flowDocument', flowChartPredefinedProcess: 'flowPredefined',
  flowChartTerminator: 'flowTerminal', flowChartConnector: 'flowConnector',
  flowChartOffpageConnector: 'flowOffpage', flowChartDelay: 'flowDelay',
  flowChartManualInput: 'flowManualInput', flowChartManualOperation: 'flowManualOperation',
  flowChartMerge: 'flowMerge', flowChartOr: 'flowOr', flowChartExtract: 'flowExtract',
  wedgeRectCallout: 'callout1', wedgeRoundRectCallout: 'callout2', wedgeEllipseCallout: 'callout3',
  mathPlus: 'mathPlus', mathMinus: 'mathMinus', mathMultiply: 'mathMultiply',
  mathDivide: 'mathDivide', mathEqual: 'mathEqual', mathNotEqual: 'mathNotEqual',
}

interface ParseContext {
  zip: JSZip
  mediaCache: Map<string, string>
}

function attr(el: Element | null, name: string): string | null {
  return el?.getAttribute(name) ?? null
}

function directChild(parent: Element, name: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (child.nodeName === name) return child
  }
  return null
}

function descendants(parent: Element, name: string): Element[] {
  return Array.from(parent.getElementsByTagName(name))
}

function firstDescendant(parent: Element, name: string): Element | null {
  return descendants(parent, name)[0] ?? null
}

/** 解析颜色（srgbClr/schemeClr） */
function parseColor(el: Element | null): string | undefined {
  if (!el) return undefined
  const srgb = firstDescendant(el, 'a:srgbClr')
  if (srgb) {
    const val = attr(srgb, 'val') ?? '000000'
    const alpha = firstDescendant(srgb, 'a:alpha')
    let hex = `#${val}`
    if (alpha) {
      const a = (parseFloat(attr(alpha, 'val') ?? '100000') / 100000) * 255
      hex += Math.round(a).toString(16).padStart(2, '0')
    }
    return hex
  }
  return undefined
}

/** 读取媒体文件为 dataURL（带缓存） */
async function mediaToDataUrl(ctx: ParseContext, relId: string, rels: Map<string, string>): Promise<string | undefined> {
  const cached = ctx.mediaCache.get(relId)
  if (cached) return cached
  const path = rels.get(relId)
  if (!path) return undefined
  try {
    const file = ctx.zip.file(path.startsWith('ppt/') ? path : `ppt/${path}`)
    if (!file) return undefined
    const base64 = await file.async('base64')
    const ext = path.split('.').pop()?.toLowerCase() ?? ''
    const mime = ext === 'png' ? 'image/png'
      : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
      : ext === 'gif' ? 'image/gif'
      : ext === 'svg' ? 'image/svg+xml'
      : ext === 'webp' ? 'image/webp'
      : ext === 'mp4' ? 'video/mp4'
      : ext === 'webm' ? 'video/webm'
      : ext === 'mp3' ? 'audio/mpeg'
      : ext === 'wav' ? 'audio/wav'
      : 'application/octet-stream'
    const dataUrl = `data:${mime};base64,${base64}`
    ctx.mediaCache.set(relId, dataUrl)
    return dataUrl
  } catch {
    return undefined
  }
}

/** 解析 r:embed 等关系 → 目标文件路径映射（baseDir 为该 rels 所属部件所在目录，如 'ppt' / 'ppt/slides'） */
function parseRels(xml: string, baseDir: string): Map<string, string> {
  const map = new Map<string, string>()
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = attr(rel, 'Id')
    const target = attr(rel, 'Target')
    if (id && target) map.set(id, resolveTarget(baseDir, target))
  }
  return map
}

/** 相对路径解析：'../media/a.png'（基于 'ppt/slides'）→ 'ppt/media/a.png' */
export function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = `${baseDir}/${target}`.split('/')
  const out: string[] = []
  for (const part of parts) {
    if (part === '..') out.pop()
    else if (part && part !== '.') out.push(part)
  }
  return out.join('/')
}

/** 文本主体 → TextElement 的 HTML 内容 */
function txBodyToHTML(txBody: Element): string {
  const paragraphs = Array.from(txBody.children).filter((c) => c.nodeName === 'a:p')
  const html: string[] = []
  for (const p of paragraphs) {
    const pPr = directChild(p, 'a:pPr')
    const algn = attr(pPr, 'algn')
    const alignStyle = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left'
    let inner = ''
    for (const r of Array.from(p.children)) {
      if (r.nodeName === 'a:br') {
        inner += '<br>'
        continue
      }
      if (r.nodeName !== 'a:r') continue
      const t = firstDescendant(r, 'a:t')
      const text = t?.textContent ?? ''
      if (!text) continue
      const rPr = firstDescendant(r, 'a:rPr')
      const styles: string[] = []
      const runStyles: string[] = []
      if (rPr) {
        const sz = attr(rPr, 'sz') // 百分之一 pt
        if (sz) {
          const pt = parseInt(sz, 10) / 100
          runStyles.push(`font-size:${Math.round(pt / 0.75)}px`)
        }
        if (attr(rPr, 'b') === '1') runStyles.push('font-weight:bold')
        if (attr(rPr, 'i') === '1') runStyles.push('font-style:italic')
        if (attr(rPr, 'u') === 'sng') runStyles.push('text-decoration:underline')
        if (attr(rPr, 'strike') === 'sng') runStyles.push('text-decoration:line-through')
        const fill = directChild(rPr, 'a:solidFill')
        const color = parseColor(fill)
        if (color) styles.push(`color:${color}`)
      }
      const styleAttr = [...styles, ...runStyles].join(';')
      inner += styleAttr ? `<span style="${styleAttr}">${escapeXMLText(text)}</span>` : escapeXMLText(text)
    }
    html.push(`<p style="text-align:${alignStyle}">${inner || '&nbsp;'}</p>`)
  }
  return html.join('') || '<p></p>'
}

function escapeXMLText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** 单个 shape/pic/graphFrame → 内部元素 */
async function parseShapeNode(
  node: Element,
  ctx: ParseContext,
  rels: Map<string, string>,
): Promise<PPTElement | null> {
  const xfrm = firstDescendant(node, 'a:xfrm')
  if (!xfrm) return null
  const off = firstDescendant(xfrm, 'a:off')
  const ext = firstDescendant(xfrm, 'a:ext')
  if (!off || !ext) return null
  const x = EMU2PX(parseInt(attr(off, 'x') ?? '0', 10))
  const y = EMU2PX(parseInt(attr(off, 'y') ?? '0', 10))
  const w = EMU2PX(parseInt(attr(ext, 'cx') ?? '0', 10))
  const h = EMU2PX(parseInt(attr(ext, 'cy') ?? '0', 10))
  const rot = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000
  const flipH = attr(xfrm, 'flipH') === '1'
  const flipV = attr(xfrm, 'flipV') === '1'
  const base = { x, y, w, h, rotate: rot ? Math.round(rot) : undefined, flipH: flipH || undefined, flipV: flipV || undefined }

  // 图片
  if (node.nodeName === 'p:pic') {
    const blip = firstDescendant(node, 'a:blip')
    const embed = attr(blip, 'r:embed')
    const src = embed ? await mediaToDataUrl(ctx, embed, rels) : undefined
    if (!src) return null
    return { id: genId('i-'), type: 'image', src, ...base, name: '图片' }
  }

  // 表格
  if (node.nodeName === 'p:graphicFrame' && firstDescendant(node, 'a:tbl')) {
    const tbl = firstDescendant(node, 'a:tbl')
    if (!tbl) return null
    const gridCols = Array.from(tbl.children).filter((c) => c.nodeName === 'a:gridCol')
    const colSizes = gridCols.map((c) => EMU2PX(parseInt(attr(c, 'w') ?? '0', 10)))
    const colTotal = colSizes.reduce((a, b) => a + b, 0) || 1
    const trs = Array.from(tbl.children).filter((c) => c.nodeName === 'a:tr')
    const rowSizes = trs.map((tr) => EMU2PX(parseInt(attr(tr, 'h') ?? '0', 10)))
    const rowTotal = rowSizes.reduce((a, b) => a + b, 0) || 1
    const cells: TableCell[][] = []
    trs.forEach((tr) => {
      const row: TableCell[] = []
      for (const tc of Array.from(tr.children).filter((c) => c.nodeName === 'a:tc')) {
        const txBody = firstDescendant(tc, 'a:txBody')
        const tcPr = directChild(tc, 'a:tcPr')
        const fill = parseColor(directChild(tcPr ?? tc, 'a:solidFill'))
        const text = txBody ? txBodyToHTML(txBody).replace(/<\/?p[^>]*>/g, '').trim() : ''
        const style: TableCell['style'] = {
          color: '#333333', align: 'left', valign: 'middle',
        }
        if (fill) style.backcolor = fill
        row.push({
          text: stripTags(text),
          colspan: attr(tc, 'gridSpan') ? parseInt(attr(tc, 'gridSpan')!, 10) : undefined,
          rowspan: attr(tc, 'rowSpan') ? parseInt(attr(tc, 'rowSpan')!, 10) : undefined,
          style,
        })
      }
      cells.push(row)
    })
    const table: TableElement = {
      id: genId('tb-'), type: 'table', x, y, w, h,
      colSizes: colSizes.map((c) => c / colTotal),
      rowSizes: rowSizes.map((r) => r / rowTotal),
      cells, name: '表格',
    }
    return table
  }

  // 连接线/形状/文本
  const prstGeom = firstDescendant(node, 'a:prstGeom')
  const prst = attr(prstGeom, 'prst') ?? 'rect'
  // p:sp 内文本体为 p:txBody（presentationml 前缀），graphicFrame 内为 a:txBody
  const txBody = firstDescendant(node, 'a:txBody') ?? firstDescendant(node, 'p:txBody')
  // spPr 在 presentationml 中为 p:spPr（p:sp/p:pic/p:cxnSp 下）
  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')

  if (prst === 'line' || prst === 'straightConnector1' || node.nodeName === 'p:cxnSp') {
    const end = [flipH ? 0 : w, flipV ? 0 : h]
    return {
      id: genId('l-'), type: 'line', x, y, w, h,
      start: [flipH ? w : 0, flipV ? h : 0],
      end: [end[0], end[1]],
      lineType: 'straight',
      color: parseColor(directChild(spPr ?? node, 'a:solidFill')) ?? '#333333',
      lineWidth: 2, lineStyle: 'solid', startArrow: '', endArrow: '', name: '线条',
    }
  }

  const shapeKey = PRST_MAP[prst] ?? 'rect'

  // 填充/边框
  let fill: string | undefined
  if (spPr) {
    for (const child of Array.from(spPr.children)) {
      if (child.nodeName === 'a:solidFill') { fill = parseColor(child) ?? fill }
      if (child.nodeName === 'a:noFill') { fill = undefined }
    }
  }

  const outline = (() => {
    if (!spPr) return undefined
    const ln = directChild(spPr, 'a:ln')
    if (!ln || attr(ln, 'w') === null) return undefined
    const color = parseColor(directChild(ln, 'a:solidFill'))
    const width = EMU2PX(parseInt(attr(ln, 'w') ?? '12700', 10))
    if (!color) return undefined
    return { color, width: Math.max(1, width), style: 'solid' as const }
  })()

  // 纯文本框（rect/textbox 且无显式填充）→ 文本元素，保留富文本样式
  if (node.nodeName === 'p:sp' && (prst === 'rect' || prst === 'textbox') && txBody && !fill) {
    return {
      id: genId('t-'), type: 'text', x, y, w, h,
      content: txBodyToHTML(txBody),
      rotate: base.rotate,
      defaultColor: '#333333', lineHeight: 1.5, padding: 8, name: '文本框',
    }
  }

  const shape: ShapeElement = {
    id: genId('s-'), type: 'shape', x, y, w, h,
    shapeKey, fill: fill ?? '#d14424',
    outline: outline ?? { color: '#00000000', width: 1, style: 'solid' },
    flipH: flipH || undefined, flipV: flipV || undefined,
    rotate: base.rotate,
    align: 'center', valign: 'middle', name: '形状',
  }

  // 形状内文本
  if (txBody) {
    const plain = txBodyToHTML(txBody).replace(/<p[^>]*>|<\/p>/g, '\n').replace(/<[^>]+>/g, '').trim()
    if (plain) {
      shape.text = plain
      const firstRun = firstDescendant(txBody, 'a:r')
      shape.defaultColor = parseColor(firstRun ? directChild(firstDescendant(firstRun, 'a:rPr') ?? firstRun, 'a:solidFill') : null) ?? '#FFFFFF'
    }
  }

  return shape
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, '').trim()
}

/** 解析单页 XML */
async function parseSlideXML(
  xml: string,
  ctx: ParseContext,
  rels: Map<string, string>,
  scale: { x: number; y: number },
): Promise<Slide> {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const spTree = firstDescendant(doc.documentElement, 'p:cSld')
    ? firstDescendant(firstDescendant(doc.documentElement, 'p:cSld')!, 'p:spTree')
    : null
  const elements: PPTElement[] = []
  if (spTree) {
    for (const node of Array.from(spTree.children)) {
      if (!['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp'].includes(node.nodeName)) continue
      const el = await parseShapeNode(node, ctx, rels)
      if (!el) continue
      // 按幻灯片尺寸比例映射到 1280 画布
      el.x = Math.round(el.x * scale.x)
      el.y = Math.round(el.y * scale.y)
      el.w = Math.round(el.w * scale.x) || 40
      el.h = Math.round(el.h * scale.y) || 30
      elements.push(el)
    }
  }
  return { id: genId('slide-'), elements, background: { type: 'solid', color: '#ffffff' } }
}

/** 主入口：解析 pptx 文件 → Presentation */
export async function importPPTX(file: File): Promise<Presentation> {
  const zip = await JSZip.loadAsync(file)
  const ctx: ParseContext = { zip, mediaCache: new Map() }

  // 幻灯片尺寸
  const presXml = await zip.file('ppt/presentation.xml')?.async('text')
  if (!presXml) throw new Error('不是有效的 PPTX 文件（缺少 presentation.xml）')
  const presDoc = new DOMParser().parseFromString(presXml, 'text/xml')
  const sldSz = firstDescendant(presDoc.documentElement, 'p:sldSz')
  const srcW = EMU2PX(parseInt(attr(sldSz, 'cx') ?? '12192000', 10))
  const srcH = EMU2PX(parseInt(attr(sldSz, 'cy') ?? '6858000', 10))

  // 页面顺序
  const relsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('text') ?? ''
  const rels = parseRels(relsXml, 'ppt')
  const sldIdLst = firstDescendant(presDoc.documentElement, 'p:sldIdLst')
  const slideIds = sldIdLst
    ? Array.from(sldIdLst.getElementsByTagName('p:sldId'))
    : []
  const slidePaths = slideIds
    .map((id) => attr(id, 'r:id'))
    .map((rid) => (rid ? rels.get(rid) : undefined))
    .filter((p): p is string => Boolean(p))

  const paths = slidePaths.length
    ? slidePaths
    : Object.keys(zip.files)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => {
          const na = parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10)
          const nb = parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10)
          return na - nb
        })

  if (!paths.length) throw new Error('PPTX 中没有幻灯片')

  const scale = { x: 1280 / (srcW || 1280), y: (1280 / (srcW || 1280)) * (srcW / srcH || 16 / 9) }
  // 目标画布 1280 x (1280/ratio)，ratio = srcW/srcH
  const ratio = srcW / srcH
  scale.x = 1280 / srcW
  scale.y = (1280 / ratio) / srcH

  const slides: Slide[] = []
  for (const path of paths) {
    const xml = await zip.file(path)?.async('text')
    if (!xml) continue
    const relPath = path.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    const relText = await zip.file(relPath)?.async('text') ?? ''
    const slideRels = relText ? parseRels(relText, 'ppt/slides') : new Map<string, string>()
    slides.push(await parseSlideXML(xml, ctx, slideRels, scale))
  }

  return {
    slides,
    theme: createDefaultTheme(),
    width: 1280,
    viewportRatio: ratio,
  }
}
