// PPTX 导出（pptxgenjs）
import PptxGenJS from 'pptxgenjs'
import type {
  ChartElement, ChartType, FormulaElement, ImageElement, LineElement, PPTElement,
  Presentation, ShapeElement, Slide, TableElement, TextElement,
} from '../../types/slides'
import { PX_PER_INCH } from '../../types/slides'
import { blendToWhite, hexToRgb } from '../utils/color'
import { normalizeChartElement } from '../chart/migrate'
import { DEFAULT_CHART_ELEMENTS } from '../chart/chartOptions'
import { renderSlideToBlob } from './image'
import { downloadBlob } from './json'

const IN = (px: number) => px / PX_PER_INCH
const PT = (px: number) => px * 0.75

/** ArrayBuffer → base64（分块避免栈溢出） */
function bufToBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf)
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}

/** 颜色归一化为 pptxgenjs 需要的 RRGGBB（无 #，透明混合白底） */
function normColor(color: string | undefined): string | undefined {
  if (!color) return undefined
  const c = color.trim()
  if (c === '#00000000' || c === 'transparent') return undefined
  if (/^#[0-9a-f]{8}$/i.test(c)) {
    const a = parseInt(c.slice(7, 9), 16) / 255
    return blendToWhite(c.slice(0, 7), a).slice(1).toUpperCase()
  }
  const rgba = c.match(/rgba?\(([^)]+)\)/i)
  if (rgba) {
    const parts = rgba[1].split(',').map((s) => parseFloat(s.trim()))
    const hex = `#${[parts[0], parts[1], parts[2]].map((n) => Math.round(n || 0).toString(16).padStart(2, '0')).join('')}`
    const alpha = parts[3] === undefined ? 1 : parts[3]
    return blendToWhite(hex, alpha).slice(1).toUpperCase()
  }
  if (/^#[0-9a-f]{3,6}$/i.test(c)) {
    let h = c.slice(1)
    if (h.length === 3) h = h.split('').map((ch) => ch + ch).join('')
    return h.toUpperCase()
  }
  return undefined
}

/** 自定义形状 key → pptxgenjs ShapeType */
const SHAPE_MAP: Record<string, string> = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle',
  rtTriangle: 'rtTriangle', parallelogram: 'parallelogram', trapezoid: 'trapezoid',
  diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', heptagon: 'heptagon',
  octagon: 'octagon', donut: 'donut', pie: 'pie', chord: 'chord', arc: 'arc',
  frame: 'frame', halfFrame: 'halfFrame', can: 'can', cube: 'cube', plaque: 'plaque',
  cross: 'cross', plus: 'mathPlus',
  arrowRight: 'rightArrow', arrowLeft: 'leftArrow', arrowUp: 'upArrow', arrowDown: 'downArrow',
  arrowLeftRight: 'leftRightArrow', arrowUpDown: 'upDownArrow', arrowQuad: 'quadArrow',
  bentArrow: 'bentArrow', curvedRightArrow: 'curvedRightArrow', chevron: 'chevron', homePlate: 'homePlate',
  star4: 'star4', star5: 'star5', star6: 'star6', star8: 'star8', star16: 'star16',
  heart: 'heart', lightning: 'lightningBolt', sun: 'sun', moon: 'moon', cloud: 'cloud',
  smile: 'smileyFace',
  flowProcess: 'flowChartProcess', flowDecision: 'flowChartDecision', flowData: 'flowChartData',
  flowDocument: 'flowChartDocument', flowPredefined: 'flowChartPredefinedProcess',
  flowTerminal: 'flowChartTerminator', flowConnector: 'flowChartConnector',
  flowOffpage: 'flowChartOffpageConnector', flowDelay: 'flowChartDelay',
  flowManualInput: 'flowChartManualInput', flowManualOperation: 'flowChartManualOperation',
  flowMerge: 'flowChartMerge', flowOr: 'flowChartOr', flowExtract: 'flowChartExtract',
  callout1: 'wedgeRectCallout', callout2: 'wedgeRoundRectCallout', callout3: 'wedgeEllipseCallout',
  mathPlus: 'mathPlus', mathMinus: 'mathMinus', mathMultiply: 'mathMultiply',
  mathDivide: 'mathDivide', mathEqual: 'mathEqual', mathNotEqual: 'mathNotEqual',
}

/* ---------- 富文本 HTML → runs ---------- */

interface TextRun {
  text: string
  options: Record<string, unknown>
}

function parseRunsFromHTML(html: string): Array<{ align?: string; runs: TextRun[] }> {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html')
  const paragraphs: Array<{ align?: string; runs: TextRun[] }> = []

  const walkInline = (node: Node, style: Record<string, unknown>, runs: TextRun[]): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? ''
      if (text) runs.push({ text, options: { ...style } })
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const el = node as HTMLElement
    if (el.tagName === 'BR') {
      runs.push({ text: '', options: { breakLine: true } })
      return
    }
    const next = { ...style }
    switch (el.tagName) {
      case 'STRONG': case 'B': next.bold = true; break
      case 'EM': case 'I': next.italic = true; break
      case 'U': next.underline = true; break
      case 'S': case 'DEL': next.strike = true; break
      case 'SUP': next.superscript = true; break
      case 'SUB': next.subscript = true; break
      case 'CODE': next.code = true; break
      case 'A': next.link = el.getAttribute('href') ?? undefined; break
      default: break
    }
    const inlineStyle = el.getAttribute('style') ?? ''
    const color = inlineStyle.match(/(?:^|;)\s*color:\s*([^;]+)/i)?.[1]
    if (color) { const c = normColor(color); if (c) next.color = c }
    const fs = inlineStyle.match(/font-size:\s*([\d.]+)px/i)?.[1]
    if (fs) next.fontSize = PT(parseFloat(fs))
    const bg = inlineStyle.match(/background(?:-color)?:\s*([^;]+)/i)?.[1]
    if (bg) { const b = normColor(bg); if (b) next.highlight = b }
    for (const child of Array.from(el.childNodes)) walkInline(child, next, runs)
  }

  const blocks = Array.from(doc.body.children)
  const blockTags = blocks.length ? blocks : [doc.body]
  for (const block of Array.from(blockTags)) {
    const el = block as HTMLElement
    const align = (el.getAttribute('style') ?? '').match(/text-align:\s*(\w+)/i)?.[1]
    const runs: TextRun[] = []
    for (const child of Array.from(el.childNodes)) walkInline(child, {}, runs)
    if (runs.length) {
      runs[runs.length - 1].options.breakLine = true
      paragraphs.push({ align, runs })
    }
  }
  if (!paragraphs.length) paragraphs.push({ runs: [] })
  return paragraphs
}

/* ---------- 元素导出 ---------- */

async function elementToImage(el: PPTElement, slide: Slide, presentation: Presentation): Promise<string | null> {
  // 用离屏渲染兜底导出（不支持的原生形状/公式等）
  try {
    const onlyElSlide: Slide = { ...slide, elements: [el] }
    const blob = await renderSlideToBlob(onlyElSlide, presentation, 'png')
    const buf = await blob.arrayBuffer()
    return bufToBase64(buf)
  } catch {
    return null
  }
}

async function ensureDataUrl(src: string): Promise<string | null> {
  if (src.startsWith('data:')) return src
  try {
    const resp = await fetch(src)
    const blob = await resp.blob()
    return await new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(reader.result as string)
      reader.onerror = reject
      reader.readAsDataURL(blob)
    })
  } catch {
    return null
  }
}

async function exportText(pptx: PptxGenJS, slide: Slide, el: TextElement, pres: Presentation): Promise<object> {
  void pptx
  void slide
  void pres
  const paragraphs = parseRunsFromHTML(el.content)
  const runs: TextRun[] = []
  for (const p of paragraphs) {
    for (const run of p.runs) {
      const opts = { ...run.options }
      if (p.align && Object.keys(opts).length <= 1) opts.align = p.align
      runs.push({ text: run.text, options: opts })
    }
  }
  const baseColor = normColor(el.defaultColor)
  return {
    type: 'text',
    el,
    textProps: {
      x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h),
      rotate: el.rotate ?? 0,
      fontSize: PT(18),
      color: baseColor,
      lineHeight: el.lineHeight ?? 1.5,
      valign: 'top',
      margin: 0,
      isTextBox: true,
      autoFit: false,
    },
    runs,
  }
}

async function exportShape(pptx: PptxGenJS, slide: Slide, el: ShapeElement, pres: Presentation): Promise<object> {
  void pptx
  const nativeKey = SHAPE_MAP[el.shapeKey]
  const fill = typeof el.fill === 'string' ? normColor(el.fill) : undefined
  const isGradient = typeof el.fill === 'object' && el.fill !== null
  if (nativeKey && !isGradient) {
    return {
      type: 'shape',
      native: nativeKey,
      props: {
        x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h),
        rotate: el.rotate ?? 0,
        fill: fill ? { color: fill } : { type: 'none' as const },
        line: el.outline && el.outline.width > 0 && normColor(el.outline.color)
          ? {
              color: normColor(el.outline.color),
              width: PT(el.outline.width),
              dashType: el.outline.style === 'dashed' ? 'dash' : el.outline.style === 'dotted' ? 'sysDot' : 'solid',
            }
          : { type: 'none' as const },
        flipH: el.flipH, flipV: el.flipV,
      },
      text: el.text
        ? {
            text: el.text,
            options: {
              fontSize: PT(el.fontSize ?? 18), color: normColor(el.defaultColor) ?? 'FFFFFF',
              align: el.align ?? 'center', valign: el.valign ?? 'middle',
              lineHeight: el.lineHeight ?? 1.2,
            },
          }
        : undefined,
    }
  }
  const data = await elementToImage(el, slide, pres)
  return {
    type: 'image',
    data,
    props: { x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h), rotate: el.rotate ?? 0 },
  }
}

async function exportImage(_slide: Slide, el: ImageElement, _pres: Presentation): Promise<object> {
  const data = await ensureDataUrl(el.src)
  void _slide
  void _pres
  return {
    type: 'image',
    data: data ? data.split(',')[1] : null,
    props: {
      x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h),
      rotate: el.rotate ?? 0, flipH: el.flipH, flipV: el.flipV,
      rounding: el.radius && el.radius > 0 ? true : false,
    },
  }
}

async function exportLine(_slide: Slide, el: LineElement): Promise<object> {
  void _slide
  return {
    type: 'line',
    props: {
      x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h),
      flipV: el.end[1] < el.start[1],
      flipH: el.end[0] < el.start[0],
      line: {
        color: normColor(el.color) ?? '333333',
        width: PT(el.lineWidth),
        dashType: el.lineStyle === 'dashed' ? 'dash' : el.lineStyle === 'dotted' ? 'sysDot' : 'solid',
        beginArrowType: el.startArrow && el.startArrow !== 'none' && el.startArrow !== 'dot' ? 'triangle' : 'none',
        endArrowType: el.endArrow && el.endArrow !== 'none' && el.endArrow !== 'dot' ? 'triangle' : 'none',
      },
    },
  }
}

async function exportTable(_slide: Slide, el: TableElement): Promise<object> {
  void _slide
  const rows = el.cells
    .filter((row): row is NonNullable<typeof row> => Boolean(row))
    .map((row) => row.filter((cell): cell is NonNullable<typeof cell> => Boolean(cell)).map((cell) => ({
      text: cell.text,
      options: {
        colspan: cell.colspan && cell.colspan > 1 ? cell.colspan : undefined,
        rowspan: cell.rowspan && cell.rowspan > 1 ? cell.rowspan : undefined,
        fill: cell.style?.backcolor ? { color: normColor(cell.style.backcolor) ?? 'FFFFFF' } : undefined,
        color: normColor(cell.style?.color) ?? '333333',
        bold: cell.style?.bold, italic: cell.style?.italic, underline: cell.style?.underline,
        fontSize: cell.style?.fontsize ? PT(cell.style.fontsize) : PT(14),
        align: cell.style?.align ?? 'left',
        valign: cell.style?.valign === 'top' ? 'top' : cell.style?.valign === 'bottom' ? 'bottom' : 'middle',
      },
    })))
  return {
    type: 'table',
    rows,
    props: {
      x: IN(el.x), y: IN(el.y), w: IN(el.w),
      colW: el.colSizes.map((c) => IN(c * el.w)),
      rowH: el.rowSizes.map((r) => IN(r * el.h)),
      border: { type: 'solid', color: normColor(el.outline?.color) ?? 'D0D0D0', pt: el.outline?.width ?? 1 },
    },
  }
}

/** 14 种图表类型 → pptxgenjs 原生图表参数；无原生支持的类型返回 null（导出为图片） */
export function chartNativeSpec(chartType: ChartType): {
  type: string
  barDir?: 'col' | 'bar'
  barGrouping?: string
  lineDataSymbol?: string
  lineSize?: number
} | null {
  // pptxgenjs 不支持雷达图
  if (chartType === 'radar') return null
  const specs: Record<string, { type: string; barDir?: 'col' | 'bar'; barGrouping?: string; lineDataSymbol?: string; lineSize?: number }> = {
    'bar-cluster': { type: 'bar', barDir: 'col' },
    'bar-stack': { type: 'bar', barDir: 'col', barGrouping: 'stacked' },
    'bar-percent': { type: 'bar', barDir: 'col', barGrouping: 'percentStacked' },
    'bar-horizontal': { type: 'bar', barDir: 'bar' },
    'bar-horizontal-stack': { type: 'bar', barDir: 'bar', barGrouping: 'stacked' },
    line: { type: 'line' },
    'line-stack': { type: 'line', barGrouping: 'stacked' },
    'line-marker': { type: 'line', lineDataSymbol: 'circle' },
    'pie': { type: 'pie' },
    'pie-doughnut': { type: 'doughnut' },
    'area': { type: 'area' },
    'area-stack': { type: 'area', barGrouping: 'stacked' },
    // pptxgenjs 无原生散点图，用无连线折线近似（既有行为）
    'scatter': { type: 'line', lineSize: 0 },
  }
  return specs[chartType] ?? { type: 'bar', barDir: 'col' }
}

async function exportChart(slide: Slide, el: ChartElement, presentation: Presentation): Promise<object> {
  // 防御性规范化：旧文档未经 loadDocument 规范化时也能正确导出
  const chart = normalizeChartElement(el)
  const elements = chart.elements ?? DEFAULT_CHART_ELEMENTS
  const asImage = async (): Promise<object> => {
    const data = await elementToImage(chart, slide, presentation)
    return {
      type: 'image',
      data,
      props: { x: IN(chart.x), y: IN(chart.y), w: IN(chart.w), h: IN(chart.h), rotate: chart.rotate ?? 0 },
    }
  }
  // 雷达图 pptxgenjs 不支持、趋势线原生图表无法表达 → 导出为图片保证所见即所得
  const native = chartNativeSpec(chart.chartType)
  if (!native || elements.trendline) return asImage()

  return {
    type: 'chart',
    chartType: native.type,
    data: chart.data.series.map((s) => ({ name: s.name, labels: chart.data.labels, values: s.values })),
    props: {
      x: IN(chart.x), y: IN(chart.y), w: IN(chart.w), h: IN(chart.h),
      barDir: native.barDir,
      barGrouping: native.barGrouping,
      lineDataSymbol: native.lineDataSymbol,
      lineSize: native.lineSize,
      chartColors: chart.chartColors ?? undefined,
      showLegend: elements.legend,
      showValue: elements.dataLabel,
      showTitle: Boolean(chart.title), title: chart.title ?? undefined,
      valGridLine: elements.gridLine ? undefined : { style: 'none' },
      catAxisLineShow: elements.axis,
      valAxisLineShow: elements.axis,
      ...(elements.axisTitle ? { showValAxisTitle: true, valAxisTitle: '数值' } : {}),
      catAxisLabelFontSize: PT(chart.fontSize ?? 14),
      valAxisLabelFontSize: PT(chart.fontSize ?? 14),
      legendFontSize: PT(chart.fontSize ?? 14),
    },
  }
}

async function exportFormula(slide: Slide, el: FormulaElement, pres: Presentation): Promise<object> {
  const data = await elementToImage(el, slide, pres)
  return {
    type: 'image',
    data,
    props: { x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h), rotate: el.rotate ?? 0 },
  }
}

async function exportMedia(el: Extract<PPTElement, { type: 'video' | 'audio' }>): Promise<object> {
  const data = await ensureDataUrl(el.src)
  return {
    type: 'media',
    mediaType: el.type,
    data: data && data.startsWith('data:') ? data.split(',')[1] : null,
    props: { x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h) },
  }
}

/* ---------- 主入口 ---------- */

export async function exportPPTX(presentation: Presentation, docName: string): Promise<void> {
  const pptx = new PptxGenJS()
  const w = presentation.width
  const h = Math.round(presentation.width / presentation.viewportRatio)
  pptx.defineLayout({ name: 'EFLINK', width: IN(w), height: IN(h) })
  pptx.layout = 'EFLINK'
  pptx.author = 'eflink-pptx'
  pptx.title = docName

  for (const slide of presentation.slides) {
    const s = pptx.addSlide()

    // 背景
    const bg = slide.background
    if (bg?.type === 'solid' && bg.color && bg.color !== '#ffffff') {
      s.background = { color: normColor(bg.color) ?? 'FFFFFF' }
    } else if (bg && (bg.type === 'gradient' || (bg.type === 'image' && bg.image?.src))) {
      const blob = await renderSlideToBlob({ ...slide, elements: [] }, presentation, 'jpeg')
      const buf = new Uint8Array(await blob.arrayBuffer())
      s.background = { data: bufToBase64(buf.buffer as ArrayBuffer) }
    }

    for (const el of slide.elements) {
      let spec: object | null = null
      switch (el.type) {
        case 'text': spec = await exportText(pptx, slide, el, presentation); break
        case 'shape': spec = await exportShape(pptx, slide, el, presentation); break
        case 'image': spec = await exportImage(slide, el, presentation); break
        case 'line': spec = await exportLine(slide, el); break
        case 'table': spec = await exportTable(slide, el); break
        case 'chart': spec = await exportChart(slide, el, presentation); break
        case 'formula': spec = await exportFormula(slide, el, presentation); break
        case 'video': case 'audio': spec = await exportMedia(el); break
      }
      if (!spec) continue
      const sp = spec as {
        type: string; native?: string; props?: Record<string, unknown>; textProps?: Record<string, unknown>
        runs?: Array<{ text: string; options?: Record<string, unknown> }>; text?: { text: string; options: Record<string, unknown> }
        rows?: Array<Array<{ text: string; options?: Record<string, unknown> }>>
        data?: string | null; chartType?: string
        mediaType?: string
      }
      switch (sp.type) {
        case 'text':
          s.addText((sp.runs ?? []).map((r) => ({ text: r.text, options: r.options ?? {} })), sp.textProps as never)
          break
        case 'shape':
          if (sp.native) {
            s.addText(sp.text ? [{ text: sp.text.text, options: sp.text.options }] : [], {
              ...(sp.props as object),
              shape: sp.native as never,
            })
          } else if (sp.data) {
            s.addImage({ ...(sp.props as object), data: sp.data })
          }
          break
        case 'image':
          if (sp.data) s.addImage({ ...(sp.props as object), data: sp.data })
          break
        case 'line':
          s.addShape('line' as never, sp.props as never)
          break
        case 'table':
          if (sp.rows) s.addTable(sp.rows as never, sp.props as never)
          break
        case 'chart':
          if (sp.chartType && sp.data) {
            s.addChart(sp.chartType as never, sp.data as never, sp.props as never)
          }
          break
        case 'media':
          if (sp.data) {
            s.addMedia({ type: (sp.mediaType ?? 'video') as never, data: sp.data, ...(sp.props as object) })
          }
          break
      }
    }

    if (slide.note) s.addNotes(slide.note)
  }

  await pptx.write({ outputType: 'blob' }).then((blob) => {
    downloadBlob(blob as Blob, `${docName || '未命名'}.pptx`)
  })
}

/** 导出前检查用的纯函数（供单测） */
export { normColor, parseRunsFromHTML, IN as pxToInch, PT as pxToPt, hexToRgb as _hexToRgb }
