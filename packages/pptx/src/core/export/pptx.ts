// PPTX 导出（pptxgenjs）
import PptxGenJS from 'pptxgenjs'
import type {
  ChartElement, ChartType, FormulaElement, ImageElement, LineElement, OoxmlFonts, PPTElement,
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

/** pptxgenjs 要求图片 data 携带 MIME base64 头（image/png;base64,...）；
 * 导出链路（split(',') / bufToBase64）产出的是裸 base64，统一在此补全。
 * base64 字符集不含分号，含 ";base64," 者必为完整 dataURL 或已带头，直接透传 */
function withImageHeader(data: string): string {
  return data.includes(';base64,') ? data : `image/png;base64,${data}`
}

/** CSS 字体栈 → pptxgenjs fontFace：取首个非通用族名并去引号。
 * PowerPoint 不识别 fallback 栈（整个字符串塞进 typeface 会被当作未知字体，
 * 中文回退到宋体），只导出真实存在的单一族名 */
const GENERIC_FAMILIES = new Set(['sans-serif', 'serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-sans-serif'])
export function fontFaceOf(stack: string | undefined): string | undefined {
  return fontFamiliesOf(stack ?? '')[0]
}

/** 内联 font-family 栈 → 族名数组（去引号、去通用族，保持栈序：latin 在前 ea 在后） */
function fontFamiliesOf(stack: string): string[] {
  return stack
    .split(',')
    .map((s) => s.trim().replace(/^['"]|['"]$/g, ''))
    .filter((n) => n && !GENERIC_FAMILIES.has(n.toLowerCase()))
}

const CJK_RE = /[㐀-䶿一-鿿　-〿＀-￯]/

/** 按 run 文本选字体：含中文用 ea 位（栈中第 2 个族名），纯西文用 latin 位（第 1 个）。
 * pptxgenjs 把 fontFace 同时写入 latin/ea/cs 三个位，中文 run 若写 latin 字体
 * （如 Arial，无 CJK 字形）会被 PowerPoint 替换回默认宋体 */
function pickFontFace(families: string[] | undefined, text: string): string | undefined {
  if (!families?.length) return undefined
  return CJK_RE.test(text) ? families[1] ?? families[0] : families[0]
}

/** run 字体 → pptxgenjs fontFace：与 theme fontScheme 字体一致时返回 +mn/+mj 占位，
 * 由 expandThemeFontPlaceholders 在产物 zip 中展开为 latin/ea/cs 引用。
 * 实证（LibreOffice）：显式写主题字体名（如微软雅黑）的中文 run 字体替换错误（回退到无关字体），
 * 经 +mn-lt/+mn-ea → theme fontScheme 解析才能命中正确替换；导入还原场景 run 字体恰来自 theme */
function runFontFace(
  families: string[] | undefined,
  text: string,
  fonts?: OoxmlFonts,
): string | undefined {
  const face = pickFontFace(families, text)
  if (!face || !fonts) return face
  const eq = (name?: string) => !!name && name.toLowerCase() === face.toLowerCase()
  if (eq(fonts.minor) || eq(fonts.minorEa)) return '+mn-lt'
  if (eq(fonts.major) || eq(fonts.majorEa)) return '+mj-lt'
  return face
}

/** pptxgenjs 三位连排（latin/ea/cs 同名同属性）→ 源文件引用形式 latin=+mn-lt / ea=cs=+mn-ea。
 * 占位符由 runFontFace 写入 fontFace（pptxgenjs 原样进 XML），此处统一展开 */
const PPTXGENJS_FONT_TRIPLET
  = '(<a:latin typeface="\\+(mn|mj)-lt" pitchFamily="34" charset="0"/><a:ea typeface="\\+\\2-lt" pitchFamily="34" charset="-122"/><a:cs typeface="\\+\\2-lt" pitchFamily="34" charset="-120"/>)'
export function expandThemeFontPlaceholders(xml: string): string {
  return xml.replace(
    new RegExp(PPTXGENJS_FONT_TRIPLET, 'g'),
    (_, _m, scope: string) =>
      `<a:latin typeface="+${scope}-lt"/><a:ea typeface="+${scope}-ea"/><a:cs typeface="+${scope}-ea"/>`,
  )
}

/** theme1.xml 注入 a:ea（pptxgenjs 只写 latin 位，且 latin 后自带空 ea）：+mn-ea 解析依赖 fontScheme 的 ea 声明。
 * latin/ea 一次重写到位，已存在的 ea（空占位）被覆盖，不会产生重复元素 */
export function injectThemeEaFonts(themeXml: string, fonts: OoxmlFonts): string {
  let out = themeXml
  for (const [scope, latin, ea] of [
    ['majorFont', fonts.major, fonts.majorEa],
    ['minorFont', fonts.minor, fonts.minorEa],
  ] as const) {
    if (!ea) continue
    const re = new RegExp(`(<a:${scope}><a:latin typeface=")[^"]*("/>)(?:<a:ea typeface="[^"]*"/>)?`)
    out = out.replace(re, `$1${latin}$2<a:ea typeface="${ea}"/>`)
  }
  return out
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

function parseRunsFromHTML(
  html: string,
  themeFonts?: OoxmlFonts,
): Array<{ align?: string; runs: TextRun[] }> {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html')
  const paragraphs: Array<{ align?: string; runs: TextRun[] }> = []

  const walkInline = (node: Node, style: Record<string, unknown>, runs: TextRun[]): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? ''
      if (text) {
        const options = { ...style }
        // 字体按文本内容逐 run 选择（中/西文用栈中不同位），fontFamilies 本身不进导出选项
        if (style.fontFamilies) options.fontFace = runFontFace(style.fontFamilies as string[], text, themeFonts)
        // 中文 run 标记 zh-CN（pptxgenjs 默认 en-US；渲染器按 run 语言选 CJK 字体回退，
        // 源文件中文 run 均为 lang="zh-CN" altLang="en-US"，缺省会导致字形度量差异）
        if (CJK_RE.test(text)) options.lang = 'zh-CN'
        delete options.fontFamilies
        runs.push({ text, options })
      }
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
    // 导入产物（TipTap 富文本）用内联样式表达加粗/斜体/下划线/删除线，标签形式兜底
    if (/font-weight:\s*(bold|[6-9]00)/i.test(inlineStyle)) next.bold = true
    if (/font-style:\s*italic/i.test(inlineStyle)) next.italic = true
    const deco = inlineStyle.match(/text-decoration(?:-line)?:\s*([^;]+)/i)?.[1] ?? ''
    if (deco.includes('underline')) next.underline = true
    if (deco.includes('line-through')) next.strike = true
    const color = inlineStyle.match(/(?:^|;)\s*color:\s*([^;]+)/i)?.[1]
    if (color) { const c = normColor(color); if (c) next.color = c }
    const fs = inlineStyle.match(/font-size:\s*([\d.]+)px/i)?.[1]
    if (fs) next.fontSize = PT(parseFloat(fs))
    const bg = inlineStyle.match(/background(?:-color)?:\s*([^;]+)/i)?.[1]
    if (bg) { const b = normColor(bg); if (b) next.highlight = b }
    const ff = inlineStyle.match(/font-family:\s*([^;]+)/i)?.[1]
    if (ff) next.fontFamilies = fontFamiliesOf(ff)
    for (const child of Array.from(el.childNodes)) walkInline(child, next, runs)
  }

  const blocks = Array.from(doc.body.children)
  const blockTags = blocks.length ? blocks : [doc.body]
  for (const block of Array.from(blockTags)) {
    const el = block as HTMLElement
    const blockStyle = el.getAttribute('style') ?? ''
    const align = blockStyle.match(/text-align:\s*(\w+)/i)?.[1]
    const runs: TextRun[] = []
    for (const child of Array.from(el.childNodes)) walkInline(child, {}, runs)
    if (runs.length) {
      runs[runs.length - 1].options.breakLine = true
      // 段级行距注入段内所有 run（pptxgenjs 从 run 选项读取段落属性）：
      // 倍数 → lineSpacingMultiple（spcPct）；px → lineSpacing（spcPts，px→pt）
      const lh = blockStyle.match(/line-height:\s*([\d.]+)(px)?/i)
      if (lh) {
        const v = parseFloat(lh[1])
        if (lh[2]) { for (const r of runs) r.options.lineSpacing = PT(v) }
        else { for (const r of runs) r.options.lineSpacingMultiple = v }
      }
      // 段级对齐注入段内所有 run（pptxgenjs 从 run 选项读段落属性，且 align 变化会拆段，
      // 段内必须统一；缺省对齐的段落不注入，避免与后续段落比较产生错误分段）
      if (align) { for (const r of runs) r.options.align = align }
      paragraphs.push({ align, runs })
    }
  }
  if (!paragraphs.length) paragraphs.push({ runs: [] })
  return paragraphs
}

/* ---------- 元素导出 ---------- */

/** 元素（含旋转）的外接框：旋转后视觉范围，中心不变。兜底图片按此框从整幅画布渲染中裁剪 */
export function rotatedBoundsOf(el: { x: number; y: number; w: number; h: number; rotate?: number }): { x: number; y: number; w: number; h: number } {
  const rad = ((el.rotate ?? 0) * Math.PI) / 180
  const cos = Math.abs(Math.cos(rad))
  const sin = Math.abs(Math.sin(rad))
  const w = el.w * cos + el.h * sin
  const h = el.w * sin + el.h * cos
  return { x: el.x + (el.w - w) / 2, y: el.y + (el.h - h) / 2, w, h }
}

/** 从画布渲染 PNG 中裁剪 [x,y,w,h]（画布坐标系）→ 裸 base64。
 * renderSlideToBlob 产出整幅画布（pixelRatio 2），直接嵌入会让整页图缩进元素框 */
async function cropCanvasPng(buf: ArrayBuffer, x: number, y: number, w: number, h: number, canvasW: number): Promise<string | null> {
  const url = URL.createObjectURL(new Blob([buf], { type: 'image/png' }))
  try {
    const img = new Image()
    // blob URL 本地解码为毫秒级；超时兜底防环境（无资源加载能力的测试环境）挂起
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('image decode timeout')), 3000)
      img.onload = () => { clearTimeout(timer); resolve(null) }
      img.onerror = () => { clearTimeout(timer); reject(new Error('image decode failed')) }
      img.src = url
    })
    const scale = img.naturalWidth / canvasW
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(w * scale))
    canvas.height = Math.max(1, Math.round(h * scale))
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.drawImage(img, x * scale, y * scale, w * scale, h * scale, 0, 0, canvas.width, canvas.height)
    return canvas.toDataURL('image/png').split(',')[1]
  } catch {
    return null
  } finally {
    URL.revokeObjectURL(url)
  }
}

async function elementToImage(el: PPTElement, slide: Slide, presentation: Presentation): Promise<string | null> {
  // 用离屏渲染兜底导出（不支持的原生形状/公式等）
  try {
    const onlyElSlide: Slide = { ...slide, elements: [el] }
    // 透明背景：兜底图内嵌回页面后叠在其他元素上，白底会盖住下层相邻文字
    const blob = await renderSlideToBlob(onlyElSlide, presentation, 'png', { transparent: true })
    // 旋转元素在画布上已按最终角度渲染：裁旋转外接框，嵌入时不再二次旋转
    const b = rotatedBoundsOf(el)
    const buf = await blob.arrayBuffer()
    return await cropCanvasPng(buf, b.x, b.y, b.w, b.h, presentation.width)
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
  const paragraphs = parseRunsFromHTML(el.content, pres.theme.ooxmlFonts)
  const runs: TextRun[] = []
  for (const p of paragraphs) {
    for (const run of p.runs) {
      runs.push({ text: run.text, options: { ...run.options } })
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
      // 行距倍数：pptxgenjs 有效选项名是 lineSpacingMultiple（spcPct），lineHeight 会被静默忽略
      lineSpacingMultiple: el.lineHeight ?? 1.5,
      // 垂直对齐随导入的 bodyPr anchor（默认 OOXML 顶对齐）；写死 top 会让 ctr 文本重排错位
      valign: el.valign ?? 'top',
      // 内边距对齐编辑器渲染（px→pt），否则 PowerPoint 默认 insets 或 0 与画布不一致
      margin: PT(el.padding ?? 8),
      isTextBox: true,
      // 与导入映射对应：autoSize（normAutofit 缩字适应框）→ shrink，autoFit（spAutoFit 框随文本）→ resize
      fit: el.autoSize ? 'shrink' : el.autoFit ? 'resize' : undefined,
    },
    runs,
  }
}

/** 导出形状（export 仅为测试暴露：path 形状走图片兜底是回归关键点） */
export async function exportShape(pptx: PptxGenJS, slide: Slide, el: ShapeElement, pres: Presentation): Promise<object> {
  void pptx
  const nativeKey = SHAPE_MAP[el.shapeKey]
  const fill = typeof el.fill === 'string' ? normColor(el.fill) : undefined
  const isGradient = typeof el.fill === 'object' && el.fill !== null
  // 自定义 path（custGeom 导入产出）无法用原生预设表达：走离屏截图兜底，避免静默退化为矩形
  if (nativeKey && !isGradient && !el.path) {
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
              fontFace: fontFaceOf(el.defaultFontName),
              align: el.align ?? 'center', valign: el.valign ?? 'middle',
              lineHeight: el.lineHeight ?? 1.2,
            },
          }
        : undefined,
    }
  }
  const data = await elementToImage(el, slide, pres)
  // 兜底图为画布裁剪（已含旋转视觉）：按旋转外接框定位，rotate 归零避免二次旋转
  const b = rotatedBoundsOf(el)
  return {
    type: 'image',
    data,
    props: { x: IN(b.x), y: IN(b.y), w: IN(b.w), h: IN(b.h), rotate: 0 },
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
    // 兜底图为画布裁剪（已含旋转视觉）：按旋转外接框定位，rotate 归零避免二次旋转
    const b = rotatedBoundsOf(chart)
    return {
      type: 'image',
      data,
      props: { x: IN(b.x), y: IN(b.y), w: IN(b.w), h: IN(b.h), rotate: 0 },
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
  // 兜底图为画布裁剪（已含旋转视觉）：按旋转外接框定位，rotate 归零避免二次旋转
  const b = rotatedBoundsOf(el)
  return {
    type: 'image',
    data,
    props: { x: IN(b.x), y: IN(b.y), w: IN(b.w), h: IN(b.h), rotate: 0 },
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

/** onProgress：每页导出完成后回调 (done, total)，供对话框显示进度 */
export async function exportPPTX(
  presentation: Presentation,
  docName: string,
  onProgress?: (done: number, total: number) => void,
): Promise<void> {
  const pptx = new PptxGenJS()
  const w = presentation.width
  const h = Math.round(presentation.width / presentation.viewportRatio)
  pptx.defineLayout({ name: 'EFLINK', width: IN(w), height: IN(h) })
  pptx.layout = 'EFLINK'
  pptx.author = 'eflink-pptx'
  pptx.title = docName
  // 主题字体还原到 latin 位（ea 位由 injectThemeEaFonts 在产物 zip 中补写）
  const fonts = presentation.theme.ooxmlFonts
  if (fonts) pptx.theme = { headFontFace: fonts.major, bodyFontFace: fonts.minor }

  for (const [i, slide] of presentation.slides.entries()) {
    const s = pptx.addSlide()

    // 背景
    const bg = slide.background
    if (bg?.type === 'solid' && bg.color && bg.color !== '#ffffff') {
      s.background = { color: normColor(bg.color) ?? 'FFFFFF' }
    } else if (bg && (bg.type === 'gradient' || (bg.type === 'image' && bg.image?.src))) {
      const blob = await renderSlideToBlob({ ...slide, elements: [] }, presentation, 'jpeg')
      const buf = new Uint8Array(await blob.arrayBuffer())
      s.background = { data: withImageHeader(bufToBase64(buf.buffer as ArrayBuffer)) }
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
            s.addImage({ ...(sp.props as object), data: withImageHeader(sp.data) })
          }
          break
        case 'image':
          if (sp.data) s.addImage({ ...(sp.props as object), data: withImageHeader(sp.data) })
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
    onProgress?.(i + 1, presentation.slides.length)
  }

  const blob = await pptx.write({ outputType: 'blob' })
  // 主题字体后处理：run 占位展开 + theme ea 注入（仅导入产物有 ooxmlFonts 时）
  const finalBlob = fonts ? await patchThemeFonts(blob as Blob, fonts) : (blob as Blob)
  downloadBlob(finalBlob, `${docName || '未命名'}.pptx`)
}

/** 产物 zip 后处理：+mn/+mj 占位展开为 latin/ea/cs 引用、theme1.xml 补写 a:ea。
 * pptxgenjs 不支持 per-run 区分 latin/ea 位，也不写 fontScheme 的 ea 声明 */
async function patchThemeFonts(blob: Blob, fonts: OoxmlFonts): Promise<Blob> {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(blob)
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith('.xml')) continue
    const file = zip.file(path)
    if (!file) continue
    let xml = await file.async('string')
    if (path === 'ppt/theme/theme1.xml') xml = injectThemeEaFonts(xml, fonts)
    xml = expandThemeFontPlaceholders(xml)
    zip.file(path, xml)
  }
  return zip.generateAsync({ type: 'blob' })
}

/** 导出前检查用的纯函数（供单测） */
export { normColor, parseRunsFromHTML, IN as pxToInch, PT as pxToPt, hexToRgb as _hexToRgb, withImageHeader }
