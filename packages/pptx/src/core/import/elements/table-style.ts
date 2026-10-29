/** ppt/tableStyles.xml 解析：tblPr(firstRow/bandRow…) + tableStyleId → 单元格有效样式。
 * 层叠顺序（低→高）：wholeTbl < band1H/2H(bandRow) < band1V/2V(bandCol) < firstCol < lastCol < firstRow < lastRow；
 * 单元格级 tcPr 填充 / run 显式样式始终优先于样式层。 */
import { attr, descendants, directChild, parseXML } from '../xml'
import { refColor, resolveColor, styleMatrixEntry } from '../styles'
import { builtinTableStyleXml } from './builtin-table-styles'
import type { PptxPackage } from '../package'
import type { PptxTheme } from '../theme'

export interface TableStylePart {
  /** 单元格填充色 */
  fill?: string
  /** 文字颜色（tcTxStyle 颜色子节点） */
  color?: string
  bold?: boolean
  /** 网格线颜色（tcBdr 首个带填充的边框） */
  border?: string
}

export interface TableStyleDef {
  whole?: TableStylePart
  firstRow?: TableStylePart
  lastRow?: TableStylePart
  firstCol?: TableStylePart
  lastCol?: TableStylePart
  band1H?: TableStylePart
  band2H?: TableStylePart
  band1V?: TableStylePart
  band2V?: TableStylePart
}

export interface TableStyleFlags {
  firstRow: boolean
  lastRow: boolean
  firstCol: boolean
  lastCol: boolean
  bandRow: boolean
  bandCol: boolean
}

export interface ResolvedTableStyle {
  flags: TableStyleFlags
  def: TableStyleDef | undefined
}

/** 包级缓存：一个 pptx 通常只有一份 tableStyles.xml，逐表重复解析浪费 */
const styleCache = new WeakMap<PptxPackage, Promise<{ defs: Map<string, TableStyleDef>; defId: string | null }>>()

/** 单条边框（a:ln 或 a:lnRef 引用主题线型）→ 颜色；noFill/缺省返回 undefined */
function edgeBorderColor(edge: Element, theme: PptxTheme): string | undefined {
  const ln = directChild(edge, 'a:ln')
  if (ln) {
    if (directChild(ln, 'a:noFill')) return undefined
    return resolveColor(directChild(ln, 'a:solidFill'), theme)
  }
  const lnRef = directChild(edge, 'a:lnRef')
  if (!lnRef) return undefined
  const idx = parseInt(attr(lnRef, 'idx') ?? '0', 10)
  if (idx < 1) return undefined
  const matrix = styleMatrixEntry(theme, 'line', idx)
  return (matrix ? resolveColor(directChild(matrix, 'a:solidFill'), theme, refColor(lnRef)) : undefined)
    ?? resolveColor(lnRef, theme) ?? undefined
}

/** 网格线优先于外框：单元格之间可见的是 inside 边框（左/右/上/下作为兜底） */
function borderColorOf(tcBdr: Element, theme: PptxTheme): string | undefined {
  for (const name of ['insideH', 'insideV', 'left', 'right', 'top', 'bottom']) {
    const edge = directChild(tcBdr, `a:${name}`)
    if (!edge) continue
    const color = edgeBorderColor(edge, theme)
    if (color) return color
  }
  return undefined
}

/** defaultFill：表级底色（a:tblBg），仅用于 wholeTbl 未声明填充时兜底 */
function parsePart(parent: Element | null, theme: PptxTheme, defaultFill?: string): TableStylePart | undefined {
  if (!parent) return undefined
  const part: TableStylePart = {}
  const tx = directChild(parent, 'a:tcTxStyle')
  if (tx) {
    if (attr(tx, 'b') === 'on') part.bold = true
    for (const child of Array.from(tx.children)) {
      if (['schemeClr', 'srgbClr', 'prstClr'].includes(child.localName)) {
        const color = resolveColor(tx, theme)
        if (color) part.color = color
        break
      }
    }
  }
  const tcStyle = directChild(parent, 'a:tcStyle')
  if (tcStyle) {
    const fillEl = directChild(tcStyle, 'a:fill')
    const fill = fillEl ? resolveColor(directChild(fillEl, 'a:solidFill'), theme) : undefined
    if (fill) part.fill = fill
    const tcBdr = directChild(tcStyle, 'a:tcBdr')
    if (tcBdr) {
      const border = borderColorOf(tcBdr, theme)
      if (border) part.border = border
    }
  }
  if (!part.fill && defaultFill) part.fill = defaultFill
  return part.fill || part.color || part.bold || part.border ? part : undefined
}

/** tableStyles.xml → 样式表（styleId 含大括号，归一化后作 key） */
async function loadStyleTable(pkg: PptxPackage, theme: PptxTheme) {
  let cached = styleCache.get(pkg)
  if (!cached) {
    cached = (async () => {
      const xml = await pkg.text('ppt/tableStyles.xml')
      const defs = new Map<string, TableStyleDef>()
      let defId: string | null = null
      if (xml) {
        const root = parseXML(xml).documentElement
        defId = attr(root, 'def')?.replace(/[{}]/g, '') ?? null
        for (const styleEl of descendants(root, 'a:tblStyle')) {
          const id = attr(styleEl, 'styleId')?.replace(/[{}]/g, '')
          if (!id) continue
          defs.set(id, {
            whole: parsePart(directChild(styleEl, 'a:wholeTbl'), theme),
            firstRow: parsePart(directChild(styleEl, 'a:firstRow'), theme),
            lastRow: parsePart(directChild(styleEl, 'a:lastRow'), theme),
            firstCol: parsePart(directChild(styleEl, 'a:firstCol'), theme),
            lastCol: parsePart(directChild(styleEl, 'a:lastCol'), theme),
            band1H: parsePart(directChild(styleEl, 'a:band1H'), theme),
            band2H: parsePart(directChild(styleEl, 'a:band2H'), theme),
            band1V: parsePart(directChild(styleEl, 'a:band1V'), theme),
            band2V: parsePart(directChild(styleEl, 'a:band2V'), theme),
          })
        }
      }
      return { defs, defId }
    })()
    styleCache.set(pkg, cached)
  }
  return cached
}

/** 内置样式缓存：同一主题色组合下各 styleId 的解析结果（null = 未支持，避免重复解析） */
const builtinCache = new Map<string, TableStyleDef | null>()

function themeSignature(theme: PptxTheme): string {
  return `${Object.entries(theme.colorMap).sort().map(([k, v]) => `${k}=${v}`).join(',')}|${Object.values(theme.schemeColors).join(',')}`
}

/** PowerPoint 自带样式 GUID 不在 tableStyles.xml 里：按规范基础定义 + 主题色替换解析 */
function builtinTableStyle(styleId: string, theme: PptxTheme): TableStyleDef | undefined {
  const id = styleId.replace(/[{}]/g, '').toUpperCase()
  const key = `${id}|${themeSignature(theme)}`
  if (builtinCache.has(key)) return builtinCache.get(key) ?? undefined
  const xml = builtinTableStyleXml(id)
  let def: TableStyleDef | undefined
  if (xml) {
    const root = parseXML(xml).documentElement
    // 表级底色（tblBg）：单元格未声明填充时透出，作为 wholeTbl 兜底
    const tblBg = directChild(root, 'a:tblBg')
    const bgFn = tblBg ? (directChild(tblBg, 'a:fillRef') ?? directChild(tblBg, 'a:solidFill')) : null
    const bgColor = bgFn ? resolveColor(bgFn, theme, refColor(bgFn)) : undefined
    def = {
      whole: parsePart(directChild(root, 'a:wholeTbl'), theme, bgColor),
      firstRow: parsePart(directChild(root, 'a:firstRow'), theme),
      lastRow: parsePart(directChild(root, 'a:lastRow'), theme),
      firstCol: parsePart(directChild(root, 'a:firstCol'), theme),
      lastCol: parsePart(directChild(root, 'a:lastCol'), theme),
      band1H: parsePart(directChild(root, 'a:band1H'), theme),
      band2H: parsePart(directChild(root, 'a:band2H'), theme),
      band1V: parsePart(directChild(root, 'a:band1V'), theme),
      band2V: parsePart(directChild(root, 'a:band2V'), theme),
    }
  }
  builtinCache.set(key, def ?? null)
  return def
}

/** tbl 元素 → 生效的表格样式（tblPr 标志位 + tableStyleId）。
 * 仅显式 tableStyleId 才套用样式：tblStyleLst@def 只定义新建表的默认样式，
 * 未引用样式的表（空 tblPr）不回退套用，否则会给纯边框表错误加填充 */
export async function resolveTableStyle(tbl: Element, pkg: PptxPackage, theme: PptxTheme): Promise<ResolvedTableStyle | undefined> {
  const tblPr = directChild(tbl, 'a:tblPr')
  if (!tblPr) return undefined
  const idRaw = directChild(tblPr, 'a:tableStyleId')?.textContent?.trim()
  if (!idRaw) return undefined
  const flags: TableStyleFlags = {
    firstRow: attr(tblPr, 'firstRow') === '1',
    lastRow: attr(tblPr, 'lastRow') === '1',
    firstCol: attr(tblPr, 'firstCol') === '1',
    lastCol: attr(tblPr, 'lastCol') === '1',
    bandRow: attr(tblPr, 'bandRow') === '1',
    bandCol: attr(tblPr, 'bandCol') === '1',
  }
  const { defs } = await loadStyleTable(pkg, theme)
  const def = defs.get(idRaw.replace(/[{}]/g, '')) ?? builtinTableStyle(idRaw, theme)
  if (!def) return undefined
  return { flags, def }
}

/** 单元格网格位 (r, c) 命中的样式层（低→高）。
 * 带状起点：表头/首列开启时从其下一行/列起算 band1（PowerPoint 表头行不参与带状循环） */
export function partsForCell(style: ResolvedTableStyle, r: number, c: number, rows: number, cols: number): TableStylePart[] {
  const { flags, def } = style
  if (!def) return []
  const parts: Array<TableStylePart | undefined> = []
  if (def.whole) parts.push(def.whole)
  if (flags.bandRow) {
    const bandIdx = r - (flags.firstRow ? 1 : 0)
    if (bandIdx >= 0) parts.push(bandIdx % 2 === 0 ? def.band1H : def.band2H)
  }
  if (flags.bandCol) {
    const bandIdx = c - (flags.firstCol ? 1 : 0)
    if (bandIdx >= 0) parts.push(bandIdx % 2 === 0 ? def.band1V : def.band2V)
  }
  if (flags.firstCol && c === 0) parts.push(def.firstCol)
  if (flags.lastCol && c === cols - 1) parts.push(def.lastCol)
  if (flags.firstRow && r === 0) parts.push(def.firstRow)
  if (flags.lastRow && r === rows - 1) parts.push(def.lastRow)
  return parts.filter((p): p is TableStylePart => Boolean(p))
}
