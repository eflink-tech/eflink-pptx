// src/core/import/elements/table.ts
/** p:graphicFrame(a:tbl) → table 元素（首 run 样式提取；gridSpan/rowSpan/vMerge/hMerge 矩阵） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { fontStackOf } from '../fonts'
import { genId } from '../../utils/id'
import { frameRotOf, parseFrameGeom } from './frame-geom'
import { lstAlgnOf } from './text'
import { partsForCell, resolveTableStyle, type ResolvedTableStyle } from './table-style'
import { addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { PptxTheme } from '../theme'
import type { TableElement, TableCell } from '../../../types/slides'

function cellTextOf(tc: Element): string {
  const txBody = firstDescendant(tc, 'a:txBody')
  if (!txBody) return ''
  return directChildren(txBody, 'a:p')
    .map((p) => directChildren(p, 'a:r').map((r) => firstDescendant(r, 'a:t')?.textContent ?? '').join(''))
    .join('\n')
    .trim()
}

/** 单元格显式样式（仅提取源文件声明的属性，表格样式层叠与默认值由调用方补齐） */
function cellStyleOf(tc: Element, theme: PptxTheme): NonNullable<TableCell['style']> {
  const style: NonNullable<TableCell['style']> = {}
  const txBody = firstDescendant(tc, 'a:txBody')
  const firstRun = txBody ? firstDescendant(txBody, 'a:r') : null
  const rPr = firstRun ? directChild(firstRun, 'a:rPr') : null
  if (rPr) {
    if (attr(rPr, 'b') === '1') style.bold = true
    if (attr(rPr, 'i') === '1') style.italic = true
    if (attr(rPr, 'u') === 'sng') style.underline = true
    const color = resolveColor(directChild(rPr, 'a:solidFill'), theme)
    if (color) style.color = color
    const sz = attr(rPr, 'sz')
    if (sz) style.fontsize = Math.round(parseInt(sz, 10) / 100 / 0.75)
    const latin = attr(directChild(rPr, 'a:latin'), 'typeface')
    const ea = attr(directChild(rPr, 'a:ea'), 'typeface')
    if (latin || ea) style.fontface = fontStackOf(latin ?? '+mn-lt', ea ?? '+mn-ea', theme)
  }
  const firstP = txBody ? firstDescendant(txBody, 'a:p') : null
  const algn = (firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null)
    ?? lstAlgnOf(txBody ? directChild(txBody, 'a:lstStyle') : null, 0)
  if (algn === 'ctr') style.align = 'center'
  else if (algn === 'r') style.align = 'right'
  const tcPr = directChild(tc, 'a:tcPr')
  if (tcPr) {
    const fill = resolveColor(directChild(tcPr, 'a:solidFill'), theme)
    if (fill) style.backcolor = fill
    // 垂直对齐：tcPr@anchor（OOXML 缺省 top；编辑器渲染默认 middle，显式化保证一致）
    const anchor = attr(tcPr, 'anchor')
    if (anchor === 'ctr') style.valign = 'middle'
    else if (anchor === 'b') style.valign = 'bottom'
    else if (anchor === 't') style.valign = 'top'
    // 边框：取四边中最常见的「颜色+宽度」组合；noFill 边视为无边框。
    // 模型单元格仅支持单一边框，四边不一致时由 parseTableEl 计入报告
    const borderNames = ['a:lnL', 'a:lnT', 'a:lnR', 'a:lnB'] as const
    const edges = new Map<string, { color: string; widthEmu: number; count: number }>()
    for (const name of borderNames) {
      const ln = directChild(tcPr, name)
      if (!ln || directChild(ln, 'a:noFill')) continue
      const color = resolveColor(directChild(ln, 'a:solidFill'), theme)
      if (!color) continue
      const widthEmu = parseInt(attr(ln, 'w') ?? '0', 10)
      const key = `${color}|${widthEmu}`
      const cur = edges.get(key)
      if (cur) cur.count += 1
      else edges.set(key, { color, widthEmu, count: 1 })
    }
    let best: { color: string; widthEmu: number; count: number } | null = null
    for (const edge of edges.values()) {
      if (!best || edge.count > best.count) best = edge
    }
    if (best) {
      style.bordercolor = best.color
      style.borderwidth = Math.max(1, Math.round(best.widthEmu / 9525))
    }
  }
  return style
}

/** a:tbl → cells 矩阵（合并原点带 colspan/rowspan，被并格为 null）；
 * tcEls 为对应源 a:tc 元素矩阵（预览层富文本用），与 cells 逐位同步，改动 cells 逻辑时必须同步维护：
 * 普通格=元素，合并占位/覆盖区=null */
export function buildCellMatrix(tbl: Element, theme: PptxTheme, tblStyle?: ResolvedTableStyle): {
  cells: Array<Array<TableCell | null>>
  tcEls: Array<Array<Element | null>>
} {
  const trs = directChildren(tbl, 'a:tr')
  const cells: Array<Array<TableCell | null>> = []
  const tcEls: Array<Array<Element | null>> = []
  const occupied: Array<Array<boolean>> = []
  for (let r = 0; r < trs.length; r += 1) {
    cells.push([])
    tcEls.push([])
    occupied.push([])
  }
  const colCount = directChildren(firstDescendant(tbl, 'a:tblGrid') ?? tbl, 'a:gridCol').length
  trs.forEach((tr, r) => {
    let c = 0
    for (const tc of directChildren(tr, 'a:tc')) {
      // hMerge/vMerge 是合并延续占位：仅占一个网格位，标记 null 后前进一列
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') {
        cells[r][c] = null
        tcEls[r][c] = null
        occupied[r][c] = true
        c += 1
        continue
      }
      // 普通格：防御性跳过已被占用的网格位
      while (occupied[r][c]) c += 1
      const colspan = parseInt(attr(tc, 'gridSpan') ?? '1', 10) || 1
      const rowspan = parseInt(attr(tc, 'rowSpan') ?? '1', 10) || 1
      const style = cellStyleOf(tc, theme)
      // 表格样式层叠（低→高，后者覆盖前者）：wholeTbl < band < first/last Row/Col；
      // 合成结果仅补单元格未显式声明的属性（tcPr 填充 / run 样式始终最高优先）
      if (tblStyle) {
        const layer: { fill?: string; color?: string; bold?: boolean; border?: string } = {}
        for (const part of partsForCell(tblStyle, r, c, trs.length, colCount)) {
          if (part.fill) layer.fill = part.fill
          if (part.color) layer.color = part.color
          if (part.bold) layer.bold = true
          if (part.border) layer.border = part.border
        }
        if (!style.backcolor && layer.fill) style.backcolor = layer.fill
        if (!style.color && layer.color) style.color = layer.color
        if (layer.bold && style.bold !== true) style.bold = true
        if (!style.bordercolor && layer.border) {
          style.bordercolor = layer.border
          style.borderwidth = 1
        }
      }
      if (!style.color) style.color = '#333333'
      if (!style.align) style.align = 'left'
      if (!style.valign) style.valign = 'middle'
      cells[r][c] = {
        text: cellTextOf(tc),
        colspan: colspan > 1 ? colspan : undefined,
        rowspan: rowspan > 1 ? rowspan : undefined,
        style,
      }
      tcEls[r][c] = tc
      // 覆盖区标记占用，并把被并格填充为 null（cells 与 tcEls 同步）
      for (let dr = 0; dr < rowspan; dr += 1) {
        for (let dc = 0; dc < colspan; dc += 1) {
          const rr = r + dr
          const cc = c + dc
          if (rr >= cells.length) continue
          occupied[rr][cc] = true
          if (dr !== 0 || dc !== 0) {
            cells[rr][cc] = null
            tcEls[rr][cc] = null
          }
        }
      }
      c += colspan
    }
  })
  return { cells, tcEls }
}

/** 行高/列宽兜底：缺失维度按已知维度均值补齐，全部缺失则等分（避免 0 比例导致渲染塌陷） */
export function sizesWithFallback(raw: number[]): number[] {
  const known = raw.filter((v) => v > 0)
  const fallback = known.length ? Math.round(known.reduce((a, b) => a + b, 0) / known.length) : 1
  return raw.map((v) => (v > 0 ? v : fallback))
}

/** 单元格富文本判定：同格内 run 样式不一致（粗细/斜体/下划线/字号/颜色/字体）时模型只能保留纯文本 */
function hasRichRuns(tc: Element, theme: PptxTheme): boolean {
  const txBody = firstDescendant(tc, 'a:txBody')
  if (!txBody) return false
  const sigs = new Set<string>()
  for (const p of directChildren(txBody, 'a:p')) {
    for (const node of Array.from(p.children)) {
      if (node.localName !== 'r' && node.localName !== 'fld') continue
      const rPr = directChild(node, 'a:rPr')
      const runColor = rPr ? resolveColor(directChild(rPr, 'a:solidFill'), theme) : null
      sigs.add([
        attr(rPr, 'b') ?? '', attr(rPr, 'i') ?? '', attr(rPr, 'u') ?? '', attr(rPr, 'strike') ?? '',
        attr(rPr, 'sz') ?? '', runColor ?? '', attr(rPr ? directChild(rPr, 'a:latin') : null, 'typeface') ?? '',
      ].join('|'))
      if (sigs.size > 1) return true
    }
  }
  return false
}

/** 单元格四边边框声明是否不一致（noFill 视作显式无边框） */
function hasEdgeMismatch(tc: Element, theme: PptxTheme): boolean {
  const tcPr = directChild(tc, 'a:tcPr')
  if (!tcPr) return false
  const sigs: string[] = []
  for (const name of ['a:lnL', 'a:lnT', 'a:lnR', 'a:lnB'] as const) {
    const ln = directChild(tcPr, name)
    if (!ln) continue
    if (directChild(ln, 'a:noFill')) { sigs.push('none'); continue }
    const color = resolveColor(directChild(ln, 'a:solidFill'), theme)
    if (!color) continue
    sigs.push(`${color}|${attr(ln, 'w') ?? '0'}`)
  }
  return new Set(sigs).size > 1
}

export async function parseTableEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  _pkg: PptxPackage,
): Promise<TableElement | null> {
  const geom = parseFrameGeom(node, xf, ctx)
  if (!geom) return null
  const { x, y, w, h } = geom

  const tbl = firstDescendant(node, 'a:tbl')
  if (!tbl) return null
  const tblStyle = await resolveTableStyle(tbl, ctx.pkg, ctx.theme)
  // 内置样式 GUID 已按规范基础定义还原；仅无法识别的 GUID 才计入报告
  const tblPr = directChild(tbl, 'a:tblPr')
  const styleId = tblPr ? directChild(tblPr, 'a:tableStyleId') : null
  if (!tblStyle && styleId?.textContent?.trim()) addSkipped(ctx.report, 'tableStyleBuiltin')
  const gridCols = directChildren(firstDescendant(node, 'a:tblGrid') ?? tbl, 'a:gridCol')
  const colSizesRaw = sizesWithFallback(gridCols.map((c) => parseInt(attr(c, 'w') ?? '0', 10)))
  const colTotal = colSizesRaw.reduce((a, b) => a + b, 0) || 1
  const trs = directChildren(tbl, 'a:tr')
  const rowSizesRaw = sizesWithFallback(trs.map((tr) => parseInt(attr(tr, 'h') ?? '0', 10)))
  const rowTotal = rowSizesRaw.reduce((a, b) => a + b, 0) || 1

  const { cells, tcEls } = buildCellMatrix(tbl, ctx.theme, tblStyle)
  // 模型限制：单元格只存单一文本与单一边框；富文本/四边不一致在此如实报告
  let rich = 0
  let edgeMismatch = false
  for (const row of tcEls) {
    for (const tc of row) {
      if (!tc) continue
      if (hasRichRuns(tc, ctx.theme)) rich += 1
      if (hasEdgeMismatch(tc, ctx.theme)) edgeMismatch = true
    }
  }
  if (rich) addSkipped(ctx.report, 'tableRichText')
  if (edgeMismatch) addSkipped(ctx.report, 'tableBorders')
  const table: TableElement = {
    id: genId('tb-'), type: 'table', x, y, w, h,
    colSizes: colSizesRaw.map((c) => c / colTotal),
    rowSizes: rowSizesRaw.map((r) => r / rowTotal),
    cells, name: '表格',
  }
  const rot = frameRotOf(firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm'), xf).rot
  if (rot) table.rotate = Math.round(rot)
  // 表格外框与单元格边框同源：取首个带边框色的单元格（含表格样式网格线），避免外灰内黑的混搭
  const firstBorder = cells.flat().find((cell) => cell?.style?.bordercolor)
  if (firstBorder?.style?.bordercolor) {
    table.outline = { color: firstBorder.style.bordercolor, width: firstBorder.style.borderwidth ?? 1, style: 'solid' }
  }
  return table
}
