// src/core/import/elements/table.ts
/** p:graphicFrame(a:tbl) → table 元素（首 run 样式提取；gridSpan/rowSpan/vMerge/hMerge 矩阵） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { genId } from '../../utils/id'
import { mapX, mapY } from '../context'
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

function cellStyleOf(tc: Element, theme: PptxTheme): TableCell['style'] {
  const style: NonNullable<TableCell['style']> = { color: '#333333', align: 'left', valign: 'middle' }
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
  }
  const firstP = txBody ? firstDescendant(txBody, 'a:p') : null
  const algn = firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null
  if (algn === 'ctr') style.align = 'center'
  else if (algn === 'r') style.align = 'right'
  const tcPr = directChild(tc, 'a:tcPr')
  const fill = tcPr ? resolveColor(directChild(tcPr, 'a:solidFill'), theme) : undefined
  if (fill) style.backcolor = fill
  return style
}

/** a:tbl → cells 矩阵（合并原点带 colspan/rowspan，被并格为 null） */
export function buildCellMatrix(tbl: Element, theme: PptxTheme): { cells: Array<Array<TableCell | null>> } {
  const trs = directChildren(tbl, 'a:tr')
  const cells: Array<Array<TableCell | null>> = []
  const occupied: Array<Array<boolean>> = []
  for (let r = 0; r < trs.length; r += 1) {
    cells.push([])
    occupied.push([])
  }
  trs.forEach((tr, r) => {
    let c = 0
    for (const tc of directChildren(tr, 'a:tc')) {
      // hMerge/vMerge 是合并延续占位：仅占一个网格位，标记 null 后前进一列
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') {
        cells[r][c] = null
        occupied[r][c] = true
        c += 1
        continue
      }
      // 普通格：防御性跳过已被占用的网格位
      while (occupied[r][c]) c += 1
      const colspan = parseInt(attr(tc, 'gridSpan') ?? '1', 10) || 1
      const rowspan = parseInt(attr(tc, 'rowSpan') ?? '1', 10) || 1
      cells[r][c] = {
        text: cellTextOf(tc),
        colspan: colspan > 1 ? colspan : undefined,
        rowspan: rowspan > 1 ? rowspan : undefined,
        style: cellStyleOf(tc, theme),
      }
      // 覆盖区标记占用，并把被并格填充为 null
      for (let dr = 0; dr < rowspan; dr += 1) {
        for (let dc = 0; dc < colspan; dc += 1) {
          const rr = r + dr
          const cc = c + dc
          if (rr >= cells.length) continue
          occupied[rr][cc] = true
          if (dr !== 0 || dc !== 0) cells[rr][cc] = null
        }
      }
      c += colspan
    }
  })
  return { cells }
}

export async function parseTableEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  _pkg: PptxPackage,
): Promise<TableElement | null> {
  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)
  const x = Math.round(mapX(xf, ex) / 9525 * ctx.scale.x)
  const y = Math.round(mapY(xf, ey) / 9525 * ctx.scale.y)
  const w = Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x))
  const h = Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y))

  const tbl = firstDescendant(node, 'a:tbl')
  if (!tbl) return null
  const gridCols = directChildren(firstDescendant(node, 'a:tblGrid') ?? tbl, 'a:gridCol')
  const colSizesRaw = gridCols.map((c) => parseInt(attr(c, 'w') ?? '0', 10))
  const colTotal = colSizesRaw.reduce((a, b) => a + b, 0) || 1
  const trs = directChildren(tbl, 'a:tr')
  const rowSizesRaw = trs.map((tr) => parseInt(attr(tr, 'h') ?? '0', 10))
  const rowTotal = rowSizesRaw.reduce((a, b) => a + b, 0) || 1

  const { cells } = buildCellMatrix(tbl, ctx.theme)
  const table: TableElement = {
    id: genId('tb-'), type: 'table', x, y, w, h,
    colSizes: colSizesRaw.map((c) => c / colTotal),
    rowSizes: rowSizesRaw.map((r) => r / rowTotal),
    cells, name: '表格',
  }
  return table
}
