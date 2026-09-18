/** p:graphicFrame(a:tbl) → 表格 SVG：合并格 rect + 复用 txBodyToHTML 的富文本单元格。
 * 边框简化：统一细边框（#BFBFBF 0.75px），OOXML 逐边 lnL/lnR/lnT/lnB 定制不还原（已知取舍）。 */
import { attr, directChildren, firstDescendant } from '../import/xml'
import { buildCellMatrix } from '../import/elements/table'
import { renderText } from './text'
import { svgEl, geomOf, boxTransform, type PreviewCtx } from './svg'

export async function renderTable(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const box = geomOf(node, ctx)
  if (!box) return null
  const tbl = firstDescendant(node, 'a:tbl')
  if (!tbl) return null

  const gridCols = directChildren(firstDescendant(node, 'a:tblGrid') ?? tbl, 'a:gridCol')
  const colSizesRaw = gridCols.map((c) => parseInt(attr(c, 'w') ?? '0', 10))
  const colTotal = colSizesRaw.reduce((a, b) => a + b, 0) || 1
  const trs = directChildren(tbl, 'a:tr')
  const rowSizesRaw = trs.map((tr) => parseInt(attr(tr, 'h') ?? '0', 10))
  const rowTotal = rowSizesRaw.reduce((a, b) => a + b, 0) || 1

  // 每列/行的相对偏移（px，基于归一化比例 → 畸形 w/h=0 时按占比兜底）
  const colX: number[] = []
  let acc = 0
  for (const w of colSizesRaw) {
    colX.push(acc)
    acc += (w / colTotal) * box.w
  }
  const rowY: number[] = []
  acc = 0
  for (const h of rowSizesRaw) {
    rowY.push(acc)
    acc += (h / rowTotal) * box.h
  }

  const { cells, tcEls } = buildCellMatrix(tbl, ctx.theme)
  // 保留 frame 旋转；flip 对表格丢弃而非镜像（与 shape.ts 同理，表格翻转语义罕见且镜像文本不可读）
  const g = svgEl('g', { transform: boxTransform({ ...box, flipH: false, flipV: false }) })
  for (let r = 0; r < cells.length; r += 1) {
    for (let c = 0; c < cells[r].length; c += 1) {
      const cell = cells[r][c]
      if (!cell) continue
      const colspan = cell.colspan ?? 1
      const rowspan = cell.rowspan ?? 1
      const x = box.x + (colX[c] ?? 0)
      const y = box.y + (rowY[r] ?? 0)
      // 跨格宽高：末列/末行右/下边界由 `?? box.w/box.h` 越界回退兜底
      const w = (colX[c + colspan] ?? box.w) - (colX[c] ?? 0)
      const h = (rowY[r + rowspan] ?? box.h) - (rowY[r] ?? 0)
      g.appendChild(svgEl('rect', {
        x, y, width: w, height: h,
        fill: cell.style?.backcolor ?? 'none',
        stroke: '#BFBFBF', 'stroke-width': 0.75,
      }))
      // 富文本单元格：复用 txBodyToHTML（XSS 加固、加粗/字号/颜色/对齐生效）
      const tc = tcEls[r][c]
      const txBody = tc ? firstDescendant(tc, 'a:txBody') : null
      if (txBody) {
        const text = await renderText(txBody, ctx, { x, y, w, h, rot: 0, flipH: false, flipV: false })
        if (text) g.appendChild(text)
      }
    }
  }
  return g
}
