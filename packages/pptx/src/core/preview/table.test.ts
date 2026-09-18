import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { buildCellMatrix } from '../import/elements/table'
import { makeCtx } from './test-utils'
import { renderTable } from './table'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

const el = (xml: string) => parseXML(xml).documentElement

const TBL = `<p:graphicFrame>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
    <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
    <a:tr h="1000000">
      <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>表头</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></a:tcPr></a:tc>
    </a:tr>
    <a:tr h="1000000">
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
    </a:tr>
  </a:tbl></a:graphicData></a:graphic>
</p:graphicFrame>`

describe('preview/renderTable', () => {
  it('合并格 rect 跨格 + 单元格富文本 foreignObject + 默认边框', async () => {
    const ctx = makeCtx()
    const g = (await renderTable(wrap(TBL), ctx))!
    const rects = Array.from(g.querySelectorAll('rect'))
    // 3 个单元格（合并格跨 2 列算 1 个 rect）
    expect(rects).toHaveLength(3)
    // 合并格：宽 = 2 列（4000000/9525 总宽的一半×2），高 = 单行
    expect(Number(rects[0].getAttribute('width'))).toBeCloseTo((2000000 / 9525) * 2, 1)
    expect(Number(rects[0].getAttribute('height'))).toBeCloseTo(1000000 / 9525, 1)
    expect(rects[0].getAttribute('fill')).toBe('#EEEEEE')
    expect(rects[0].getAttribute('stroke')).toBe('#BFBFBF')
    // 富文本进入 foreignObject（复用 txBodyToHTML：加粗/字号/颜色生效）
    const fo = g.querySelector('foreignObject')!
    expect(fo.firstElementChild!.innerHTML).toContain('表头')
    expect(fo.firstElementChild!.innerHTML).toContain('font-weight:bold')
  })

  it('无 p:xfrm 时返回 null', async () => {
    const ctx = makeCtx()
    const node = wrap(`<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic></p:graphicFrame>`)
    expect(await renderTable(node, ctx)).toBeNull()
  })
})

describe('buildCellMatrix 增量（tcEls 源元素矩阵）', () => {
  it('返回 {cells, tcEls} 形状一致；原点为 a:tc 元素，被并格为 null', () => {
    const tbl = el(`<a:tbl xmlns:a="urn:a">
      <a:tr h="500000">
        <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
      <a:tr h="500000">
        <a:tc hMerge="1"/>
        <a:tc><a:txBody><a:p><a:r><a:t>D</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
    </a:tbl>`)
    const { cells, tcEls } = buildCellMatrix(tbl, { schemeColors: {}, majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {} })
    // 两矩阵形状一致（2×2）
    expect(cells).toHaveLength(2)
    expect(tcEls).toHaveLength(2)
    for (let r = 0; r < cells.length; r += 1) {
      expect(tcEls[r]).toHaveLength(cells[r].length)
    }
    // 原点是源 a:tc 元素
    expect(tcEls[0][0]?.nodeName).toBe('a:tc')
    expect(tcEls[0][1]?.nodeName).toBe('a:tc')
    // hMerge 占位与无格处为 null
    expect(tcEls[1][0]).toBeNull()
    expect(tcEls[1][1]?.nodeName).toBe('a:tc')
  })

  it('gridSpan/rowSpan 覆盖区 tcEls 同步为 null', () => {
    const tbl = el(`<a:tbl xmlns:a="urn:a">
      <a:tr h="500000">
        <a:tc gridSpan="2" rowSpan="2"><a:txBody><a:p><a:r><a:t>合并</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
      <a:tr h="500000">
        <a:tc hMerge="1"/><a:tc vMerge="1"/>
        <a:tc><a:txBody><a:p><a:r><a:t>D</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
    </a:tbl>`)
    const { cells, tcEls } = buildCellMatrix(tbl, { schemeColors: {}, majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {} })
    expect(tcEls[0][0]?.nodeName).toBe('a:tc')
    expect(tcEls[0][1]).toBeNull()
    expect(tcEls[1][0]).toBeNull()
    expect(tcEls[1][1]).toBeNull()
    expect(tcEls[0][2]?.nodeName).toBe('a:tc')
    expect(tcEls[1][2]?.nodeName).toBe('a:tc')
    expect(cells[0][0]?.colspan).toBe(2)
    expect(cells[0][0]?.rowspan).toBe(2)
  })
})
