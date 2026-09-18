import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseTableEl, buildCellMatrix } from './table'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

const ctxBase = (pkg: PptxPackage) => ({
  pkg,
  partPath: 'ppt/slides/slide1.xml',
  theme,
  report: { skipped: {} } as ImportReport,
  scale: { x: 1, y: 1 },
  placeholders: new Map(),
})

describe('buildCellMatrix', () => {
  // 合并格 gridSpan=2 rowSpan=2 占据 (0,0)-(1,1)，B/D 落在第 2 列对齐
  it('gridSpan/rowSpan/vMerge/hMerge → colspan/rowspan + null 填充', () => {
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
    const { cells } = buildCellMatrix(tbl, theme)
    expect(cells[0][0]?.text).toBe('合并')
    expect(cells[0][0]?.colspan).toBe(2)
    expect(cells[0][0]?.rowspan).toBe(2)
    expect(cells[0][1]).toBeNull()
    expect(cells[0][2]?.text).toBe('B')
    expect(cells[1][0]).toBeNull()
    expect(cells[1][1]).toBeNull()
    expect(cells[1][2]?.text).toBe('D')
  })
})

describe('parseTableEl', () => {
  it('首 run 样式进 style 字段；行高列宽归一化', async () => {
    const pkg = await makePkg()
    const xml = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblGrid><a:gridCol w="952500"/><a:gridCol w="952500"/></a:tblGrid>
          <a:tr h="952500">
            <a:tc><a:txBody><a:p><a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>单元格</a:t></a:r></a:p></a:txBody></a:tc>
            <a:tc><a:txBody><a:p><a:r><a:t>普通</a:t></a:r></a:p></a:txBody></a:tc>
          </a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`
    const result = await parseTableEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'table') throw new Error('expected table')
    expect(result.colSizes).toEqual([0.5, 0.5])
    expect(result.cells[0][0]?.style?.bold).toBe(true)
    expect(result.cells[0][0]?.style?.color).toBe('#D14424')
    expect(result.cells[0][0]?.style?.fontsize).toBe(27)
    expect(result.cells[0][1]?.text).toBe('普通')
  })
})
