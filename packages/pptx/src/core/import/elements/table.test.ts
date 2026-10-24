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

async function makePkg(tableStyles?: string): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  if (tableStyles) zip.file('ppt/tableStyles.xml', tableStyles)
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

  it('tcPr 显式黑色边框 → bordercolor/borderwidth + 表级 outline 同色', async () => {
    const pkg = await makePkg()
    const xml = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblGrid><a:gridCol w="952500"/><a:gridCol w="952500"/></a:tblGrid>
          <a:tr h="952500">
            <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody>
              <a:tcPr anchor="ctr">
                <a:lnL w="12700"><a:solidFill><a:srgbClr val="080000"/></a:solidFill></a:lnL>
                <a:lnT w="12700"><a:solidFill><a:srgbClr val="080000"/></a:solidFill></a:lnT>
                <a:lnR w="12700"><a:solidFill><a:srgbClr val="080000"/></a:solidFill></a:lnR>
                <a:lnB w="12700"><a:solidFill><a:srgbClr val="080000"/></a:solidFill></a:lnB>
                <a:noFill/>
              </a:tcPr></a:tc>
            <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
          </a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`
    const result = await parseTableEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'table') throw new Error('expected table')
    expect(result.cells[0][0]?.style?.bordercolor).toBe('#080000')
    expect(result.cells[0][0]?.style?.borderwidth).toBe(1)
    expect(result.cells[0][0]?.style?.valign).toBe('middle')
    expect(result.outline).toEqual({ color: '#080000', width: 1, style: 'solid' })
  })

  it('tblPr(firstRow+bandRow)+tableStyleId → 表头覆盖 wholeTbl、数据行交替带', async () => {
    const pkg = await makePkg(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{S1}">
  <a:tblStyle styleId="{S1}" styleName="t">
    <a:wholeTbl>
      <a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle>
      <a:tcStyle>
        <a:tcBdr><a:insideH><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH></a:tcBdr>
        <a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
      </a:tcStyle>
    </a:wholeTbl>
    <a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
    <a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>
    <a:firstRow>
      <a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle>
      <a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>`)
    const cell = '<a:tc><a:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>'
    const xml = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{S1}</a:tableStyleId></a:tblPr>
          <a:tblGrid><a:gridCol w="952500"/></a:tblGrid>
          <a:tr h="635000">${cell}</a:tr>
          <a:tr h="635000">${cell}</a:tr>
          <a:tr h="635000">${cell}</a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`
    const result = await parseTableEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'table') throw new Error('expected table')
    const header = result.cells[0][0]?.style
    const band1 = result.cells[1][0]?.style
    const band2 = result.cells[2][0]?.style
    // 表头：firstRow 纯 accent1 覆盖 wholeTbl 淡色；文字 lt1 白色加粗（run 无显式声明）
    expect(header?.backcolor?.toUpperCase()).toBe('#4472C4')
    expect(header?.color?.toUpperCase()).toBe('#FFFFFF')
    expect(header?.bold).toBe(true)
    // 带状：r=1 命中 band1H（tint40 深一档），r=2 回落 wholeTbl（tint20 浅一档）
    expect(band1?.backcolor).not.toBe(band2?.backcolor)
    // 网格线来自 wholeTbl lt1 白
    expect(band1?.bordercolor?.toUpperCase()).toBe('#FFFFFF')
  })

  it('空 tblPr（未引用 tableStyleId）不套用 def 样式', async () => {
    const pkg = await makePkg(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{S1}">
  <a:tblStyle styleId="{S1}" styleName="t">
    <a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>
  </a:tblStyle>
</a:tblStyleLst>`)
    const xml = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblPr/>
          <a:tblGrid><a:gridCol w="952500"/></a:tblGrid>
          <a:tr h="952500"><a:tc><a:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`
    const result = await parseTableEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'table') throw new Error('expected table')
    expect(result.cells[0][0]?.style?.backcolor).toBeUndefined()
  })
})
