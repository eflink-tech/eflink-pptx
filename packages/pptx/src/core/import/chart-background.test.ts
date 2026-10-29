import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { importPPTXDetailed } from './index'
import { previewPPTXDetailed } from '../preview'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const FRAME = `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="952500" y="476250"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart r:id="rId7"/>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

/** chartSpace：spPr 为图表区背景，plotSpPr 为绘图区背景 */
function chartXml(spPr: string, plotSpPr = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <c:chart><c:plotArea>
    ${plotSpPr}
    <c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S1</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
  </c:plotArea></c:chart>
  ${spPr}
</c:chartSpace>`
}

async function buildPptx(chartRootXml: string, chartRels = ''): Promise<File> {
  const zip = new JSZip()
  const p = `xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:c="${C_NS}"`
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${p}>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${p}><p:cSld><p:spTree>
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr/>
  ${FRAME}
</p:spTree></p:cSld></p:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`)
  zip.file('ppt/charts/chart1.xml', chartRootXml)
  zip.file('ppt/charts/_rels/chart1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">${chartRels}</Relationships>`)
  zip.file('ppt/media/image1.png', PNG, { base64: true })
  return new File([await zip.generateAsync({ type: 'blob' })], 'test.pptx')
}

describe('图表区背景', () => {
  it('chartSpace 纯色背景写入 backgroundColor', async () => {
    const file = await buildPptx(chartXml('<c:spPr><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></c:spPr>'))
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped).toEqual({})
    const chart = presentation.slides[0].elements.find((e) => e.type === 'chart')
    expect(chart).toMatchObject({ backgroundColor: '#FFCC00' })
  })

  it('chartSpace 渐变背景写入 backgroundGradient', async () => {
    const file = await buildPptx(chartXml(`<c:spPr><a:gradFill><a:gsLst>
      <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
      <a:gs pos="100000"><a:srgbClr val="3366FF"/></a:gs>
    </a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr>`))
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped).toEqual({})
    const chart = presentation.slides[0].elements.find((e) => e.type === 'chart')
    expect(chart).toMatchObject({
      backgroundGradient: { type: 'linear', rotate: 90, colors: [{ pos: 0, color: '#FFFFFF' }, { pos: 1, color: '#3366FF' }] },
    })
  })

  it('chartSpace 图片背景按部件 rels 解析为 data URL', async () => {
    const file = await buildPptx(
      chartXml('<c:spPr><a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>'),
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
    )
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped).toEqual({})
    const chart = presentation.slides[0].elements.find((e) => e.type === 'chart')
    expect((chart as { backgroundImage?: string }).backgroundImage?.startsWith('data:image/png;base64,')).toBe(true)
  })

  it('绘图区背景无对应字段时报告 chartBackground', async () => {
    const file = await buildPptx(chartXml('', '<c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr>'))
    const { report } = await importPPTXDetailed(file)
    expect(report.skipped.chartBackground).toBe(1)
  })

  it('预览 SVG 绘制图表区背景', async () => {
    const file = await buildPptx(chartXml('<c:spPr><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></c:spPr>'))
    const { pages } = await previewPPTXDetailed(file)
    const fills = Array.from(pages[0].querySelectorAll('rect')).map((r) => r.getAttribute('fill'))
    expect(fills).toContain('#FFCC00')
  })
})
