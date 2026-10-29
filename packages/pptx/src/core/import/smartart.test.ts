import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { importPPTXDetailed } from './index'
import { previewPPTXDetailed } from '../preview'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** 单页 PPTX：slideXml 为页面 spTree 内容；parts 为附加部件 */
async function buildPptx(slideXml: string, relsXml: string, parts: Record<string, string> = {}): Promise<File> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="${P_NS}" xmlns:r="${R_NS}">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:dgm="${DGM_NS}" xmlns:mc="${MC_NS}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    ${slideXml}
  </p:spTree></p:cSld>
</p:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">${relsXml}</Relationships>`)
  for (const [path, xml] of Object.entries(parts)) zip.file(path, xml)
  zip.file('ppt/media/image1.png', PNG, { base64: true })
  return new File([await zip.generateAsync({ type: 'blob' })], 'test.pptx')
}

const DIAGRAM_FRAME = `
  <p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="952500" y="476250"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
      <dgm:relIds r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`

const DIAGRAM_RELS = `
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>`

const DATA_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}">
  <dgm:ptLst/>
  <dgm:bg><a:solidFill><a:srgbClr val="339933"/></a:solidFill></dgm:bg>
</dgm:dataModel>`

const DRAWING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}">
  <dsp:spTree>
    <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
    <dsp:grpSpPr/>
    <dsp:sp modelId="{11111111-1111-1111-1111-111111111111}">
      <dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="476250" y="238125"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
      </dsp:spPr>
      <dsp:style>
        <a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
        <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
        <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
      </dsp:style>
      <dsp:txBody>
        <a:bodyPr anchor="ctr"/><a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Diagram</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`

describe('SmartArt 缓存绘制回退', () => {
  it('按 dsp:drawing 展开为形状 + dgm:bg 背景，位置叠加框架偏移', async () => {
    const file = await buildPptx(DIAGRAM_FRAME, DIAGRAM_RELS, {
      'ppt/diagrams/data1.xml': DATA_XML,
      'ppt/diagrams/drawing1.xml': DRAWING_XML,
    })
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped.smartartDrawing).toBe(1)
    expect(report.skipped.smartartUnsupported).toBeUndefined()
    const els = presentation.slides[0].elements
    expect(els).toHaveLength(2)
    // 背景矩形：框架 off/ext（952500,476250,3810000,1905000）→ px
    expect(els[0]).toMatchObject({ type: 'shape', x: 100, y: 50, w: 400, h: 200, fill: '#339933' })
    // 绘制部件本地坐标 (50,25) + 框架偏移 (100,50)
    expect(els[1]).toMatchObject({ type: 'shape', x: 150, y: 75, w: 200, h: 100, text: 'Diagram' })
    expect((els[1] as { defaultColor?: string }).defaultColor).toBe('#FFFFFF')
  })

  it('预览渲染同样包含背景与文字', async () => {
    const file = await buildPptx(DIAGRAM_FRAME, DIAGRAM_RELS, {
      'ppt/diagrams/data1.xml': DATA_XML,
      'ppt/diagrams/drawing1.xml': DRAWING_XML,
    })
    const { pages, report } = await previewPPTXDetailed(file)
    expect(report.skipped.smartartDrawing).toBe(1)
    const svg = pages[0]
    const fills = Array.from(svg.querySelectorAll('rect')).map((r) => r.getAttribute('fill'))
    expect(fills).toContain('#339933')
    expect(svg.textContent).toContain('Diagram')
  })

  it('绘制部件缺失时报告 smartartUnsupported', async () => {
    const file = await buildPptx(DIAGRAM_FRAME, DIAGRAM_RELS, { 'ppt/diagrams/data1.xml': DATA_XML })
    const { report } = await importPPTXDetailed(file)
    expect(report.skipped.smartartUnsupported).toBe(1)
  })
})

const OLE_FRAME = `
  <p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="5" name="Object 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1905000"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
      <mc:AlternateContent>
        <mc:Choice Requires="v">
          <p:oleObj spid="_x0000_s1" name="Image" r:id="rId9" imgW="100" imgH="100" progId="Photoshop.Image.8"><p:embed/></p:oleObj>
        </mc:Choice>
        <mc:Fallback>
          <p:oleObj name="Image" r:id="rId9" imgW="100" imgH="100" progId="Photoshop.Image.8">
            <p:embed/>
            <p:pic>
              <p:nvPicPr><p:cNvPr id="0" name="Object 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
              <p:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
              <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            </p:pic>
          </p:oleObj>
        </mc:Fallback>
      </mc:AlternateContent>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`

const OLE_RELS = `
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`

describe('OLE 预览图', () => {
  it('有 p:pic 预览帧时按图片导入并报告 olePreviewOnly', async () => {
    const file = await buildPptx(OLE_FRAME, OLE_RELS)
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped.olePreviewOnly).toBe(1)
    expect(report.skipped.oleObjectUnsupported).toBeUndefined()
    const img = presentation.slides[0].elements.find((e) => e.type === 'image')
    expect(img).toMatchObject({ type: 'image', x: 100, y: 100, w: 200, h: 200 })
    expect((img as { src: string }).src.startsWith('data:image/png;base64,')).toBe(true)
  })

  it('无预览图时报告 oleObjectUnsupported', async () => {
    const noPic = OLE_FRAME.replace(/<p:pic>[\s\S]*?<\/p:pic>/, '')
    const file = await buildPptx(noPic, OLE_RELS)
    const { report } = await importPPTXDetailed(file)
    expect(report.skipped.oleObjectUnsupported).toBe(1)
    expect(report.skipped.olePreviewOnly).toBeUndefined()
  })
})

describe('批注与内置表格样式报告', () => {
  it('存在批注关系时报告 commentsUnsupported', async () => {
    const file = await buildPptx('', `
      <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>`)
    const { report } = await importPPTXDetailed(file)
    expect(report.skipped.commentsUnsupported).toBe(1)
  })

  it('表引用可识别的内置样式 GUID 时按规范基础定义还原（不报告）', async () => {
    const table = `
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1"><a:tableStyleId>{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
          </a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>`
    const file = await buildPptx(table, '')
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped.tableStyleBuiltin).toBeUndefined()
    // firstRow：dk1 填充 + lt1 白字（默认主题）
    const cell = presentation.slides[0].elements.find((e) => e.type === 'table')
    const style = (cell as { cells: Array<Array<{ style: { backcolor?: string; color?: string } }>> }).cells[0][0].style
    expect(style.backcolor).toBe('#000000')
    expect(style.color).toBe('#FFFFFF')
  })

  it('表引用未知样式 GUID 时报告 tableStyleBuiltin', async () => {
    const table = `
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1"><a:tableStyleId>{00000000-0000-0000-0000-000000000000}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
          </a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>`
    const file = await buildPptx(table, '')
    const { report } = await importPPTXDetailed(file)
    expect(report.skipped.tableStyleBuiltin).toBe(1)
  })
})
