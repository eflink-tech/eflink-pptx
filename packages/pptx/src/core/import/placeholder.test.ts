import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { importPPTXDetailed } from './index'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** 单页 PPTX + 一个版式：slideSpTree 为页面 spTree 内容，layoutSpTree 为版式 spTree 内容 */
async function buildPptx(slideSpTree: string, layoutSpTree: string, slideRels = ''): Promise<File> {
  const zip = new JSZip()
  const p = `xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"`
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
  ${slideSpTree}
</p:spTree></p:cSld></p:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  ${slideRels}
</Relationships>`)
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${p}><p:cSld><p:spTree>
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr/>
  ${layoutSpTree}
</p:spTree></p:cSld></p:sldLayout>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}"/>`)
  zip.file('ppt/media/image1.png', PNG, { base64: true })
  return new File([await zip.generateAsync({ type: 'blob' })], 'test.pptx')
}

const CUST_GEOM = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>
<a:pathLst><a:path w="9000000" h="5000000">
  <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
  <a:lnTo><a:pt x="9000000" y="0"/></a:lnTo>
  <a:lnTo><a:pt x="9000000" y="5000000"/></a:lnTo>
  <a:lnTo><a:pt x="0" y="5000000"/></a:lnTo>
  <a:close/>
</a:path></a:pathLst></a:custGeom>`

function layoutPlaceholder(idx: string, geomXml: string): string {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Picture Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="${idx}"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1200000" y="800000"/><a:ext cx="9000000" cy="5000000"/></a:xfrm>
    ${geomXml}
  </p:spPr>
</p:sp>`
}

describe('占位符外观继承', () => {
  it('空 spPr 的形状占位符从版式继承 custGeom、填充与位置', async () => {
    const slide = `<p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Picture Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
    </p:sp>`
    const file = await buildPptx(slide, layoutPlaceholder('13', `${CUST_GEOM}<a:solidFill><a:srgbClr val="FFD966"/></a:solidFill>`))
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped).toEqual({})
    const shape = presentation.slides[0].elements[0]
    expect(shape).toMatchObject({ type: 'shape', x: 126, y: 84, w: 945, h: 525, fill: '#FFD966' })
    expect((shape as { path?: string }).path).toBeTruthy()
  })

  it('图片占位符继承 prstGeom 圆角与位置', async () => {
    const slide = `<p:pic>
      <p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="14"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr/>
    </p:pic>`
    const file = await buildPptx(
      slide,
      layoutPlaceholder('14', '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>'),
      '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
    )
    const { presentation } = await importPPTXDetailed(file)
    const img = presentation.slides[0].elements[0]
    expect(img).toMatchObject({ type: 'image', x: 126, y: 84, w: 945, h: 525 })
    // 圆角半径 = min(w,h) × 0.5
    expect((img as { radius?: number }).radius).toBe(Math.round(525 * 0.5))
  })

  it('图片占位符继承的 custGeom 裁剪无法表达时报告 imageClipGeom', async () => {
    const slide = `<p:pic>
      <p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="15"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr/>
    </p:pic>`
    const file = await buildPptx(
      slide,
      layoutPlaceholder('15', CUST_GEOM),
      '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
    )
    const { presentation, report } = await importPPTXDetailed(file)
    expect(report.skipped.imageClipGeom).toBe(1)
    expect(presentation.slides[0].elements[0]).toMatchObject({ type: 'image', x: 126, y: 84, w: 945, h: 525 })
  })
})
