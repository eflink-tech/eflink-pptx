import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { importPPTXDetailed, importPPTX } from './index'

/** 构造最小合法 PPTX（一页：一个文本框 + 一个矩形形状，无版式/母版） */
async function buildMinimalPptx(): Promise<File> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="2400" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>标题文字</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Rect 2"/></p:nvSpPr>
      <p:spPr>
        <a:xfrm rot="1800000"><a:off x="6096000" y="2743200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        <a:ln w="25400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)

  const blob = await zip.generateAsync({ type: 'blob' })
  return new File([blob], 'test.pptx')
}

/** 构造两页 PPTX（第二页 XML 畸形），用于验证单页失败容错 */
async function buildTwoSlidesOneBroken(): Promise<File> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>正常页</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`)
  zip.file('ppt/slides/slide2.xml', '<p:sld><p:cSld>畸形 XML 未闭合')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)

  const blob = await zip.generateAsync({ type: 'blob' })
  return new File([blob], 'two-slides.pptx')
}

describe('importPPTXDetailed', () => {
  it('最小包：页面数、元素坐标按 scale 折算、报告无跳过项', async () => {
    const { presentation, report } = await importPPTXDetailed(await buildMinimalPptx())
    expect(presentation.slides).toHaveLength(1)
    expect(presentation.width).toBe(1280)
    expect(presentation.viewportRatio).toBeCloseTo(16 / 9, 3)
    // 无版式/母版 → 主题走 Office 默认色板
    expect(presentation.theme.colors).toHaveLength(10)
    expect(presentation.theme.colors[0]).toBe('#4472C4')

    const els = presentation.slides[0].elements
    expect(els).toHaveLength(2)

    // 12192000 EMU = 1280px → scale 1：914400 EMU = 96px
    const text = els[0]
    expect(text.type).toBe('text')
    expect(text.x).toBe(96)
    expect(text.y).toBe(96)
    expect(text.type === 'text' && text.content).toContain('标题文字')

    const shape = els[1]
    expect(shape.type).toBe('shape')
    if (shape.type === 'shape') {
      expect(shape.shapeKey).toBe('roundRect')
      expect(shape.rotate).toBe(30)
      expect(shape.fill).toBe('#42A5F5')
    }

    expect(report.skipped).toEqual({})
  })

  it('importPPTX 兼容入口仅返回 Presentation', async () => {
    const pres = await importPPTX(await buildMinimalPptx())
    expect(pres.slides).toHaveLength(1)
  })

  it('缺 presentation.xml → 抛中文错误', async () => {
    const zip = new JSZip()
    const blob = await zip.generateAsync({ type: 'blob' })
    await expect(importPPTXDetailed(new File([blob], 'bad.pptx')))
      .rejects.toThrow('不是有效的 PPTX 文件')
  })

  it('单页 XML 畸形：其余页正常、报告计入 slideParseFailed', async () => {
    const { presentation, report } = await importPPTXDetailed(await buildTwoSlidesOneBroken())
    expect(presentation.slides).toHaveLength(1)
    const els = presentation.slides[0].elements
    expect(els).toHaveLength(1)
    expect(els[0].type === 'text' && els[0].content).toContain('正常页')
    expect(report.skipped['slideParseFailed']).toBe(1)
  })

  it('sldIdLst 缺失时按 slideN.xml 编号兜底', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`)
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
    zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
    for (const n of [1, 2]) {
      zip.file(`ppt/slides/slide${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="T"/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>页${n}</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`)
    }

    const { presentation } = await importPPTXDetailed(new File(
      [await zip.generateAsync({ type: 'blob' })], 'fallback.pptx',
    ))
    expect(presentation.slides).toHaveLength(2)
    expect(presentation.slides[0].elements[0].type === 'text' && presentation.slides[0].elements[0].content).toContain('页1')
  })
})
