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

/** 构造 strict OOXML 包：strict 关系 URI + 非常规前缀（S/D/L/M/T），含嵌入字体声明 */
async function buildStrictPptx(): Promise<File> {
  const PML = 'http://purl.oclc.org/ooxml/presentationml/main'
  const DML = 'http://purl.oclc.org/ooxml/drawingml/main'
  const REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships'
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pre:presentation xmlns:pre="${PML}" xmlns:r="${REL}">
  <pre:embeddedFontLst><pre:embeddedFont/></pre:embeddedFontLst>
  <pre:sldIdLst><pre:sldId id="256" r:id="rId1"/></pre:sldIdLst>
  <pre:sldSz cx="12192000" cy="6858000"/>
</pre:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<S:sld xmlns:S="${PML}" xmlns:D="${DML}">
  <S:cSld><S:spTree>
    <S:nvGrpSpPr><S:cNvPr id="1" name=""/></S:nvGrpSpPr>
    <S:grpSpPr/>
    <S:sp>
      <S:nvSpPr><S:cNvPr id="2" name="T"/></S:nvSpPr>
      <S:spPr>
        <D:xfrm><D:off x="914400" y="914400"/><D:ext cx="4572000" cy="914400"/></D:xfrm>
        <D:prstGeom prst="rect"><D:avLst/></D:prstGeom>
        <D:noFill/>
      </S:spPr>
      <S:txBody><D:bodyPr/><D:p><D:r><D:rPr sz="2400"><D:solidFill><D:srgbClr val="112233"/></D:solidFill></D:rPr><D:t>严格模式</D:t></D:r></D:p></S:txBody>
    </S:sp>
  </S:spTree></S:cSld>
</S:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<L:sldLayout xmlns:L="${PML}" xmlns:T="${DML}">
  <L:cSld><L:spTree>
    <L:nvGrpSpPr><L:cNvPr id="1" name=""/></L:nvGrpSpPr><L:grpSpPr/>
    <L:sp>
      <L:nvSpPr><L:cNvPr id="7" name="Decoration"/></L:nvSpPr>
      <L:spPr>
        <T:xfrm><T:off x="0" y="0"/><T:ext cx="914400" cy="914400"/></T:xfrm>
        <T:prstGeom prst="rect"><T:avLst/></T:prstGeom>
        <T:solidFill><T:srgbClr val="FFCC00"/></T:solidFill>
      </L:spPr>
      <L:txBody><T:bodyPr/><T:p><T:endParaRPr/></T:p></L:txBody>
    </L:sp>
  </L:spTree></L:cSld>
</L:sldLayout>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<M:sldMaster xmlns:M="${PML}" xmlns:T="${DML}">
  <M:cSld><M:spTree>
    <M:nvGrpSpPr><M:cNvPr id="1" name=""/></M:nvGrpSpPr><M:grpSpPr/>
  </M:spTree></M:cSld>
</M:sldMaster>`)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/theme" Target="../theme/theme1.xml"/>
</Relationships>`)
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Strict">
  <a:themeElements>
    <a:clrScheme name="Strict">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="AA00BB"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Strict">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`)
  return new File([await zip.generateAsync({ type: 'blob' })], 'strict.pptx')
}

describe('strict OOXML 兼容', () => {
  it('strict 关系 URI + 非常规前缀仍能解析内容、版式装饰与主题，并报告嵌入字体', async () => {
    const { presentation, report } = await importPPTXDetailed(await buildStrictPptx())
    const slide = presentation.slides[0]
    const texts = slide.elements.filter((el) => el.type === 'text')
    expect(texts).toHaveLength(1)
    expect(texts[0].type === 'text' && texts[0].content).toContain('严格模式')
    expect(texts[0].type === 'text' && texts[0].content).toContain('112233')
    // 版式装饰（FFCC00 矩形）经 strict slideLayout 关系链导入
    expect(slide.elements.some((el) => el.type === 'shape')).toBe(true)
    // 母版 → 主题经 strict theme 关系解析
    expect(presentation.theme.colors[0]).toBe('#AA00BB')
    expect(report.skipped.embeddedFont).toBe(1)
  })
})

describe('graphicFrame 未导入内容报告', () => {
  it('SmartArt(diagram) 与 OLE 容器分别计入报告', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="urn:p" xmlns:r="urn:r">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
    zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:slide" Target="slides/slide1.xml"/>
</Relationships>`)
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>
    <p:graphicFrame>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="urn:drawingml/2006/diagram"/></a:graphic>
    </p:graphicFrame>
    <p:graphicFrame>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="urn:presentationml/2006/ole"/></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`)
    zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)

    const { presentation, report } = await importPPTXDetailed(new File(
      [await zip.generateAsync({ type: 'blob' })], 'frames.pptx',
    ))
    expect(presentation.slides[0].elements).toHaveLength(0)
    expect(report.skipped.smartartUnsupported).toBe(1)
    expect(report.skipped.oleObjectUnsupported).toBe(1)
  })
})
