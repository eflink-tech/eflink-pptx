import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { previewPPTXDetailed, previewPPTX } from './index'

/** 构造最小合法 PPTX（一页：一个文本框 + 一个圆角矩形形状，无版式/母版），与一期导入测试 fixture 保持一致 */
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

/** 构造两页 PPTX（第二页 XML 畸形未闭合），用于验证单页失败容错 */
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

describe('previewPPTX', () => {
  it('最小包：1 页 SVG，viewBox 源画布 px，元素齐全，报告无跳过', async () => {
    const { pages, report } = await previewPPTXDetailed(await buildMinimalPptx())
    expect(pages).toHaveLength(1)
    // sldSz 12192000×6858000 EMU → 1280×720 px 源画布
    expect(pages[0].getAttribute('viewBox')).toBe('0 0 1280 720')
    // 文本框 → foreignObject；圆角矩形形状 → path
    expect(pages[0].querySelector('foreignObject')).not.toBeNull()
    expect(pages[0].querySelector('path')).not.toBeNull()
    expect(report.skipped).toEqual({})
    await expect(previewPPTX(await buildMinimalPptx())).resolves.toHaveLength(1)
  })

  it('坏页容错：畸形第二页计入 slideParseFailed 不中断', async () => {
    const { pages, report } = await previewPPTXDetailed(await buildTwoSlidesOneBroken())
    expect(pages).toHaveLength(1)
    expect(report.skipped).toEqual({ slideParseFailed: 1 })
  })

  it('非 pptx → 明确中文错误', async () => {
    const zip = new JSZip()
    const file = new File([await zip.generateAsync({ type: 'blob' })], 'empty.pptx')
    await expect(previewPPTXDetailed(file)).rejects.toThrow('缺少 presentation.xml')
  })

  it('sldIdLst 为空且无 slide 部件 → 抛「没有幻灯片」中文错误', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`)
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
    zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst/>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)
    const file = new File([await zip.generateAsync({ type: 'blob' })], 'no-slides.pptx')
    await expect(previewPPTXDetailed(file)).rejects.toThrow('没有幻灯片')
  })
})
