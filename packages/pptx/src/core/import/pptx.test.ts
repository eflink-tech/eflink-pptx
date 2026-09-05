import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { importPPTX, resolveTarget } from './pptx'

/** 构造最小合法 PPTX（一页：一个文本框 + 一个矩形形状） */
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

describe('importPPTX', () => {
  it('解析最小 PPTX：页面尺寸映射与文本/形状还原', async () => {
    const pres = await importPPTX(await buildMinimalPptx())
    expect(pres.slides).toHaveLength(1)
    expect(pres.width).toBe(1280)
    // 16:9 源尺寸 → viewportRatio 16/9
    expect(pres.viewportRatio).toBeCloseTo(16 / 9, 3)

    const els = pres.slides[0].elements
    expect(els).toHaveLength(2)

    // 文本框（914400 EMU = 96px；12192000 EMU = 1280px → scale 1280/1280 = 1）
    const text = els[0]
    expect(text.type).toBe('text')
    expect(text.x).toBe(96)
    expect(text.y).toBe(96)
    expect(text.type === 'text' && text.content).toContain('标题文字')
    expect(text.type === 'text' && text.content).toContain('font-weight:bold')
    expect(text.type === 'text' && text.content).toContain('D14424')
    expect(text.type === 'text' && text.content).toContain('text-align:center')

    // 形状：roundRect + 旋转 30°（1800000/60000）+ 填充与边框
    const shape = els[1]
    expect(shape.type).toBe('shape')
    if (shape.type === 'shape') {
      expect(shape.shapeKey).toBe('roundRect')
      expect(shape.rotate).toBe(30)
      expect(shape.fill).toBe('#42A5F5')
      expect(shape.outline?.width).toBeGreaterThanOrEqual(2)
    }
  })

  it('非 PPTX 文件抛错', async () => {
    const zip = new JSZip()
    const blob = await zip.generateAsync({ type: 'blob' })
    await expect(importPPTX(new File([blob], 'bad.pptx'))).rejects.toThrow()
  })
})

describe('resolveTarget', () => {
  it('presentation 层相对路径', () => {
    expect(resolveTarget('ppt', 'slides/slide1.xml')).toBe('ppt/slides/slide1.xml')
  })
  it('slide 层 ../ 媒体路径', () => {
    expect(resolveTarget('ppt/slides', '../media/image1.png')).toBe('ppt/media/image1.png')
  })
  it('绝对路径去前导斜杠', () => {
    expect(resolveTarget('ppt', '/ppt/media/a.png')).toBe('ppt/media/a.png')
  })
})
