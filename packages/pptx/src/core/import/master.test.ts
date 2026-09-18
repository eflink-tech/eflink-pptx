import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseBackgroundFill, collectPlaceholders, findAncestry, parseSlideAncestry } from './master'
import { PptxPackage } from './package'
import type { ImportReport } from '../../types/slides'

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout xmlns:p="urn:p"/>')
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

describe('master.ts', () => {
  it('findAncestry：slide → layout → master 链', async () => {
    const pkg = await makePkg()
    const chain = await findAncestry(pkg, 'ppt/slides/slide1.xml')
    expect(chain.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml')
    expect(chain.masterPath).toBe('ppt/slideMasters/slideMaster1.xml')
  })

  it('collectPlaceholders：ph 的 idx/type 作 key', () => {
    const tree = el(`<p:spTree xmlns:p="urn:p" xmlns:a="urn:a">
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3"/><p:nvPr><p:ph idx="12"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4"/></p:nvSpPr>
        <p:spPr/>
      </p:sp>
    </p:spTree>`)
    const map = collectPlaceholders(tree)
    expect(map.get('title')).toEqual({ x: 100, y: 200, w: 300, h: 400 })
    expect(map.get('12')).toEqual({ x: 1, y: 2, w: 3, h: 4 })
    expect(map.size).toBe(2)
  })

  it('parseBackgroundFill：solid → Background；无 bg 返回 undefined', () => {
    const theme = { schemeColors: { lt1: '#FFFFFF' } as Record<string, string>, majorFont: '', minorFont: '', colorMap: {} }
    const bg = el(`<p:bg xmlns:p="urn:p" xmlns:a="urn:a">
      <p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:effectLst/></p:bgPr>
    </p:bg>`)
    const result = parseBackgroundFill(bg, theme)
    expect(result).toEqual({ type: 'solid', color: '#112233' })
    expect(parseBackgroundFill(null, theme)).toBeUndefined()
  })

  it('parseBackgroundFill：blipFill → image 背景（src 由 ancestry 层按 rels 补充）', () => {
    const theme = { schemeColors: {} as Record<string, string>, majorFont: '', minorFont: '', colorMap: {} }
    const bg = el(`<p:bg xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:bgPr><a:blipFill><a:blip r:embed="rId9"/></a:blipFill><a:effectLst/></p:bgPr>
    </p:bg>`)
    expect(parseBackgroundFill(bg, theme)).toEqual({ type: 'image' })
  })

  it('parseSlideAncestry：背景图片 src 解析失败 → 不保留悬空 image 背景，回退 lt1 兜底', async () => {
    const zip = new JSZip()
    // 版式带图片背景，但 blip r:embed 指向不存在的 rel → src 无法解析
    zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
  <p:cSld>
    <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId9"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree/>
  </p:cSld>
</p:sldLayout>`)
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`)
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"><p:cSld><p:spTree/></p:cSld></p:sldMaster>')
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const report: ImportReport = { skipped: {} }
    const ancestry = await parseSlideAncestry(pkg, 'ppt/slides/slide1.xml', report, { x: 1, y: 1 })
    expect(ancestry.background).toEqual({ type: 'solid', color: '#FFFFFF' })
  })
})
