import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from '../import/package'
import { renderSlide } from './slide'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const rel = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`
const rels = (inner: string) => `${XML_DECL}<Relationships xmlns="${RELS_NS}">${inner}</Relationships>`

/** 最小 pptx：presentation → slide1 → layout1 → master1 → theme1；母版含背景 + 装饰矩形 + title 占位符 */
async function buildPptx(): Promise<Blob> {
  const zip = new JSZip()

  zip.file('[Content_Types].xml', `${XML_DECL}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)

  zip.file('_rels/.rels', rels(rel('rId1', 'officeDocument', 'ppt/presentation.xml')))

  zip.file('ppt/presentation.xml', `${XML_DECL}
<p:presentation xmlns:p="${P_NS}" xmlns:r="${R_NS}">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`)

  zip.file('ppt/_rels/presentation.xml.rels', rels(rel('rId1', 'slide', 'slides/slide1.xml')))

  zip.file('ppt/slideMasters/slideMaster1.xml', `${XML_DECL}
<p:sldMaster xmlns:p="${P_NS}" xmlns:a="${A_NS}">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FDF6EC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="装饰矩形"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="6096000"/><a:ext cx="12192000" cy="762000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="标题占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>`)

  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels(
    rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')
    + rel('rId2', 'theme', '../theme/theme1.xml'),
  ))

  zip.file('ppt/slideLayouts/slideLayout1.xml', `${XML_DECL}
<p:sldLayout xmlns:p="${P_NS}" xmlns:a="${A_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`)

  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels(rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')))

  zip.file('ppt/theme/theme1.xml', `${XML_DECL}
<a:theme xmlns:a="${A_NS}" name="T">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`)

  // slide 占位符无 xfrm：位置继承 master 的 title 占位符（838200,365125 EMU = 88px,38.33px）
  zip.file('ppt/slides/slide1.xml', `${XML_DECL}
<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="标题"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>标题文本</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`)

  zip.file('ppt/slides/_rels/slide1.xml.rels', rels(rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')))

  return zip.generateAsync({ type: 'blob' })
}

describe('preview/renderSlide', () => {
  it('页面组装：viewBox 源画布尺寸 / 母版背景与装饰合入 / 占位符位置继承', async () => {
    const pkg = await PptxPackage.load(await buildPptx())
    const svg = (await renderSlide(pkg, 'ppt/slides/slide1.xml', 12192000, 6858000, 0, { skipped: {} }))!
    expect(svg.getAttribute('viewBox')).toBe('0 0 1280 720')
    expect(svg.getAttribute('width')).toBe('1280')
    expect(svg.querySelector('rect[fill="#FDF6EC"]')).not.toBeNull() // 母版背景 rect
    expect(svg.querySelector('path[fill="#FF0000"]')).not.toBeNull() // 母版装饰矩形
    const fos = Array.from(svg.querySelectorAll('foreignObject'))
    expect(fos.some((f) => f.getAttribute('x') === '88')).toBe(true) // 占位符位置继承（838200/9525）
  })

  it('slide 部件缺失 → null', async () => {
    const pkg = await PptxPackage.load(await buildPptx())
    expect(await renderSlide(pkg, 'ppt/slides/missing.xml', 12192000, 6858000, 0, { skipped: {} })).toBeNull()
  })

  it('master XML 畸形 → 单部件降级不拖垮整页（默认主题 lt1 兜底背景，不抛错）', async () => {
    const zip = await JSZip.loadAsync(await buildPptx())
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"><p:cSld>') // 故意畸形
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    // master 降级后占位符位置表为空：slide 占位符无位置被保守丢弃（与一期 parseSlideAncestry 语义一致）
    const svg = (await renderSlide(pkg, 'ppt/slides/slide1.xml', 12192000, 6858000, 0, { skipped: {} }))!
    expect(svg.getAttribute('viewBox')).toBe('0 0 1280 720')
    expect(svg.querySelector('rect[fill="#FFFFFF"]')).not.toBeNull() // 默认主题 lt1 兜底背景
    expect(svg.querySelector('foreignObject')).toBeNull() // 无继承位置的占位符不渲染
  })
})
