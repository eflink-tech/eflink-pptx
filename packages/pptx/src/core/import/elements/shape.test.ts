import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseShapeEl } from './shape'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makeCtx(slideXml: string) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slideXml)
  const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
  return {
    pkg,
    ctx: {
      pkg,
      partPath: 'ppt/slides/slide1.xml',
      theme,
      report: { skipped: {} } as ImportReport,
      scale: { x: 1, y: 1 },
      placeholders: new Map(),
    },
  }
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('parseShapeEl', () => {
  it('渐变填充 + 阴影 + custGeom path', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Custom"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
      <a:custGeom><a:pathLst><a:path w="1000" h="1000">
        <a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1000" y="1000"/></a:lnTo>
      </a:path></a:pathLst></a:custGeom>
      <a:gradFill>
        <a:gsLst>
          <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
        </a:gsLst>
        <a:lin ang="2700000"/>
      </a:gradFill>
      <a:ln><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
      <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
  </p:sp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const sp = el(xml).getElementsByTagName('p:sp')[0]
    const result = await parseShapeEl(sp, ctx, IDENTITY_XFORM, pkg)
    expect(result).not.toBeNull()
    if (result?.type !== 'shape') throw new Error('expected shape')
    expect(result.path).toBe('M0,0 L100,100')
    expect(typeof result.fill).toBe('object')
    const grad = result.fill as { type: string; colors: Array<{ pos: number; color: string }>; rotate: number }
    expect(grad.colors[0].color).toBe('#FF0000')
    expect(grad.colors[1].color).toBe('#0000FF')
    expect(grad.rotate).toBe(45)
    expect(result.shadow?.blur).toBe(5)
    expect(result.shadow?.color).toBe('#00000066')
  })

  it('线条：prstDash 与箭头映射', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:cxnSp>
    <p:nvSpPr><p:cNvPr id="2" name="Line"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>
    </p:spPr>
  </p:cxnSp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const cxn = el(xml).getElementsByTagName('p:cxnSp')[0]
    const result = await parseShapeEl(cxn, ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'line') throw new Error('expected line')
    expect(result.lineWidth).toBe(2)
    expect(result.color).toBe('#4472C4')
  })

  it('带边框透明矩形（noFill + ln + txBody）判为形状并保留 outline', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Bordered Box"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>边框内文字</a:t></a:r></a:p></p:txBody>
  </p:sp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const sp = el(xml).getElementsByTagName('p:sp')[0]
    const result = await parseShapeEl(sp, ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'shape') throw new Error('expected shape')
    // 透明填充 + 边框保留
    expect(result.fill).toBe('#00000000')
    expect(result.outline?.color).toBe('#333333')
    expect(result.outline?.width).toBeGreaterThanOrEqual(1)
    expect(result.text).toContain('边框内文字')
  })

  it('形状文本剥标签后解码实体：空段 &nbsp; 不残留、转义字符还原', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Entity Box"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="00AA66"/></a:solidFill>
    </p:spPr>
    <p:txBody><a:bodyPr/>
      <a:p><a:endParaRPr/></a:p>
      <a:p><a:r><a:rPr lang="zh-CN" sz="1800"/><a:t>卡卡 &amp; KK &lt;202X&gt;</a:t></a:r></a:p>
      <a:p><a:endParaRPr/></a:p>
    </p:txBody>
  </p:sp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const sp = el(xml).getElementsByTagName('p:sp')[0]
    const result = await parseShapeEl(sp, ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'shape') throw new Error('expected shape')
    expect(result.text).toBe('卡卡 & KK <202X>')
    expect(result.text).not.toContain('&nbsp;')
    expect(result.text).not.toContain('&amp;')
    expect(result.text).not.toContain('&lt;')
  })

  it('占位符继承：无 xfrm 时从 placeholders 取位置折算 px', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
    <p:spPr>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
  </p:sp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    ctx.placeholders.set('title', { x: 952500, y: 952500, w: 4762500, h: 952500 })
    const sp = el(xml).getElementsByTagName('p:sp')[0]
    const result = await parseShapeEl(sp, ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'shape') throw new Error('expected shape')
    // scale {x:1,y:1} 时 px = EMU / 9525
    expect(result.x).toBe(100)
    expect(result.y).toBe(100)
    expect(result.w).toBe(500)
    expect(result.h).toBe(100)
  })
})
