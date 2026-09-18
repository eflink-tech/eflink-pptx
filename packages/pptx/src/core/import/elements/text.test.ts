import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { txBodyToHTML } from './text'
import { PptxPackage } from '../package'
import type { PptxTheme } from '../theme'

const theme: PptxTheme = {
  schemeColors: { tx1: '#000000', accent1: '#4472C4' },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('txBodyToHTML', () => {
  it('基础 run：字号/加粗/颜色/对齐', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>标题</a:t></a:r>
      </a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('标题')
    expect(r.html).toContain('font-weight:bold')
    expect(r.html).toContain('color:#D14424')
    expect(r.html).toContain('text-align:center')
    expect(r.autoSize).toBe(false)
  })

  it('normAutofit → autoSize；vert=eaVert → vertical', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr vert="eaVert"><a:normAutofit fontScale="92500"/></a:bodyPr>
      <a:p><a:r><a:t>竖排</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.autoSize).toBe(true)
    expect(r.vertical).toBe(true)
  })

  it('项目符号段落聚合为 ul；编号聚合为 ol', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>项一</a:t></a:r></a:p>
      <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>项二</a:t></a:r></a:p>
      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>步一</a:t></a:r></a:p>
      <a:p><a:r><a:t>普通</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('<ul>')
    expect(r.html).toContain('<li')
    expect(r.html).toContain('<ol>')
    expect(r.html).toContain('普通')
    // 普通段落在列表之后
    expect(r.html.indexOf('</ol>')).toBeLessThan(r.html.indexOf('普通'))
  })

  it('空段落输出占位段落', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a"><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('<p')
  })

  it('typeface 含引号时转义，不逃逸 style 属性', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:latin typeface="x&quot; onmouseover=&quot;alert(1)"/></a:rPr><a:t>恶意</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    // 引号必须转义为 &quot;，不允许出现未转义的属性逃逸
    expect(r.html).toContain('&quot;')
    expect(r.html).not.toMatch(/onmouseover="/)
  })

  it('超链接 External：白名单协议产出 <a href>', async () => {
    const pkg = await makePkg('https://example.com')
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>链接</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme, pkg, 'ppt/slides/slide1.xml')
    expect(r.html).toContain('<a href="https://example.com">')
  })

  it('超链接 External：非白名单协议（javascript:）回退为 span，不产出 href', async () => {
    const pkg = await makePkg('javascript:alert(1)')
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>恶意</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme, pkg, 'ppt/slides/slide1.xml')
    expect(r.html).not.toContain('href')
  })
})

describe('txBodyToHTML 字距与行距', () => {
  it('rPr@spc → letter-spacing', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r><a:rPr lang="zh-CN" spc="300"/><a:t>AB</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('letter-spacing:4px') // 300/100/0.75 = 4
  })

  it('a:lnSpc a:spcPct → line-height 倍数', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN"/><a:t>x</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('line-height:1.5')
  })

  it('a:lnSpc a:spcPts → line-height px', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:lnSpc><a:spcPts val="2000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN"/><a:t>x</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('line-height:27px') // 2000/100/0.75 = 26.67 → 四舍五入 27
  })

  it('bullet 段落 + lnSpc → li 携带 line-height', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="zh-CN"/><a:t>项</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('<ul>')
    expect(r.html).toContain('<li style="text-align:left;line-height:2">')
  })

  it('lnSpc val 非法/非正数时跳过，不产出 line-height', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:lnSpc><a:spcPct val="abc"/></a:lnSpc></a:pPr><a:r><a:t>一</a:t></a:r></a:p>
      <a:p><a:pPr><a:lnSpc><a:spcPts val="-500"/></a:lnSpc></a:pPr><a:r><a:t>二</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).not.toContain('line-height')
  })
})

/** 构造含 slide1.xml 与 External 超链接 rels 的真实包 */
async function makePkg(linkTarget: string): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${linkTarget}" TargetMode="External"/>
</Relationships>`)
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}
