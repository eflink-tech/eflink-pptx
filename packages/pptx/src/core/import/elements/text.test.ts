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

  it('bodyPr insets 折算 padding；spAutoFit → autoFitShape', async () => {
    // 四边 0（常见于设计稿导出）：无内边距，且框随文本增高
    const zero = await txBodyToHTML(el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr>
      <a:p><a:r><a:t>零边距</a:t></a:r></a:p>
    </p:txBody>`), theme)
    expect(zero.padding).toBe(0)
    expect(zero.autoFitShape).toBe(true)
    // 缺省 insets：OOXML 默认 tIns/bIns=45720 → (45720+45720)/2/9525 = 4.8 → 5
    const def = await txBodyToHTML(el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r><a:t>默认</a:t></a:r></a:p>
    </p:txBody>`), theme)
    expect(def.padding).toBe(5)
    expect(def.autoFitShape).toBe(false)
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
    // 字体名剔除引号/分号（fontStackOf），不允许出现未转义的属性逃逸
    expect(r.html).not.toContain('" onmouseover')
    expect(r.html).not.toMatch(/onmouseover="/)
  })

  it('run 的 a:latin + a:ea → 完整字体栈（latin、ea、中文回退）', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>字体</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain(`font-family:'Arial', '微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif`)
  })

  it('主题字体引用 +mn-lt / +mn-ea 解析为主题字体', async () => {
    const themeEa: PptxTheme = { ...theme, minorEaFont: '汉仪雅酷黑' }
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>正文</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, themeEa)
    expect(r.html).toContain(`font-family:'Calibri', '汉仪雅酷黑', 'PingFang SC', 'Microsoft YaHei', sans-serif`)
  })

  it('a:fld 字段（页码/日期）渲染为文本 run，不丢弃', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:fld id="{GUID}" type="slidenum"><a:rPr sz="1800"/><a:t>3</a:t></a:fld></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('3')
    expect(r.html).toContain('font-size:24px')
  })

  it('lstStyle lvl1pPr defRPr 作为 run 无显式样式时的默认值', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:lstStyle>
        <a:lvl1pPr><a:defRPr sz="1600"><a:solidFill><a:schemeClr val="bg2"><a:lumMod val="25000"/></a:schemeClr></a:solidFill>
          <a:latin typeface="阿里巴巴普惠体"/><a:ea typeface="阿里巴巴普惠体"/></a:defRPr>
        </a:lvl1pPr>
      </a:lstStyle>
      <a:p><a:r><a:t>默认样式</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('默认样式')
    expect(r.html).toContain('font-size:21.33px') // 1600/100/0.75 = 21.33（保留小数，避免导出往返误差）
    expect(r.html).toContain('font-family:')
    expect(r.html).toContain('阿里巴巴普惠体')
  })

  it('run rPr 显式属性优先于 lstStyle 默认值', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:lstStyle>
        <a:lvl1pPr><a:defRPr sz="1600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:lvl1pPr>
      </a:lstStyle>
      <a:p><a:r><a:rPr sz="3200"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>覆盖</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('font-size:42.67px') // 3200/100/0.75 = 42.67（保留小数）
    expect(r.html).toContain('color:#00FF00')
    expect(r.html).not.toContain('color:#FF0000')
  })

  it('lstStyle 按段落 lvl 选取对应级别默认值', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:lstStyle>
        <a:lvl1pPr><a:defRPr sz="1600"/></a:lvl1pPr>
        <a:lvl2pPr><a:defRPr sz="2800"/></a:lvl2pPr>
      </a:lstStyle>
      <a:p><a:pPr lvl="1"/><a:r><a:t>二级</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('font-size:37.33px') // 2800/100/0.75 = 37.33（保留小数）
  })

  it('超链接 External：白名单协议产出 <a href>', async () => {
    const pkg = await makePkg('https://example.com')
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <a:bodyPr/>
      <a:p><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>链接</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme, pkg, 'ppt/slides/slide1.xml')
    // run 未声明字体时回退主题字体（含 font-family 内联样式），故只断言 href 前缀
    expect(r.html).toContain('<a href="https://example.com"')
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
    expect(r.html).toContain('line-height:26.67px') // 2000/100/0.75 = 26.67（保留小数，避免导出往返误差）.67 → 四舍五入 27
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

  it('bodyPr anchor → 垂直对齐（ctr/b/t，未声明为 null）', async () => {
    const mk = (anchor: string) => el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr${anchor}/>
      <a:p><a:r><a:t>字</a:t></a:r></a:p>
    </p:txBody>`)
    expect((await txBodyToHTML(mk(' anchor="ctr"'), theme)).anchor).toBe('ctr')
    expect((await txBodyToHTML(mk(' anchor="b"'), theme)).anchor).toBe('b')
    expect((await txBodyToHTML(mk(' anchor="t"'), theme)).anchor).toBe('t')
    expect((await txBodyToHTML(mk(''), theme)).anchor).toBeNull()
  })

  it('空心字（noFill + a:ln 描边）→ 用描边色近似为文字色', async () => {
    // 设计稿常用空心描边字（如目录页 Catalogue）：编辑器不支持文字描边，取描边色保证不丢色
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:r>
        <a:rPr sz="4400"><a:ln w="6350"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln><a:noFill/></a:rPr>
        <a:t>Catalogue</a:t>
      </a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('color:#4472C4')
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
