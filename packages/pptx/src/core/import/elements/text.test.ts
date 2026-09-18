import { describe, expect, it } from 'vitest'
import { txBodyToHTML } from './text'
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
})
