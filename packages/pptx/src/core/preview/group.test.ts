import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './test-utils'
import { renderGroup } from './group'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

/** 组：父框 200×100px（rot 90°），子空间 100×50px，内部红矩形铺满子空间 */
const GRP = `<p:grpSp>
  <p:grpSpPr>
    <a:xfrm rot="5400000">
      <a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/>
      <a:chOff x="0" y="0"/><a:chExt cx="952500" cy="476250"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>
</p:grpSp>`

describe('preview/renderGroup', () => {
  it('SVG transform 表达 child→parent 映射 + 组合旋转（无需降级）', async () => {
    const ctx = makeCtx()
    const g = (await renderGroup(wrap(GRP), ctx))!
    const tf = g.getAttribute('transform')!
    expect(tf).toContain('translate(100,100)')
    expect(tf).toContain('scale(2,2)') // 1905000/9525=200px ÷ 952500/9525=100px；纵向同理 100/50
    expect(tf).toContain('translate(0,0)') // -chOff=0，String(-0) === '0'
    expect(tf).toContain('rotate(90,200,150)') // 绕父框中心（100+200/2, 100+100/2）
    expect(ctx.report.skipped.groupRotation).toBeUndefined() // 不计入降级
    expect(g.querySelector('path')).not.toBeNull()
  })

  it('嵌套组合递归展开', async () => {
    const ctx = makeCtx()
    const outer = wrap(`<p:grpSp><p:grpSpPr><a:xfrm>
      <a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/>
    </a:xfrm></p:grpSpPr>${GRP}</p:grpSp>`)
    const g = (await renderGroup(outer, ctx))!
    expect(g.querySelectorAll('g').length).toBeGreaterThanOrEqual(1)
  })
})
