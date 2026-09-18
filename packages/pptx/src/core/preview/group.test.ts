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
    // 完整串断言：rotate 在串首（OOXML p'=R·T·S·T'·p），child(50,25) 子空间中心映射后恰为组中心 (200,150) 不动点
    expect(g.getAttribute('transform')).toBe('rotate(90,200,150) translate(100,100) scale(2,2) translate(0,0)')
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
