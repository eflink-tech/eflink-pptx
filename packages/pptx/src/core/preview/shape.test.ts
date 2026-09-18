import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { renderShape, renderLine } from './shape'
import { makeCtx } from './test-utils'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

describe('preview/renderShape', () => {
  it('预设形状：嵌套 svg + SHAPE_PATHS 路径 + solidFill 填充 + 描边', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="4762500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
      <a:ln w="25400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.tagName).toBe('g')
    const nested = g.querySelector('svg')!
    expect(nested.getAttribute('viewBox')).toBe('0 0 100 100')
    expect(nested.getAttribute('preserveAspectRatio')).toBe('none')
    expect(nested.getAttribute('x')).toBe('100')
    expect(nested.getAttribute('width')).toBe('500')
    const path = nested.querySelector('path')!
    expect(path.getAttribute('fill')).toBe('#42A5F5')
    expect(Number(path.getAttribute('stroke-width'))).toBeCloseTo(2.6667, 3) // 25400/9525
  })

  it('渐变填充：defs 登记 linearGradient 全停站，fill 引用 url(#id)', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill><a:gsLst>
        <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
      </a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    const path = g.querySelector('path')!
    const url = path.getAttribute('fill')!
    expect(url.startsWith('url(#grad-')).toBe(true)
    const grad = ctx.defs.querySelector('linearGradient')!
    expect(grad.querySelectorAll('stop')).toHaveLength(3)
    expect(grad.querySelectorAll('stop')[1].getAttribute('stop-color')).toBe('#00FF00')
    // ang=90°：渐变方向向下（y1 < y2）
    expect(Number(grad.getAttribute('y1'))).toBeLessThan(Number(grad.getAttribute('y2')))
  })

  it('阴影：defs 登记 feDropShadow（blur/2 为 stdDeviation）', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.querySelector('path')!.getAttribute('filter')!.startsWith('url(#shadow-')).toBe(true)
    const drop = ctx.defs.querySelector('feDropShadow')!
    expect(Number(drop.getAttribute('stdDeviation'))).toBeCloseTo(50800 / 9525 / 2, 3)
    // dir=2700000（45°）：dist=38100 EMU→4px，dy = 4·sin45° ≈ 2.828（dx≈dy）
    expect(Number(drop.getAttribute('dy'))).toBeCloseTo((38100 / 9525) * Math.sin(Math.PI / 4), 3)
  })

  it('noFill → fill:none；rot 上到 g、flip 移到嵌套 svg（文本不随 flip 镜像）', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm rot="1800000" flipH="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.querySelector('path')!.getAttribute('fill')).toBe('none')
    // 外层 g 只保留 rotate（PowerPoint 翻转形状不镜像文字，flip 不能放在含 foreignObject 的外层 g 上）
    expect(g.getAttribute('transform')).toBe('rotate(30,50,50)')
    // flip 移到嵌套 svg：boxTransform({rot:0, flipH:true, x:0,y:0,w:100,h:100})，rotate 先于 flip（p' = R·F·p）由先 rot 后 flip 的结构保证
    expect(g.querySelector('svg')!.getAttribute('transform')).toBe('translate(100,0) scale(-1,1) translate(0,0)')
  })

  it('flipH + txBody：外层 g 不含 scale(-1,1)，嵌套 svg 含；foreignObject 直挂 g 不镜像', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>文字</a:t></a:r></a:p></p:txBody><p:spPr>
      <a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.getAttribute('transform') ?? '').not.toContain('scale(-1,1)') // 无 rot 时 g 甚至无 transform 属性
    const nested = g.querySelector('svg')!
    expect(nested.getAttribute('transform')!).toContain('scale(-1,1)')
    const fo = g.querySelector('foreignObject')!
    expect((fo.firstElementChild as HTMLElement).innerHTML).toContain('文字')
  })

  it('p:cxnSp → line + 箭头 marker', () => {
    const ctx = makeCtx()
    const node = wrap(`<p:cxnSp><p:spPr>
      <a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        <a:tailEnd type="triangle"/></a:ln>
    </p:spPr></p:cxnSp>`)
    const line = renderLine(node, ctx)!
    expect(line.getAttribute('x1')).toBe('0')
    expect(line.getAttribute('y1')).toBe('100') // flipV
    expect(line.getAttribute('stroke')).toBe('#FF0000')
    expect(line.getAttribute('marker-end')!.startsWith('url(#marker-')).toBe(true)
    expect(ctx.defs.querySelector('marker')).not.toBeNull()
  })
})
