import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './test-utils'
import { renderChart } from './chart'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

/** 最简 PptxPackage 桩：parseChartEl 只用到 relTarget / text */
function stubPkg(chartPart: string): never {
  return {
    relTarget: async () => 'ppt/charts/chart1.xml',
    text: async (p: string) => (p === 'ppt/charts/chart1.xml' ? chartPart : null),
    rels: async () => new Map(),
    mediaDataUrl: async () => undefined,
  } as never
}

/** graphicFrame 模板：c:chart 引用 rId1（p:xfrm 630×420 px，源画布坐标） */
const FRAME = `<p:graphicFrame>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

/** 预期绘图区（plotRect）：box 整数化后 630×420，留轴标签/标题边距 */
const P = { x: 36, y: 24, w: 630 - 44, h: 420 - 48 }

/** 通用系列片段：strCache 类目 + numCache 数值 */
const serBody = (name: string, cats: string[], vals: number[]) => `<c:ser>
  <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:cat><c:strRef><c:strCache>
    ${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('\n    ')}
  </c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:numCache>
    ${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('\n    ')}
  </c:numCache></c:numRef></c:val>
</c:ser>`

const chartPart = (kindXml: string, withTitle: boolean) => `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart>
    ${withTitle ? '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>月度报表</a:t></a:r></a:p></c:rich></c:tx></c:title>' : ''}
    <c:plotArea>${kindXml}</c:plotArea>
  </c:chart>
</c:chartSpace>`

/** 柱状图 part（2 系列 10/20/30、5/15/25 × 3 类目 A/B/C，带标题） */
const barPart = (dir: 'col' | 'bar', grouping: string) => chartPart(`
  <c:barChart>
    <c:barDir val="${dir}"/><c:grouping val="${grouping}"/>
    ${serBody('系列一', ['A', 'B', 'C'], [10, 20, 30])}
    ${serBody('系列二', ['A', 'B', 'C'], [5, 15, 25])}
  </c:barChart>`, true)
const BAR_PART = barPart('col', 'clustered')
const BAR_STACK_PART = barPart('col', 'stacked')
const BAR_HSTACK_PART = barPart('bar', 'stacked')

/** 折线图 part（2 系列，同上数据；marker 控制是否带数据点标记） */
const linePart = (withMarker: boolean) => chartPart(`
  <c:lineChart>
    ${withMarker ? '<c:marker val="1"/>' : ''}<c:grouping val="standard"/>
    ${serBody('系列一', ['A', 'B', 'C'], [10, 20, 30])}
    ${serBody('系列二', ['A', 'B', 'C'], [5, 15, 25])}
  </c:lineChart>`, false)
const LINE_PART = linePart(false)
const LINE_MARKER_PART = linePart(true)

/** 面积图 part（2 系列，非堆叠） */
const AREA_PART = chartPart(`
  <c:areaChart>
    <c:grouping val="standard"/>
    ${serBody('系列一', ['A', 'B', 'C'], [10, 20, 30])}
    ${serBody('系列二', ['A', 'B', 'C'], [5, 15, 25])}
  </c:areaChart>`, false)

/** 散点图 part（1 系列，xVal 1/3，yVal 5/7——xVal/yVal 必须在 ser 内部） */
const SCATTER_PART = chartPart(`
  <c:scatterChart>
    <c:ser>
      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>散点</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:xVal><c:numRef><c:numCache>
        <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt>
      </c:numCache></c:numRef></c:xVal>
      <c:yVal><c:numRef><c:numCache>
        <c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt>
      </c:numCache></c:numRef></c:yVal>
    </c:ser>
  </c:scatterChart>`, false)

/** 饼图/圆环 part（1 系列，指定数值） */
const piePart = (vals: number[], kind: 'pieChart' | 'doughnutChart' = 'pieChart') => chartPart(`
  <c:${kind}>
    <c:varyColors val="1"/>
    ${serBody('占比', ['A', 'B', 'C'], vals)}
  </c:${kind}>`, false)
const PIE_PART = piePart([40, 35, 25])
const DOUGHNUT_PART = piePart([40, 35, 25], 'doughnutChart')
const PIE_FULL_PART = piePart([100])

/** 雷达图 part（预览层降级占位） */
const RADAR_PART = chartPart(`
  <c:radarChart>
    <c:radarStyle val="marker"/>
    ${serBody('维度', ['A', 'B', 'C'], [60, 80, 70])}
  </c:radarChart>`, false)

describe('preview/renderChart', () => {
  it('柱状图：每系列×类目一个 rect（toBe 6），标题渲染 + 最高柱坐标精确', async () => {
    const ctx = makeCtx({ pkg: stubPkg(BAR_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('rect')).toHaveLength(6)
    const title = g.querySelector('text')
    expect(title?.textContent).toBe('月度报表')
    // 最高柱：类目 3（i=2）系列 1（si=0，v=30）；clustered 双系列各占半带宽，maxV = 30×1.1 = 33
    const rects = Array.from(g.querySelectorAll('rect'))
    const band = P.w / 3
    const half = band / 2
    const h = (P.h * 30) / (30 * 1.1)
    expect(Number(rects[4].getAttribute('x'))).toBeCloseTo(P.x + band * 2 + half * 0.1, 1)
    expect(Number(rects[4].getAttribute('width'))).toBeCloseTo(half * 0.8, 1)
    expect(Number(rects[4].getAttribute('y'))).toBeCloseTo(P.y + P.h - h, 1)
    expect(Number(rects[4].getAttribute('height'))).toBeCloseTo(h, 1)
  })

  it('柱状堆叠图：顶不溢出绘图区（累计和 55 → maxV 60.5）', async () => {
    const ctx = makeCtx({ pkg: stubPkg(BAR_STACK_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const rects = Array.from(g.querySelectorAll('rect'))
    expect(rects).toHaveLength(6)
    const maxV = 55 * 1.1
    // 最顶层段（类目 3 系列累计 55）顶部 y
    expect(Number(rects[5].getAttribute('y'))).toBeCloseTo(P.y + P.h - (P.h * 55) / maxV, 1)
    for (const r of rects) {
      const y = Number(r.getAttribute('y'))
      const h = Number(r.getAttribute('height'))
      expect(y).toBeGreaterThanOrEqual(P.y - 0.01)
      expect(y + h).toBeLessThanOrEqual(P.y + P.h + 0.01)
    }
  })

  it('横向堆叠图：同行各段沿值轴递增、width 对应各自 v/maxV×p.w、占满整行', async () => {
    const ctx = makeCtx({ pkg: stubPkg(BAR_HSTACK_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const rects = Array.from(g.querySelectorAll('rect'))
    expect(rects).toHaveLength(6)
    const maxV = 55 * 1.1
    // 类目 1 两段：系列一(10) 在左，系列二(5) 接在其右
    expect(Number(rects[0].getAttribute('x'))).toBeCloseTo(P.x, 1)
    expect(Number(rects[0].getAttribute('width'))).toBeCloseTo((10 / maxV) * P.w, 1)
    expect(Number(rects[1].getAttribute('x'))).toBeCloseTo(P.x + (10 / maxV) * P.w, 1)
    expect(Number(rects[1].getAttribute('width'))).toBeCloseTo((5 / maxV) * P.w, 1)
    // 占满整行：y=类目带起点、高=band
    expect(Number(rects[1].getAttribute('y'))).toBeCloseTo(P.y, 1)
    expect(Number(rects[1].getAttribute('height'))).toBeCloseTo(P.h / 3, 1)
  })

  it('折线图：每系列一条 polyline（3 个点对），marker 变体补 6 个数据点圆', async () => {
    const ctx = makeCtx({ pkg: stubPkg(LINE_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const lines = Array.from(g.querySelectorAll('polyline'))
    expect(lines).toHaveLength(2)
    expect(lines[0].getAttribute('stroke')).toBe('#4472C4')
    expect(lines[1].getAttribute('stroke')).toBe('#ED7D31')
    expect(lines[0].getAttribute('points')!.split(' ')).toHaveLength(3)
    expect(g.querySelectorAll('circle')).toHaveLength(0)

    const ctx2 = makeCtx({ pkg: stubPkg(LINE_MARKER_PART) })
    const g2 = (await renderChart(wrap(FRAME), ctx2))!
    // 2 系列 × 3 数据点 = 6 个 marker 圆
    expect(g2.querySelectorAll('circle')).toHaveLength(6)
  })

  it('面积图：每系列一个闭合 path（fill-opacity 0.5），且画坐标轴', async () => {
    const ctx = makeCtx({ pkg: stubPkg(AREA_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const paths = Array.from(g.querySelectorAll('path'))
    expect(paths).toHaveLength(2)
    expect(paths[0].getAttribute('fill')).toBe('#4472C4')
    expect(paths[0].getAttribute('fill-opacity')).toBe('0.5')
    expect(paths[0].getAttribute('d')!.startsWith('M')).toBe(true)
    expect(paths[0].getAttribute('d')!.endsWith('Z')).toBe(true)
    // 笛卡尔类画轴线
    expect(g.querySelectorAll('line')).toHaveLength(2)
  })

  it('散点图：每数据项一个圆，落在绘图区内', async () => {
    const ctx = makeCtx({ pkg: stubPkg(SCATTER_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const dots = Array.from(g.querySelectorAll('circle'))
    expect(dots).toHaveLength(2)
    for (const c of dots) {
      const cx = Number(c.getAttribute('cx'))
      const cy = Number(c.getAttribute('cy'))
      expect(cx).toBeGreaterThanOrEqual(P.x)
      expect(cx).toBeLessThanOrEqual(P.x + P.w)
      expect(cy).toBeGreaterThanOrEqual(P.y)
      expect(cy).toBeLessThanOrEqual(P.y + P.h)
    }
  })

  it('饼图：每数据项一个扇形 path，不画坐标轴；圆环为扇环（双弧）', async () => {
    const ctx = makeCtx({ pkg: stubPkg(PIE_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('path')).toHaveLength(3)
    // 饼图无坐标轴
    expect(g.querySelectorAll('line')).toHaveLength(0)

    const ctx2 = makeCtx({ pkg: stubPkg(DOUGHNUT_PART) })
    const g2 = (await renderChart(wrap(FRAME), ctx2))!
    const paths = Array.from(g2.querySelectorAll('path'))
    expect(paths).toHaveLength(3)
    // 扇环路径 = 外弧 + 内弧（两个 A 命令）
    for (const p of paths) expect(p.getAttribute('d')!.match(/A /g)).toHaveLength(2)
  })

  it('单片满圆饼：特判画整圆而非退化 path', async () => {
    const ctx = makeCtx({ pkg: stubPkg(PIE_FULL_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('path')).toHaveLength(0)
    expect(g.querySelector('circle')).not.toBeNull()
  })

  it('边界：labels 为空（缺 c:cat）按单类目兜底不崩溃；负值 clamp 到基线', async () => {
    const noCat = BAR_PART.replace(/<c:cat>[\s\S]*?<\/c:cat>/g, '')
    const ctx = makeCtx({ pkg: stubPkg(noCat) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    // labels=[] → n=1，仍渲染 2 系列
    expect(g.querySelectorAll('rect')).toHaveLength(2)

    const negPart = BAR_PART.replace('<c:v>30</c:v>', '<c:v>-30</c:v>')
    const ctx2 = makeCtx({ pkg: stubPkg(negPart) })
    const g2 = (await renderChart(wrap(FRAME), ctx2))!
    for (const r of g2.querySelectorAll('rect')) {
      expect(Number(r.getAttribute('height'))).toBeGreaterThanOrEqual(0)
      expect(Number(r.getAttribute('width'))).toBeGreaterThanOrEqual(0)
      expect(Number.isFinite(Number(r.getAttribute('y')))).toBe(true)
    }
  })

  it('雷达图：降级占位框', async () => {
    const ctx = makeCtx({ pkg: stubPkg(RADAR_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelector('rect[fill="#F2F2F2"]')).not.toBeNull()
  })

  it('缺失图表 part → null 且计入报告', async () => {
    const ctx = makeCtx({ pkg: stubPkg('') })
    // text 返回空串为 falsy → parseChartEl 走 missingChart 分支
    expect(await renderChart(wrap(FRAME), ctx)).toBeNull()
    expect(ctx.report.skipped.missingChart).toBe(1)
  })
})
