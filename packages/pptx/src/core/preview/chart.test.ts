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

/** graphicFrame 模板：c:chart 引用 rId1（标题在 chart part 内，graphicFrame 侧无开关） */
const FRAME = `<p:graphicFrame>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

/** 柱状图 part：标题「月度报表」+ 2 系列（系列一 10/20/30，系列二 5/15/25）× 3 类目（A/B/C） */
const BAR_PART = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>月度报表</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列一</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列二</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

/** 饼图 part：1 系列 3 个数据项（40/35/25） */
const PIE_PART = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>占比</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>35</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`

/** 雷达图 part：1 系列 3 个数据项（预览层降级占位） */
const RADAR_PART = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart><c:plotArea>
    <c:radarChart>
      <c:radarStyle val="marker"/>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>维度</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>80</c:v></c:pt><c:pt idx="2"><c:v>70</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:radarChart>
  </c:plotArea></c:chart>
</c:chartSpace>`

describe('preview/renderChart', () => {
  it('柱状图：每系列×类目一个 rect，标题渲染', async () => {
    const ctx = makeCtx({ pkg: stubPkg(BAR_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('rect').length).toBeGreaterThanOrEqual(6) // 2 系列 × 3 类目
    const title = g.querySelector('text')
    expect(title?.textContent).toBe('月度报表')
  })

  it('饼图：每数据项一个扇形 path', async () => {
    const ctx = makeCtx({ pkg: stubPkg(PIE_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('path').length).toBe(3)
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
