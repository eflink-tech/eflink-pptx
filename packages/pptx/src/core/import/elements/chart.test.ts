import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseChartEl, mapChartType } from './chart'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

const CHART_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列一</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt idx="0"><c:v>一月</c:v></c:pt><c:pt idx="1"><c:v>二月</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`

async function makeCtx(chartXml: string = CHART_XML) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="urn:p" xmlns:c="urn:c"/>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`)
  zip.file('ppt/charts/chart1.xml', chartXml)
  const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
  const ctx = {
    pkg,
    partPath: 'ppt/slides/slide1.xml',
    theme,
    report: { skipped: {} } as ImportReport,
    scale: { x: 1, y: 1 },
    placeholders: new Map(),
  }
  return { pkg, ctx }
}

const FRAME_XML = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c">
  <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3" xmlns:r="urn:r"/>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('mapChartType', () => {
  it('barChart 组合映射', () => {
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'clustered' })).toBe('bar-cluster')
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'stacked' })).toBe('bar-stack')
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'percentStacked' })).toBe('bar-percent')
    expect(mapChartType('barChart', { barDir: 'bar', grouping: 'clustered' })).toBe('bar-horizontal')
    expect(mapChartType('barChart', { barDir: 'bar', grouping: 'stacked' })).toBe('bar-horizontal-stack')
  })
  it('line/pie/area/scatter/radar 映射', () => {
    expect(mapChartType('lineChart', { marker: true })).toBe('line-marker')
    expect(mapChartType('lineChart', {})).toBe('line')
    expect(mapChartType('lineChart', { grouping: 'stacked' })).toBe('line-stack')
    expect(mapChartType('pieChart', {})).toBe('pie')
    expect(mapChartType('doughnutChart', {})).toBe('pie-doughnut')
    expect(mapChartType('areaChart', {})).toBe('area')
    expect(mapChartType('areaChart', { grouping: 'stacked' })).toBe('area-stack')
    expect(mapChartType('scatterChart', {})).toBe('scatter')
    expect(mapChartType('radarChart', {})).toBe('radar')
  })
})

describe('parseChartEl', () => {
  it('解析 rels 关联的 chart part → ChartElement', async () => {
    const { pkg, ctx } = await makeCtx()
    const result = await parseChartEl(el(FRAME_XML), ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'chart') throw new Error('expected chart')
    expect(result.chartType).toBe('bar-cluster')
    expect(result.data.labels).toEqual(['一月', '二月'])
    expect(result.data.series).toEqual([{ name: '系列一', values: [10, 25] }])
  })

  it('scatterChart：c:xVal/c:yVal → labels= xVal、values= yVal', async () => {
    const SCATTER_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart><c:plotArea>
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
    </c:scatterChart>
  </c:plotArea></c:chart>
</c:chartSpace>`
    const { pkg, ctx } = await makeCtx(SCATTER_XML)
    const result = await parseChartEl(el(FRAME_XML), ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'chart') throw new Error('expected chart')
    expect(result.chartType).toBe('scatter')
    expect(result.data.labels).toEqual(['1', '3'])
    expect(result.data.series).toEqual([{ name: '散点', values: [5, 7] }])
  })
})
