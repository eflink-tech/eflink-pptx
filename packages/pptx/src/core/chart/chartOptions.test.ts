import { describe, expect, it } from 'vitest'
import {
  buildChartOption, CHART_GROUPS, CHART_PALETTES, CHART_TYPE_IDS,
  DEFAULT_CHART_ELEMENTS, SAMPLE_CHART_DATA,
} from './chartOptions'
import type { ChartData } from '../../types/slides'

const data: ChartData = {
  labels: ['一月', '二月', '三月'],
  series: [
    { name: '销售额', values: [120, 210, 170] },
    { name: '利润', values: [90, 150, 160] },
  ],
}

function seriesOf(typeId: string): Array<Record<string, unknown>> {
  const option = buildChartOption(data, typeId)
  return option.series as Array<Record<string, unknown>>
}

describe('CHART_GROUPS 注册表', () => {
  it('对齐 eflink-excel：7 组 14 种类型', () => {
    expect(CHART_GROUPS.map((g) => g.group)).toEqual(['柱状图', '折线图', '饼图', '条形图', '面积图', '散点图', '雷达图'])
    expect(CHART_TYPE_IDS.size).toBe(14)
  })
})

describe('buildChartOption 各类型构建', () => {
  it('柱状图变体：堆积与百分比堆积', () => {
    const cluster = seriesOf('bar-cluster')
    expect(cluster).toHaveLength(2)
    expect(cluster[0].stack).toBeUndefined()

    const stacked = seriesOf('bar-stack')
    expect(stacked[0].stack).toBe('total')

    const percent = seriesOf('bar-percent') as Array<{ data: number[]; stack?: string }>
    // 第一类 120+90=210 → 销售额占比 57.14
    expect(percent[0].data[0]).toBeCloseTo(57.14, 1)
  })

  it('折线图：带标记变体与面积图', () => {
    const marker = seriesOf('line-marker') as Array<Record<string, unknown>>
    expect(marker[0].symbol).toBe('circle')
    expect(seriesOf('area')[0].areaStyle).toBeDefined()
    expect(seriesOf('area-stack')[0].stack).toBe('total')
  })

  it('条形图横向', () => {
    const option = buildChartOption(data, 'bar-horizontal')
    const xAxis = option.xAxis as Record<string, unknown>
    const yAxis = option.yAxis as Record<string, unknown>
    expect(xAxis.type).toBe('value')
    expect(yAxis.type).toBe('category')
  })

  it('饼图取第一个系列，圆环有内半径', () => {
    const pie = buildChartOption(data, 'pie')
    const pieSeries = (pie.series as Array<Record<string, unknown>>)[0]
    expect(pieSeries.type).toBe('pie')
    expect(pieSeries.radius).toBe('62%')
    expect(pieSeries.data).toEqual([
      { name: '一月', value: 120 },
      { name: '二月', value: 210 },
      { name: '三月', value: 170 },
    ])

    const doughnut = buildChartOption(data, 'pie-doughnut')
    expect((doughnut.series as Array<Record<string, unknown>>)[0].radius).toEqual(['30%', '64%'])
  })

  it('散点图：数值 labels 走数值 x，否则用序号', () => {
    const numeric = buildChartOption({ labels: ['1', '2', '3'], series: [{ name: 'y', values: [4, 5, 6] }] }, 'scatter')
    expect(((numeric.series as Array<Record<string, unknown>>)[0].data as number[][])[0]).toEqual([1, 4])

    const categorical = buildChartOption(data, 'scatter')
    expect(((categorical.series as Array<Record<string, unknown>>)[0].data as number[][])[0]).toEqual([0, 120])
  })

  it('雷达图：指标取 labels，每个系列一个多边形', () => {
    const option = buildChartOption(data, 'radar')
    expect((option.radar as { indicator: unknown[] }).indicator).toHaveLength(3)
    const radarSeries = (option.series as Array<Record<string, unknown>>)[0]
    expect((radarSeries.data as unknown[]).length).toBe(2)
  })

  it('趋势线：为每个系列追加虚线系列且不改变原系列', () => {
    const series = seriesOf('bar-cluster') // 无趋势线
    expect(series).toHaveLength(2)
    const withTrend = buildChartOption(data, 'bar-cluster', false, {
      elements: { ...DEFAULT_CHART_ELEMENTS, trendline: true },
    })
    expect((withTrend.series as unknown[]).length).toBe(4)
  })

  it('元素开关：mini 模式隐藏图例，dataLabel 控制标签', () => {
    const mini = buildChartOption(data, 'bar-cluster', true)
    expect(((mini.legend as { show?: boolean })).show).toBe(false)

    const labeled = buildChartOption(data, 'bar-cluster', false, {
      elements: { ...DEFAULT_CHART_ELEMENTS, dataLabel: true },
    })
    const labeledSeries = labeled.series as Array<Record<string, unknown>>
    expect((labeledSeries[0].label as { show?: boolean }).show).toBe(true)
  })

  it('配色：自定义 chartColors 优先于色板', () => {
    const option = buildChartOption(data, 'bar-cluster', false, {
      chartColors: ['#ff0000', '#00ff00'],
      paletteIndex: 2,
    })
    expect(option.color).toEqual(['#ff0000', '#00ff00'])

    const paletted = buildChartOption(data, 'bar-cluster', false, { paletteIndex: 1 })
    expect(paletted.color).toEqual(CHART_PALETTES[1])
  })

  it('示例数据对全部注册类型可构建（缩略图冒烟）', () => {
    for (const { types } of CHART_GROUPS) {
      for (const { id } of types) {
        const option = buildChartOption(SAMPLE_CHART_DATA, id, true)
        expect(option.series, id).toBeTruthy()
      }
    }
  })
})
