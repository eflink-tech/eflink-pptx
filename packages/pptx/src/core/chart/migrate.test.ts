import { describe, expect, it } from 'vitest'
import { normalizeChartElement, normalizePresentation } from './migrate'
import { DEFAULT_CHART_ELEMENTS } from './chartOptions'
import type { ChartElement } from '../../types/slides'

function legacyChart(patch: Partial<ChartElement> = {}): ChartElement {
  return {
    id: 'c-1', type: 'chart', x: 0, y: 0, w: 560, h: 380,
    chartType: 'bar' as ChartElement['chartType'],
    data: { labels: ['一'], series: [{ name: '系列一', values: [1] }] },
    showLegend: true, showLabel: true,
    ...patch,
  } as ChartElement
}

describe('normalizeChartElement', () => {
  it('历史类型 id 映射为 ECharts 版类型', () => {
    expect(normalizeChartElement(legacyChart()).chartType).toBe('bar-cluster')
    expect(normalizeChartElement(legacyChart({ chartType: 'barH' as ChartElement['chartType'] })).chartType).toBe('bar-horizontal')
    expect(normalizeChartElement(legacyChart({ chartType: 'doughnut' as ChartElement['chartType'] })).chartType).toBe('pie-doughnut')
  })

  it('已有类型保持不变', () => {
    expect(normalizeChartElement(legacyChart({ chartType: 'radar' })).chartType).toBe('radar')
    expect(normalizeChartElement(legacyChart({ chartType: 'bar-percent' })).chartType).toBe('bar-percent')
  })

  it('未知类型兜底为簇状柱状图', () => {
    expect(normalizeChartElement(legacyChart({ chartType: 'nope' as ChartElement['chartType'] })).chartType).toBe('bar-cluster')
  })

  it('旧开关字段迁移为 elements', () => {
    const normalized = normalizeChartElement(legacyChart())
    expect(normalized.elements?.legend).toBe(true)
    expect(normalized.elements?.dataLabel).toBe(true)
    expect(normalized.elements?.gridLine).toBe(DEFAULT_CHART_ELEMENTS.gridLine)
  })

  it('无旧字段时使用默认 elements，paletteIndex 补默认值', () => {
    const bare = normalizeChartElement(legacyChart({ showLegend: undefined, showLabel: undefined }))
    expect(bare.elements).toEqual(DEFAULT_CHART_ELEMENTS)
    expect(bare.paletteIndex).toBe(0)
  })

  it('已有 elements 不被旧字段覆盖', () => {
    const normalized = normalizeChartElement(legacyChart({
      elements: { ...DEFAULT_CHART_ELEMENTS, legend: false },
      showLegend: true,
    }))
    expect(normalized.elements?.legend).toBe(false)
  })
})

describe('normalizePresentation', () => {
  it('规范化演示文稿中的图表元素，非图表元素不动', () => {
    const textEl = { id: 't-1', type: 'text' as const, x: 0, y: 0, w: 100, h: 40, content: 'hi' }
    const presentation = {
      id: 'p-1', name: '测试', width: 1280, viewportRatio: 9 / 16,
      slides: [{ id: 's-1', elements: [textEl, legacyChart()] }],
    } as unknown as Parameters<typeof normalizePresentation>[0]
    const normalized = normalizePresentation(presentation)
    const chart = normalized.slides[0]!.elements[1] as ChartElement
    expect(chart.chartType).toBe('bar-cluster')
    expect(normalized.slides[0]!.elements[0]).toBe(textEl)
  })
})
