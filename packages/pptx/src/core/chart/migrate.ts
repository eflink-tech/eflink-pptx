// 旧图表数据兼容：历史文档中的 8 种基础类型与独立开关字段，统一规范化为
// ECharts 版的 14 种类型 + elements 开关。loadDocument / parseJSONFile 载入时调用，
// ChartRenderer 内部也会做一次防御性规范化（覆盖缩略图、播放器、导出等所有渲染路径）。
import { DEFAULT_CHART_ELEMENTS, CHART_TYPE_IDS, type ChartElements } from './chartOptions'
import type { ChartElement, ChartType, Presentation, Slide } from '../../types/slides'

/** 历史 chartType → 新类型 id */
const LEGACY_CHART_TYPE_MAP: Record<string, ChartType> = {
  bar: 'bar-cluster',
  barH: 'bar-horizontal',
  doughnut: 'pie-doughnut',
}

function normalizeChartType(chartType: ChartType): ChartType {
  if (chartType in LEGACY_CHART_TYPE_MAP) return LEGACY_CHART_TYPE_MAP[chartType]!
  // 未知/非法类型兜底为簇状柱状图，避免渲染崩溃
  return CHART_TYPE_IDS.has(chartType) ? chartType : 'bar-cluster'
}

export function normalizeChartElement<T extends ChartElement>(el: T): T {
  const chartType = normalizeChartType(el.chartType)
  const elements: ChartElements = {
    ...DEFAULT_CHART_ELEMENTS,
    ...el.elements,
    // 旧文档的独立开关仅在未写过 elements 时生效
    ...(el.elements ? {} : { legend: el.showLegend ?? DEFAULT_CHART_ELEMENTS.legend, dataLabel: el.showLabel ?? DEFAULT_CHART_ELEMENTS.dataLabel }),
  }
  return {
    ...el,
    chartType,
    elements,
    paletteIndex: el.paletteIndex ?? 0,
  }
}

function normalizeSlide(slide: Slide): Slide {
  if (!slide.elements.some((el) => el.type === 'chart')) return slide
  return {
    ...slide,
    elements: slide.elements.map((el) => (el.type === 'chart' ? normalizeChartElement(el) : el)),
  }
}

/** 规范化整个演示文稿中的图表元素（旧文档载入 / JSON 导入时调用） */
export function normalizePresentation(presentation: Presentation): Presentation {
  return { ...presentation, slides: presentation.slides.map(normalizeSlide) }
}
