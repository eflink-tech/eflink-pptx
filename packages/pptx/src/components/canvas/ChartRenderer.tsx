// ECharts 图表渲染（与 eflink-excel 共用同一套配置构建，视觉一致）
// 使用 SVG 渲染器：html-to-image 截图管线（缩略图/图片导出/降级导出）可直接捕获，且任意缩放清晰
import { useEffect, useMemo, useRef } from 'react'
import * as echarts from 'echarts'
import type { ChartElement } from '../../types/slides'
import { buildChartOption } from '../../core/chart/chartOptions'
import { normalizeChartElement } from '../../core/chart/migrate'

export const DEFAULT_CHART_COLORS = [
  '#d14424', '#42a5f5', '#7cb342', '#f5c56b', '#8e6bbf',
  '#26c6da', '#ef5350', '#66bb6a', '#ffa726', '#5e6db8',
]

interface Props {
  el: ChartElement
}

export function ChartRenderer({ el }: Props) {
  // 防御性规范化：旧文档/导入数据未经 loadDocument 规范化时也能正确渲染
  const chartEl = useMemo(() => normalizeChartElement(el), [el])
  const hostRef = useRef<HTMLDivElement>(null)
  const chartRef = useRef<echarts.ECharts | null>(null)

  const option = useMemo(
    () => buildChartOption(chartEl.data, chartEl.chartType, false, {
      elements: chartEl.elements,
      paletteIndex: chartEl.paletteIndex,
      title: chartEl.title,
      chartColors: chartEl.chartColors,
      fontSize: chartEl.fontSize,
      gridColor: chartEl.gridColor,
      barBorderRadius: chartEl.barBorderRadius,
    }),
    [chartEl],
  )

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const chart = chartRef.current ?? echarts.init(host, null, { renderer: 'svg', width: el.w, height: el.h })
    chartRef.current = chart
    chart.setOption(option, true)
  }, [el.w, el.h, option])

  useEffect(() => {
    chartRef.current?.resize({ width: el.w, height: el.h })
  }, [el.w, el.h])

  useEffect(() => () => {
    chartRef.current?.dispose()
    chartRef.current = null
  }, [])

  return <div ref={hostRef} className="size-full overflow-hidden" />
}
