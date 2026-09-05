/**
 * 图表类型注册表与 ECharts 配置构建（移植自 eflink-excel，视觉与交互保持一致）。
 * 插入面板的缩略图预览与画布渲染共用同一份类型定义和构建逻辑，保证「所见即所插」。
 * 数据约定：labels 为类别名，series 为若干命名数值系列；
 * 饼图/圆环取 labels + 第一个系列；散点图 x 取 labels（可数值化时用数值轴，否则用序号）。
 */
import type { EChartsOption } from 'echarts'
import type { ChartData } from '../../types/slides'

export interface ChartTypeSpec {
  id: string
  label: string
}

export interface ChartGroup {
  group: string
  types: ChartTypeSpec[]
}

/** 图表类型注册表（顺序即面板展示顺序，与 eflink-excel 一致） */
export const CHART_GROUPS: ChartGroup[] = [
  {
    group: '柱状图',
    types: [
      { id: 'bar-cluster', label: '簇状柱状图' },
      { id: 'bar-stack', label: '堆积柱状图' },
      { id: 'bar-percent', label: '百分比堆积柱状图' },
    ],
  },
  {
    group: '折线图',
    types: [
      { id: 'line', label: '折线图' },
      { id: 'line-stack', label: '堆积折线图' },
      { id: 'line-marker', label: '带数据标记的折线图' },
    ],
  },
  {
    group: '饼图',
    types: [
      { id: 'pie', label: '饼图' },
      { id: 'pie-doughnut', label: '圆环图' },
    ],
  },
  {
    group: '条形图',
    types: [
      { id: 'bar-horizontal', label: '条形图' },
      { id: 'bar-horizontal-stack', label: '堆积条形图' },
    ],
  },
  {
    group: '面积图',
    types: [
      { id: 'area', label: '面积图' },
      { id: 'area-stack', label: '堆积面积图' },
    ],
  },
  {
    group: '散点图',
    types: [{ id: 'scatter', label: '散点图' }],
  },
  {
    group: '雷达图',
    types: [{ id: 'radar', label: '雷达图' }],
  },
]

/** 全部图表类型 id 集合（用于类型归属判断） */
export const CHART_TYPE_IDS = new Set(CHART_GROUPS.flatMap((g) => g.types.map((t) => t.id)))

/** 图表选择面板缩略图的示例数据 */
export const SAMPLE_CHART_DATA: ChartData = {
  labels: ['一月', '二月', '三月'],
  series: [
    { name: '销售额', values: [120, 210, 170] },
    { name: '利润', values: [90, 150, 160] },
  ],
}

/** 图表元素开关（对应设置面板的「图表元素」，与 eflink-excel 一致） */
export interface ChartElements {
  legend: boolean // 图例
  gridLine: boolean // 网格线
  axis: boolean // 坐标轴
  trendline: boolean // 趋势线（柱状/折线类，最小二乘拟合）
  chartTitle: boolean // 图表标题
  dataLabel: boolean // 数据标签
  axisTitle: boolean // 轴标题
}

export const DEFAULT_CHART_ELEMENTS: ChartElements = {
  legend: true,
  gridLine: true,
  axis: true,
  trendline: false,
  chartTitle: false,
  dataLabel: false,
  axisTitle: false,
}

/** 图表配色方案（与「配色方案」面板对应，与 eflink-excel 一致） */
export const CHART_PALETTES: string[][] = [
  ['#5b8ff9', '#61ddaa', '#f6bd16', '#7262fd', '#78d3f8', '#9661bc'],
  ['#3370ff', '#34c724', '#ff9f1a', '#f5556a', '#8a5cf6', '#00b5d8'],
  ['#1f4e79', '#2e75b6', '#9dc3e6', '#ffd966', '#c55a11', '#70ad47'],
  ['#334155', '#64748b', '#0ea5e9', '#22c55e', '#f59e0b', '#ef4444'],
]

/** 渲染前的构建元信息 */
export interface ChartBuildMeta {
  elements?: Partial<ChartElements>
  paletteIndex?: number
  title?: string
  /** 自定义系列配色（优先于配色方案） */
  chartColors?: string[]
  /** 轴/图例字号 */
  fontSize?: number
  /** 网格线颜色 */
  gridColor?: string
  /** 柱状圆角 */
  barBorderRadius?: number
}

/** 线性趋势线（最小二乘） */
function trendData(data: number[]): number[] {
  const n = data.length
  if (n < 2) return [...data]
  let sx = 0
  let sy = 0
  let sxy = 0
  let sxx = 0
  data.forEach((y, x) => {
    sx += x
    sy += y
    sxy += x * y
    sxx += x * x
  })
  const denom = n * sxx - sx * sx
  if (!denom) return [...data]
  const k = (n * sxy - sx * sy) / denom
  const b = (sy - k * sx) / n
  return data.map((_, x) => +(k * x + b).toFixed(2))
}

/** 散点图 x 值：labels 全部可数值化时用数值，否则用序号 */
function scatterXValues(labels: string[]): number[] {
  const numeric = labels.map((l) => Number.parseFloat(l))
  if (labels.length > 0 && numeric.every((n) => Number.isFinite(n))) return numeric
  return labels.map((_, i) => i)
}

/**
 * 构建图表 ECharts 配置。
 * @param data 图表数据（labels + series）
 * @param typeId CHART_GROUPS 中注册的类型 id
 * @param mini 缩略图模式：压缩边距、隐藏图例、缩小字号
 * @param meta 元素开关 / 配色 / 标题等
 */
export function buildChartOption(
  data: ChartData,
  typeId: string,
  mini = false,
  meta?: ChartBuildMeta,
): EChartsOption {
  const el: ChartElements = { ...DEFAULT_CHART_ELEMENTS, ...meta?.elements }
  const customColors = meta?.chartColors?.filter(Boolean)
  const palette = customColors?.length
    ? customColors
    : CHART_PALETTES[Math.min(meta?.paletteIndex ?? 0, CHART_PALETTES.length - 1)]
  const baseFontSize = mini ? 8 : meta?.fontSize ?? 12
  const base: EChartsOption = {
    animation: false,
    silent: true,
    backgroundColor: 'transparent',
    color: palette,
  }

  // echarts 6 图例文字默认颜色在透明背景导出时不可见，需显式给色
  const textStyle = { color: '#1f2329', fontSize: baseFontSize }
  const gridLineColor = meta?.gridColor ?? 'rgba(0,0,0,0.08)'
  const categoryAxis = (categories: string[]) => ({
    type: 'category' as const,
    data: categories,
    axisLabel: { fontSize: baseFontSize },
  })
  const valueAxis = () => ({
    type: 'value' as const,
    axisLabel: { fontSize: baseFontSize },
    splitLine: { lineStyle: { color: gridLineColor } },
  })
  const grid = mini
    ? { left: 30, right: 6, top: 8, bottom: 16 }
    : { left: 64, right: 28, top: 48, bottom: 40 }

  const { labels, series } = data

  // 元素开关：标题开启时在顶部显示标题文字（图例相应下移，避免与标题重叠）
  const titleOption = el.chartTitle
    ? { title: { text: meta?.title || '图表标题', left: 'center', top: 6, textStyle: { fontSize: baseFontSize + 2, color: '#1f2329' } } }
    : {}
  const legendTop = el.chartTitle ? 34 : 8

  // 雷达图：指标取 labels，每个系列一个多边形
  if (typeId === 'radar') {
    const maxByRow = Math.max(...series.flatMap((s) => s.values), 1)
    return {
      ...base,
      ...titleOption,
      legend: { show: !mini && el.legend, bottom: mini ? undefined : 4, textStyle },
      radar: {
        indicator: labels.map((name) => ({ name, max: Math.ceil(maxByRow * 1.2) })),
        radius: mini ? '62%' : '66%',
        axisName: { fontSize: baseFontSize },
      },
      series: [
        {
          type: 'radar',
          data: series.map((s) => ({ value: s.values, name: s.name })),
          areaStyle: { opacity: 0.15 },
        },
      ],
    }
  }

  // 饼图 / 圆环图：labels 为扇区名，第一个系列为数值
  if (typeId === 'pie' || typeId === 'pie-doughnut') {
    const pieData = labels.map((name, i) => ({ name, value: series[0]?.values[i] ?? 0 }))
    return {
      ...base,
      ...titleOption,
      legend: { show: !mini && el.legend, orient: 'vertical', right: mini ? 0 : 10, top: 'middle', textStyle },
      series: [
        {
          type: 'pie',
          radius: typeId === 'pie-doughnut' ? ['30%', '64%'] : '62%',
          center: mini ? ['50%', '50%'] : ['38%', '52%'],
          data: pieData,
          label: { show: el.dataLabel && !mini, formatter: '{b}: {d}%' },
        },
      ],
    }
  }

  // 散点图：双数值轴，x 取 labels（可数值化）或序号
  if (typeId === 'scatter') {
    const xs = scatterXValues(labels)
    return {
      ...base,
      ...titleOption,
      grid,
      legend: { show: !mini && el.legend, top: mini ? undefined : legendTop, textStyle },
      xAxis: { ...valueAxis(), splitLine: { show: el.gridLine, lineStyle: { color: gridLineColor } } },
      yAxis: { ...valueAxis(), splitLine: { show: el.gridLine, lineStyle: { color: gridLineColor } } },
      series: series.map((s) => ({
        type: 'scatter',
        name: s.name,
        data: s.values.map((y, i) => [xs[i], y]),
        symbolSize: mini ? 4 : 12,
      })),
    }
  }

  // 其余：柱状 / 折线 / 条形 / 面积（按变体处理）
  const stacked = typeId === 'bar-stack' || typeId === 'line-stack' || typeId === 'area-stack' || typeId === 'bar-horizontal-stack'
  const percent = typeId === 'bar-percent'
  const horizontal = typeId === 'bar-horizontal' || typeId === 'bar-horizontal-stack'

  let chartSeries: Record<string, unknown>[] = []
  if (typeId.startsWith('bar')) {
    const radius = meta?.barBorderRadius ?? 0
    chartSeries = series.map((s) => ({
      type: 'bar',
      name: s.name,
      data: s.values,
      barMaxWidth: mini ? 8 : 48,
      ...(radius ? { itemStyle: { borderRadius: radius } } : {}),
      ...(stacked || percent ? { stack: 'total' } : {}),
    }))
  } else if (typeId.startsWith('area')) {
    chartSeries = series.map((s) => ({
      type: 'line',
      name: s.name,
      data: s.values,
      areaStyle: {},
      ...(stacked ? { stack: 'total' } : {}),
      showSymbol: false,
    }))
  } else {
    const withSymbol = typeId === 'line-marker'
    chartSeries = series.map((s) => ({
      type: 'line',
      name: s.name,
      data: s.values,
      ...(stacked ? { stack: 'total' } : {}),
      ...(withSymbol ? { symbol: 'circle', symbolSize: mini ? 3 : 8, label: { show: !mini } } : {}),
    }))
  }

  if (percent) {
    // 百分比堆积：每个类别内部归一化为百分比
    const totals = labels.map((_, r) => series.reduce((sum, s) => sum + (s.values[r] ?? 0), 0))
    chartSeries = chartSeries.map((s) => ({
      ...s,
      data: (s.data as number[]).map((v, r) => (totals[r] ? +((v / totals[r]) * 100).toFixed(2) : 0)),
    }))
  }

  if (el.dataLabel && !mini) {
    chartSeries = chartSeries.map((s) => ({ ...s, label: { show: true } }))
  }

  if (el.trendline && (typeId.startsWith('bar') || typeId.startsWith('line') || typeId.startsWith('area'))) {
    // 线性趋势线：每个系列一条虚线（不进图例）
    chartSeries.push(
      ...series.map((s) => ({
        type: 'line',
        name: `趋势(${s.name})`,
        data: trendData(s.values),
        showSymbol: false,
        lineStyle: { type: 'dashed' as const, opacity: 0.7 },
        ...(typeId.startsWith('area') ? { areaStyle: { opacity: 0 } } : {}),
      })),
    )
  }

  const categoryAxisOpt = {
    ...categoryAxis(labels),
    axisLine: { show: el.axis },
    axisTick: { show: el.axis },
    axisLabel: { ...categoryAxis(labels).axisLabel, show: el.axis },
  }
  const valueAxisOpt = {
    ...valueAxis(),
    axisLabel: { ...valueAxis().axisLabel, show: el.axis },
    splitLine: { ...valueAxis().splitLine, show: el.gridLine },
    ...(el.axisTitle ? { name: '数值' } : {}),
  }
  return {
    ...base,
    ...titleOption,
    grid,
    legend: { show: !mini && el.legend, top: mini ? undefined : legendTop, textStyle, data: series.map((s) => s.name) },
    ...(horizontal
      ? { xAxis: valueAxisOpt, yAxis: { ...categoryAxisOpt, data: labels } }
      : { xAxis: categoryAxisOpt, yAxis: valueAxisOpt }),
    series: chartSeries as EChartsOption['series'],
  }
}
