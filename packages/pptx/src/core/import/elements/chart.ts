// src/core/import/elements/chart.ts
/** p:graphicFrame(c:chart) → chart 元素（OOXML chart part 缓存数据 → ChartData） */
import { attr, directChild, directChildren, firstDescendant, parseXML } from '../xml'
import { genId } from '../../utils/id'
import { addSkipped } from '../context'
import { resolveColor } from '../styles'
import { parseGradient } from './shape'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import { frameRotOf, parseFrameGeom } from './frame-geom'
import type { ChartElement, ChartType, Gradient } from '../../../types/slides'

interface ChartKindHints {
  barDir?: string
  grouping?: string
  marker?: boolean
}

/** chart 部件节点名 + 提示 → 内部 ChartType */
export function mapChartType(kindNode: string, hints: ChartKindHints): ChartType | null {
  const grouping = hints.grouping ?? 'clustered'
  switch (kindNode) {
    case 'barChart':
      if (hints.barDir === 'bar') {
        return grouping === 'stacked' || grouping === 'percentStacked' ? 'bar-horizontal-stack' : 'bar-horizontal'
      }
      return grouping === 'stacked' ? 'bar-stack' : grouping === 'percentStacked' ? 'bar-percent' : 'bar-cluster'
    case 'lineChart':
      if (grouping === 'stacked') return 'line-stack'
      return hints.marker ? 'line-marker' : 'line'
    case 'pieChart': return 'pie'
    case 'doughnutChart': return 'pie-doughnut'
    case 'areaChart': return grouping === 'stacked' ? 'area-stack' : 'area'
    case 'scatterChart': return 'scatter'
    case 'radarChart': return 'radar'
    default: return null
  }
}

function cacheValues(ref: Element | null, tag: string): string[] {
  const cache = ref ? firstDescendant(ref, tag) : null
  if (!cache) return []
  return directChildren(cache, 'c:pt')
    .map((pt) => firstDescendant(pt, 'c:v')?.textContent ?? '')
    .filter((v) => v !== '')
}

export async function parseChartEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<ChartElement | null> {
  const chartRef = firstDescendant(node, 'c:chart')
  const rId = attr(chartRef, 'r:id')
  if (!rId) return null
  const chartPath = await pkg.relTarget(ctx.partPath, rId)
  const xml = chartPath ? await pkg.text(chartPath) : null
  if (!xml) {
    addSkipped(ctx.report, 'missingChart')
    return null
  }
  const root = parseXML(xml).documentElement
  const plotArea = firstDescendant(root, 'c:plotArea')
  if (!plotArea) {
    addSkipped(ctx.report, 'missingChart')
    return null
  }
  // 图表区背景（chartSpace 级 c:spPr）：纯色/渐变/图片可还原；
  // 绘图区（plotArea）背景是内层区域，模型无对应字段，如实报告
  let background: { color?: string; gradient?: Gradient; image?: string } | undefined
  const chartSpPr = directChild(root, 'c:spPr')
  if (chartSpPr) {
    const solid = directChild(chartSpPr, 'a:solidFill')
    const grad = directChild(chartSpPr, 'a:gradFill')
    const blipFill = directChild(chartSpPr, 'a:blipFill')
    if (solid) {
      const color = resolveColor(solid, ctx.theme)
      if (color) background = { color }
    } else if (grad) {
      const gradient = parseGradient(grad, ctx)
      if (gradient) background = { gradient }
    } else if (blipFill) {
      const embedId = attr(directChild(blipFill, 'a:blip'), 'r:embed')
      const target = embedId && chartPath ? await pkg.relTarget(chartPath, embedId) : null
      const src = target ? await pkg.mediaDataUrl(target) : undefined
      if (src) background = { image: src }
      else addSkipped(ctx.report, 'chartBackground')
    } else if (directChild(chartSpPr, 'a:pattFill')) {
      addSkipped(ctx.report, 'chartBackground')
    }
  }
  const plotSpPr = directChild(plotArea, 'c:spPr')
  if (plotSpPr && (directChild(plotSpPr, 'a:solidFill') || directChild(plotSpPr, 'a:gradFill')
    || directChild(plotSpPr, 'a:blipFill') || directChild(plotSpPr, 'a:pattFill'))) {
    addSkipped(ctx.report, 'chartBackground')
  }

  const KIND_NODES = ['barChart', 'lineChart', 'pieChart', 'doughnutChart', 'areaChart', 'scatterChart', 'radarChart']
  // 图表类型节点是 plotArea 的直接子级，按本地名探测（不依赖前缀绑定）
  const kindNodes = Array.from(plotArea.children).filter((c) => KIND_NODES.includes(c.localName))
  const kindNode = kindNodes[0] ?? null
  if (!kindNode) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  // 组合图（柱+线等多类型叠加）：模型单类型，导入首个类型并报告
  if (kindNodes.length > 1) addSkipped(ctx.report, 'chartCombo')
  const sers = directChildren(kindNode, 'c:ser')
  const markerOnSer = sers.some((ser) => {
    const sym = attr(directChild(directChild(ser, 'c:marker') ?? ser, 'c:symbol'), 'val')
    return Boolean(sym) && sym !== 'none'
  })
  const hints: ChartKindHints = {
    barDir: attr(directChild(kindNode, 'c:barDir'), 'val') ?? undefined,
    grouping: attr(directChild(kindNode, 'c:grouping'), 'val') ?? undefined,
    // 折线标记：标准位置在 ser/c:marker/c:symbol；兼容部分生成器写在类型节点上的 c:marker@val
    marker: markerOnSer || attr(directChild(kindNode, 'c:marker'), 'val') === '1',
  }
  const chartType = mapChartType(kindNode.localName, hints)
  if (!chartType) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }

  // 数值缓存 → 数字，过滤 NaN/Infinity 等非法值
  const toNumbers = (vals: string[]): number[] => vals.map(Number).filter(Number.isFinite)
  // 数值型标签：仅保留可解析为有限数字的项
  const numericOnly = (vals: string[]): string[] => vals.filter((v) => Number.isFinite(Number(v)))
  // scatter 使用 c:xVal/c:yVal，其余用 c:cat/c:val
  if (!sers.length) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  const isScatter = kindNode.localName === 'scatterChart'
  let labels: string[]
  if (isScatter) {
    labels = numericOnly(cacheValues(firstDescendant(sers[0], 'c:xVal'), 'c:numCache'))
  } else {
    labels = cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:strCache')
    if (!labels.length) labels = numericOnly(cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:numCache'))
  }
  const series = sers.map((ser) => ({
    name: cacheValues(firstDescendant(ser, 'c:tx'), 'c:strCache')[0] ?? '系列',
    values: isScatter
      ? toNumbers(cacheValues(firstDescendant(ser, 'c:yVal'), 'c:numCache'))
      : toNumbers(cacheValues(firstDescendant(ser, 'c:val'), 'c:numCache')),
  }))

  const geom = parseFrameGeom(node, xf, ctx)
  if (!geom) return null

  const chart: ChartElement = {
    id: genId('ch-'), type: 'chart', chartType,
    data: { labels, series },
    x: geom.x, y: geom.y, w: geom.w, h: geom.h,
    name: '图表',
  }
  if (background?.color) chart.backgroundColor = background.color
  if (background?.gradient) chart.backgroundGradient = background.gradient
  if (background?.image) chart.backgroundImage = background.image
  const rot = frameRotOf(firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm'), xf).rot
  if (rot) chart.rotate = Math.round(rot)
  const titleEl = firstDescendant(root, 'c:title')
  const title = titleEl
    ? directChildren(titleEl, 'c:tx').map((tx) => firstDescendant(tx, 'a:t')?.textContent ?? '').join('').trim()
    : ''
  if (title) chart.title = title

  // 系列配色：c:spPr/a:solidFill（柱/面积/饼）或 a:ln/a:solidFill（折线/散点）
  const seriesColorOf = (ser: Element): string | undefined => {
    const spPr = directChild(ser, 'c:spPr')
    if (!spPr) return undefined
    const solid = resolveColor(directChild(spPr, 'a:solidFill'), ctx.theme)
    if (solid) return solid
    const ln = directChild(spPr, 'a:ln')
    return ln ? resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? undefined : undefined
  }
  const seriesColors = sers.map(seriesColorOf)
  // 仅当全部系列都有显式色时写入，避免与配色方案色板错位
  if (seriesColors.length && seriesColors.every((c): c is string => Boolean(c))) chart.chartColors = seriesColors

  // 图表元素开关（模型字段全量必填）
  const legendEl = firstDescendant(root, 'c:legend')
  const legendDelete = legendEl ? directChild(legendEl, 'c:delete') : null
  const legendVisible = Boolean(legendEl) && attr(legendDelete, 'val') !== '1'
  const axes = ['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx']
    .map((tag) => firstDescendant(root, tag))
    .filter((ax): ax is Element => Boolean(ax))
  // 笛卡尔图表缺省带坐标轴（源文件常省略轴声明）；饼图/圆环/雷达为极坐标不含轴
  const polar = ['pieChart', 'doughnutChart', 'radarChart'].includes(kindNode.localName)
  chart.elements = {
    legend: legendVisible,
    gridLine: Boolean(firstDescendant(root, 'c:majorGridlines')),
    axis: axes.length > 0 || !polar,
    trendline: Boolean(firstDescendant(kindNode, 'c:trendline')),
    chartTitle: Boolean(title),
    dataLabel: Boolean(firstDescendant(kindNode, 'c:dLbls')),
    axisTitle: axes.some((ax) => Boolean(firstDescendant(ax, 'c:title'))),
  }

  // 网格线颜色 / 全局字号（c:txPr defRPr）
  const gridLines = firstDescendant(root, 'c:majorGridlines')
  const gridLn = gridLines ? directChild(directChild(gridLines, 'c:spPr') ?? gridLines, 'a:ln') : null
  const gridColor = gridLn ? resolveColor(directChild(gridLn, 'a:solidFill'), ctx.theme) : null
  if (gridColor) chart.gridColor = gridColor
  const txPr = firstDescendant(root, 'c:txPr')
  const defRPr = txPr ? firstDescendant(txPr, 'a:defRPr') : null
  const szRaw = attr(defRPr, 'sz')
  if (szRaw) chart.fontSize = Math.round(parseInt(szRaw, 10) / 100 / 0.75)
  return chart
}
