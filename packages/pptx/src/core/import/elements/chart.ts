// src/core/import/elements/chart.ts
/** p:graphicFrame(c:chart) → chart 元素（OOXML chart part 缓存数据 → ChartData） */
import { attr, directChild, directChildren, firstDescendant, parseXML } from '../xml'
import { genId } from '../../utils/id'
import { addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import { parseFrameGeom } from './frame-geom'
import type { ChartElement, ChartType } from '../../../types/slides'

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

  const KIND_NODES = ['barChart', 'lineChart', 'pieChart', 'doughnutChart', 'areaChart', 'scatterChart', 'radarChart']
  // 图表类型节点是 plotArea 的直接子级，按本地名探测（不依赖前缀绑定）
  const kindNode = Array.from(plotArea.children).find((c) => KIND_NODES.includes(c.localName)) ?? null
  if (!kindNode) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  const hints: ChartKindHints = {
    barDir: attr(directChild(kindNode, 'c:barDir'), 'val') ?? undefined,
    grouping: attr(directChild(kindNode, 'c:grouping'), 'val') ?? undefined,
    marker: attr(directChild(kindNode, 'c:marker'), 'val') === '1',
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
  const sers = directChildren(kindNode, 'c:ser')
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
  const titleEl = firstDescendant(root, 'c:title')
  const title = titleEl
    ? directChildren(titleEl, 'c:tx').map((tx) => firstDescendant(tx, 'a:t')?.textContent ?? '').join('').trim()
    : ''
  if (title) chart.title = title
  return chart
}
