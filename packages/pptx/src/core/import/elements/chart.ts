// src/core/import/elements/chart.ts
/** p:graphicFrame(c:chart) → chart 元素（OOXML chart part 缓存数据 → ChartData） */
import { attr, directChild, directChildren, firstDescendant, parseXML } from '../xml'
import { genId } from '../../utils/id'
import { mapX, mapY, addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
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

  // scatter 使用 c:xVal/c:yVal，其余用 c:cat/c:val
  const sers = directChildren(kindNode, 'c:ser')
  if (!sers.length) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  const isScatter = kindNode.localName === 'scatterChart'
  let labels: string[]
  if (isScatter) {
    labels = cacheValues(firstDescendant(sers[0], 'c:xVal'), 'c:numCache')
  } else {
    labels = cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:strCache')
    if (!labels.length) labels = cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:numCache')
  }
  const series = sers.map((ser) => ({
    name: cacheValues(firstDescendant(ser, 'c:tx'), 'c:strCache')[0] ?? '系列',
    values: isScatter
      ? cacheValues(firstDescendant(ser, 'c:yVal'), 'c:numCache').map(Number)
      : cacheValues(firstDescendant(ser, 'c:val'), 'c:numCache').map(Number),
  }))

  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)

  const chart: ChartElement = {
    id: genId('ch-'), type: 'chart', chartType,
    data: { labels, series },
    x: Math.round(mapX(xf, ex) / 9525 * ctx.scale.x),
    y: Math.round(mapY(xf, ey) / 9525 * ctx.scale.y),
    w: Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x)),
    h: Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y)),
    name: '图表',
  }
  const titleEl = firstDescendant(root, 'c:title')
  const title = titleEl
    ? directChildren(titleEl, 'c:tx').map((tx) => firstDescendant(tx, 'a:t')?.textContent ?? '').join('').trim()
    : ''
  if (title) chart.title = title
  return chart
}
