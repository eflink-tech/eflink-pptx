// src/core/preview/chart.ts
/** p:graphicFrame(c:chart) → 图表 SVG：复用一期 parseChartEl（scale 恒 {1,1}，返回坐标即源画布 px），
 * 按 chartType 直渲 SVG 图元；radar 降级占位框、bar-percent 按堆叠近似（已知取舍）。
 * 绘制坐标统一以 parseChartEl 返回的 x/y/w/h 为准（与一期导入模型同源，避免双轨几何分叉）。 */
import { parseChartEl } from '../import/elements/chart'
import { IDENTITY_XFORM } from '../import/context'
import { svgEl, type PreviewCtx } from './svg'
import type { ChartElement } from '../../types/slides'

/** Office 默认图表主题色（系列循环取色） */
const CHART_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

type BoxLike = { x: number; y: number; w: number; h: number }

export async function renderChart(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const chart = await parseChartEl(node, ctx, IDENTITY_XFORM, ctx.pkg)
  if (!chart) return null
  const box: BoxLike = { x: chart.x, y: chart.y, w: chart.w, h: chart.h }
  const g = svgEl('g')
  if (chart.chartType === 'radar') {
    // 雷达图结构复杂（放射网格 + 面积叠加），预览降级为占位框
    g.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h, fill: '#F2F2F2', stroke: '#BFBFBF', 'stroke-width': 1 }))
    const t = svgEl('text', { x: box.x + box.w / 2, y: box.y + box.h / 2, 'text-anchor': 'middle', 'font-size': 12, fill: '#999999' })
    t.textContent = '雷达图暂不支持预览'
    g.appendChild(t)
    return g
  }
  drawAxis(g, box)
  const t = chart.chartType
  if (t === 'pie' || t === 'pie-doughnut') drawPie(g, chart, box, t === 'pie-doughnut')
  else if (t.startsWith('bar')) drawBars(g, chart, box, t.includes('horizontal'), t.includes('stack') || t === 'bar-percent')
  else if (t.startsWith('area')) drawArea(g, chart, box, t === 'area-stack')
  else if (t === 'scatter') drawScatter(g, chart, box)
  else drawLine(g, chart, box, t === 'line-stack', t === 'line-marker')
  if (chart.title) {
    const title = svgEl('text', { x: box.x + box.w / 2, y: box.y + 16, 'text-anchor': 'middle', 'font-size': 14, fill: '#333333' })
    title.textContent = chart.title
    g.appendChild(title)
  }
  return g
}

/** 绘图区：留出轴标签与标题空间 */
function plotRect(box: BoxLike): BoxLike {
  return { x: box.x + 36, y: box.y + 24, w: box.w - 44, h: box.h - 48 }
}

/** 值轴上限：最大值 ×1.1 留白；全 0 时兜底为 1 防除零 */
function valueMax(chart: ChartElement): number {
  return Math.max(0, ...chart.data.series.flatMap((s) => s.values)) * 1.1 || 1
}

function drawAxis(g: SVGElement, box: BoxLike): void {
  const p = plotRect(box)
  g.appendChild(svgEl('line', { x1: p.x, y1: p.y, x2: p.x, y2: p.y + p.h, stroke: '#999999', 'stroke-width': 1 }))
  g.appendChild(svgEl('line', { x1: p.x, y1: p.y + p.h, x2: p.x + p.w, y2: p.y + p.h, stroke: '#999999', 'stroke-width': 1 }))
}

/** 柱状图：类目分带（band），非堆叠按系列均分带宽；horizontal 时换轴；stacked/bar-percent 按累计高度叠放 */
function drawBars(g: SVGElement, chart: ChartElement, box: BoxLike, horizontal: boolean, stacked: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  const band = (horizontal ? p.h : p.w) / n
  const groupSize = stacked ? 1 : series.length
  for (let i = 0; i < n; i += 1) {
    let acc = 0
    for (let si = 0; si < series.length; si += 1) {
      const v = series[si].values[i] ?? 0
      const color = CHART_COLORS[si % CHART_COLORS.length]
      if (horizontal) {
        const seg = (v / maxV) * p.w
        const y = stacked
          ? p.y + band * i + (band * acc) / maxV
          : p.y + band * i + (band / groupSize) * si + (band / groupSize) * 0.1
        const h = stacked ? (band * v) / maxV : (band / groupSize) * 0.8
        g.appendChild(svgEl('rect', { x: p.x, y, width: seg, height: h, fill: color }))
      } else {
        const w = stacked ? band : (band / groupSize) * 0.8
        const x = stacked ? p.x + band * i : p.x + band * i + (band / groupSize) * si + (band / groupSize) * 0.1
        const h = (p.h * v) / maxV
        const y = stacked ? p.y + p.h - (p.h * (acc + v)) / maxV : p.y + p.h - h
        g.appendChild(svgEl('rect', { x, y, width: w, height: h, fill: color }))
      }
      acc += v
    }
  }
}

/** 折线图：polyline 逐系列连线；stacked 按低层系列累计取值；marker 时补数据点圆 */
function drawLine(g: SVGElement, chart: ChartElement, box: BoxLike, stacked: boolean, marker: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  for (let si = 0; si < series.length; si += 1) {
    const color = CHART_COLORS[si % CHART_COLORS.length]
    const pts: string[] = []
    for (let i = 0; i < n; i += 1) {
      const v = (series[si].values[i] ?? 0)
        + (stacked ? series.slice(0, si).reduce((s, sr) => s + (sr.values[i] ?? 0), 0) : 0)
      pts.push(`${p.x + (p.w * i) / Math.max(1, n - 1)},${p.y + p.h - (v / maxV) * p.h}`)
    }
    g.appendChild(svgEl('polyline', { points: pts.join(' '), fill: 'none', stroke: color, 'stroke-width': 2 }))
    if (marker) {
      for (const pt of pts) {
        const [x, y] = pt.split(',')
        g.appendChild(svgEl('circle', { cx: x, cy: y, r: 3, fill: color }))
      }
    }
  }
}

/** 面积图：闭合 path；stacked 时按系列自下而上累计，绘制顺序反转保证下层可见 */
function drawArea(g: SVGElement, chart: ChartElement, box: BoxLike, stacked: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  const cum = new Array<number>(n).fill(0)
  const rows = series.map((sr) => sr.values.map((v, i) => (stacked ? (cum[i] += v ?? 0) : v ?? 0)))
  const order = stacked ? [...rows.keys()].reverse() : [...rows.keys()]
  for (const si of order) {
    const pts = rows[si].map((v, i) => `${p.x + (p.w * i) / Math.max(1, n - 1)},${p.y + p.h - (v / maxV) * p.h}`)
    if (!pts.length) continue
    const d = `M ${pts[0]} L ${pts.join(' L ')} L ${p.x + p.w},${p.y + p.h} L ${p.x},${p.y + p.h} Z`
    g.appendChild(svgEl('path', { d, fill: CHART_COLORS[si % CHART_COLORS.length], 'fill-opacity': stacked ? 1 : 0.5 }))
  }
}

/** 散点图：labels 为 xVal 数值串，第一系列 values 为 y；按数据范围线性映射到绘图区 */
function drawScatter(g: SVGElement, chart: ChartElement, box: BoxLike): void {
  const p = plotRect(box)
  const xs = chart.data.labels.map(Number).filter(Number.isFinite)
  const ys = chart.data.series.flatMap((s) => s.values).filter(Number.isFinite)
  const minX = Math.min(0, ...xs)
  const maxX = Math.max(1, ...xs)
  const minY = Math.min(0, ...ys)
  const maxY = Math.max(1, ...ys) * 1.1
  for (let si = 0; si < chart.data.series.length; si += 1) {
    const color = CHART_COLORS[si % CHART_COLORS.length]
    const vals = chart.data.series[si].values
    xs.forEach((x, i) => {
      const y = vals[i]
      if (!Number.isFinite(y)) return
      g.appendChild(svgEl('circle', {
        cx: p.x + ((x - minX) / (maxX - minX || 1)) * p.w,
        cy: p.y + p.h - ((y - minY) / (maxY - minY || 1)) * p.h,
        r: 4, fill: color,
      }))
    })
  }
}

/** 饼图/圆环图：取第一系列，12 点钟方向起顺时针分扇 */
function drawPie(g: SVGElement, chart: ChartElement, box: BoxLike, doughnut: boolean): void {
  const series = chart.data.series[0]
  if (!series) return
  const total = series.values.reduce((a, b) => a + b, 0) || 1
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const r = Math.min(box.w, box.h) * 0.4
  const rIn = doughnut ? r * 0.6 : 0
  let a0 = -Math.PI / 2
  for (let i = 0; i < series.values.length; i += 1) {
    const a1 = a0 + ((series.values[i] ?? 0) / total) * Math.PI * 2
    g.appendChild(svgEl('path', { d: slicePath(cx, cy, r, rIn, a0, a1), fill: CHART_COLORS[i % CHART_COLORS.length] }))
    a0 = a1
  }
}

/** 扇环路径：外弧顺时针 + 内弧逆时针（rIn=0 时为扇形） */
function slicePath(cx: number, cy: number, r: number, rIn: number, a0: number, a1: number): string {
  const large = a1 - a0 > Math.PI ? 1 : 0
  const x0 = cx + r * Math.cos(a0)
  const y0 = cy + r * Math.sin(a0)
  const x1 = cx + r * Math.cos(a1)
  const y1 = cy + r * Math.sin(a1)
  if (rIn <= 0) return `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`
  const ix0 = cx + rIn * Math.cos(a1)
  const iy0 = cy + rIn * Math.sin(a1)
  const ix1 = cx + rIn * Math.cos(a0)
  const iy1 = cy + rIn * Math.sin(a0)
  return `M ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} L ${ix0} ${iy0} A ${rIn} ${rIn} 0 ${large} 0 ${ix1} ${iy1} Z`
}
