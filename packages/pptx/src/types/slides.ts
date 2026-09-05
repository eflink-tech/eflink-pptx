// 数据模型：演示文稿 / 幻灯片 / 元素（判别联合）
// 画布坐标系：px，默认 1280x720（16:9），导出 PPTX 时按 96px = 1inch 换算

/** 画布宽度（px），16:9 标准 */
export const SLIDE_WIDTH = 1280
/** 画布高度（px） */
export const SLIDE_HEIGHT = 720
/** px -> inch（pptxgenjs 导出用） */
export const PX_PER_INCH = 96

export type ElementTypes =
  | 'text'
  | 'image'
  | 'shape'
  | 'line'
  | 'table'
  | 'chart'
  | 'video'
  | 'audio'
  | 'formula'

/* ---------- 元素公共属性 ---------- */

export interface ElementBase {
  id: string
  type: ElementTypes
  x: number
  y: number
  w: number
  h: number
  /** 旋转角度（deg，顺时针） */
  rotate?: number
  /** 透明度 0-1 */
  opacity?: number
  /** 锁定后不可拖拽/编辑 */
  lock?: boolean
  /** 组合 id（同一 groupId 的元素一起选中/移动） */
  groupId?: string
  /** 超链接 */
  link?: ElementLink
  /** 元素动画 */
  anim?: ElementAnim
  /** 批注内容（单层） */
  comment?: string
  /** 元素名称（大纲/无障碍） */
  name?: string
}

export interface ElementLink {
  type: 'web' | 'slide'
  /** web: URL；slide: 幻灯片 id */
  target: string
}

export type AnimType = 'in' | 'emphasis' | 'out'
export type AnimTrigger = 'click' | 'withPrevious' | 'afterPrevious'

export interface ElementAnim {
  type: AnimType
  /** 效果关键字，如 fade / flyInLeft / zoomIn / pulse 等 */
  effect: string
  duration: number
  delay: number
  trigger: AnimTrigger
}

/* ---------- 背景 ---------- */

export interface Background {
  type: 'solid' | 'gradient' | 'image'
  color?: string
  gradient?: Gradient
  image?: BackgroundImage
}

export interface Gradient {
  type: 'linear' | 'radial'
  colors: GradientColor[]
  rotate?: number
}

export interface GradientColor {
  pos: number
  color: string
}

export interface BackgroundImage {
  src: string
  size: 'cover' | 'contain' | 'repeat'
}

/* ---------- 文本元素 ---------- */

export interface TextElement extends ElementBase {
  type: 'text'
  /** 富文本 HTML（TipTap 产出，含内联样式） */
  content: string
  defaultFontName?: string
  defaultColor?: string
  /** 行高倍数 */
  lineHeight?: number
  /** 字间距（px） */
  charSpace?: number
  /** 文字描边（px） */
  textStroke?: { color: string; width: number }
  /** 文字阴影 */
  shadow?: ShadowEffect
  /** 竖排 */
  vertical?: boolean
  /** 自动收缩字号以适应容器 */
  autoSize?: boolean
  /** 内边距（px） */
  padding?: number
  outline?: { color: string; width: number; style: string }
  fill?: string
}

export interface ShadowEffect {
  h: number
  v: number
  blur: number
  color: string
}

/* ---------- 图片元素 ---------- */

export interface ImageElement extends ElementBase {
  type: 'image'
  src: string
  /** 裁剪（相对原始图片的比例矩形） */
  clip?: { x: number; y: number; w: number; h: number }
  /** 裁剪容器形状 key（shape 库） */
  clipShape?: string
  radius?: number
  flipH?: boolean
  flipV?: boolean
  shadow?: ShadowEffect
  outline?: { color: string; width: number; style: string }
}

/* ---------- 形状元素 ---------- */

export interface ShapeElement extends ElementBase {
  type: 'shape'
  /** 预设形状 key（见 core/render/shape.ts SHAPE_PATHS） */
  shapeKey: string
  fill?: string | Gradient
  pattern?: string
  outline?: { color: string; width: number; style: string }
  flipH?: boolean
  flipV?: boolean
  shadow?: ShadowEffect
  /** 形状内嵌文本 */
  text?: string
  defaultColor?: string
  defaultFontName?: string
  fontSize?: number
  fontWeight?: string
  align?: 'left' | 'center' | 'right'
  valign?: 'top' | 'middle' | 'bottom'
  lineHeight?: number
}

/* ---------- 线条元素 ---------- */

export type LineType = 'straight' | 'broken' | 'curve'
export type LineStyle = 'solid' | 'dashed' | 'dotted'
export type LineArrow = '' | 'arrow' | 'triangle' | 'dot' | 'none'

/** 线条端点附着：0=上 1=右 2=下 3=左（元素 bbox 四边中点） */
export type LineAnchorIndex = 0 | 1 | 2 | 3

export interface LineAttach {
  elementId: string
  anchor: LineAnchorIndex
}

export interface LineElement extends ElementBase {
  type: 'line'
  /** 相对元素框的起止点（px） */
  start: [number, number]
  end: [number, number]
  /** 折线/曲线的中间控制点（相对元素框 px） */
  points?: Array<[number, number]>
  /** 起点附着到图形锚点 */
  startAttach?: LineAttach
  /** 终点附着到图形锚点 */
  endAttach?: LineAttach
  lineType: LineType
  color: string
  lineWidth: number
  lineStyle: LineStyle
  startArrow: LineArrow
  endArrow: LineArrow
}

/* ---------- 表格元素 ---------- */

export interface TableCell {
  text: string
  colspan?: number
  rowspan?: number
  style?: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    color?: string
    backcolor?: string
    fontsize?: number
    fontface?: string
    align?: 'left' | 'center' | 'right'
    valign?: 'top' | 'middle' | 'bottom'
  }
}

export interface TableElement extends ElementBase {
  type: 'table'
  /** 行高比例（总高 = h） */
  rowSizes: number[]
  /** 列宽比例（总宽 = w） */
  colSizes: number[]
  /** 数据矩阵（合并单元格以 colspan/rowspan 表达，被合并处填 null 不存储，渲染时跳过） */
  cells: Array<Array<TableCell | null>>
  outline?: { color: string; width: number; style: string }
  theme?: TableTheme
  autoSize?: boolean
}

export interface TableTheme {
  color: string[]
  /** 表头样式开关 */
  headColor: string
}

/* ---------- 图表元素 ---------- */

/** 图表类型 id（与 eflink-excel 的 CHART_GROUPS 注册表一致，14 种 / 7 组） */
export type ChartType =
  | 'bar-cluster'
  | 'bar-stack'
  | 'bar-percent'
  | 'line'
  | 'line-stack'
  | 'line-marker'
  | 'pie'
  | 'pie-doughnut'
  | 'bar-horizontal'
  | 'bar-horizontal-stack'
  | 'area'
  | 'area-stack'
  | 'scatter'
  | 'radar'

export interface ChartData {
  labels: string[]
  series: Array<{
    name: string
    values: number[]
  }>
}

export interface ChartElement extends ElementBase {
  type: 'chart'
  chartType: ChartType
  data: ChartData
  /** 系列配色（可空则用配色方案色板） */
  chartColors?: string[]
  /** 配色方案索引（CHART_PALETTES，chartColors 未设置时生效） */
  paletteIndex?: number
  /** 图表元素开关（图例/网格线/坐标轴/趋势线/标题/数据标签/轴标题） */
  elements?: {
    legend: boolean
    gridLine: boolean
    axis: boolean
    trendline: boolean
    chartTitle: boolean
    dataLabel: boolean
    axisTitle: boolean
  }
  /** 柱状/条状圆角 */
  barBorderRadius?: number
  /** @deprecated 旧字段，已迁移至 elements.legend */
  showLegend?: boolean
  /** @deprecated 旧字段，已迁移至 elements.dataLabel */
  showLabel?: boolean
  title?: string
  fontSize?: number
  gridColor?: string
}

/* ---------- 音视频元素 ---------- */

export interface VideoElement extends ElementBase {
  type: 'video'
  src: string
  poster?: string
  loop: boolean
  autoPlay: boolean
}

export interface AudioElement extends ElementBase {
  type: 'audio'
  src: string
  loop: boolean
  autoPlay: boolean
  /** 可自定义图标色 */
  color?: string
}

/* ---------- 公式元素 ---------- */

export interface FormulaElement extends ElementBase {
  type: 'formula'
  latex: string
  /** 字体色 */
  color?: string
  fontSize?: number
}

/* ---------- 元素联合 ---------- */

export type PPTElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | LineElement
  | TableElement
  | ChartElement
  | VideoElement
  | AudioElement
  | FormulaElement

/* ---------- 幻灯片 / 演示文稿 ---------- */

export type TransitionPreset =
  | 'none'
  | 'fade'
  | 'slideLeft'
  | 'slideRight'
  | 'slideUp'
  | 'slideDown'
  | 'wipeLeft'
  | 'wipeRight'
  | 'zoomIn'
  | 'zoomOut'
  | 'blinds'
  | 'flip'

export interface SlideTransition {
  preset: TransitionPreset
  duration: number
  easing: string
}

export interface Slide {
  id: string
  elements: PPTElement[]
  background?: Background
  transition?: SlideTransition
  /** 演讲者备注 */
  note?: string
  /** 分节名 */
  section?: string
}

export interface Theme {
  /** 主题色板（图表/默认取色用） */
  colors: string[]
  /** 默认背景 */
  background: Background
  fontName: string
  fontColor: string
}

export interface Presentation {
  slides: Slide[]
  theme: Theme
  /** 画布宽度 px（高度 = width / viewportRatio） */
  width: number
  viewportRatio: number
}

/* ---------- 创建空演示文稿 ---------- */

export function createSlide(id: string): Slide {
  return { id, elements: [] }
}

export function createDefaultTheme(): Theme {
  return {
    colors: [
      '#d14424', '#e6935c', '#f5c56b', '#7cb342',
      '#42a5f5', '#5e6db8', '#8e6bbf', '#455a64',
      '#f2f2f2', '#666666',
    ],
    background: { type: 'solid', color: '#ffffff' },
    fontName: 'Microsoft YaHei',
    fontColor: '#333333',
  }
}

export function createPresentation(id: string): Presentation {
  return {
    slides: [createSlide(id)],
    theme: createDefaultTheme(),
    width: SLIDE_WIDTH,
    viewportRatio: SLIDE_WIDTH / SLIDE_HEIGHT,
  }
}
