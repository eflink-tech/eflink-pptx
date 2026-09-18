import { describe, expect, it, vi } from 'vitest'
import { chartNativeSpec, normColor, parseRunsFromHTML, pxToInch as IN, pxToPt as PT, exportShape, withImageHeader } from './pptx'
import type { ChartType, Presentation, ShapeElement, Slide } from '../../types/slides'
import { createDefaultTheme } from '../../types/slides'

// 离屏渲染依赖浏览器 canvas，测试中 mock 掉，仅验证分支走向
vi.mock('./image', () => ({
  renderSlideToBlob: vi.fn(async () => new Blob(['fake-png'])),
}))

describe('chartNativeSpec 图表导出映射', () => {
  it('柱状/条形图方向与堆积方式', () => {
    expect(chartNativeSpec('bar-cluster')).toMatchObject({ type: 'bar', barDir: 'col' })
    expect(chartNativeSpec('bar-stack')).toMatchObject({ type: 'bar', barDir: 'col', barGrouping: 'stacked' })
    expect(chartNativeSpec('bar-percent')).toMatchObject({ type: 'bar', barDir: 'col', barGrouping: 'percentStacked' })
    expect(chartNativeSpec('bar-horizontal')).toMatchObject({ type: 'bar', barDir: 'bar' })
    expect(chartNativeSpec('bar-horizontal-stack')).toMatchObject({ type: 'bar', barDir: 'bar', barGrouping: 'stacked' })
  })

  it('折线/面积/饼图变体', () => {
    expect(chartNativeSpec('line')).toMatchObject({ type: 'line' })
    expect(chartNativeSpec('line-stack')).toMatchObject({ type: 'line', barGrouping: 'stacked' })
    expect(chartNativeSpec('line-marker')).toMatchObject({ type: 'line', lineDataSymbol: 'circle' })
    expect(chartNativeSpec('area-stack')).toMatchObject({ type: 'area', barGrouping: 'stacked' })
    expect(chartNativeSpec('pie')).toMatchObject({ type: 'pie' })
    expect(chartNativeSpec('pie-doughnut')).toMatchObject({ type: 'doughnut' })
  })

  it('散点用无连线折线近似，雷达图返回 null 走图片导出', () => {
    expect(chartNativeSpec('scatter')).toMatchObject({ type: 'line', lineSize: 0 })
    expect(chartNativeSpec('radar')).toBeNull()
  })

  it('全部 14 种类型均有映射定义', () => {
    const allTypes: ChartType[] = [
      'bar-cluster', 'bar-stack', 'bar-percent', 'line', 'line-stack', 'line-marker',
      'pie', 'pie-doughnut', 'bar-horizontal', 'bar-horizontal-stack',
      'area', 'area-stack', 'scatter', 'radar',
    ]
    for (const t of allTypes) expect(chartNativeSpec(t)).toBeDefined()
  })
})

describe('normColor', () => {
  it('hex6 转大写无 #', () => {
    expect(normColor('#d14424')).toBe('D14424')
  })
  it('hex3 扩展', () => {
    expect(normColor('#f00')).toBe('FF0000')
  })
  it('rgba 混合白底', () => {
    // 50% 红混合白 → (255,128,128)
    expect(normColor('rgba(255,0,0,0.5)')).toBe('FF8080')
  })
  it('hex8 透明度混合', () => {
    expect(normColor('#00000000')).toBeUndefined()
    expect(normColor('#ff000080')).toBe('FF7F7F')
  })
  it('transparent 返回 undefined', () => {
    expect(normColor('transparent')).toBeUndefined()
  })
})

describe('尺寸换算', () => {
  it('px → inch', () => {
    expect(IN(96)).toBeCloseTo(1)
    expect(IN(1280)).toBeCloseTo(13.3333, 3)
  })
  it('px → pt', () => {
    expect(PT(16)).toBe(12)
  })
})

describe('parseRunsFromHTML', () => {
  it('解析内联样式', () => {
    const result = parseRunsFromHTML('<p><span style="color:#d14424;font-size:32px">标题</span></p>')
    expect(result).toHaveLength(1)
    expect(result[0].runs[0].text).toBe('标题')
    expect(result[0].runs[0].options.color).toBe('D14424')
    expect(result[0].runs[0].options.fontSize).toBe(24)
  })

  it('解析加粗斜体下划线', () => {
    const result = parseRunsFromHTML('<p><strong>粗</strong><em>斜</em><u>下</u></p>')
    const runs = result[0].runs
    expect(runs[0].options.bold).toBe(true)
    expect(runs[1].options.italic).toBe(true)
    expect(runs[2].options.underline).toBe(true)
  })

  it('多段落生成 breakLine', () => {
    const result = parseRunsFromHTML('<p>一</p><p>二</p>')
    expect(result).toHaveLength(2)
    expect(result[0].runs[0].options.breakLine).toBe(true)
  })

  it('段落对齐', () => {
    const result = parseRunsFromHTML('<p style="text-align:center">居中</p>')
    expect(result[0].align).toBe('center')
  })

  it('空内容兜底', () => {
    const result = parseRunsFromHTML('')
    expect(result).toHaveLength(1)
  })
})

describe('withImageHeader 图片 data 补 base64 头', () => {
  it('裸 base64 → 补 image/png;base64 头（pptxgenjs 校验要求）', () => {
    const bare = 'iVBORw0KGgoAAAANSUhEUg=='
    expect(withImageHeader(bare)).toBe(`image/png;base64,${bare}`)
  })

  it('完整 dataURL / 已带头 → 原样透传', () => {
    const url = 'data:image/jpeg;base64,/9j/4AAQ'
    expect(withImageHeader(url)).toBe(url)
    expect(withImageHeader('image/jpeg;base64,/9j/4AAQ')).toBe('image/jpeg;base64,/9j/4AAQ')
  })
})

describe('exportShape 自定义 path 分支', () => {
  const slide: Slide = { id: 's1', elements: [] }
  const pres: Presentation = {
    slides: [slide], theme: createDefaultTheme(), width: 1280, viewportRatio: 16 / 9,
  }
  // pptx 实例在 exportShape 中被显式 void，测试传空对象即可
  const dummyPptx = {} as Parameters<typeof exportShape>[0]

  it('path 形状不走原生分支，落图片兜底（避免 custGeom 静默退化矩形）', async () => {
    const el: ShapeElement = {
      id: 'a', type: 'shape', x: 0, y: 0, w: 100, h: 100,
      shapeKey: 'rect', path: 'M0,0 L100,100', fill: '#FF0000',
    }
    const spec = await exportShape(dummyPptx, slide, el, pres)
    expect(spec).toMatchObject({ type: 'image' })
  })

  it('普通预设形状仍走原生分支', async () => {
    const el: ShapeElement = {
      id: 'b', type: 'shape', x: 0, y: 0, w: 100, h: 100,
      shapeKey: 'rect', fill: '#FF0000',
    }
    const spec = await exportShape(dummyPptx, slide, el, pres)
    expect(spec).toMatchObject({ type: 'shape', native: 'rect' })
  })
})
