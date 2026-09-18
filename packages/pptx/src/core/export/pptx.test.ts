import { describe, expect, it, vi } from 'vitest'
import { chartNativeSpec, normColor, parseRunsFromHTML, pxToInch as IN, pxToPt as PT, exportShape, withImageHeader, exportPPTX, fontFaceOf, rotatedBoundsOf, expandThemeFontPlaceholders, injectThemeEaFonts } from './pptx'
import type { ChartType, Presentation, ShapeElement, Slide } from '../../types/slides'
import { createDefaultTheme } from '../../types/slides'

// 离屏渲染依赖浏览器 canvas，测试中 mock 掉，仅验证分支走向
vi.mock('./image', () => ({
  renderSlideToBlob: vi.fn(async () => new Blob(['fake-png'])),
}))

// pptxgenjs 与下载在单测中 mock 掉，仅验证导出主流程的回调行为
vi.mock('pptxgenjs', () => ({
  default: class {
    slides: unknown[] = []
    defineLayout = vi.fn()
    addSlide = () => {
      const s = {
        addText: vi.fn(), addShape: vi.fn(), addTable: vi.fn(), addChart: vi.fn(),
        addImage: vi.fn(), addMedia: vi.fn(), addNotes: vi.fn(), background: undefined,
      }
      this.slides.push(s)
      return s
    }
    write = vi.fn(async () => new Blob(['x']))
  },
}))
vi.mock('./json', () => ({ downloadBlob: vi.fn() }))

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

  it('解析内联 font-weight/font-style/text-decoration（导入产物用内联样式而非标签）', () => {
    const result = parseRunsFromHTML(
      '<p><span style="font-weight:bold">粗</span><span style="font-style:italic">斜</span>' +
      '<span style="text-decoration:underline">下</span><span style="text-decoration:line-through">删</span></p>',
    )
    const runs = result[0].runs
    expect(runs[0].options.bold).toBe(true)
    expect(runs[1].options.italic).toBe(true)
    expect(runs[2].options.underline).toBe(true)
    expect(runs[3].options.strike).toBe(true)
  })

  it('中文 run 标记 lang=zh-CN（渲染器按语言选 CJK 字体回退，源文件中文 run 均为 zh-CN）', () => {
    const result = parseRunsFromHTML(
      '<p><span style="font-family:\'微软雅黑\', \'Microsoft YaHei\', sans-serif">中文</span>' +
      '<span style="font-family:\'微软雅黑\', \'Microsoft YaHei\', sans-serif">ABC</span></p>',
    )
    expect(result[0].runs[0].options.lang).toBe('zh-CN')
    expect(result[0].runs[1].options.lang).toBeUndefined()
  })

  it('run 字体与 theme 字体一致 → +mn 引用占位（LibreOffice 实证：显式写主题字体名 CJK 回退错误）', () => {
    const themeFonts = { major: 'Arial', majorEa: '微软雅黑', minor: 'Arial', minorEa: '微软雅黑' }
    // 导入产物：源 run latin=+mn-lt ea=+mn-ea → 栈 [Arial, 微软雅黑]；中文 run 取 ea 位
    const stack = `'Arial', '微软雅黑', 'PingFang SC', sans-serif`
    const result = parseRunsFromHTML(`<p><span style="font-family:${stack}">中文</span><span style="font-family:${stack}">ABC</span></p>`, themeFonts)
    expect(result[0].runs[0].options.fontFace).toBe('+mn-lt')
    expect(result[0].runs[1].options.fontFace).toBe('+mn-lt')
    // 与 theme 不一致的字体保持显式名（用户改字体的 run，PowerPoint 正常处理）
    const custom = parseRunsFromHTML(`<p><span style="font-family:'阿里巴巴普惠体', sans-serif">中文</span></p>`, themeFonts)
    expect(custom[0].runs[0].options.fontFace).toBe('阿里巴巴普惠体')
    // 无 theme 字体信息（新建演示）不产生占位
    const plain = parseRunsFromHTML(`<p><span style="font-family:${stack}">中文</span></p>`)
    expect(plain[0].runs[0].options.fontFace).toBe('微软雅黑')
  })

  it('段级 line-height → 行距导出选项（倍数 / px→pt）', () => {
    // 倍数（spcPct）：0.8 直接透传
    const multiple = parseRunsFromHTML('<p style="line-height:0.8">标</p>')
    expect(multiple[0].runs[0].options.lineSpacingMultiple).toBe(0.8)
    // px（spcPts）：27px → 20.25pt
    const points = parseRunsFromHTML('<p style="line-height:27px">行</p>')
    expect(points[0].runs[0].options.lineSpacing).toBeCloseTo(20.25)
    // 段内多个 run 均携带
    const multi = parseRunsFromHTML('<p style="line-height:1.5"><span style="color:#FF0000">红</span>黑</p>')
    expect(multi[0].runs[0].options.lineSpacingMultiple).toBe(1.5)
    expect(multi[0].runs[1].options.lineSpacingMultiple).toBe(1.5)
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

  it('段落对齐注入段内所有 run（含带格式 run；pptxgenjs 从 run 选项读段落属性，且 align 变化会拆段）', () => {
    const result = parseRunsFromHTML(
      '<p style="text-align:center"><span style="color:#FF0000">红</span>黑</p>' +
      '<p>左对齐段</p>',
    )
    expect(result[0].runs[0].options.align).toBe('center')
    expect(result[0].runs[1].options.align).toBe('center')
    // 未声明对齐的段落不注入
    expect(result[1].runs[0].options.align).toBeUndefined()
  })

  it('内联 font-family → fontFace：中文 run 取 ea 位，纯西文取 latin 位', () => {
    // 模拟导入产物的栈序：latin（微软雅黑）→ ea（PingFang SC）→ 系统回退
    const html = `<p><span style="font-family:'微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif">中文</span><span style="font-family:'微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif">ABC</span></p>`
    const result = parseRunsFromHTML(html)
    expect(result[0].runs[0].options.fontFace).toBe('PingFang SC')
    expect(result[0].runs[1].options.fontFace).toBe('微软雅黑')
    // fontFamilies 是中间产物，不应泄漏进导出选项
    expect(result[0].runs[0].options.fontFamilies).toBeUndefined()
  })

  it('空内容兜底', () => {
    const result = parseRunsFromHTML('')
    expect(result).toHaveLength(1)
  })
})

describe('fontFaceOf 字体栈 → fontFace', () => {
  it('取首个非通用族名并去引号（PowerPoint 不识别 fallback 栈）', () => {
    expect(fontFaceOf(`'Microsoft YaHei', 'PingFang SC', sans-serif`)).toBe('Microsoft YaHei')
    expect(fontFaceOf('"DengXian", serif')).toBe('DengXian')
    expect(fontFaceOf('等线')).toBe('等线')
  })

  it('纯通用族 / 空值返回 undefined', () => {
    expect(fontFaceOf('sans-serif, serif')).toBeUndefined()
    expect(fontFaceOf(undefined)).toBeUndefined()
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

describe('exportPPTX 导出进度回调', () => {
  it('每页导出后回调 (done, total)（对话框据此显示导出中 (n/m)）', async () => {
    const slide: Slide = { id: 's1', elements: [] }
    const pres: Presentation = {
      slides: [slide, { ...slide, id: 's2' }],
      theme: createDefaultTheme(), width: 1280, viewportRatio: 16 / 9,
    }
    const onProgress = vi.fn()
    await exportPPTX(pres, '测试', onProgress)
    expect(onProgress).toHaveBeenCalledTimes(2)
    expect(onProgress).toHaveBeenNthCalledWith(1, 1, 2)
    expect(onProgress).toHaveBeenNthCalledWith(2, 2, 2)
  })
})

describe('rotatedBoundsOf 元素旋转外接框', () => {
  it('无旋转 → 原框', () => {
    expect(rotatedBoundsOf({ x: 100, y: 50, w: 200, h: 100, rotate: 0 }))
      .toEqual({ x: 100, y: 50, w: 200, h: 100 })
    expect(rotatedBoundsOf({ x: 100, y: 50, w: 200, h: 100 }))
      .toEqual({ x: 100, y: 50, w: 200, h: 100 })
  })

  it('90° 旋转 → 宽高互换、中心不变', () => {
    const b = rotatedBoundsOf({ x: 100, y: 50, w: 200, h: 100, rotate: 90 })
    expect(b.w).toBeCloseTo(100)
    expect(b.h).toBeCloseTo(200)
    expect(b.x).toBeCloseTo(100 + (200 - 100) / 2) // 中心 x=200 不变
    expect(b.y).toBeCloseTo(50 + (100 - 200) / 2)
  })

  it('45° 旋转 → 外接框扩大', () => {
    const b = rotatedBoundsOf({ x: 0, y: 0, w: 100, h: 100, rotate: 45 })
    expect(b.w).toBeCloseTo(100 * Math.SQRT2)
    expect(b.x).toBeCloseTo(-((100 * Math.SQRT2 - 100) / 2))
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
    // 兜底图必须透明背景渲染：白底图会盖住下层相邻文字（目录页角框遮挡「01.」事故）
    const { renderSlideToBlob } = await import('./image')
    expect(renderSlideToBlob).toHaveBeenCalledWith(
      expect.objectContaining({ elements: [el] }), pres, 'png', { transparent: true },
    )
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

describe('theme 字体后处理', () => {
  it('expandThemeFontPlaceholders：+mn/+mj 占位展开为 latin/ea/cs 引用（源文件 run 形式）', () => {
    const runXml = '<a:rPr lang="zh-CN" sz="9600" b="1" dirty="0">'
      + '<a:latin typeface="+mn-lt" pitchFamily="34" charset="0"/>'
      + '<a:ea typeface="+mn-lt" pitchFamily="34" charset="-122"/>'
      + '<a:cs typeface="+mn-lt" pitchFamily="34" charset="-120"/></a:rPr>'
    const out = expandThemeFontPlaceholders(runXml)
    expect(out).toContain('<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-ea"/>')
    expect(out).not.toContain('pitchFamily')
    // major 占位同理展开为 +mj 形式
    const mj = expandThemeFontPlaceholders('<a:latin typeface="+mj-lt" pitchFamily="34" charset="0"/><a:ea typeface="+mj-lt" pitchFamily="34" charset="-122"/><a:cs typeface="+mj-lt" pitchFamily="34" charset="-120"/>')
    expect(mj).toContain('<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-ea"/>')
    // 非占位字体不动
    expect(expandThemeFontPlaceholders('<a:latin typeface="Arial" pitchFamily="34" charset="0"/>')).toContain('typeface="Arial"')
  })

  it('injectThemeEaFonts：theme1.xml 注入 a:ea（pptxgenjs 只写 latin 位）', () => {
    // pptxgenjs 在 latin 后自带空 ea：注入必须替换而非追加（重复 a:ea 为脏 XML）
    const themeXml = '<a:fontScheme><a:majorFont><a:latin typeface="Arial Light"/><a:ea typeface=""/></a:majorFont>'
      + '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/></a:minorFont></a:fontScheme>'
    const out = injectThemeEaFonts(themeXml, { major: 'Arial Light', majorEa: '微软雅黑', minor: 'Arial', minorEa: '微软雅黑' })
    expect(out).toContain('<a:majorFont><a:latin typeface="Arial Light"/><a:ea typeface="微软雅黑"/></a:majorFont>')
    expect(out).toContain('<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/></a:minorFont>')
    expect(out.match(/<a:ea /g)).toHaveLength(2)
    // 无空 ea 的 theme 直接插入
    const plain = '<a:fontScheme><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>'
    expect(injectThemeEaFonts(plain, { major: 'Arial', minor: 'Arial', minorEa: '微软雅黑' }))
      .toContain('<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/></a:minorFont>')
    // ea 缺失时不注入
    const noEa = injectThemeEaFonts(themeXml, { major: 'Arial Light', minor: 'Arial' })
    expect(noEa).toBe(themeXml)
  })
})
