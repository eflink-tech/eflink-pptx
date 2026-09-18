import { describe, expect, it } from 'vitest'
import { fontStackOf } from './fonts'
import { DEFAULT_SCHEME, type PptxTheme } from './theme'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri Light',
  minorFont: 'Calibri',
  majorEaFont: '汉仪长文简',
  minorEaFont: '微软雅黑',
  colorMap: {},
}

describe('fontStackOf', () => {
  it('西文 + 东亚字体 → latin 在前、ea 次之、中文回退栈兜底', () => {
    expect(fontStackOf('Arial', '微软雅黑', theme)).toBe(
      `'Arial', '微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('仅 latin 时直接跟中文回退栈', () => {
    expect(fontStackOf('Calibri', null, theme)).toBe(
      `'Calibri', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('仅 ea 时省略 latin 位', () => {
    expect(fontStackOf(null, '宋体', theme)).toBe(
      `'宋体', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('主题引用 +mn-lt / +mj-ea 解析为主题字体', () => {
    expect(fontStackOf('+mn-lt', '+mj-ea', theme)).toBe(
      `'Calibri', '汉仪长文简', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('主题引用 +mj-lt / +mn-ea 对应 major/minor 各归其位', () => {
    expect(fontStackOf('+mj-lt', '+mn-ea', theme)).toBe(
      `'Calibri Light', '微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('重复字体去重（latin 与 ea 相同时只保留一次）', () => {
    expect(fontStackOf('微软雅黑', '微软雅黑', theme)).toBe(
      `'微软雅黑', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('字体名与回退栈重复时去重', () => {
    expect(fontStackOf('PingFang SC', null, theme)).toBe(
      `'PingFang SC', 'Microsoft YaHei', sans-serif`,
    )
  })

  it('无任何字体时仅返回回退栈', () => {
    expect(fontStackOf(null, null, theme)).toBe(`'PingFang SC', 'Microsoft YaHei', sans-serif`)
  })

  it('字体名含引号/分号时转义，防止逃逸 style 属性注入', () => {
    const stack = fontStackOf('a"b;c', null, theme)
    expect(stack).not.toContain('"')
    expect(stack).not.toContain(';a')
  })

  it('主题引用带未知尾缀（+mn-xx）按 latin 位解析 minor 字体', () => {
    expect(fontStackOf('+mn-xx', null, theme)).toContain("'Calibri'")
  })
})
