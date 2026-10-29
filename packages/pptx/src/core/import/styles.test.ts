import { describe, expect, it } from 'vitest'
import { resolveColor, resolveColorOf } from './styles'
import type { PptxTheme } from './theme'

const theme: PptxTheme = {
  schemeColors: { dk1: '#000000', lt1: '#FFFFFF', accent1: '#4472C4' },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

function hexToChannels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ]
}

describe('styles.resolveColor', () => {
  it('srgbClr 直接取值', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"/></a:solidFill>`), theme)).toBe('#FF0000')
  })

  it('schemeClr 走主题色表', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"/></a:solidFill>`), theme)).toBe('#4472C4')
  })

  it('alpha 修饰 → 8 位 hex', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF000080')
  })

  it('lumMod/lumOff：白底 Darker 25%（lumMod 75000）→ #BFBFBF', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="lt1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>`), theme)).toBe('#BFBFBF')
  })

  it('tint/shade 近似', () => {
    // shade 50% 黑色不变、白色减半
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#808080')
  })

  it('alpha 与 lumMod 组合（顺序无关）：先亮度变换再叠加 alpha', () => {
    const lumOnly = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>`), theme)
    // 修饰符顺序 1：lumMod 在前
    const lumThenAlpha = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"><a:lumMod val="50000"/><a:alpha val="50000"/></a:schemeClr></a:solidFill>`), theme)
    // 修饰符顺序 2：alpha 在前
    const alphaThenLum = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"><a:alpha val="50000"/><a:lumMod val="50000"/></a:schemeClr></a:solidFill>`), theme)
    expect(lumOnly).toMatch(/^#[0-9A-F]{6}$/)
    expect(lumThenAlpha).toBe(`${lumOnly}80`)
    expect(alphaThenLum).toBe(`${lumOnly}80`)
  })

  it('lumMod/lumOff（Lighter 50%）：白底保持 #FFFFFF，accent1 各通道变浅', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="lt1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill>`), theme)).toBe('#FFFFFF')
    const lighter = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill>`), theme)
    expect(lighter).toMatch(/^#[0-9A-F]{6}$/)
    expect(lighter).not.toBe('#4472C4')
    const before = hexToChannels('#4472C4')
    const after = hexToChannels(lighter!)
    expect(after[0]).toBeGreaterThan(before[0])
    expect(after[1]).toBeGreaterThan(before[1])
    expect(after[2]).toBeGreaterThan(before[2])
  })

  it('prstClr 全量表：长名/短别名都命中', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:prstClr val="dkSlateBlue"/></a:solidFill>`), theme)).toBe('#483D8B')
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:prstClr val="lightGoldenrodYellow"/></a:solidFill>`), theme)).toBe('#FAFAD2')
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:prstClr val="notAColor"/></a:solidFill>`), theme)).toBeUndefined()
  })

  it('alphaMod/alphaOff：作用在默认不透明或已有 alpha 上', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:alphaMod val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF000080')
    // 先 alpha 50% 再 alphaMod 50% → 25%
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:alphaMod val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF000040')
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:alphaOff val="-75000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF000040')
  })

  it('单通道修饰：red/redMod/redOff 只改 R 通道', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="000000"><a:red val="100000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF0000')
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="804020"><a:redMod val="0"/></a:srgbClr></a:solidFill>`), theme)).toBe('#004020')
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="000000"><a:greenOff val="50000"/><a:blue val="100000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#0080FF')
  })

  it('hueOff/satOff：色相旋转与饱和度偏移', () => {
    // 纯红 hueOff 120° → 纯绿（HSL 往返允许 1 通道误差）
    const green = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:hueOff val="33333"/></a:srgbClr></a:solidFill>`), theme)!
    const [, g1] = hexToChannels(green)
    expect(g1).toBeGreaterThan(250)
    // satMod 0 → 去饱和为灰（亮度不变）
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:satMod val="0"/></a:srgbClr></a:solidFill>`), theme)).toBe('#808080')
    // satOff 提升饱和度：浅红变成更纯的红
    const deeper = resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="F08080"><a:satOff val="40000"/></a:srgbClr></a:solidFill>`), theme)!
    expect(hexToChannels(deeper)[1]).toBeLessThan(0x80)
  })

  it('非法 hex 返回 undefined（不产生 NaN 传播）', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="GGHHII"/></a:solidFill>`), theme)).toBeUndefined()
  })

  it('空容器/无色子节点返回 undefined', () => {
    expect(resolveColor(el(`<a:ln xmlns:a="urn:a"/>`), theme)).toBeUndefined()
    expect(resolveColor(null, theme)).toBeUndefined()
    expect(resolveColorOf(el(`<a:noFill xmlns:a="urn:a"/>`), theme)).toBeUndefined()
  })
})
