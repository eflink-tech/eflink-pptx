import { describe, expect, it } from 'vitest'
import { resolveColor, resolveColorOf } from './styles'
import type { PptxTheme } from './theme'

const theme: PptxTheme = {
  schemeColors: { dk1: '#000000', lt1: '#FFFFFF', accent1: '#4472C4' } as Record<string, string>,
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
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

  it('空容器/无色子节点返回 undefined', () => {
    expect(resolveColor(el(`<a:ln xmlns:a="urn:a"/>`), theme)).toBeUndefined()
    expect(resolveColor(null, theme)).toBeUndefined()
    expect(resolveColorOf(el(`<a:noFill xmlns:a="urn:a"/>`), theme)).toBeUndefined()
  })
})
