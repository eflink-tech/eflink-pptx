import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import type { ImportReport } from '../../types/slides'
import { SVG_NS, svgEl, emu2pxF, sz2px, geomOf, boxTransform, type PreviewCtx } from './svg'

/** 构造最小预览上下文（结构兼容 ParseContext，pkg 仅占位） */
export function makeCtx(overrides?: Partial<PreviewCtx>): PreviewCtx {
  const defs = document.createElementNS(SVG_NS, 'defs')
  let n = 0
  return {
    pkg: {} as never,
    partPath: 'ppt/slides/slide1.xml',
    theme: { schemeColors: {}, majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {} },
    report: { skipped: {} } as ImportReport,
    scale: { x: 1, y: 1 },
    placeholders: new Map(),
    defs,
    uid: (prefix: string) => `${prefix}-${n++}`,
    ...overrides,
  }
}

describe('preview/svg 助手', () => {
  it('emu2pxF：EMU→px 浮点不取整', () => {
    expect(emu2pxF(9525)).toBe(1)
    expect(emu2pxF(12192000)).toBe(1280)
    expect(emu2pxF(9524)).toBeCloseTo(0.9999, 4) // 9524/9525 = 0.999895…，浮点不取整
    expect(emu2pxF(9524)).toBeLessThan(1)
  })

  it('sz2px：字号（1/100 pt）→ px', () => {
    expect(sz2px(1800)).toBe(24)
    expect(sz2px(2400)).toBe(32)
  })

  it('svgEl：构建命名空间元素，undefined/null 属性跳过', () => {
    const el = svgEl('rect', { x: 1, fill: undefined, stroke: null, width: '10' })
    expect(el.namespaceURI).toBe('http://www.w3.org/2000/svg')
    expect(el.getAttribute('x')).toBe('1')
    expect(el.getAttribute('width')).toBe('10')
    expect(el.hasAttribute('fill')).toBe(false)
    expect(el.hasAttribute('stroke')).toBe(false)
  })

  it('geomOf：a:xfrm → Box（px 浮点 + rot/flip）', () => {
    const xml = `<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:sp><p:spPr>
        <a:xfrm rot="1800000" flipH="1"><a:off x="952500" y="190500"/><a:ext cx="4762500" cy="952500"/></a:xfrm>
      </p:spPr></p:sp></root>`
    const sp = parseXML(xml).documentElement.firstElementChild as Element
    const box = geomOf(sp, makeCtx())
    expect(box).not.toBeNull()
    expect(box!.x).toBe(100)
    expect(box!.y).toBe(20)
    expect(box!.w).toBe(500)
    expect(box!.h).toBe(100)
    expect(box!.rot).toBe(30)
    expect(box!.flipH).toBe(true)
    expect(box!.flipV).toBe(false)
  })

  it('geomOf：无 xfrm 时回退占位符位置表（EMU→px）', () => {
    const xml = `<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:sp><p:nvSpPr><p:ph type="title"/></p:nvSpPr><p:spPr/></p:sp></root>`
    const sp = parseXML(xml).documentElement.firstElementChild as Element
    const placeholders = new Map([['title', { x: 952500, y: 952500, w: 9525000, h: 952500 }]])
    const box = geomOf(sp, makeCtx({ placeholders }))
    expect(box).toEqual({ x: 100, y: 100, w: 1000, h: 100, rot: 0, flipH: false, flipV: false })
  })

  it('boxTransform：rot 绕自身中心、flip 平移翻转', () => {
    const t1 = boxTransform({ x: 100, y: 100, w: 500, h: 100, rot: 30, flipH: false, flipV: false })
    expect(t1).toBe('rotate(30,350,150)')
    const t2 = boxTransform({ x: 100, y: 100, w: 500, h: 100, rot: 0, flipH: true, flipV: false })
    expect(t2).toBe('translate(600,100) scale(-1,1) translate(-100,-100)')
    expect(boxTransform({ x: 0, y: 0, w: 10, h: 10, rot: 0, flipH: false, flipV: false })).toBeUndefined()
  })
})
