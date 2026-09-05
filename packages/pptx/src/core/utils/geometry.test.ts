import { describe, expect, it } from 'vitest'
import { applyResize, rotatedRectBounds, unionBounds } from './geometry'

describe('applyResize', () => {
  it('无旋转：拖动右下角扩大', () => {
    const r = applyResize({ x: 100, y: 100, w: 200, h: 100 }, 0, 'se', 50, 30)
    expect(r.w).toBe(250)
    expect(r.h).toBe(130)
    expect(r.x).toBe(100)
    expect(r.y).toBe(100)
  })

  it('无旋转：拖动左上角，右下角固定', () => {
    const r = applyResize({ x: 100, y: 100, w: 200, h: 100 }, 0, 'nw', -40, -20)
    expect(r.x).toBe(60)
    expect(r.y).toBe(80)
    expect(r.x + r.w).toBe(300)
    expect(r.y + r.h).toBe(200)
  })

  it('最小尺寸限制', () => {
    const r = applyResize({ x: 100, y: 100, w: 200, h: 100 }, 0, 'w', 500, 0)
    expect(r.w).toBe(10)
  })

  it('旋转 90°：向右拖动改变的是高度方向', () => {
    // 元素旋转 90° 后，画布 x 位移对应本地 -y
    const r = applyResize({ x: 100, y: 100, w: 200, h: 100 }, 90, 'se', 50, 0)
    // 本地位移 ldx = dx*cos90 = 0, ldy = -dx*sin90 = -50 → h 缩小 50
    expect(r.h).toBeCloseTo(50)
  })

  it('keepRatio 保持宽高比', () => {
    const r = applyResize({ x: 0, y: 0, w: 200, h: 100 }, 0, 'se', 100, 0, true)
    expect(r.w / r.h).toBeCloseTo(2)
  })
})

describe('rotatedRectBounds', () => {
  it('无旋转返回原矩形', () => {
    expect(rotatedRectBounds({ x: 10, y: 20, w: 30, h: 40 }, 0)).toEqual({ x: 10, y: 20, w: 30, h: 40 })
  })

  it('旋转 45° 包围盒变大', () => {
    const b = rotatedRectBounds({ x: 0, y: 0, w: 100, h: 100 }, 45)
    expect(b.w).toBeGreaterThan(100)
    expect(b.h).toBeGreaterThan(100)
  })
})

describe('unionBounds', () => {
  it('合并多个矩形', () => {
    const u = unionBounds([
      { x: 0, y: 0, w: 100, h: 100 },
      { x: 200, y: 50, w: 50, h: 50 },
    ])
    expect(u).toEqual({ x: 0, y: 0, w: 250, h: 100 })
  })
})
