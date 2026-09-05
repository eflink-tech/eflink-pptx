import { describe, expect, it } from 'vitest'
import { alignElements, calcSnap, collectSnapTargets, distributeElements } from './align'
import type { PPTElement } from '../../types/slides'

function textEl(id: string, x: number, y: number, w = 100, h = 50): PPTElement {
  return { id, type: 'text', x, y, w, h, content: '<p>x</p>' }
}

describe('alignElements', () => {
  it('多选左对齐取最小 x', () => {
    const els = [textEl('a', 50, 0), textEl('b', 120, 0)]
    alignElements(els, ['a', 'b'], 'left', 1280, 720)
    expect(els[0].x).toBe(50)
    expect(els[1].x).toBe(50)
  })

  it('单选右对齐参照画布', () => {
    const els = [textEl('a', 0, 0, 100, 50)]
    alignElements(els, ['a'], 'right', 1280, 720)
    expect(els[0].x).toBe(1280 - 100)
  })

  it('垂直居中', () => {
    const els = [textEl('a', 0, 0, 100, 50)]
    alignElements(els, ['a'], 'vcenter', 1280, 720)
    expect(els[0].y).toBe((720 - 50) / 2)
  })
})

describe('distributeElements', () => {
  it('横向等距分布', () => {
    const els = [textEl('a', 0, 0, 100, 50), textEl('b', 30, 0, 100, 50), textEl('c', 400, 0, 100, 50)]
    distributeElements(els, ['a', 'b', 'c'], 'horizontal')
    const xs = els.map((e) => e.x)
    // 首尾不动，中间均匀
    expect(xs[0]).toBe(0)
    expect(xs[2]).toBe(400)
    expect(xs[1]).toBeCloseTo((0 + 400) / 2 - 50 + (100 + 100) / 2 - 100 + 50)
    const gap1 = xs[1] - (xs[0] + 100)
    const gap2 = xs[2] - (xs[1] + 100)
    expect(gap1).toBeCloseTo(gap2)
  })

  it('少于 3 个元素不生效', () => {
    const els = [textEl('a', 0, 0), textEl('b', 30, 0)]
    distributeElements(els, ['a', 'b'], 'vertical')
    expect(els[1].y).toBe(0)
  })
})

describe('calcSnap', () => {
  const targets = collectSnapTargets([textEl('other', 400, 300)], [], 1280, 720)

  it('接近其他元素左边时吸附', () => {
    const snap = calcSnap({ x: 403, y: 100, w: 100, h: 50 }, targets)
    expect(snap.dx).toBe(-3)
    expect(snap.guides.some((g) => g.orient === 'v')).toBe(true)
  })

  it('接近画布中心时吸附', () => {
    const snap = calcSnap({ x: 643, y: 100, w: 100, h: 50 }, targets)
    // 自身中心 693 vs 画布中心 640 → 不在阈值内；自身左边 643 vs 640 → 吸附 -3
    expect(snap.dx).toBe(-3)
  })

  it('远离目标时不吸附', () => {
    const snap = calcSnap({ x: 900, y: 100, w: 100, h: 50 }, targets)
    expect(snap.dx).toBe(0)
    expect(snap.dy).toBe(0)
    expect(snap.guides).toHaveLength(0)
  })
})
