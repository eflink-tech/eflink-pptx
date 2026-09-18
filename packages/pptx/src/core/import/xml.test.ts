import { describe, expect, it } from 'vitest'
import { attr, directChild, descendants, firstDescendant } from './xml'

const doc = new DOMParser().parseFromString(
  `<root xmlns:a="urn:a"><item id="1" w="25400"><a:ln w="12700"><a:srgbClr val="FF0000"/></a:ln></item></root>`,
  'text/xml',
)

describe('xml helpers', () => {
  it('attr 读取属性，缺失返回 null', () => {
    const item = directChild(doc.documentElement, 'item')!
    expect(attr(item, 'id')).toBe('1')
    expect(attr(item, 'missing')).toBeNull()
    expect(attr(null, 'id')).toBeNull()
  })

  it('directChild 只找直接子级', () => {
    const item = directChild(doc.documentElement, 'item')!
    expect(directChild(item, 'a:ln')).not.toBeNull()
    expect(directChild(doc.documentElement, 'a:ln')).toBeNull()
  })

  it('descendants/firstDescendant 递归查找', () => {
    const clr = firstDescendant(doc.documentElement, 'a:srgbClr')
    expect(attr(clr, 'val')).toBe('FF0000')
    expect(descendants(doc.documentElement, 'a:srgbClr')).toHaveLength(1)
    expect(firstDescendant(doc.documentElement, 'nope')).toBeNull()
  })
})
