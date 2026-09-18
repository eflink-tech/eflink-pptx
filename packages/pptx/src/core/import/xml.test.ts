import { describe, expect, it } from 'vitest'
import { attr, directChild, directChildren, descendants, firstDescendant, parseXML } from './xml'

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

  it('directChildren 过滤直接子级，无匹配返回空数组', () => {
    expect(directChildren(doc.documentElement, 'item')).toHaveLength(1)
    expect(directChildren(doc.documentElement, '不存在的名字')).toEqual([])
  })

  it('parseXML 对非法 XML 显式抛错（而非静默返回 parsererror 文档）', () => {
    expect(() => parseXML('<root><unclosed></root>')).toThrow(/XML 解析失败/)
    expect(() => parseXML('不是 XML 的普通文本')).toThrow(/XML 解析失败/)
  })
})
