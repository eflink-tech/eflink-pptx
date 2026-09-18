import { describe, expect, it } from 'vitest'
import { custGeomToPath, getShapeKey } from './geometry'

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('geometry', () => {
  it('getShapeKey：已知预设映射内部 key，未知回退 rect', () => {
    expect(getShapeKey('roundRect')).toBe('roundRect')
    expect(getShapeKey('rightArrow')).toBe('arrowRight')
    expect(getShapeKey('nonExistentPrst')).toBe('rect')
  })

  it('custGeomToPath：moveTo/lnTo/close 归一化到 0-100', () => {
    const cust = el(`<a:custGeom xmlns:a="urn:a">
      <a:pathLst>
        <a:path w="1000" h="500">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="1000" y="0"/></a:lnTo>
          <a:lnTo><a:pt x="500" y="500"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>`)
    expect(custGeomToPath(cust, 200, 100)).toBe('M0,0 L100,0 L50,100 Z')
  })

  it('custGeomToPath：cubicBezTo 与缺省 path 尺寸（用元素尺寸）', () => {
    const cust = el(`<a:custGeom xmlns:a="urn:a">
      <a:pathLst>
        <a:path>
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:cubicBezTo><a:pt x="0" y="952500"/><a:pt x="952500" y="952500"/><a:pt x="952500" y="0"/></a:cubicBezTo>
        </a:path>
      </a:pathLst>
    </a:custGeom>`)
    // 元素 96x96px → 952500 EMU；路径缺省 w/h 用元素尺寸
    expect(custGeomToPath(cust, 96, 96)).toBe('M0,0 C0,100 100,100 100,0')
  })

  it('custGeomToPath：无 pathLst 返回 null', () => {
    expect(custGeomToPath(el(`<a:custGeom xmlns:a="urn:a"/>`), 10, 10)).toBeNull()
  })
})
