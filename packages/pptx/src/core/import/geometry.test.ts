import { describe, expect, it } from 'vitest'
import { custGeomToPath, getShapeKey, presetGeomToPath } from './geometry'

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
    // 元素 96x96px → 914400 EMU；路径缺省 w/h 用元素尺寸；坐标超界不裁剪（渲染端 viewBox 负责裁剪显示）
    expect(custGeomToPath(cust, 96, 96)).toBe('M0,0 C0,104.17 104.17,104.17 104.17,0')
  })

  it('custGeomToPath：无 pathLst 返回 null', () => {
    expect(custGeomToPath(el(`<a:custGeom xmlns:a="urn:a"/>`), 10, 10)).toBeNull()
  })

  it('presetGeomToPath：snip1Rect 顶部切角矩形（adj 相对 min(w,h)）', () => {
    // adj=50000、71x14：切角 7px → x1 = 7/71*100 ≈ 9.86
    const withAdj = el(`<a:prstGeom xmlns:a="urn:a" prst="snip1Rect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>`)
    expect(presetGeomToPath(withAdj, 'snip1Rect', 71, 14))
      .toBe('M9.86,0 L90.14,0 L100,100 L0,100 Z')
    // 无 avLst → OOXML 默认 adj=16667：切角 14*16667/100000 ≈ 2.33px → x1 ≈ 3.29
    const noAdj = el(`<a:prstGeom xmlns:a="urn:a" prst="snip1Rect"><a:avLst/></a:prstGeom>`)
    expect(presetGeomToPath(noAdj, 'snip1Rect', 71, 14))
      .toBe('M3.29,0 L96.71,0 L100,100 L0,100 Z')
    // 非支持预设返回 null（走既有 shapeKey 映射）
    expect(presetGeomToPath(null, 'rect', 10, 10)).toBeNull()
  })

  it('presetGeomToPath：frame 四边边框（adj1/adj2 = 左上/右下厚度）', () => {
    // 1280x720、adj1=2130：厚度 720*2130/100000 ≈ 15.34px → tx=15.34/1280*100 ≈ 1.2、ty ≈ 2.13
    const f = el(`<a:prstGeom xmlns:a="urn:a" prst="frame"><a:avLst><a:gd name="adj1" fmla="val 2130"/><a:gd name="adj2" fmla="val 2130"/></a:avLst></a:prstGeom>`)
    expect(presetGeomToPath(f, 'frame', 1280, 720))
      .toBe('M0,0 L100,0 L100,100 L0,100 Z M1.2,2.13 L1.2,97.87 L98.8,97.87 L98.8,2.13 Z')
    // 源文件常只声明 adj1：其余边跟随 adj1（均匀边框），实证 LibreOffice/PowerPoint 均如此
    const only1 = el(`<a:prstGeom xmlns:a="urn:a" prst="frame"><a:avLst><a:gd name="adj1" fmla="val 2130"/></a:avLst></a:prstGeom>`)
    expect(presetGeomToPath(only1, 'frame', 1280, 720))
      .toBe('M0,0 L100,0 L100,100 L0,100 Z M1.2,2.13 L1.2,97.87 L98.8,97.87 L98.8,2.13 Z')
    // 无 avLst → ECMA 默认 adj=12500：厚度 720*0.125=90px → tx=90/1280*100=7.03、ty=12.5
    const d = el(`<a:prstGeom xmlns:a="urn:a" prst="frame"><a:avLst/></a:prstGeom>`)
    expect(presetGeomToPath(d, 'frame', 1280, 720))
      .toBe('M0,0 L100,0 L100,100 L0,100 Z M7.03,12.5 L7.03,87.5 L92.97,87.5 L92.97,12.5 Z')
  })
})
