import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseSpTreeNode } from './index'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('parseGroupEl（经 parseSpTreeNode 分发）', () => {
  it('组合子元素坐标按 off/chOff/ext/chExt 折算并打平', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="urn:p"/>`)
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const ctx = {
      pkg,
      partPath: 'ppt/slides/slide1.xml',
      theme,
      report: { skipped: {} } as ImportReport,
      scale: { x: 1, y: 1 },
      placeholders: new Map(),
    }
    // 组合：off=(0,0) ext=(2000000,1000000) chOff=(0,0) chExt=(1000000,1000000) → x 放大 2 倍
    // 子元素 chOff 空间 x=0,y=0,w=500000,h=500000 → 绝对 x=0,y=0,w=1000000,h=1000000 EMU → px(96dpi) w=105,h=105
    const grp = el(`<p:grpSp xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGrpSpPr><p:cNvPr id="1" name="组"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>
        <a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/>
        <a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="子矩形"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>
    </p:grpSp>`)
    const els = await parseSpTreeNode(grp, ctx, IDENTITY_XFORM, pkg)
    expect(els).toHaveLength(1)
    const shape = els[0]
    if (shape.type !== 'shape') throw new Error('expected shape')
    expect(shape.w).toBe(105) // 1000000 EMU / 9525 = 104.98 → 105
    expect(shape.shapeKey).toBe('roundRect')
    expect(shape.fill).toBe('#42A5F5')
  })
})
