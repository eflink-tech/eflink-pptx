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

  it('组合非零 chOff：子元素按 off-chOff 补偿折算', async () => {
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
    // off=(1000000,500000) ext=(2000000,1000000) chOff=(100000,50000) chExt=(1000000,1000000)
    // 子在 chOff 空间 (100000,50000) 大小 (500000,250000) → 绝对 (1000000,500000) 大小 (1000000,250000) EMU
    // px(96dpi)：x=104.99→105, y=52.49→52, w=104.99→105, h=26.25→26
    const grp = el(`<p:grpSp xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGrpSpPr><p:cNvPr id="1" name="组"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>
        <a:off x="1000000" y="500000"/><a:ext cx="2000000" cy="1000000"/>
        <a:chOff x="100000" y="50000"/><a:chExt cx="1000000" cy="1000000"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="子矩形"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="100000" y="50000"/><a:ext cx="500000" cy="250000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>
    </p:grpSp>`)
    const els = await parseSpTreeNode(grp, ctx, IDENTITY_XFORM, pkg)
    expect(els).toHaveLength(1)
    const shape = els[0]
    if (shape.type !== 'shape') throw new Error('expected shape')
    expect(shape.x).toBe(105)
    expect(shape.y).toBe(52)
    expect(shape.w).toBe(105)
    expect(shape.h).toBe(26)
  })

  it('组合旋转：子元素绕组合中心映射并累计旋转角', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="urn:p"/>`)
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const report: ImportReport = { skipped: {} }
    const ctx = {
      pkg,
      partPath: 'ppt/slides/slide1.xml',
      theme,
      report,
      scale: { x: 1, y: 1 },
      placeholders: new Map(),
    }
    // rot="600000" = 10°：子元素中心绕组合中心旋转，自身累计旋转 10°
    const grp = el(`<p:grpSp xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGrpSpPr><p:cNvPr id="1" name="组"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm rot="600000">
        <a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/>
        <a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="子矩形"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>
    </p:grpSp>`)
    const els = await parseSpTreeNode(grp, ctx, IDENTITY_XFORM, pkg)
    expect(els).toHaveLength(1)
    const shape = els[0]
    if (shape.type !== 'shape') throw new Error('expected shape')
    expect(shape.rotate).toBe(10)
    // 子中心 (250000,250000) → 缩放 (500000,250000) → 绕组合中心 (1000000,500000) 旋转 10°
    // → (551008,166974)，盒 1000000×500000 → px x=5,y=-9,w=105,h=52
    expect(shape.x).toBe(5)
    expect(shape.y).toBe(-9)
    expect(shape.w).toBe(105)
    expect(shape.h).toBe(52)
    expect(report.skipped.groupRotation).toBeUndefined()
  })

  it('组合翻转：子元素位置镜像并标记 flipH', async () => {
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
    const grp = el(`<p:grpSp xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGrpSpPr><p:cNvPr id="1" name="组"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm flipH="1">
        <a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/>
        <a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="子矩形"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="200000" cy="200000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>
    </p:grpSp>`)
    const els = await parseSpTreeNode(grp, ctx, IDENTITY_XFORM, pkg)
    const shape = els[0]
    if (shape.type !== 'shape') throw new Error('expected shape')
    expect(shape.flipH).toBe(true)
    expect(shape.x).toBe(84) // 子中心镜像到 x=900000 → 盒左 800000/9525
    expect(shape.y).toBe(0)
  })
})
