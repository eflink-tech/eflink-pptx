import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from '../import/package'
import { parseXML } from '../import/xml'
import { makeCtx } from './test-utils'
import { renderPicture } from './picture'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

/** 构造含 1x1 PNG（红色）媒体 + slide1 rels 的包，返回挂好 pkg 的 ctx */
async function makePictureCtx(): Promise<ReturnType<typeof makeCtx>> {
  const zip = new JSZip()
  // 1x1 红色 PNG
  const pngB64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  zip.file('ppt/media/a.png', pngB64, { base64: true })
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/a.png"/>
</Relationships>`)
  const blob = await zip.generateAsync({ type: 'blob' })
  const pkg = await PptxPackage.load(new File([blob], 't.pptx'))
  return makeCtx({ pkg: pkg as never })
}

const PIC = `<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`

describe('preview/renderPicture', () => {
  it('基本图片：image + preserveAspectRatio=none', async () => {
    const ctx = await makePictureCtx()
    const g = (await renderPicture(wrap(PIC), ctx))!
    const img = g.querySelector('image')!
    expect(img.getAttribute('href')).toContain('data:image/png')
    expect(img.getAttribute('preserveAspectRatio')).toBe('none')
    expect(Number(img.getAttribute('width'))).toBeCloseTo(100, 1)
  })

  it('srcRect 裁剪：图片放大映射 + clipPath 裁回框内', async () => {
    const ctx = await makePictureCtx()
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/><a:srcRect l="20000" t="10000" r="20000" b="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="6000" cy="3000"/></a:xfrm></p:spPr></p:pic>`)
    const g2 = (await renderPicture(node, ctx))!
    const img = g2.querySelector('image')!
    // 裁剪区宽 60% → 全图宽 = 框宽/0.6；x = 框x - l*全图宽
    expect(Number(img.getAttribute('width'))).toBeGreaterThan(6000 / 9525)
    expect(Number(img.getAttribute('height'))).toBeGreaterThan(3000 / 9525)
    expect(img.getAttribute('clip-path')!.startsWith('url(#clip-')).toBe(true)
    const clipRect = ctx.defs.querySelector('clipPath rect')!
    expect(Number(clipRect.getAttribute('width'))).toBeCloseTo(6000 / 9525, 1)
  })

  it('缺失图片 → null 且计入报告', async () => {
    // 走真实包路径：rIdX 不存在 → relTarget 返回 null（而非空包桩抛 TypeError）
    const ctx = await makePictureCtx()
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdX"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic>`)
    expect(await renderPicture(node, ctx)).toBeNull()
    expect(ctx.report.skipped.missingImage).toBe(1)
  })

  it('视频：海报帧 + 居中播放标记', async () => {
    const ctx = await makePictureCtx()
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="v"/><p:nvPr><a:videoFile r:link="rId1"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr></p:pic>`)
    const g = (await renderPicture(node, ctx))!
    expect(g.querySelector('image')).not.toBeNull() // 海报帧照常渲染
    expect(g.querySelector('circle')).not.toBeNull() // 播放标记
    expect(g.querySelector('path')).not.toBeNull()
  })
})
