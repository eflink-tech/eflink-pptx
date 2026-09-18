import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parsePictureEl } from './picture'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/a.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/v.mp4"/>
</Relationships>`)
  zip.file('ppt/media/a.png', 'PNGDATA')
  zip.file('ppt/media/v.mp4', 'MP4DATA')
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

const ctxBase = (pkg: PptxPackage) => ({
  pkg,
  partPath: 'ppt/slides/slide1.xml',
  theme,
  report: { skipped: {} } as ImportReport,
  scale: { x: 1, y: 1 },
  placeholders: new Map(),
})

describe('parsePictureEl', () => {
  it('图片：src + srcRect 裁剪', async () => {
    const pkg = await makePkg()
    const xml = `<p:pic xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:nvPicPr><p:cNvPr id="2" name="Img"/></p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="rId1"/>
        <a:srcRect l="10000" t="20000" r="10000" b="20000"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>`
    const result = await parsePictureEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'image') throw new Error('expected image')
    expect(result.src).toMatch(/^data:image\/png;base64,/)
    expect(result.clip).toEqual({ x: 0.1, y: 0.2, w: 0.8, h: 0.6 })
  })

  it('图片：xfrm rot → rotate（1/60000 deg）', async () => {
    const pkg = await makePkg()
    const xml = `<p:pic xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:nvPicPr><p:cNvPr id="2" name="Img"/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
      <p:spPr><a:xfrm rot="19829678"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
    </p:pic>`
    const result = await parsePictureEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'image') throw new Error('expected image')
    // 19829678/60000 ≈ 330.49°，取整为 330
    expect(result.rotate).toBe(330)
  })

  it('视频：videoFile → video 元素，blip 作海报帧', async () => {
    const pkg = await makePkg()
    const xml = `<p:pic xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:nvPicPr><p:cNvPr id="2" name="Vid"/>
        <p:nvPr><a:videoFile r:link="rId2"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
    </p:pic>`
    const result = await parsePictureEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'video') throw new Error('expected video')
    expect(result.src).toMatch(/^data:video\/mp4;base64,/)
    expect(result.poster).toMatch(/^data:image\/png;base64,/)
    expect(result.loop).toBe(false)
  })
})
