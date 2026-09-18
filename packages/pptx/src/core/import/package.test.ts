import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage, resolveTarget } from './package'

async function buildPkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation/>')
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`)
  zip.file('ppt/media/image1.png', 'PNGDATA')
  const blob = await zip.generateAsync({ type: 'blob' })
  return PptxPackage.load(blob)
}

describe('PptxPackage', () => {
  it('text 读取部件文本，缺失返回 null', async () => {
    const pkg = await buildPkg()
    expect(await pkg.text('ppt/presentation.xml')).toContain('presentation')
    expect(await pkg.text('ppt/missing.xml')).toBeNull()
  })

  it('rels 解析关系并区分 External', async () => {
    const pkg = await buildPkg()
    const rels = await pkg.rels('ppt/slides/slide1.xml')
    expect(rels.get('rId1')?.target).toBe('ppt/media/image1.png')
    expect(rels.get('rId2')?.mode).toBe('External')
  })

  it('relTarget 解析为包内绝对路径', async () => {
    const pkg = await buildPkg()
    expect(await pkg.relTarget('ppt/slides/slide1.xml', 'rId1')).toBe('ppt/media/image1.png')
    expect(await pkg.relTarget('ppt/slides/slide1.xml', 'rIdX')).toBeNull()
  })

  it('mediaDataUrl 产出 base64 data URL 并缓存', async () => {
    const pkg = await buildPkg()
    const url = await pkg.mediaDataUrl('ppt/media/image1.png')
    expect(url).toMatch(/^data:image\/png;base64,/)
    expect(await pkg.mediaDataUrl('ppt/media/none.png')).toBeUndefined()
  })

  it('resolveTarget 处理相对/绝对路径', () => {
    expect(resolveTarget('ppt/slides', '../media/a.png')).toBe('ppt/media/a.png')
    expect(resolveTarget('ppt', '/ppt/media/a.png')).toBe('ppt/media/a.png')
  })

  it('根部件 rels：按 _rels/<name>.rels 解析并记录关系 Type', async () => {
    const zip = new JSZip()
    zip.file('presentation.xml', '<p:presentation/>')
    zip.file('_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>`)
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const rels = await pkg.rels('presentation.xml')
    expect(rels.get('rId1')?.target).toBe('slideMasters/slideMaster1.xml')
    expect(rels.get('rId1')?.type).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster')
  })
})
