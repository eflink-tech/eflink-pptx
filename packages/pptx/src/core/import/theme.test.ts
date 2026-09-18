import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from './package'
import { parseThemeForMaster, resolveSchemeColor } from './theme'

async function buildPkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>`)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`)
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`)
  const blob = await zip.generateAsync({ type: 'blob' })
  return PptxPackage.load(blob)
}

describe('parseThemeForMaster', () => {
  it('解析主题色表 / 字体 / clrMap，sysClr 取 lastClr', async () => {
    const theme = await parseThemeForMaster(await buildPkg(), 'ppt/slideMasters/slideMaster1.xml')
    expect(theme.schemeColors.dk1).toBe('#000000')
    expect(theme.schemeColors.lt1).toBe('#FFFFFF')
    expect(theme.schemeColors.accent1).toBe('#4472C4')
    expect(theme.majorFont).toBe('Calibri Light')
    expect(theme.minorFont).toBe('Calibri')
    expect(theme.colorMap.bg1).toBe('lt1')
  })

  it('resolveSchemeColor 过 clrMap 再查色表；缺失回退黑色', async () => {
    const theme = await parseThemeForMaster(await buildPkg(), 'ppt/slideMasters/slideMaster1.xml')
    expect(resolveSchemeColor('bg1', theme)).toBe('#FFFFFF')
    expect(resolveSchemeColor('accent2', theme)).toBe('#ED7D31')
    expect(resolveSchemeColor('nope', theme)).toBe('#000000')
  })

  it('缺 theme 部件时回退默认色表', async () => {
    const zip = new JSZip()
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:x"/>')
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const theme = await parseThemeForMaster(pkg, 'ppt/slideMasters/slideMaster1.xml')
    expect(theme.schemeColors.accent1).toBe('#4472C4')
  })
})
