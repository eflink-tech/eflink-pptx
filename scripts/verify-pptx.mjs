// 校验导出的 PPTX 为合法 OOXML 包（生成 → 下载 → 解包检查）
import { chromium } from '@playwright/test'
import { mkdirSync } from 'node:fs'

const PORT = 5180
mkdirSync('test-results/pptx-check', { recursive: true })

// 复用已运行的 dev server；否则报错退出
async function main() {
  const browser = await chromium.launch()
  const page = await browser.newPage()
  await page.goto(`http://localhost:${PORT}/`)
  await page.getByTestId('insert-text').click()
  await page.getByTestId('insert-table').click()

  await page.getByTitle('导出').click()
  const downloadPromise = page.waitForEvent('download', { timeout: 30000 })
  await page.getByTestId('export-pptx').click()
  const download = await downloadPromise
  const path = 'test-results/pptx-check/export.pptx'
  await download.saveAs(path)
  await browser.close()

  // 用 unzip 检查结构（python zipfile 更稳）
  const { execSync } = await import('node:child_process')
  const list = execSync(`python3 -c "
import zipfile
z = zipfile.ZipFile('${path}')
names = z.namelist()
required = ['[Content_Types].xml', 'ppt/presentation.xml']
slides = [n for n in names if n.startswith('ppt/slides/slide')]
print('OK' if all(r in names for r in required) and slides else 'MISSING')
print('slides:', len(slides))
import re
pres = z.read('ppt/presentation.xml').decode('utf-8')
print('has_sldSz:', 'sldSz' in pres)
slide1 = z.read([n for n in sorted(slides)][0]).decode('utf-8')
print('has_table:', 'graphicFrame' in slide1 or 'tbl' in slide1)
print('has_text:', '默认文本' in slide1)
"`).toString()
  console.log(list)
}
main().catch((e) => { console.error(e); process.exit(1) })
