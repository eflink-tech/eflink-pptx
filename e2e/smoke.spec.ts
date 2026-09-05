import { expect, type Page, test } from '@playwright/test'

// 画布内元素（排除缩略图中的同源渲染副本）
function canvas(page: Page) {
  return page.getByTestId('canvas-slide')
}

test('编辑器核心流程', async ({ page }) => {
  await page.goto('/')
  await expect(page).toHaveTitle(/易飞演示文稿/)

  // 顶栏与画布就绪
  await expect(page.getByTestId('topbar')).toBeVisible()
  await expect(page.getByTestId('thumbnail-panel')).toBeVisible()
  await expect(page.getByTestId('style-panel')).toBeVisible()

  // 插入文本框：画布出现元素且选中（8 个缩放柄）
  await page.getByTestId('insert-text').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(1)
  await expect(page.locator('.element-handle')).toHaveCount(8)

  // 插入表格与图表（图表需在类型面板中选定一种）
  await page.getByTestId('insert-table').click()
  await expect(canvas(page).locator('.pptx-element table')).toHaveCount(1)
  await page.getByTestId('insert-chart').click()
  await page.getByTestId('chart-picker-panel').getByTitle('簇状柱状图').click()
  await expect(canvas(page).locator('.pptx-element svg').first()).toBeVisible()

  // 撤销/重做可用
  await page.getByTitle('撤销').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(2)
  await page.getByTitle('重做').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(3)

  // 页面管理：新建一页
  await page.getByTestId('add-slide').click()
  await expect(page.locator('[data-testid^="thumb-"]')).toHaveCount(2)
  await expect(page.getByTestId('slide-nav')).toHaveText('2 / 2')

  // 切回第 1 页并选中元素
  await page.locator('[data-testid="thumb-0"]').click()
  await canvas(page).locator('.pptx-element').first().click({ force: true })
  await expect(page.locator('.element-handle')).toHaveCount(8)

  // AI 面板开关
  await page.getByTestId('ai-toggle').click()
  await expect(page.getByTestId('ai-panel')).toBeVisible()
  await page.getByTestId('ai-toggle').click()

  // 放映模式进出
  await page.getByTestId('play').click()
  await expect(page.getByTestId('player')).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(page.getByTestId('player')).not.toBeVisible()
})

test('刷新后自动恢复文档', async ({ page }) => {
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(1)
  // 触发 Ctrl+S 立即保存
  await page.keyboard.press('ControlOrMeta+s')

  await page.reload()
  await page.waitForLoadState('domcontentloaded')
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(1)
})

test('模板库应用', async ({ page }) => {
  await page.goto('/')
  await page.getByTitle('模板库').click()
  await expect(page.locator('.modal-pop')).toBeVisible()
  await page.getByRole('button', { name: '应用整套' }).first().click()
  await expect(page.locator('.modal-pop')).not.toBeVisible()
  await expect(page.locator('[data-testid^="thumb-"]')).toHaveCount(3)
})

test('导出 PPTX / JSON / PNG 下载', async ({ page }) => {
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(1)

  // JSON 导出
  await page.getByTitle('导出').click()
  await expect(page.locator('.modal-pop')).toBeVisible()
  const jsonDownload = page.waitForEvent('download', { timeout: 15000 })
  await page.getByRole('button', { name: '导出 JSON' }).click()
  const json = await jsonDownload
  expect(json.suggestedFilename()).toContain('.json')

  // PPTX 导出
  const pptxDownload = page.waitForEvent('download', { timeout: 30000 })
  await page.getByTestId('export-pptx').click()
  const pptx = await pptxDownload
  expect(pptx.suggestedFilename()).toContain('.pptx')

  // PNG 导出（当前页）
  const pngDownload = page.waitForEvent('download', { timeout: 30000 })
  await page.getByRole('button', { name: '仅当前页' }).click()
  const png = await pngDownload
  expect(png.suggestedFilename()).toContain('.png')
  await page.keyboard.press('Escape')
})

test('查找替换', async ({ page }) => {
  await page.goto('/')
  // 新插入文本默认内容为「默认文本」
  await page.getByTestId('insert-text').click()
  await expect(canvas(page).locator('.pptx-element')).toHaveCount(1)

  // 打开查找替换
  await page.getByTitle('查找替换（Ctrl+F）').click()
  await page.getByPlaceholder('查找内容').fill('默认')
  await page.getByRole('button', { name: '查找', exact: true }).click()
  await page.getByPlaceholder('替换为（可留空）').fill('你好')
  await page.getByRole('button', { name: '全部替换' }).click()
  // 画布中的文本被替换
  await expect(canvas(page).locator('.pptx-element').first()).toContainText('你好文本')
})

test('PPTX 导出→导入往返', async ({ page }) => {
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await page.getByTestId('insert-table').click()
  await expect(canvas(page).locator('.pptx-element table')).toHaveCount(1)

  // 导出 PPTX
  await page.getByTitle('导出').click()
  const dl = page.waitForEvent('download', { timeout: 30000 })
  await page.getByTestId('export-pptx').click()
  const download = await dl
  const filePath = await download.path()

  // 关闭导出框，导入回来
  await page.keyboard.press('Escape')
  await page.getByTitle('导入（PPTX/JSON）').click()
  await page.setInputFiles('input[type="file"]', filePath!)
  // 导入完成后画布应有表格与文本
  await expect(canvas(page).locator('.pptx-element table')).toHaveCount(1, { timeout: 15000 })
  await expect(canvas(page).locator('.pptx-element').first()).toContainText('默认文本')
})
