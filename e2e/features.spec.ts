import { expect, type Page, test } from '@playwright/test'

// 深度功能自测套件：插入/编辑/文本/表格/图表/页面/动画/放映/文件管理/导出全链路
// 每个用例独立上下文（localStorage/IndexedDB 隔离），全部监听页面运行时错误

function canvasEl(page: Page) {
  return page.getByTestId('canvas-slide').locator('.pptx-element')
}

function watchErrors(page: Page) {
  const errors: string[] = []
  page.on('pageerror', (err) => errors.push(err.message))
  return errors
}

function expectNoErrors(errors: string[]) {
  expect(errors, `页面运行时错误: ${errors.join(' | ')}`).toEqual([])
}

/** 插入一个图表并选中；expectedTotal 为画布上应有的元素总数 */
async function insertChart(page: Page, typeLabel = '簇状柱状图', expectedTotal = 1) {
  await page.getByTestId('insert-chart').click()
  await page.getByTestId('chart-picker-panel').getByTitle(typeLabel).click()
  await expect(canvasEl(page)).toHaveCount(expectedTotal)
  await canvasEl(page).last().click()
  await expect(page.locator('.element-handle')).toHaveCount(8)
}

const TINY_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

test('插入形状/线条/公式/图片/音频', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')

  // 形状：矩形
  await page.getByTestId('insert-shape').click()
  await page.locator('button[title="矩形"]').click()
  await expect(canvasEl(page)).toHaveCount(1)
  await expect(canvasEl(page).locator('svg path').first()).toBeVisible()

  // 线条：实线直线
  await page.getByTestId('insert-line').click()
  await page.getByTestId('line-preset-menu').locator('button[title="straight-solid"]').click()
  await expect(canvasEl(page)).toHaveCount(2)

  // 公式：直接插入默认 LaTeX
  await page.getByTestId('insert-formula').click()
  await expect(canvasEl(page)).toHaveCount(3)

  // 图片：动态 file input → filechooser
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('insert-image').click(),
  ])
  await chooser.setFiles({ name: 'dot.png', mimeType: 'image/png', buffer: TINY_PNG })
  await expect(canvasEl(page).locator('img')).toHaveCount(1)

  // 音频：同上（dataURL 插入）
  const [audioChooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('insert-audio').click(),
  ])
  await audioChooser.setFiles({ name: 'a.mp3', mimeType: 'audio/mpeg', buffer: Buffer.from('ID3', 'utf-8') })
  await expect(canvasEl(page)).toHaveCount(5)

  expectNoErrors(errors)
})

test('元素拖拽移动、方向键微调、缩放与旋转', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)

  // 拖拽移动
  const before = await canvasEl(page).boundingBox()
  expect(before).toBeTruthy()
  const cx = before!.x + before!.width / 2
  const cy = before!.y + before!.height / 2
  await page.mouse.move(cx, cy)
  await page.mouse.down()
  await page.mouse.move(cx + 60, cy + 40, { steps: 5 })
  await page.mouse.up()
  const afterMove = await canvasEl(page).boundingBox()
  expect(Math.abs(afterMove!.x - before!.x)).toBeGreaterThan(30)
  expect(Math.abs(afterMove!.y - before!.y)).toBeGreaterThan(15)

  // 方向键微调（选中态下 x += 1 画布px）
  await canvasEl(page).click()
  const beforeNudge = await canvasEl(page).boundingBox()
  await page.keyboard.press('ArrowRight')
  const afterNudge = await canvasEl(page).boundingBox()
  expect(afterNudge!.x).toBeGreaterThan(beforeNudge!.x)

  // 缩放：拖拽 se 手柄（顺序 nw,n,ne,e,se,s,sw,w → 下标 4）
  const beforeResize = await canvasEl(page).boundingBox()
  const handle = page.locator('.element-handle').nth(4)
  const h = await handle.boundingBox()
  await page.mouse.move(h!.x + h!.width / 2, h!.y + h!.height / 2)
  await page.mouse.down()
  await page.mouse.move(h!.x + 50, h!.y + 40, { steps: 4 })
  await page.mouse.up()
  const afterResize = await canvasEl(page).boundingBox()
  expect(afterResize!.width).toBeGreaterThan(beforeResize!.width)
  expect(afterResize!.height).toBeGreaterThan(beforeResize!.height)

  // 旋转：拖拽旋转柄
  const r = await page.locator('[title="旋转"]').boundingBox()
  await page.mouse.move(r!.x + r!.width / 2, r!.y + r!.height / 2)
  await page.mouse.down()
  await page.mouse.move(r!.x + 80, r!.y + 60, { steps: 6 })
  await page.mouse.up()
  const transform = await canvasEl(page).getAttribute('style')
  expect(transform).toContain('rotate')

  expectNoErrors(errors)
})

test('快捷键：原位复制/粘贴/删除/撤销/重做', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)
  await canvasEl(page).click()

  // Ctrl+D 原位复制
  await page.keyboard.press('ControlOrMeta+d')
  await expect(canvasEl(page)).toHaveCount(2)

  // Ctrl+C / Ctrl+V
  await page.keyboard.press('ControlOrMeta+c')
  await page.keyboard.press('ControlOrMeta+v')
  await expect(canvasEl(page)).toHaveCount(3)

  // Delete 删除选中（粘贴后克隆为选中态）
  await page.keyboard.press('Delete')
  await expect(canvasEl(page)).toHaveCount(2)

  // 撤销一步应恢复删除前的 3 个元素
  await page.getByTitle('撤销').click()
  await expect(canvasEl(page)).toHaveCount(3)

  // 重做回到删除后的 2 个
  await page.getByTitle('重做').click()
  await expect(canvasEl(page)).toHaveCount(2)

  expectNoErrors(errors)
})

test('右键菜单：锁定/解锁与组合', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await page.getByTestId('insert-shape').click()
  await page.locator('button[title="矩形"]').click()
  await expect(canvasEl(page)).toHaveCount(2)

  // 组合：Ctrl+A 全选后 Ctrl+G，再右键应出现「取消组合」
  await page.keyboard.press('ControlOrMeta+a')
  await page.keyboard.press('ControlOrMeta+g')
  await canvasEl(page).first().click({ button: 'right', force: true })
  await expect(page.getByTestId('context-menu')).toBeVisible()
  await expect(page.getByTestId('context-menu').getByText('取消组合')).toBeVisible()
  await page.keyboard.press('Escape')

  // 右键单个元素（最上层）→ 锁定 → 选中后无操作柄
  await canvasEl(page).last().click()
  await canvasEl(page).last().click({ button: 'right', force: true })
  await page.getByTestId('context-menu').getByText('锁定', { exact: true }).click()
  await expect(page.getByTestId('context-menu')).not.toBeVisible()
  await canvasEl(page).last().click({ force: true })
  await expect(page.locator('.element-handle')).toHaveCount(0)

  expectNoErrors(errors)
})

test('文本富文本编辑：进入编辑、输入、加粗', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)

  // 双击进入编辑
  await canvasEl(page).dblclick()
  const editor = page.locator('.pptx-richtext-editor')
  await expect(editor).toBeVisible()
  await expect(page.getByTestId('text-toolbar')).toBeVisible()
  // prod 构建下挂载快于 tiptap 异步聚焦（autofocus: 'end'），等焦点落定再输入，避免按键丢失
  await expect(editor.locator('.ProseMirror')).toBeFocused()

  // 全选输入新内容
  await page.keyboard.press('ControlOrMeta+a')
  await page.keyboard.type('季度经营汇报')
  await expect(canvasEl(page)).toContainText('季度经营汇报')

  // 加粗 → strong 标签（先重新全选，折叠光标下加粗只作用于后续输入）
  await page.keyboard.press('ControlOrMeta+a')
  await page.getByTestId('text-toolbar').getByTitle('加粗').click()
  await expect(canvasEl(page).locator('strong')).toHaveCount(1)

  // Esc 退出编辑
  await page.keyboard.press('Escape')
  await expect(editor).not.toBeVisible()
  await expect(canvasEl(page)).toContainText('季度经营汇报')

  expectNoErrors(errors)
})

test('表格编辑：单元格输入', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-table').click()
  await expect(canvasEl(page).locator('table')).toHaveCount(1)

  // 单击第一个单元格进入就地编辑（contentEditable）
  await canvasEl(page).locator('td').first().click()
  const cellEditor = canvasEl(page).locator('.pptx-table-cell-editor')
  await expect(cellEditor).toBeVisible()
  await cellEditor.fill('季度')
  // 选中态下样式面板应出现表格设置
  await expect(page.getByTestId('style-panel')).toContainText('表格')
  // 点击画布空白处触发失焦保存
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })
  await expect(canvasEl(page).locator('table')).toContainText('季度')

  expectNoErrors(errors)
})

test('图表：数据编辑、系列增删、切换行列、配色、趋势线、类型切换', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await insertChart(page)
  const panel = page.getByTestId('style-panel')
  const chartSvg = page.locator('[data-testid="canvas-slide"] svg')

  // 修改第一个数值 → 数据编辑器生效
  const firstValue = panel.locator('input[type="number"]').first()
  await firstValue.fill('999')
  await expect(firstValue).toHaveValue('999')

  // 添加系列 → 图例出现 系列3（与默认系列1/系列2 命名一致）
  await panel.getByRole('button', { name: '添加', exact: true }).click()
  await expect(chartSvg.getByText('系列3').first()).toBeVisible()

  // 切换行列 → 图例变为 一月~四月
  await panel.getByRole('button', { name: '切换行列' }).click()
  await expect(chartSvg.getByText('一月').first()).toBeVisible()

  // 配色方案 2 → 首系列变 #3370ff
  await panel.getByTitle('配色方案 2').click()
  await expect(chartSvg.locator('[fill="#3370ff"]').first()).toBeVisible()

  // 趋势线 → 出现虚线
  const trendToggle = panel.locator('label', { hasText: '趋势线' }).locator('button[role="switch"]')
  await trendToggle.click()
  await expect(chartSvg.locator('[stroke-dasharray]').first()).toBeVisible()

  // 类型切换为堆积柱状图不崩溃
  await page.getByTestId('chart-type').selectOption({ label: '堆积柱状图' })
  await expect(canvasEl(page)).toHaveCount(1)

  expectNoErrors(errors)
})

test('页面背景渐变与画布比例', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')

  // 背景改渐变
  await page.getByTestId('bg-type').selectOption({ label: '渐变' })
  const bgHtml = await page.getByTestId('canvas-slide').evaluate((el) => el.innerHTML)
  expect(bgHtml).toContain('gradient')

  // 比例 4:3 → 画布宽高比变化
  const before = await page.getByTestId('canvas-slide').boundingBox()
  await page.getByTestId('canvas-ratio').selectOption({ label: '4 : 3' })
  await page.waitForTimeout(300)
  const after = await page.getByTestId('canvas-slide').boundingBox()
  const ratioAfter = after!.width / after!.height
  expect(Math.abs(ratioAfter - 4 / 3)).toBeLessThan(0.15)
  expect(Math.abs(before!.width / before!.height - ratioAfter)).toBeGreaterThan(0.1)

  // 切回 16:9
  await page.getByTestId('canvas-ratio').selectOption({ label: '16 : 9' })

  expectNoErrors(errors)
})

test('切换动画与元素动画面板', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)

  // 页面切换动画：先点击画布空白处取消选中，面板才显示页面设置区
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })
  await page.getByTestId('transition-effect').selectOption({ label: '淡入淡出' })

  // 元素动画：选中元素 → 动画页签
  await canvasEl(page).click()
  await page.getByTestId('style-panel').getByRole('button', { name: '动画', exact: true }).click()
  await page.getByTestId('anim-type').selectOption({ label: '进入' })
  const effectSelect = page.getByTestId('anim-effect')
  await expect(effectSelect).toBeVisible()
  await effectSelect.selectOption({ index: 0 })

  expectNoErrors(errors)
})

test('放映与演讲者视图', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()

  // 放映
  await page.getByTestId('play').click()
  await expect(page.getByTestId('player')).toBeVisible()
  await page.keyboard.press('ArrowRight')
  await page.keyboard.press('Escape')
  await expect(page.getByTestId('player')).not.toBeVisible()

  // 演讲者视图
  await page.getByTitle('演讲者视图').click()
  await expect(page.getByTestId('presenter')).toBeVisible()
  await expect(page.getByTestId('presenter-notes')).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(page.getByTestId('presenter')).not.toBeVisible()

  expectNoErrors(errors)
})

test('文档管理：新建/副本', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)

  await page.getByTitle('文档管理').click()
  await expect(page.locator('.modal-pop')).toBeVisible()

  // 新建演示文稿（会自动切换并关闭弹窗）
  await page.getByRole('button', { name: '＋ 新建演示文稿' }).click()
  await page.waitForTimeout(600)
  await expect(page.locator('.modal-pop')).not.toBeVisible()
  await expect(canvasEl(page)).toHaveCount(0)

  // 重新打开：文档列表应含原文档；创建副本后数量增加
  await page.getByTitle('文档管理').click()
  await expect(page.locator('.modal-pop')).toBeVisible()
  const docsBefore = await page.locator('button[title="创建副本"]').count()
  await page.getByTitle('创建副本').first().click()
  await page.waitForTimeout(600)
  await expect(page.locator('button[title="创建副本"]').count()).resolves.toBeGreaterThanOrEqual(docsBefore + 1)

  await page.keyboard.press('Escape')
  expectNoErrors(errors)
})

test('主题配色一键应用', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()

  await page.getByTitle('主题配色').click()
  await expect(page.locator('.modal-pop')).toBeVisible()
  await page.getByRole('button', { name: '应用+换背景' }).first().click()
  await page.waitForTimeout(500)

  expectNoErrors(errors)
})

test('公式元素与 LaTeX 编辑对话框', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-formula').click()
  await expect(canvasEl(page)).toHaveCount(1)

  // 选中 → 样式面板公式区 → 点击进入编辑
  await canvasEl(page).click()
  await page.getByTestId('style-panel').getByRole('button').filter({ hasText: 'a^2+b^2=c^2' }).click()
  await expect(page.locator('math-field')).toBeVisible()
  await page.getByRole('button', { name: '确认' }).click()
  await page.waitForTimeout(300)

  expectNoErrors(errors)
})

test('图表导出降级路径：趋势线与雷达图导出为图片', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')

  // 趋势线开启（原生图表不支持 → 元素降级为图片导出）
  await insertChart(page)
  const trendToggle = page.getByTestId('style-panel').locator('label', { hasText: '趋势线' }).locator('button[role="switch"]')
  await trendToggle.click()
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })

  await page.getByTitle('导出').click()
  const dl1 = page.waitForEvent('download', { timeout: 60000 })
  await page.getByTestId('export-pptx').click()
  const pptx = await dl1
  expect(pptx.suggestedFilename()).toContain('.pptx')
  await page.keyboard.press('Escape')

  // 雷达图（pptxgenjs 无原生支持 → 图片导出）
  await insertChart(page, '雷达图', 2)
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })
  await page.getByTitle('导出').click()
  const dl2 = page.waitForEvent('download', { timeout: 60000 })
  await page.getByTestId('export-pptx').click()
  const pptx2 = await dl2
  expect(pptx2.suggestedFilename()).toContain('.pptx')

  expectNoErrors(errors)
})

test('缩略图面板开关与画布网格', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('thumbnail-panel').click()
  await page.getByTestId('thumbnail-panel').click()
  // 网格开关（底栏）
  await page.getByTitle('网格').first().click()
  expectNoErrors(errors)
})

test('对齐排列与组合快捷键', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await page.getByTestId('insert-shape').click()
  await page.locator('button[title="圆形"]').click()
  await expect(canvasEl(page)).toHaveCount(2)

  // 全选后对齐：Ctrl+L 左对齐（画布左缘）、Ctrl+T 顶对齐
  await page.keyboard.press('ControlOrMeta+a')
  await page.keyboard.press('ControlOrMeta+l')
  await page.keyboard.press('ControlOrMeta+t')
  await page.waitForTimeout(200)
  // 两个元素应靠到同一左缘/顶缘（画布坐标一致）
  const boxes = [
    await canvasEl(page).first().boundingBox(),
    await canvasEl(page).last().boundingBox(),
  ]
  expect(Math.abs(boxes[0]!.x - boxes[1]!.x)).toBeLessThan(2)
  expect(Math.abs(boxes[0]!.y - boxes[1]!.y)).toBeLessThan(2)

  // 排列：置顶/置底（右键菜单）
  await canvasEl(page).last().click({ button: 'right', force: true })
  await page.getByTestId('context-menu').getByText('置底').click()
  await expect(page.getByTestId('context-menu')).not.toBeVisible()

  expectNoErrors(errors)
})

test('批注与演讲者备注', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTestId('insert-text').click()
  await expect(canvasEl(page)).toHaveCount(1)
  await canvasEl(page).click()

  // 批注页签：元素批注输入
  await page.getByTestId('style-panel').getByRole('button', { name: '批注', exact: true }).click()
  const commentBox = page.getByTestId('style-panel').locator('textarea').first()
  await expect(commentBox).toBeVisible()
  await commentBox.fill('这里需要配图')
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })

  // 演讲者备注（页面级，分区默认折叠需先展开）
  const noteSection = page.getByTestId('style-panel').getByText('演讲者备注（本页）')
  await noteSection.click()
  const noteBox = page.getByTestId('style-panel').getByPlaceholder('本页备注…')
  await expect(noteBox).toBeVisible()
  await noteBox.fill('开场白：欢迎各位')
  await page.getByTestId('canvas-slide').click({ position: { x: 5, y: 5 } })

  expectNoErrors(errors)
})

test('快捷键帮助对话框', async ({ page }) => {
  const errors = watchErrors(page)
  await page.goto('/')
  await page.getByTitle('快捷键').click()
  await expect(page.locator('.modal-pop')).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(page.locator('.modal-pop')).not.toBeVisible()
  expectNoErrors(errors)
})
