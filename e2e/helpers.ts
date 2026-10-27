import type { Page } from '@playwright/test'

/**
 * 顶栏菜单操作。office-kit OfficeMenuBar 的一级按钮与下拉动作项
 * 都是显式 role="menuitem"（不能用 getByRole('button')）；
 * 带 checked 态的动作项（网格/缩略图等开关）role 是 "menuitemcheckbox"。
 * name 默认子串匹配，可省略尾部的「…」。
 */
export async function menuClick(page: Page, group: string, leaf?: string): Promise<void> {
  await page.getByRole('menuitem', { name: group }).click()
  if (leaf) {
    await page
      .getByRole('menuitem', { name: leaf })
      .or(page.getByRole('menuitemcheckbox', { name: leaf }))
      .click()
  }
}
