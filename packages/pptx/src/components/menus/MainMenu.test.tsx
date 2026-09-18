// 主菜单组件测试：分类展开 / 二级菜单 / 叶子分类触发
import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, fireEvent, render, screen } from '@testing-library/react'
import { MainMenu } from './MainMenu'
import { useUIStore } from '../../store/uiStore'

// vitest 未开启 globals，testing-library 不会自动注册 cleanup，需显式清理避免用例间 DOM 残留
afterEach(() => {
  cleanup()
})

describe('MainMenu', () => {
  it('点击后展开五个一级分类', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    expect(screen.getByText('文件')).toBeTruthy()
    expect(screen.getByText('设计')).toBeTruthy()
    expect(screen.getByText('视图')).toBeTruthy()
    expect(screen.getByText('放映')).toBeTruthy()
    expect(screen.getByText('快捷键')).toBeTruthy()
  })

  it('悬停文件分类展开二级菜单（新建文档/保存）', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.mouseEnter(screen.getByText('文件'))
    expect(screen.getByText('新建文档')).toBeTruthy()
    expect(screen.getByText('保存')).toBeTruthy()
  })

  it('点击快捷键叶子分类打开快捷键弹窗', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.click(screen.getByText('快捷键'))
    expect(useUIStore.getState().modal).toBe('hotkey')
  })
})
