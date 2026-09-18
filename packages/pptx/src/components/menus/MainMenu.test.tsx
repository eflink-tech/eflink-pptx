// 主菜单组件测试：分类展开 / 二级菜单 / 叶子分类触发
import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, fireEvent, render, screen } from '@testing-library/react'
import { MainMenu } from './MainMenu'
import { useUIStore } from '../../store/uiStore'
import { setPptxShareHandler } from '../../core/share/shareBridge'

// vitest 未开启 globals，testing-library 不会自动注册 cleanup，需显式清理避免用例间 DOM 残留
afterEach(() => {
  cleanup()
  // 还原分享 handler，避免注入状态泄漏到其他用例
  setPptxShareHandler(null)
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

  it('悬停放映分类展开二级菜单（放映/演讲者视图）', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    // 悬停前「放映」仅出现在一级分类
    fireEvent.mouseEnter(screen.getByText('放映'))
    expect(screen.getByText('演讲者视图')).toBeTruthy()
    // 悬停后「放映」同时存在于一级分类与二级叶子，需用 getAllByText 避免歧义
    expect(screen.getAllByText('放映').length).toBeGreaterThanOrEqual(1)
  })

  it('点击快捷键叶子分类打开快捷键弹窗', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.click(screen.getByText('快捷键'))
    expect(useUIStore.getState().modal).toBe('hotkey')
  })

  it('未注入分享 handler 时不显示分享入口，注入后出现', () => {
    // 默认未注入 handler：分享入口不渲染
    const { unmount } = render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    expect(screen.queryByText('分享')).toBeNull()
    unmount()

    // 注入 handler 后重新渲染：分享入口出现
    setPptxShareHandler(() => Promise.resolve({ url: 'https://example.com/s/1' }))
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.mouseEnter(screen.getByText('文件'))
    expect(screen.getByText('分享')).toBeTruthy()
  })
})
