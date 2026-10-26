// 画布右键菜单：元素操作 / 空白操作
import { useEffect, useRef } from 'react'
import { useI18n } from '@eflink-tech/office-kit'
import { useEditorStore } from '../../store/editorStore'

export interface ContextMenuState {
  x: number
  y: number
  /** 命中的元素 id；空白处为 null */
  elementId: string | null
}

export function CanvasContextMenu({ menu, onClose }: { menu: ContextMenuState; onClose: () => void }) {
  const { t } = useI18n()
  const store = useEditorStore.getState()
  const slide = store.presentation.slides[store.slideIndex]
  const el = menu.elementId ? slide?.elements.find((e) => e.id === menu.elementId) : null
  const selectedIds = el
    ? (el.groupId ? slide.elements.filter((x) => x.groupId === el.groupId).map((x) => x.id) : [el.id])
    : store.selectedIds

  // Esc 关闭菜单。onClose 走 ref：监听器只挂载一次，
  // 避免其它 handler（如快捷键改选中）在事件派发中途触发重渲染、解绑监听导致 Escape 失效
  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCloseRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  const run = (fn: () => void) => () => {
    fn()
    onClose()
  }

  const items: Array<{ label: string; action: () => void; danger?: boolean; disabled?: boolean }> = []

  if (el) {
    items.push(
      { label: t('pptx.ctx.copy'), action: run(() => { store.setSelected(selectedIds); useEditorStore.getState().copySelection() }) },
      { label: t('pptx.ctx.duplicate'), action: run(() => { store.setSelected(selectedIds); useEditorStore.getState().duplicateElements(selectedIds) }) },
      { label: t('pptx.arrange.delete'), danger: true, action: run(() => { store.setSelected(selectedIds); store.pushHistory(); useEditorStore.getState().removeElements(selectedIds) }) },
    )
    items.push({ label: 'sep1', action: () => undefined })
    items.push(
      { label: t('pptx.arrange.toFront'), action: run(() => { store.pushHistory(); store.setElementLevel(selectedIds, 'top') }) },
      { label: t('pptx.arrange.moveUp'), action: run(() => { store.pushHistory(); store.setElementLevel(selectedIds, 'up') }) },
      { label: t('pptx.arrange.moveDown'), action: run(() => { store.pushHistory(); store.setElementLevel(selectedIds, 'down') }) },
      { label: t('pptx.arrange.toBack'), action: run(() => { store.pushHistory(); store.setElementLevel(selectedIds, 'bottom') }) },
    )
    items.push({ label: 'sep2', action: () => undefined })
    items.push(
      { label: el.lock ? t('pptx.ctx.unlock') : t('pptx.ctx.lock'), action: run(() => { store.pushHistory(); store.updateElements(selectedIds, (d) => { d.lock = !d.lock }) }) },
      { label: t('pptx.ctx.editText'), disabled: el.type !== 'text' && el.type !== 'shape', action: run(() => store.setEditingId(el.id)) },
    )
    // 已分组的元素显示「取消组合」；未分组的多选才显示「组合」
    if (el.groupId) items.push({ label: t('pptx.arrange.ungroup'), action: run(() => { store.pushHistory(); store.setGroup(selectedIds, undefined) }) })
    else if (selectedIds.length > 1) items.push({ label: t('pptx.arrange.group'), action: run(() => { store.pushHistory(); store.setGroup(selectedIds, `g-${Date.now().toString(36)}`) }) })
  } else {
    items.push(
      { label: t('pptx.ctx.paste'), disabled: !store.clipboard, action: run(() => { store.pushHistory(); store.pasteClipboard() }) },
      { label: t('pptx.ctx.selectAll'), disabled: !slide?.elements.length, action: run(() => store.selectAll()) },
      { label: 'sep3', action: () => undefined },
      { label: t('pptx.thumb.new'), action: run(() => { store.pushHistory(); store.addSlide(store.slideIndex + 1) }) },
    )
  }

  const style: React.CSSProperties = {
    left: Math.min(menu.x, window.innerWidth - 170),
    top: Math.min(menu.y, window.innerHeight - items.length * 30 - 16),
  }

  return (
    <>
      <div className="fixed inset-0 z-[90]" onMouseDown={onClose} onContextMenu={(e) => { e.preventDefault(); onClose() }} />
      <div className="fixed z-[100] w-40 rounded-md border border-gray-200 bg-white py-1 text-sm shadow-xl" style={style} data-testid="context-menu">
        {items.map((item) => item.label.startsWith('sep')
          ? <div key={item.label} className="my-1 h-px bg-gray-100" />
          : (
              <button
                key={item.label}
                disabled={item.disabled}
                className={`block w-full px-3 py-1.5 text-left hover:bg-gray-100 disabled:opacity-30 ${item.danger ? 'text-red-500' : 'text-gray-700'}`}
                onMouseDown={(e) => e.preventDefault()}
                onClick={item.action}
              >
                {item.label}
              </button>
            ))}
      </div>
    </>
  )
}
