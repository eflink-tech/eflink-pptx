// 全局快捷键
import { useEffect } from 'react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import { saveDoc } from './persistence'
import { alignElements, type AlignMode } from '../utils/align'

function isTextTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null
  if (!el) return false
  return Boolean(
    el.closest?.('.pptx-richtext-editor')
    || el.tagName === 'INPUT'
    || el.tagName === 'TEXTAREA'
    || el.isContentEditable,
  )
}

/** 方向键微调 */
function nudge(dx: number, dy: number): void {
  const store = useEditorStore.getState()
  if (!store.selectedIds.length) return
  store.pushHistory()
  store.updateElements(store.selectedIds, (el) => {
    el.x += dx
    el.y += dy
  })
}

/** Ctrl+方向 对齐快捷映射 */
function alignBy(mode: AlignMode): void {
  const store = useEditorStore.getState()
  if (!store.selectedIds.length) return
  store.pushHistory()
  // 必须在 immer producer 内修改，直接改冻结快照会抛只读错误
  useEditorStore.setState((s) => {
    const slide = s.presentation.slides[s.slideIndex]
    alignElements(slide.elements, store.selectedIds, mode, store.presentation.width, Math.round(store.presentation.width / store.presentation.viewportRatio))
  })
}

export function useGlobalHotkeys(): void {
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const ui = useUIStore.getState()
      const store = useEditorStore.getState()
      const inText = isTextTarget(e.target)
      const mod = e.ctrlKey || e.metaKey

      // Esc 优先级最高
      if (e.key === 'Escape') {
        if (ui.modal) { ui.closeModal(); return }
        if (ui.playerMode !== 'off') return // 放映组件自行处理
        if (store.editingId) { store.setEditingId(null); return }
        if (store.selectedIds.length) { store.setSelected([]); return }
        return
      }

      if (ui.playerMode !== 'off') return
      if (ui.modal) return

      // 放映快捷键：F5 从当前页放映，Shift+F5 从头放映
      if (e.key === 'F5') {
        e.preventDefault()
        ui.setPlayerMode('playing', e.shiftKey ? 0 : store.slideIndex)
        return
      }

      // 保存（无论是否在输入态都拦截）
      if (mod && (e.key === 's' || e.key === 'S')) {
        e.preventDefault()
        if (store.docId) saveDoc(store.docId, store.docName, store.presentation).then(store.markSaved)
        return
      }

      // 查找替换
      if (mod && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault()
        ui.openModal('findReplace')
        return
      }

      if (inText) {
        // 编辑文本时仅保留撤销/重做交给 TipTap
        return
      }

      if (mod && !e.shiftKey && e.key.toLowerCase() === 'z') { e.preventDefault(); store.undo(); return }
      if ((mod && e.shiftKey && e.key.toLowerCase() === 'z') || (mod && e.key.toLowerCase() === 'y')) { e.preventDefault(); store.redo(); return }
      if (mod && e.key.toLowerCase() === 'c') {
        e.preventDefault()
        if (store.selectedIds.length) store.copySelection()
        else store.copyCurrentSlide()
        return
      }
      if (mod && e.key.toLowerCase() === 'x') {
        e.preventDefault()
        if (store.selectedIds.length) { store.pushHistory(); store.cutSelection() }
        return
      }
      if (mod && e.key.toLowerCase() === 'v') {
        e.preventDefault()
        if (!store.clipboard && !store.slideClipboard) return
        // 粘贴入画布/页面属于可撤销变更，先推历史
        store.pushHistory()
        if (!store.selectedIds.length && store.slideClipboard) store.pasteSlide()
        else store.pasteClipboard()
        return
      }
      if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); if (store.selectedIds.length) { store.pushHistory(); store.duplicateElements(store.selectedIds) } return }
      if (mod && e.key.toLowerCase() === 'a') { e.preventDefault(); store.selectAll(); return }
      if (mod && e.key.toLowerCase() === 'g') {
        e.preventDefault()
        if (store.selectedIds.length > 1) store.setGroup(store.selectedIds, `g-${Date.now().toString(36)}`)
        else if (store.selectedIds.length === 1) {
          const el = store.presentation.slides[store.slideIndex].elements.find((x) => x.id === store.selectedIds[0])
          if (el?.groupId) store.setGroup(store.selectedIds, undefined)
        }
        return
      }

      switch (e.key) {
        case 'Delete':
        case 'Backspace':
          if (store.selectedIds.length) {
            e.preventDefault()
            store.pushHistory()
            store.removeElements(store.selectedIds)
          }
          break
        case 'ArrowLeft': e.preventDefault(); nudge(e.shiftKey ? -10 : -1, 0); break
        case 'ArrowRight': e.preventDefault(); nudge(e.shiftKey ? 10 : 1, 0); break
        case 'ArrowUp': {
          e.preventDefault()
          if (store.selectedIds.length) { nudge(0, e.shiftKey ? -10 : -1) }
          else if (store.slideIndex > 0) { useEditorStore.setState({ slideIndex: store.slideIndex - 1, selectedIds: [] }) }
          break
        }
        case 'ArrowDown': {
          e.preventDefault()
          if (store.selectedIds.length) { nudge(0, e.shiftKey ? 10 : 1) }
          else if (store.slideIndex < store.presentation.slides.length - 1) { useEditorStore.setState({ slideIndex: store.slideIndex + 1, selectedIds: [] }) }
          break
        }
        case 'PageUp': {
          e.preventDefault()
          const s = useEditorStore.getState()
          if (s.slideIndex > 0) { s.pushHistory(); useEditorStore.setState({ slideIndex: s.slideIndex - 1, selectedIds: [] }) }
          break
        }
        case 'PageDown': {
          e.preventDefault()
          const s = useEditorStore.getState()
          if (s.slideIndex < s.presentation.slides.length - 1) { s.pushHistory(); useEditorStore.setState({ slideIndex: s.slideIndex + 1, selectedIds: [] }) }
          break
        }
        case 'l': case 'L': if (mod && store.selectedIds.length) { e.preventDefault(); alignBy('left') } break
        case 'r': case 'R': if (mod && store.selectedIds.length) { e.preventDefault(); alignBy('right') } break
        case 't': case 'T': if (mod && store.selectedIds.length) { e.preventDefault(); alignBy('top') } break
        case 'b': case 'B': if (mod && store.selectedIds.length) { e.preventDefault(); alignBy('bottom') } break
        default: break
      }
    }

    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])
}
