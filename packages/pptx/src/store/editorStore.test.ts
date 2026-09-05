import { beforeEach, describe, expect, it } from 'vitest'
import { useEditorStore } from './editorStore'
import { createPresentation } from '../types/slides'
import type { PPTElement } from '../types/slides'

function textEl(id: string, x = 0, y = 0): PPTElement {
  return { id, type: 'text', x, y, w: 100, h: 50, content: '<p>x</p>' }
}

function freshState() {
  useEditorStore.setState({
    docId: 'doc-test',
    docName: '测试',
    presentation: createPresentation('s0'),
    slideIndex: 0,
    selectedIds: [],
    editingId: null,
    history: [],
    future: [],
    clipboard: null,
    dirty: false,
  })
}

describe('editorStore 历史记录', () => {
  beforeEach(freshState)

  it('pushHistory 后 undo 恢复、redo 重做', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a'))
    s.pushHistory()
    s.addElement(textEl('b'))
    expect(useEditorStore.getState().presentation.slides[0].elements).toHaveLength(2)

    useEditorStore.getState().undo()
    expect(useEditorStore.getState().presentation.slides[0].elements).toHaveLength(1)

    useEditorStore.getState().redo()
    expect(useEditorStore.getState().presentation.slides[0].elements).toHaveLength(2)
  })

  it('pushHistory 清空 redo 栈', () => {
    const s = useEditorStore.getState()
    s.pushHistory()
    s.addElement(textEl('a'))
    s.undo()
    s.pushHistory()
    expect(useEditorStore.getState().future).toHaveLength(0)
  })

  it('undo 后页索引收敛', () => {
    const s = useEditorStore.getState()
    s.addSlide()
    s.addSlide()
    useEditorStore.setState({ slideIndex: 2 })
    s.undo()
    const st = useEditorStore.getState()
    expect(st.slideIndex).toBeLessThan(st.presentation.slides.length)
  })
})

describe('editorStore 元素操作', () => {
  beforeEach(freshState)

  it('duplicateElements 生成新 id 并偏移位置', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a', 100, 100))
    const ids = s.duplicateElements(['a'])
    const st = useEditorStore.getState()
    expect(st.presentation.slides[0].elements).toHaveLength(2)
    const copy = st.presentation.slides[0].elements.find((el) => el.id === ids[0])
    expect(copy?.x).toBe(124)
    expect(ids[0]).not.toBe('a')
  })

  it('setElementLevel 置顶置底', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a'))
    s.addElement(textEl('b'))
    s.addElement(textEl('c'))
    s.setElementLevel(['a'], 'top')
    let els = useEditorStore.getState().presentation.slides[0].elements
    expect(els[els.length - 1].id).toBe('a')
    s.setElementLevel(['a'], 'bottom')
    els = useEditorStore.getState().presentation.slides[0].elements
    expect(els[0].id).toBe('a')
  })

  it('setElementLevel 上移一层/下移一层', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a'))
    s.addElement(textEl('b'))
    s.addElement(textEl('c'))
    // a b c → 上移 b → a c b
    s.setElementLevel(['b'], 'up')
    expect(useEditorStore.getState().presentation.slides[0].elements.map((e) => e.id)).toEqual(['a', 'c', 'b'])
    // 下移 b → a b c
    s.setElementLevel(['b'], 'down')
    expect(useEditorStore.getState().presentation.slides[0].elements.map((e) => e.id)).toEqual(['a', 'b', 'c'])
    // 已在顶层再上移不变
    s.setElementLevel(['c'], 'up')
    expect(useEditorStore.getState().presentation.slides[0].elements.map((e) => e.id)).toEqual(['a', 'b', 'c'])
    // 已在底层再下移不变
    s.setElementLevel(['a'], 'down')
    expect(useEditorStore.getState().presentation.slides[0].elements.map((e) => e.id)).toEqual(['a', 'b', 'c'])
  })

  it('组合与取消组合', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a'))
    s.addElement(textEl('b'))
    s.setGroup(['a', 'b'], 'g1')
    const els = useEditorStore.getState().presentation.slides[0].elements
    expect(els.every((el) => el.groupId === 'g1')).toBe(true)
    s.setGroup(['a', 'b'], undefined)
    expect(useEditorStore.getState().presentation.slides[0].elements.every((el) => !el.groupId)).toBe(true)
  })

  it('复制粘贴产生偏移副本', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a', 10, 10))
    useEditorStore.setState({ selectedIds: ['a'] })
    s.copySelection()
    s.pasteClipboard()
    const els = useEditorStore.getState().presentation.slides[0].elements
    expect(els).toHaveLength(2)
    expect(els[1].x).toBe(34)
  })
})

describe('editorStore 页面操作', () => {
  beforeEach(freshState)

  it('addSlide/deleteSlides/moveSlide', () => {
    const s = useEditorStore.getState()
    s.addSlide()
    s.addSlide()
    expect(useEditorStore.getState().presentation.slides).toHaveLength(3)
    expect(useEditorStore.getState().slideIndex).toBe(2)

    s.moveSlide(2, 0)
    expect(useEditorStore.getState().slideIndex).toBe(0)

    s.deleteSlides([0])
    expect(useEditorStore.getState().presentation.slides).toHaveLength(2)
  })

  it('至少保留一页：全选删除不生效', () => {
    const s = useEditorStore.getState()
    s.deleteSlides([0])
    expect(useEditorStore.getState().presentation.slides).toHaveLength(1)
  })

  it('copySlide 复制到当前页之后', () => {
    const s = useEditorStore.getState()
    s.addElement(textEl('a'))
    s.copySlide(0)
    const st = useEditorStore.getState()
    expect(st.presentation.slides).toHaveLength(2)
    expect(st.slideIndex).toBe(1)
    expect(st.presentation.slides[1].elements[0].id).not.toBe('a')
  })
})
