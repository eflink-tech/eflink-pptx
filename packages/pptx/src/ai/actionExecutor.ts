// AI 产物 → 编辑器数据（应用大纲/整套页面到文档）
import { useEditorStore } from '../store/editorStore'
import type { AIGenSlide, OutlineResult } from './types'
import { sanitizeElements } from './systemPrompt'
import { genId } from '../core/utils/id'
import type { ShapeElement, TextElement } from '../types/slides'

/** 把大纲写入当前文档（每页放标题+要点文本，供用户再细化或让 AI 生成版式） */
export function applyOutline(outline: OutlineResult): void {
  const store = useEditorStore.getState()
  store.pushHistory()

  const slides = outline.slides.map((item) => {
    const elements: Array<TextElement> = []
    elements.push({
      id: genId('t-'), type: 'text', x: 100, y: 180, w: 1080, h: 100,
      content: `<p>${escapeHtml(item.title)}</p>`,
      defaultColor: '#d14424', lineHeight: 1.3, padding: 8, name: '标题',
    })
    if (item.points.length) {
      const body = item.points.map((p) => `<p>· ${escapeHtml(p)}</p>`).join('')
      elements.push({
        id: genId('t-'), type: 'text', x: 100, y: 320, w: 1080, h: 280,
        content: body, defaultColor: '#555555', lineHeight: 1.8, padding: 8, name: '正文',
      })
    }
    return { id: genId('slide-'), elements, background: { type: 'solid' as const, color: '#ffffff' } }
  })

  store.replacePresentation({
    ...structuredClone(store.presentation),
    slides,
  })
  useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
}

/** 把 AI 生成的整套页面应用到当前文档 */
export function applyGeneratedDeck(genSlides: AIGenSlide[]): void {
  const store = useEditorStore.getState()
  store.pushHistory()

  const slides = genSlides.map((gen) => ({
    id: genId('slide-'),
    elements: sanitizeElements(gen.elements ?? []).map((el): TextElement | ShapeElement => {
      if (el.type === 'shape') {
        return {
          id: genId('s-'), type: 'shape', x: el.x, y: el.y, w: el.w, h: el.h,
          shapeKey: 'rect', fill: el.fill ?? '#fbeae5',
          outline: { color: '#00000000', width: 1, style: 'solid' },
          align: 'center', valign: 'middle', name: '装饰',
        }
      }
      const isTitle = el.role === 'title'
      return {
        id: genId('t-'), type: 'text', x: el.x, y: el.y, w: el.w, h: el.h,
        content: `<p>${escapeHtml(el.text)}</p>`,
        defaultColor: isTitle ? '#d14424' : '#444444',
        lineHeight: 1.5, padding: 8, name: isTitle ? '标题' : '正文',
      }
    }),
    background: { type: 'solid' as const, color: '#ffffff' },
  }))

  store.replacePresentation({
    ...structuredClone(store.presentation),
    slides,
  })
  useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '</p><p>')
}
