// 查找替换（作用于全部页面的文本元素）
import { useState } from 'react'
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'

function stripHTML(html: string): string {
  const div = document.createElement('div')
  div.innerHTML = html
  return div.textContent ?? ''
}

export function FindReplaceDialog() {
  const { t } = useI18n()
  const [find, setFind] = useState('')
  const [replace, setReplace] = useState('')
  const [hits, setHits] = useState<Array<{ slideIndex: number; elId: string; preview: string }>>([])
  const [searched, setSearched] = useState(false)
  const toast = useToastStore.getState().toast

  const doFind = () => {
    if (!find.trim()) return
    const store = useEditorStore.getState()
    const results: Array<{ slideIndex: number; elId: string; preview: string }> = []
    store.presentation.slides.forEach((slide, slideIndex) => {
      for (const el of slide.elements) {
        if (el.type !== 'text') continue
        const plain = stripHTML(el.content)
        if (plain.includes(find)) {
          results.push({ slideIndex, elId: el.id, preview: plain.slice(Math.max(0, plain.indexOf(find) - 10), plain.indexOf(find) + find.length + 16) })
        }
      }
    })
    setHits(results)
    setSearched(true)
    useUIStore.getState().setFindHits(results.filter((r) => r.slideIndex === store.slideIndex).map((r) => r.elId))
  }

  const gotoHit = (hit: { slideIndex: number; elId: string }) => {
    useEditorStore.setState({ slideIndex: hit.slideIndex, selectedIds: [hit.elId] })
    useUIStore.getState().setFindHits([hit.elId])
  }

  const doReplaceAll = () => {
    if (!find.trim()) return
    const store = useEditorStore.getState()
    let count = 0
    const pres = structuredClone(store.presentation)
    for (const slide of pres.slides) {
      for (const el of slide.elements) {
        if (el.type !== 'text') continue
        if (el.content.includes(find)) {
          count += el.content.split(find).length - 1
          el.content = el.content.replaceAll(find, replace)
        }
      }
    }
    if (count) {
      store.pushHistory()
      store.replacePresentation(pres)
      useEditorStore.setState({ slideIndex: store.slideIndex })
      toast(t('pptx.find.replaced', { count }), 'success')
      doFind()
    } else {
      toast(t('pptx.find.nothingToReplace'), 'error')
    }
  }

  return (
    <Modal title={t('pptx.find.title')}>
      <div className="space-y-3">
        <div className="flex gap-2">
          <input
            autoFocus
            value={find}
            onChange={(e) => setFind(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && doFind()}
            placeholder={t('pptx.find.placeholder')}
            className="flex-1 rounded border border-gray-200 px-2.5 py-1.5 text-sm focus:border-[#d14424] focus:outline-none"
          />
          <button className="rounded bg-[#d14424] px-4 py-1.5 text-sm text-white hover:bg-[#b93a1d]" onClick={doFind}>
            {t('action.find')}
          </button>
        </div>
        <div className="flex gap-2">
          <input
            value={replace}
            onChange={(e) => setReplace(e.target.value)}
            placeholder={t('pptx.find.replacePlaceholder')}
            className="flex-1 rounded border border-gray-200 px-2.5 py-1.5 text-sm focus:border-[#d14424] focus:outline-none"
          />
          <button className="rounded border border-[#d14424] px-4 py-1.5 text-sm text-[#d14424] hover:bg-[#fbeae5]" onClick={doReplaceAll}>
            {t('pptx.find.replaceAll')}
          </button>
        </div>

        {searched && (
          <div className="max-h-56 space-y-1 overflow-y-auto rounded-lg bg-gray-50 p-2">
            {hits.length === 0
              ? <div className="p-2 text-center text-xs text-gray-400">{t('pptx.find.noHits', { query: find })}</div>
              : hits.map((hit, i) => (
                <button
                  key={i}
                  className="block w-full rounded px-2 py-1.5 text-left text-xs hover:bg-white"
                  onClick={() => gotoHit(hit)}
                >
                  <span className="mr-2 text-gray-400">{t('pptx.find.slide', { index: hit.slideIndex + 1 })}</span>
                  <span className="text-gray-600">…{hit.preview}…</span>
                </button>
              ))}
          </div>
        )}
        <div className="text-[11px] text-gray-400">{t('pptx.find.summary', { count: hits.length })}</div>
      </div>
    </Modal>
  )
}

// 供快捷键入口使用
export function openFindReplace(): void {
  useUIStore.getState().openModal('findReplace')
}
