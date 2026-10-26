// 导出对话框：PPTX / 图片 / JSON / 打印
import { useState } from 'react'
import { Loader2 } from 'lucide-react'
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { exportJSON } from '../../core/export/json'
import { exportImages, renderSlideToBlob } from '../../core/export/image'
import { exportPPTX } from '../../core/export/pptx'

/** 执行中按钮的标签：spinner +「导出中」+ 进度（如有） */
function BusyLabel({ progress }: { progress?: string }) {
  const { t } = useI18n()
  return (
    <span className="inline-flex items-center gap-1.5">
      <Loader2 size={14} className="animate-spin" />
      {t('pptx.export.busy')}{progress ? ` (${progress})` : '…'}
    </span>
  )
}

export function ExportDialog() {
  const { t } = useI18n()
  const [busy, setBusy] = useState('')
  const [progress, setProgress] = useState('')
  const toast = useToastStore.getState().toast
  const closeModal = useUIStore.getState().closeModal

  const printPDF = async () => {
    const store = useEditorStore.getState()
    const imgs: string[] = []
    for (let i = 0; i < store.presentation.slides.length; i++) {
      setProgress(`${i + 1}/${store.presentation.slides.length}`)
      const blob = await renderSlideToBlob(store.presentation.slides[i], store.presentation, 'jpeg')
      imgs.push(URL.createObjectURL(blob))
    }
    const win = window.open('', '_blank')
    if (!win) {
      toast(t('pptx.export.popupBlocked'), 'error')
      return
    }
    win.document.write(`<!DOCTYPE html><html><head><title>${store.docName}</title><style>
      body{margin:0}img{display:block;width:100vw;height:auto;page-break-after:always}
      @page{size:landscape;margin:0}
    </style></head><body>${imgs.map((src) => `<img src="${src}">`).join('')}</body></html>`)
    win.document.close()
    win.onload = () => win.print()
    setTimeout(() => win.print(), 600)
  }

  const run = async (key: string, fn: () => Promise<void> | void) => {
    setBusy(key)
    try {
      await fn()
      if (key !== 'pptx') toast(t('pptx.export.done'), 'success')
    } catch (error) {
      toast(t('pptx.export.failed', { message: error instanceof Error ? error.message : t('pptx.export.unknownError') }), 'error')
    } finally {
      setBusy('')
      setProgress('')
    }
  }

  return (
    <Modal title={t('pptx.export.title')}>
      <div className="space-y-2.5">
        <button
          className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#d14424] disabled:opacity-50"
          disabled={Boolean(busy)}
          onClick={() => void run('pptx', async () => {
            const store = useEditorStore.getState()
            await exportPPTX(store.presentation, store.docName, (done, total) => setProgress(`${done}/${total}`))
          })}
          data-testid="export-pptx"
        >
          <div className="flex items-center gap-2 text-sm font-medium text-gray-700">
            {busy === 'pptx' ? <BusyLabel progress={progress} /> : t('pptx.export.pptx')}
          </div>
          <div className="mt-0.5 text-xs text-gray-400">{t('pptx.export.pptxDesc')}</div>
        </button>

        <div className="rounded-lg border border-gray-200 px-4 py-3">
          <div className="text-sm font-medium text-gray-700">{t('pptx.export.images')}</div>
          <div className="mt-0.5 mb-2 text-xs text-gray-400">{t('pptx.export.imagesDesc')}</div>
          <div className="flex gap-2">
            <button
              className="rounded border border-gray-200 px-3 py-1 text-xs hover:bg-gray-50 disabled:opacity-50"
              disabled={Boolean(busy)}
              onClick={() => void run('png', () => exportImages(
                useEditorStore.getState().presentation.slides,
                useEditorStore.getState().presentation,
                'png',
                useEditorStore.getState().docName,
                (done, total) => setProgress(`${done}/${total}`),
              ))}
            >
              {busy === 'png' ? <BusyLabel progress={progress} /> : 'PNG'}
            </button>
            <button
              className="rounded border border-gray-200 px-3 py-1 text-xs hover:bg-gray-50 disabled:opacity-50"
              disabled={Boolean(busy)}
              onClick={() => void run('jpeg', () => exportImages(
                useEditorStore.getState().presentation.slides,
                useEditorStore.getState().presentation,
                'jpeg',
                useEditorStore.getState().docName,
              ))}
            >
              JPG
            </button>
            <button
              className="rounded border border-gray-200 px-3 py-1 text-xs hover:bg-gray-50 disabled:opacity-50"
              disabled={Boolean(busy)}
              onClick={() => void run('cur', async () => {
                const store = useEditorStore.getState()
                await exportImages([store.presentation.slides[store.slideIndex]], store.presentation, 'png', store.docName)
              })}
            >
              {t('pptx.export.current')}
            </button>
          </div>
        </div>

        <button
          className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#d14424] disabled:opacity-50"
          disabled={Boolean(busy)}
          onClick={() => run('json', async () => {
            const store = useEditorStore.getState()
            exportJSON(store.presentation, store.docName)
          })}
        >
          <div className="text-sm font-medium text-gray-700">{t('pptx.export.json')}</div>
          <div className="mt-0.5 text-xs text-gray-400">{t('pptx.export.jsonDesc')}</div>
        </button>

        <button
          className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#d14424] disabled:opacity-50"
          disabled={Boolean(busy)}
          onClick={() => void run('pdf', async () => {
            await printPDF()
            closeModal()
          })}
        >
          <div className="flex items-center gap-2 text-sm font-medium text-gray-700">
            {busy === 'pdf' ? <BusyLabel progress={progress} /> : t('pptx.export.print')}
          </div>
          <div className="mt-0.5 text-xs text-gray-400">{t('pptx.export.printDesc')}</div>
        </button>
      </div>
    </Modal>
  )
}
