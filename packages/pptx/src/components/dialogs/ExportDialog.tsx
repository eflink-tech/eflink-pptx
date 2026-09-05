// 导出对话框：PPTX / 图片 / JSON / 打印
import { useState } from 'react'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { exportJSON } from '../../core/export/json'
import { exportImages, renderSlideToBlob } from '../../core/export/image'
import { exportPPTX } from '../../core/export/pptx'

export function ExportDialog() {
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
      toast('打印窗口被拦截，请允许弹出窗口', 'error')
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
      if (key !== 'pptx') toast('导出完成', 'success')
    } catch (error) {
      toast(`导出失败：${error instanceof Error ? error.message : '未知错误'}`, 'error')
    } finally {
      setBusy('')
      setProgress('')
    }
  }

  return (
    <Modal title="导出">
      <div className="space-y-2.5">
        <button
          className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#d14424] disabled:opacity-50"
          disabled={Boolean(busy)}
          onClick={() => void run('pptx', () => exportPPTX(useEditorStore.getState().presentation, useEditorStore.getState().docName))}
          data-testid="export-pptx"
        >
          <div className="text-sm font-medium text-gray-700">导出 PPTX</div>
          <div className="mt-0.5 text-xs text-gray-400">文本/形状/图片/表格/图表转为原生 PPT 对象，可用 PowerPoint / WPS 打开</div>
        </button>

        <div className="rounded-lg border border-gray-200 px-4 py-3">
          <div className="text-sm font-medium text-gray-700">导出图片</div>
          <div className="mt-0.5 mb-2 text-xs text-gray-400">多页时打包为 zip</div>
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
              PNG{busy === 'png' && progress ? ` (${progress})` : ''}
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
              仅当前页
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
          <div className="text-sm font-medium text-gray-700">导出 JSON</div>
          <div className="mt-0.5 text-xs text-gray-400">完整工程数据，可在本工具重新导入继续编辑</div>
        </button>

        <button
          className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#d14424] disabled:opacity-50"
          disabled={Boolean(busy)}
          onClick={() => void run('pdf', async () => {
            await printPDF()
            closeModal()
          })}
        >
          <div className="text-sm font-medium text-gray-700">打印 / 导出 PDF{busy === 'pdf' && progress ? ` (${progress})` : ''}</div>
          <div className="mt-0.5 text-xs text-gray-400">通过浏览器打印对话框另存为 PDF</div>
        </button>
      </div>
    </Modal>
  )
}
