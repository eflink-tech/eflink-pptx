// 文档管理对话框：本地多文档（新建/切换/重命名/复制/删除）
import { useEffect, useState } from 'react'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import {
  createDoc, deleteDoc, duplicateDoc, listDocs, loadStartupDoc, renameDoc, saveDoc,
} from '../../core/editor/persistence'

interface DocItem { id: string; name: string; updatedAt: number }

function formatTime(ts: number): string {
  const d = new Date(ts)
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
}

export function FileManagerDialog() {
  const [docs, setDocs] = useState<DocItem[]>([])
  const [renaming, setRenaming] = useState<string | null>(null)
  const currentId = useEditorStore((s) => s.docId)
  const toast = useToastStore.getState().toast
  const closeModal = useUIStore.getState().closeModal

  const refresh = () => {
    void listDocs().then(setDocs)
  }
  useEffect(refresh, [])

  const switchTo = async (id: string) => {
    // 先保存当前
    const store = useEditorStore.getState()
    await saveDoc(store.docId, store.docName, store.presentation)
        localStorage.setItem('eflink-pptx-last-doc', id)
    localStorage.removeItem('eflink-pptx-mirror')
    const doc = await loadStartupDoc()
    store.loadDocument(doc)
    toast(`已切换到「${doc.name}」`, 'success')
    closeModal()
  }

  const onNew = async () => {
    const store = useEditorStore.getState()
    await saveDoc(store.docId, store.docName, store.presentation)
    const doc = await createDoc('未命名演示文稿')
    store.loadDocument(doc)
    toast('已新建文档', 'success')
    closeModal()
  }

  const onCopy = async (id: string) => {
    const copy = await duplicateDoc(id)
    if (copy) {
      toast('已创建副本', 'success')
      refresh()
    }
  }

  const onDelete = async (id: string) => {
    if (docs.length <= 1 && id === currentId) {
      toast('至少保留一个文档', 'error')
      return
    }
    await deleteDoc(id)
    if (id === currentId) {
      const docsLeft = await listDocs()
      if (docsLeft.length) {
        localStorage.setItem('eflink-pptx-last-doc', docsLeft[0].id)
        localStorage.removeItem('eflink-pptx-mirror')
                const doc = await loadStartupDoc()
        useEditorStore.getState().loadDocument(doc)
      }
    }
    toast('已删除', 'success')
    refresh()
  }

  const onRename = async (id: string, name: string) => {
    await renameDoc(id, name)
    if (id === currentId) useEditorStore.getState().renameDocument(name)
    setRenaming(null)
    refresh()
  }

  return (
    <Modal title="文档管理" width={560}>
      <button
        className="mb-3 w-full rounded-lg border-2 border-dashed border-gray-300 py-2.5 text-sm text-gray-500 hover:border-[#d14424] hover:text-[#d14424]"
        onClick={() => void onNew()}
        data-testid="file-new"
      >
        ＋ 新建演示文稿
      </button>
      <div className="space-y-1.5">
        {docs.length === 0 && <div className="py-6 text-center text-xs text-gray-400">暂无历史文档（当前文档保存后出现在这里）</div>}
        {docs.map((doc) => (
          <div key={doc.id} className={`flex items-center gap-2 rounded-lg border px-3 py-2 ${doc.id === currentId ? 'border-[#d14424] bg-[#fbeae5]/40' : 'border-gray-200'}`}>
            {renaming === doc.id
              ? (
                  <input
                    autoFocus
                    defaultValue={doc.name}
                    className="flex-1 rounded border border-gray-300 px-2 py-1 text-sm"
                    onBlur={(e) => void onRename(doc.id, e.target.value.trim() || doc.name)}
                    onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
                  />
                )
              : (
                  <button className="flex-1 truncate text-left text-sm text-gray-700" onClick={() => void switchTo(doc.id)} title="点击切换">
                    {doc.name}
                    {doc.id === currentId && <span className="ml-2 text-xs text-[#d14424]">当前</span>}
                  </button>
                )}
            <span className="shrink-0 text-xs text-gray-400">{formatTime(doc.updatedAt)}</span>
            <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title="重命名" onClick={() => setRenaming(doc.id)}>✎</button>
            <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title="创建副本" onClick={() => void onCopy(doc.id)}>⧉</button>
            <button className="rounded p-1 text-gray-400 hover:bg-red-50 hover:text-red-500" title="删除" onClick={() => void onDelete(doc.id)}>🗑</button>
          </div>
        ))}
      </div>
    </Modal>
  )
}
