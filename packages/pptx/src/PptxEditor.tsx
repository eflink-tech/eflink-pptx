// src/PptxEditor.tsx
// 组件入口：复刻原应用启动流程（恢复上次文档 + 订阅自动保存），
// 宿主渲染 <PptxEditor /> 即获得完整编辑器（顶栏 / 缩略图 / 画布 / 样式面板 / 放映 / AI 助手）。
import { useEffect, useState, type JSX } from 'react'
import { AppLayout } from './components/layout/AppLayout'
import { loadStartupDoc, saveDoc, scheduleAutosave } from './core/editor/persistence'
import { useEditorStore } from './store/editorStore'

export function PptxEditor(): JSX.Element {
  const [ready, setReady] = useState(false)

  useEffect(() => {
    let cancelled = false
    let unsubscribe: (() => void) | undefined
    void (async () => {
      const doc = await loadStartupDoc()
      if (cancelled) return
      useEditorStore.getState().loadDocument(doc)

      // 自动保存：订阅 dirty 标记，引用比较检测 presentation 变化，立即写 localStorage 镜像，Dexie 节流写
      let lastPresentation = useEditorStore.getState().presentation
      unsubscribe = useEditorStore.subscribe((state) => {
        if (!state.dirty) return
        if (state.presentation === lastPresentation) return
        lastPresentation = state.presentation
        const { docId, docName, markSaved } = useEditorStore.getState()
        if (!docId) return // docId 未初始化时不写入（避免 JSON.stringify 丢弃 undefined 字段）
        // localStorage 镜像：立即同步写入，确保刷新可恢复
        try {
          localStorage.setItem('eflink-pptx-mirror', JSON.stringify({ id: docId, name: docName, presentation: state.presentation }))
          localStorage.setItem('eflink-pptx-last-doc', docId)
        } catch { /* 存储满等异常忽略 */ }
        // Dexie：节流写入（持久化 + 多文档管理）
        scheduleAutosave(() => {
          saveDoc(docId, docName, useEditorStore.getState().presentation).then(markSaved)
        })
      })
      setReady(true)
    })().catch((err) => {
      console.error('启动编辑器失败', err)
      if (!cancelled) setReady(true)
    })

    return () => {
      cancelled = true
      unsubscribe?.()
    }
  }, [])

  if (!ready) {
    return <div className="flex h-full items-center justify-center text-sm text-slate-400">加载中…</div>
  }
  return <AppLayout />
}
