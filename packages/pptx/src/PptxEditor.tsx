// src/PptxEditor.tsx
// 组件入口：复刻原应用启动流程（恢复上次文档 + 订阅本地草稿镜像），
// 宿主渲染 <PptxEditor /> 即获得完整编辑器（顶栏 / 缩略图 / 画布 / 样式面板 / 放映 / AI 助手）。
// 保存策略：内容变化立即写 localStorage 镜像（毫秒级防崩溃，本地草稿，不清除 dirty）；
// 云端保存仅由手动触发（⌘S/Ctrl+S、顶栏保存按钮、分享前强制保存、window bridge save）。
import { useEffect, useState, type JSX } from 'react'
import { AppLayout } from './components/layout/AppLayout'
import { loadStartupDoc, saveDoc } from './core/editor/persistence'
import { useEditorStore } from './store/editorStore'

/** window bridge 契约：宿主/自动化查询脏状态、手动触发云端保存、丢弃本地草稿 */
interface EflinkEditorBridge {
  isDirty: () => boolean
  save: () => Promise<void>
  discard: () => void
}

/**
 * 组件入口：渲染即获得完整编辑器（顶栏 / 缩略图 / 画布 / 样式面板 / 放映 / AI 助手）
 * @param bootDocId 指定启动文档 id（如分享查看页注入的只读快照），优先于本地"上次文档"
 */
export function PptxEditor({ bootDocId }: { bootDocId?: string } = {}): JSX.Element {
  const [ready, setReady] = useState(false)

  useEffect(() => {
    let cancelled = false
    let unsubscribe: (() => void) | undefined
    void (async () => {
      const doc = await loadStartupDoc(bootDocId)
      if (cancelled) return
      useEditorStore.getState().loadDocument(doc)

      // 本地草稿镜像：订阅 dirty 标记，引用比较检测 presentation 变化，
      // 立即同步写 localStorage 镜像（毫秒级，防崩溃丢稿；不调用远端、不清除 dirty）
      let lastPresentation = useEditorStore.getState().presentation
      unsubscribe = useEditorStore.subscribe((state) => {
        if (!state.dirty) return
        if (state.presentation === lastPresentation) return
        lastPresentation = state.presentation
        const { docId, docName } = useEditorStore.getState()
        if (!docId) return // docId 未初始化时不写入（避免 JSON.stringify 丢弃 undefined 字段）
        try {
          localStorage.setItem('eflink-pptx-mirror', JSON.stringify({ id: docId, name: docName, presentation: state.presentation }))
          localStorage.setItem('eflink-pptx-last-doc', docId)
        } catch { /* 存储满等异常忽略 */ }
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
    // bootDocId 仅作为启动参数读取一次，后续变化不重新启动
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // window bridge：挂载后注册，卸载删除（save 为云端保存入口，成功后 dirty=false）
  useEffect(() => {
    if (!ready) return
    const bridge: EflinkEditorBridge = {
      isDirty: () => useEditorStore.getState().dirty,
      save: () => {
        const { docId, docName, presentation, markSaved } = useEditorStore.getState()
        if (!docId) return Promise.resolve()
        return saveDoc(docId, docName, presentation).then(markSaved)
      },
      discard: () => {
        try { localStorage.removeItem('eflink-pptx-mirror') } catch { /* 忽略 */ }
      },
    }
    ;(window as unknown as Record<string, unknown>).__eflinkEditorBridge = bridge
    return () => {
      delete (window as unknown as Record<string, unknown>).__eflinkEditorBridge
    }
  }, [ready])

  if (!ready) {
    return <div className="flex h-full items-center justify-center text-sm text-slate-400">加载中…</div>
  }
  return <AppLayout />
}
