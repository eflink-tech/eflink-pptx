// 顶部工具栏
import { useRef, useState } from 'react'
import logoUrl from '../../assets/pptx-eflink-logo.png'
import {
  Undo2, Redo2, MonitorPlay, MonitorSpeaker, Sparkles, Grid3x3, LayoutTemplate,
  Palette, Upload, Download, FolderOpen, FilePlus2, Save, PanelLeft, Keyboard, Search, ArrowLeft,
} from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { createDoc, saveDoc } from '../../core/editor/persistence'
import { getEditorBackHref } from '../../core/editor/chrome'
import { insertImageFile } from '../../core/editor/media'
import { InsertMenu } from '../menus/InsertMenu'

function ToolButton({ icon, label, onClick, disabled, active }: {
  icon: React.ReactNode
  label: string
  onClick: () => void
  disabled?: boolean
  active?: boolean
}) {
  return (
    <button
      className={`flex items-center gap-1 rounded-md px-2 py-1.5 text-xs transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${
        active ? 'bg-[#fbeae5] text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'
      }`}
      title={label}
      onClick={onClick}
      disabled={disabled}
    >
      {icon}
    </button>
  )
}

export function TopBar() {
  const undoDepth = useEditorStore((s) => s.history.length)
  const redoDepth = useEditorStore((s) => s.future.length)
  const docName = useEditorStore((s) => s.docName)
  const ui = useUIStore()
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [nameEditing, setNameEditing] = useState(false)
  const backHref = getEditorBackHref()

  return (
    <div className="flex h-12 shrink-0 items-center gap-1 border-b border-gray-200 bg-white px-3" data-testid="topbar">
      {backHref && (
        <a
          href={backHref}
          title="返回"
          className="mr-1 flex size-8 shrink-0 items-center justify-center rounded-md text-gray-500 transition-colors hover:bg-gray-100 hover:text-gray-700"
        >
          <ArrowLeft size={18} />
        </a>
      )}
      <img src={logoUrl} alt="易飞演示文稿" className="mr-1.5 h-8 w-8 rounded-full" />
      <span className="mr-2 text-base font-bold text-[#d14424]">易飞演示文稿</span>

      <ToolButton icon={<FilePlus2 size={17} />} label="新建文档" onClick={async () => {
        const doc = await createDoc('未命名演示文稿')
        useEditorStore.getState().loadDocument(doc)
      }} />
      <ToolButton icon={<FolderOpen size={17} />} label="文档管理" onClick={() => ui.openModal('fileManager')} />
      <ToolButton icon={<Save size={17} />} label="保存（Ctrl+S）" onClick={() => {
        const s = useEditorStore.getState()
        if (!s.docId) { useToastStore.getState().toast('文档未初始化，无法保存', 'error'); return }
        saveDoc(s.docId, s.docName, s.presentation).then(s.markSaved)
        useToastStore.getState().toast('已保存', 'success')
      }} />

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <ToolButton icon={<Undo2 size={17} />} label="撤销" disabled={undoDepth === 0} onClick={() => useEditorStore.getState().undo()} />
      <ToolButton icon={<Redo2 size={17} />} label="重做" disabled={redoDepth === 0} onClick={() => useEditorStore.getState().redo()} />

      <div className="mx-1 h-6 w-px bg-gray-200" />
      <div className="flex-1" />

      <InsertMenu />

      <div className="flex-1" />
      <div className="mx-1 h-6 w-px bg-gray-200" />

      <ToolButton icon={<LayoutTemplate size={17} />} label="模板库" onClick={() => ui.openModal('template')} />
      <ToolButton icon={<Palette size={17} />} label="主题配色" onClick={() => ui.openModal('theme')} />
      <ToolButton icon={<Upload size={17} />} label="导入（PPTX/JSON）" onClick={() => ui.openModal('import')} />
      <ToolButton icon={<Download size={17} />} label="导出" onClick={() => ui.openModal('export')} />
      <ToolButton icon={<Search size={17} />} label="查找替换（Ctrl+F）" onClick={() => ui.openModal('findReplace')} />

      <div className="mx-1 h-6 w-px bg-gray-200" />
      <ToolButton icon={<Keyboard size={17} />} label="快捷键" onClick={() => ui.openModal('hotkey')} />

      <input
        type="file"
        ref={fileInputRef}
        accept="image/*"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) insertImageFile(file)
          e.target.value = ''
        }}
      />

      <button
        className={`flex items-center gap-1 rounded-md px-2.5 py-1.5 text-xs ${ui.aiPanelVisible ? 'bg-[#d14424] text-white' : 'bg-[#fbeae5] text-[#d14424] hover:bg-[#f6d9d0]'}`}
        title="AI 助手"
        onClick={() => ui.toggleAIPanel()}
        data-testid="ai-toggle"
      >
        <Sparkles size={15} />
      </button>

      <div className="mx-1 h-6 w-px bg-gray-200" />
      <ToolButton icon={<Grid3x3 size={17} />} label="网格" active={ui.gridVisible} onClick={() => ui.toggleGrid()} />
      <ToolButton icon={<PanelLeft size={17} />} label="缩略图面板" active={ui.thumbnailsVisible} onClick={() => ui.toggleThumbnails()} />

      <button
        className="ml-1 flex items-center gap-1 rounded-md bg-[#d14424] px-3 py-1.5 text-xs text-white hover:bg-[#b93a1d]"
        title="放映"
        onClick={() => ui.setPlayerMode('playing', useEditorStore.getState().slideIndex)}
        data-testid="play"
      >
        <MonitorPlay size={15} />
      </button>
      <button
        className="ml-1 flex items-center gap-1 rounded-md border border-[#d14424] px-2.5 py-1.5 text-xs text-[#d14424] hover:bg-[#fbeae5]"
        onClick={() => ui.setPlayerMode('presenter', useEditorStore.getState().slideIndex)}
        title="演讲者视图"
        data-testid="play-presenter"
      >
        <MonitorSpeaker size={15} />
      </button>

      {/* 文档名 */}
      <div className="ml-2 max-w-[180px] truncate text-xs text-gray-500" data-testid="doc-name"
        onDoubleClick={() => setNameEditing(true)}
      >
        {nameEditing
          ? (
              <input
                autoFocus
                className="w-36 rounded border border-gray-300 px-1 text-xs"
                defaultValue={docName}
                onBlur={(e) => { useEditorStore.getState().renameDocument(e.target.value.trim() || docName); setNameEditing(false) }}
                onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
              />
            )
          : docName}
      </div>
    </div>
  )
}
