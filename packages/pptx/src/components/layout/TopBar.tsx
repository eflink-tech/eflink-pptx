// 顶部工具栏：品牌区 + 主菜单 + 高频图标（撤销/重做/插入/AI/放映）
// 文件类与低频功能入口已收进 MainMenu（文件/设计/视图/放映/快捷键）
import { ArrowLeft, Undo2, Redo2, MonitorPlay, Sparkles } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import { getEditorBackHref } from '../../core/editor/chrome'
import { InsertMenu } from '../menus/InsertMenu'
import { MainMenu } from '../menus/MainMenu'
import logoUrl from '../../assets/pptx-eflink-logo.png'

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
  const ui = useUIStore()
  const backHref = getEditorBackHref()

  return (
    <div className="relative flex h-12 shrink-0 items-center gap-1 border-b border-gray-200 bg-white px-3" data-testid="topbar">
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

      <MainMenu />

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <ToolButton icon={<Undo2 size={17} />} label="撤销" disabled={undoDepth === 0} onClick={() => useEditorStore.getState().undo()} />
      <ToolButton icon={<Redo2 size={17} />} label="重做" disabled={redoDepth === 0} onClick={() => useEditorStore.getState().redo()} />

      <div className="mx-1 h-6 w-px bg-gray-200" />

      {/* 插入工具栏绝对定位水平居中：不依赖两侧按钮组宽度（flex-1 撑开只保证居中于两侧组之间） */}
      <div className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2">
        <InsertMenu />
      </div>

      <div className="ml-auto flex items-center gap-1">
        <div className="mx-1 h-6 w-px bg-gray-200" />

        <button
        className={`flex items-center gap-1 rounded-md px-2.5 py-1.5 text-xs ${ui.aiPanelVisible ? 'bg-[#d14424] text-white' : 'bg-[#fbeae5] text-[#d14424] hover:bg-[#f6d9d0]'}`}
        title="AI 助手"
        onClick={() => ui.toggleAIPanel()}
        data-testid="ai-toggle"
      >
        <Sparkles size={15} />
      </button>

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <button
        className="ml-1 flex items-center gap-1 rounded-md bg-[#d14424] px-3 py-1.5 text-xs text-white hover:bg-[#b93a1d]"
        title="放映"
        onClick={() => ui.setPlayerMode('playing', useEditorStore.getState().slideIndex)}
        data-testid="play"
      >
        <MonitorPlay size={15} />
      </button>
      </div>

      {/* 文档名与保存状态指示在底部状态栏（BottomBar） */}
    </div>
  )
}
