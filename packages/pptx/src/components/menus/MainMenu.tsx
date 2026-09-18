// 主菜单：级联二级菜单（对齐 word/excel 交互——悬停一级分类展开二级，带 ✓ 开关态与快捷键提示）
// 分类按 eflink 实际功能定制：文件 / 设计 / 视图 / 放映 / 快捷键（叶子）；仅列真实可用功能
import { Fragment, useEffect, useRef, useState } from 'react'
import {
  Check, ChevronDown, ChevronRight, Download, Eye, FilePlus2, FileText,
  Grid3x3, Keyboard, LayoutTemplate, Menu, MessageCircle, MonitorPlay,
  MonitorSpeaker, Palette, PanelLeft, Save, Search, Share2, Upload,
  type LucideIcon,
} from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { createDoc, saveDoc, type LoadedDoc } from '../../core/editor/persistence'
import { getPptxShareHandler } from '../../core/share/shareBridge'
import { ShareDialog } from '../common/ShareDialog'

interface MenuLeaf {
  key: string
  label: string
  icon?: LucideIcon
  shortcut?: string
  /** 开关型菜单项：选中时右侧显示 ✓ */
  checked?: boolean
  action?: () => void
}

interface MenuCategory {
  key: string
  label: string
  icon: LucideIcon
  /** 分组之间渲染分隔线 */
  groups: MenuLeaf[][]
  /** 叶子分类（无子菜单时直接点击触发，如快捷键） */
  leaf?: MenuLeaf
}

export function MainMenu() {
  const [open, setOpen] = useState(false)
  const [activeCat, setActiveCat] = useState<string | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const ui = useUIStore()

  // 分享弹窗（doc 为点击"分享"时刻的文档快照，弹窗期间编辑不影响本次分享内容）
  const [shareOpen, setShareOpen] = useState(false)
  const [shareDoc, setShareDoc] = useState<LoadedDoc | null>(null)
  // 分享前强制保存：先落库最新内容，再捕获当前文档；保存失败则中止分享（避免分享远端旧数据）
  const openShare = async () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast('文档未初始化，无法分享', 'error'); return }
    try {
      await saveDoc(s.docId, s.docName, s.presentation)
      s.markSaved()
    } catch {
      useToastStore.getState().toast('保存失败，无法分享', 'error')
      return
    }
    setShareDoc({ id: s.docId, name: s.docName, presentation: s.presentation })
    setShareOpen(true)
  }

  function closeAll() {
    setOpen(false)
    setActiveCat(null)
  }

  // Escape 关闭整组菜单；点击组件外部关闭
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) closeAll()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeAll()
    }
    window.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const saveNow = () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast('文档未初始化，无法保存', 'error'); return }
    saveDoc(s.docId, s.docName, s.presentation)
      .then(() => { s.markSaved(); useToastStore.getState().toast('已保存', 'success') })
      .catch(() => useToastStore.getState().toast('保存失败，请重试', 'error'))
  }

  const createNew = async () => {
    const doc = await createDoc('未命名演示文稿')
    useEditorStore.getState().loadDocument(doc)
  }

  const play = (mode: 'playing' | 'presenter') => {
    useUIStore.getState().setPlayerMode(mode, useEditorStore.getState().slideIndex)
  }

  // 分享入口仅在宿主注入分享实现后出现（纯组件独立运行时不显示）
  const canShare = getPptxShareHandler() !== null

  const categories: MenuCategory[] = [
    {
      key: 'file',
      label: '文件',
      icon: FileText,
      groups: [
        [
          { key: 'new', label: '新建文档', icon: FilePlus2, action: () => void createNew() },
          { key: 'save', label: '保存', icon: Save, shortcut: 'Ctrl S', action: saveNow },
        ],
        [
          { key: 'import', label: '导入（PPTX/JSON）…', icon: Upload, action: () => ui.openModal('import') },
          { key: 'export', label: '导出…', icon: Download, action: () => ui.openModal('export') },
        ],
        ...(canShare ? [[
          { key: 'share', label: '分享', icon: Share2, action: () => void openShare() },
          { key: 'feedback', label: '反馈', icon: MessageCircle, action: () => window.open('/contact', '_blank') },
        ] as MenuLeaf[]] : []),
      ],
    },
    {
      key: 'design',
      label: '设计',
      icon: Palette,
      groups: [
        [
          { key: 'template', label: '模板库…', icon: LayoutTemplate, action: () => ui.openModal('template') },
          { key: 'theme', label: '主题配色…', icon: Palette, action: () => ui.openModal('theme') },
        ],
      ],
    },
    {
      key: 'view',
      label: '视图',
      icon: Eye,
      groups: [
        [
          { key: 'grid', label: '网格', icon: Grid3x3, checked: ui.gridVisible, action: () => ui.toggleGrid() },
          { key: 'thumbnails', label: '缩略图面板', icon: PanelLeft, checked: ui.thumbnailsVisible, action: () => ui.toggleThumbnails() },
        ],
        [
          { key: 'findReplace', label: '查找替换…', icon: Search, shortcut: 'Ctrl F', action: () => ui.openModal('findReplace') },
        ],
      ],
    },
    {
      key: 'play',
      label: '放映',
      icon: MonitorPlay,
      groups: [
        [
          { key: 'play', label: '放映', icon: MonitorPlay, action: () => play('playing') },
          { key: 'presenter', label: '演讲者视图', icon: MonitorSpeaker, action: () => play('presenter') },
        ],
      ],
    },
    {
      key: 'hotkey',
      label: '快捷键',
      icon: Keyboard,
      /** 快捷键是叶子分类：点击直接打开对话框，不需要子菜单 */
      groups: [],
      leaf: {
        key: 'hotkey-list',
        label: '查看全部快捷键',
        icon: Keyboard,
        action: () => ui.openModal('hotkey'),
      },
    },
  ]

  const renderLeaf = (leaf: MenuLeaf) => (
    <button
      key={leaf.key}
      type="button"
      onClick={() => {
        leaf.action?.()
        closeAll()
      }}
      className="flex w-full items-center gap-2.5 px-3 py-[7px] text-left text-[13px] text-[#1f2329] transition-colors hover:bg-[#f2f3f4]"
    >
      {leaf.icon ? <leaf.icon size={16} className="shrink-0 text-[#51565f]" /> : <span className="w-4 shrink-0" />}
      <span className="min-w-0 flex-1 truncate">{leaf.label}</span>
      {leaf.checked && <Check size={13} className="shrink-0 text-[#1f2329]" />}
      {leaf.shortcut && <span className="shrink-0 text-xs text-[#8f959e]">{leaf.shortcut}</span>}
    </button>
  )

  return (
    <div ref={rootRef} className="relative shrink-0">
      <button
        type="button"
        data-testid="main-menu"
        title="主菜单"
        onClick={() => {
          setOpen((v) => !v)
          setActiveCat(null)
        }}
        className={`flex items-center gap-0.5 rounded-md px-2 py-1.5 text-xs transition-colors ${
          open ? 'bg-gray-100 text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'
        }`}
      >
        <Menu size={17} />
        <ChevronDown size={12} />
      </button>

      {open && (
        <div className="absolute left-0 top-full z-[9999] mt-1 w-[132px] rounded-lg border border-black/[0.08] bg-white py-1 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
          {categories.map((cat) => (
            <div
              key={cat.key}
              className="relative"
              onMouseEnter={() => {
                // 叶子分类（快捷键）无子菜单：悬停时收起已展开的子菜单
                setActiveCat(cat.leaf ? null : cat.key)
              }}
            >
              <button
                type="button"
                onClick={() => {
                  // 叶子分类直接触发
                  if (cat.leaf) {
                    cat.leaf.action?.()
                    closeAll()
                    return
                  }
                  setActiveCat(cat.key)
                }}
                className={`flex w-full items-center justify-between px-3 py-2 text-left text-[13px] transition-colors ${
                  activeCat === cat.key
                    ? 'bg-[#f2f3f4] text-[#1f2329]'
                    : 'text-[#1f2329] hover:bg-[#f2f3f4]'
                } ${cat.leaf ? 'cursor-pointer' : ''}`}
              >
                <span className="flex items-center gap-2">
                  <cat.icon size={16} className="shrink-0 text-[#51565f]" />
                  {cat.label}
                </span>
                {!cat.leaf && <ChevronRight size={12} className="shrink-0 text-[#8f959e]" />}
              </button>

              {activeCat === cat.key && (
                <div className="absolute left-full top-0 z-[9999] ml-0.5 w-[224px] rounded-lg border border-black/[0.08] bg-white py-1 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
                  {cat.groups.map((group, gi) => (
                    <Fragment key={gi}>
                      {gi > 0 && <div className="my-1 h-px bg-black/[0.05]" />}
                      {group.map((leaf) => renderLeaf(leaf))}
                    </Fragment>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      <ShareDialog open={shareOpen} doc={shareDoc} onClose={() => setShareOpen(false)} />
    </div>
  )
}
