// 顶部：统一 OfficeHeader（返回 + 品牌 + 横向菜单 + AI/放映）+ 第二行工具栏（撤销/重做 + 插入）
import { useState } from 'react'
import {
  BarChart3, Download, FilePlus2, Grid3x3,
  Image as ImageIcon, Keyboard, LayoutTemplate, MessageCircle, MonitorPlay,
  MonitorSpeaker, Music, Palette, PanelLeft, Redo2, Save, Search, Share2,
  Shapes, Sigma, Spline, Table2, Type, Undo2, Upload, Video, Sparkles,
} from 'lucide-react'
import { OfficeHeader, useI18n, type OfficeHeaderBrand, type OfficeMenu } from '@eflink-tech/office-kit'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { getEditorBackHref } from '../../core/editor/chrome'
import { createDoc, saveDoc, type LoadedDoc } from '../../core/editor/persistence'
import { getPptxShareHandler } from '../../core/share/shareBridge'
import { insertImageFile } from '../../core/editor/media'
import {
  createChartElement, createFormulaElement, createShapeElement, createTableElement,
  createTextElement, createVideoElement, createAudioElement,
} from '../../core/schema/factory'
import { createLineElementFromPreset, LINE_PRESETS } from '../../core/schema/linePresets'
import { InsertMenu } from '../menus/InsertMenu'
import { ShareDialog } from '../common/ShareDialog'
import logoUrl from '../../assets/pptx-doculuma-logo.png'
import type { PPTElement } from '../../types/slides'

function insertCentered(factory: (x: number, y: number) => PPTElement): void {
  const store = useEditorStore.getState()
  const slideW = store.presentation.width
  const slideH = Math.round(slideW / store.presentation.viewportRatio)
  store.pushHistory()
  store.addElement(factory(slideW / 2 - 150, slideH / 2 - 120))
}

function pickFile(accept: string, onPick: (file: File) => void): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = accept
  input.onchange = () => {
    const file = input.files?.[0]
    if (file) onPick(file)
    input.value = ''
  }
  input.click()
}

function pickMedia(kind: 'video' | 'audio'): void {
  pickFile(kind === 'video' ? 'video/*' : 'audio/*', (file) => {
    const reader = new FileReader()
    reader.onload = () => {
      const store = useEditorStore.getState()
      const src = reader.result as string
      const slideW = store.presentation.width
      const slideH = Math.round(slideW / store.presentation.viewportRatio)
      store.pushHistory()
      if (kind === 'video') {
        const w = slideW * 0.5
        store.addElement(createVideoElement((slideW - w) / 2, (slideH - w * 0.5625) / 2, src))
      } else {
        store.addElement(createAudioElement(slideW / 2 - 32, slideH / 2 - 32, src))
      }
    }
    reader.readAsDataURL(file)
  })
}

function ToolButton({ icon, label, onClick, disabled }: {
  icon: React.ReactNode
  label: string
  onClick: () => void
  disabled?: boolean
}) {
  return (
    <button
      className="flex items-center gap-1 rounded-md px-2 py-1.5 text-xs text-gray-600 transition-colors hover:bg-gray-100 disabled:cursor-not-allowed disabled:opacity-40"
      title={label}
      aria-label={label}
      onClick={onClick}
      disabled={disabled}
    >
      {icon}
    </button>
  )
}

export function TopBar({ branding }: { branding?: OfficeHeaderBrand | false } = {}) {
  const { t } = useI18n()
  const undoDepth = useEditorStore((s) => s.history.length)
  const redoDepth = useEditorStore((s) => s.future.length)
  const ui = useUIStore()
  const backHref = getEditorBackHref()

  const [shareOpen, setShareOpen] = useState(false)
  const [shareDoc, setShareDoc] = useState<LoadedDoc | null>(null)

  const saveNow = () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast(t('pptx.toast.notReadySave'), 'error'); return }
    saveDoc(s.docId, s.docName, s.presentation)
      .then(() => { s.markSaved(); useToastStore.getState().toast(t('pptx.toast.saved'), 'success') })
      .catch(() => useToastStore.getState().toast(t('pptx.toast.saveFailed'), 'error'))
  }

  const openShare = async () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast(t('pptx.toast.notReadyShare'), 'error'); return }
    try {
      await saveDoc(s.docId, s.docName, s.presentation)
      s.markSaved()
    } catch {
      useToastStore.getState().toast(t('pptx.toast.shareSaveFailed'), 'error')
      return
    }
    setShareDoc({ id: s.docId, name: s.docName, presentation: s.presentation })
    setShareOpen(true)
  }

  const play = (mode: 'playing' | 'presenter') => {
    useUIStore.getState().setPlayerMode(mode, useEditorStore.getState().slideIndex)
  }

  const canShare = getPptxShareHandler() !== null

  const menus: OfficeMenu[] = [
    {
      key: 'file',
      label: t('shell.menu.file'),
      items: [
        [
          { key: 'new', label: t('pptx.file.new'), icon: <FilePlus2 size={16} />, onSelect: () => void createDoc(t('pptx.untitled')).then((doc) => useEditorStore.getState().loadDocument(doc)) },
          { key: 'save', label: t('action.save'), icon: <Save size={16} />, shortcut: 'Ctrl S', onSelect: saveNow },
        ],
        [
          { key: 'import', label: t('pptx.file.import'), icon: <Upload size={16} />, onSelect: () => ui.openModal('import') },
          { key: 'export', label: t('pptx.file.export'), icon: <Download size={16} />, onSelect: () => ui.openModal('export') },
        ],
        ...(canShare ? [[
          { key: 'share', label: t('shell.share'), icon: <Share2 size={16} />, onSelect: () => void openShare() },
        ]] : []),
      ],
    },
    {
      key: 'edit',
      label: t('shell.menu.edit'),
      items: [
        [
          { key: 'undo', label: t('action.undo'), icon: <Undo2 size={16} />, shortcut: 'Ctrl Z', disabled: undoDepth === 0, onSelect: () => useEditorStore.getState().undo() },
          { key: 'redo', label: t('action.redo'), icon: <Redo2 size={16} />, shortcut: 'Ctrl Y', disabled: redoDepth === 0, onSelect: () => useEditorStore.getState().redo() },
        ],
        [
          { key: 'findReplace', label: t('pptx.edit.findReplace'), icon: <Search size={16} />, shortcut: 'Ctrl F', onSelect: () => ui.openModal('findReplace') },
        ],
      ],
    },
    {
      key: 'insert',
      label: t('shell.menu.insert'),
      items: [
        [
          { key: 'text', label: t('pptx.insert.text'), icon: <Type size={16} />, onSelect: () => insertCentered((x, y) => createTextElement(x, y)) },
          { key: 'shape', label: t('pptx.insert.shape'), icon: <Shapes size={16} />, onSelect: () => insertCentered((x, y) => createShapeElement(x, y, 'rect')) },
          { key: 'line', label: t('pptx.insert.line'), icon: <Spline size={16} />, onSelect: () => {
            const store = useEditorStore.getState()
            const slideW = store.presentation.width
            const slideH = Math.round(slideW / store.presentation.viewportRatio)
            const preset = LINE_PRESETS['straight-solid']
            store.pushHistory()
            store.addElement(createLineElementFromPreset((slideW - preset.w) / 2, (slideH - Math.max(preset.h, 24)) / 2, 'straight-solid'))
          } },
        ],
        [
          { key: 'table', label: t('pptx.insert.table'), icon: <Table2 size={16} />, onSelect: () => insertCentered((x, y) => createTableElement(x, y)) },
          { key: 'chart', label: t('pptx.insert.chart'), icon: <BarChart3 size={16} />, onSelect: () => insertCentered((x, y) => createChartElement(x, y, 'bar-cluster', t)) },
          { key: 'image', label: t('pptx.insert.image'), icon: <ImageIcon size={16} />, onSelect: () => pickFile('image/*', (file) => void insertImageFile(file)) },
        ],
        [
          { key: 'video', label: t('pptx.insert.video'), icon: <Video size={16} />, onSelect: () => pickMedia('video') },
          { key: 'audio', label: t('pptx.insert.audio'), icon: <Music size={16} />, onSelect: () => pickMedia('audio') },
          { key: 'formula', label: t('pptx.insert.formula'), icon: <Sigma size={16} />, onSelect: () => insertCentered((x, y) => createFormulaElement(x, y)) },
        ],
      ],
    },
    {
      key: 'page',
      label: t('shell.menu.page'),
      items: [
        [
          { key: 'template', label: t('pptx.page.templates'), icon: <LayoutTemplate size={16} />, onSelect: () => ui.openModal('template') },
          { key: 'theme', label: t('pptx.page.theme'), icon: <Palette size={16} />, onSelect: () => ui.openModal('theme') },
        ],
      ],
    },
    {
      key: 'view',
      label: t('shell.menu.view'),
      items: [
        [
          { key: 'grid', label: t('pptx.view.grid'), icon: <Grid3x3 size={16} />, checked: ui.gridVisible, onSelect: () => ui.toggleGrid() },
          { key: 'thumbnails', label: t('pptx.view.thumbnails'), icon: <PanelLeft size={16} />, checked: ui.thumbnailsVisible, onSelect: () => ui.toggleThumbnails() },
        ],
        [
          { key: 'play', label: t('pptx.view.play'), icon: <MonitorPlay size={16} />, onSelect: () => play('playing') },
          { key: 'presenter', label: t('pptx.view.presenter'), icon: <MonitorSpeaker size={16} />, onSelect: () => play('presenter') },
        ],
        [
          { key: 'findReplace', label: t('pptx.edit.findReplace'), icon: <Search size={16} />, shortcut: 'Ctrl F', onSelect: () => ui.openModal('findReplace') },
        ],
      ],
    },
    {
      key: 'help',
      label: t('shell.menu.help'),
      items: [
        [
          { key: 'shortcuts', label: t('pptx.help.shortcuts'), icon: <Keyboard size={16} />, onSelect: () => ui.openModal('hotkey') },
          { key: 'feedback', label: t('pptx.help.feedback'), icon: <MessageCircle size={16} />, onSelect: () => window.open('/contact', '_blank') },
        ],
      ],
    },
  ]

  return (
    <>
      <OfficeHeader
        brand={branding === false ? false : (branding ?? { logo: logoUrl, name: t('pptx.brand') })}
        backHref={backHref}
        menus={menus}
        actions={
          <>
            <button
              className={`flex h-7 shrink-0 items-center gap-1 rounded-md px-2 text-xs transition-colors ${
                ui.aiPanelVisible ? 'bg-[#fbeae5] text-[#d14424]' : 'text-[#4c5158] hover:bg-black/[0.06] hover:text-[#1f2329]'
              }`}
              title={t('pptx.toolbar.ai')}
              aria-label={t('pptx.toolbar.ai')}
              onClick={() => ui.toggleAIPanel()}
              data-testid="ai-toggle"
            >
              <Sparkles size={14} />
              {t('pptx.toolbar.ai')}
            </button>
            <button
              className="flex h-7 shrink-0 items-center gap-1 rounded-md bg-[#d14424] px-2.5 text-xs text-white transition-colors hover:bg-[#b93a1d]"
              title={t('pptx.toolbar.play')}
              aria-label={t('pptx.toolbar.play')}
              onClick={() => play('playing')}
              data-testid="play"
            >
              <MonitorPlay size={14} />
              {t('pptx.toolbar.play')}
            </button>
          </>
        }
      />
      <div className="relative z-30 flex h-10 shrink-0 items-center gap-1 border-b border-gray-200 bg-white px-3" data-testid="topbar">
        <ToolButton icon={<Undo2 size={17} />} label={t('action.undo')} disabled={undoDepth === 0} onClick={() => useEditorStore.getState().undo()} />
        <ToolButton icon={<Redo2 size={17} />} label={t('action.redo')} disabled={redoDepth === 0} onClick={() => useEditorStore.getState().redo()} />
        <div className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2">
          <InsertMenu />
        </div>
      </div>
      <ShareDialog open={shareOpen} doc={shareDoc} onClose={() => setShareOpen(false)} />
    </>
  )
}
