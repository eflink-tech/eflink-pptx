// 插入菜单：文本/形状/线条/表格/图表/音视频/公式
import { useEffect, useRef, useState } from 'react'
import { ChevronDown, Type, Spline, Table2, BarChart3, Video, Music, Sigma } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import {
  createAudioElement, createChartElement, createFormulaElement,
  createShapeElement, createTableElement, createTextElement, createVideoElement,
} from '../../core/schema/factory'
import { createLineElementFromPreset, LINE_PRESETS, PRESETS_BY_GROUP, type LinePresetId } from '../../core/schema/linePresets'
import { LinePresetIcon } from './LinePresetIcon'
import { ChartPickerPanel } from './ChartPickerPanel'
import { insertImageFile } from '../../core/editor/media'
import type { ChartType, PPTElement } from '../../types/slides'
import { SHAPE_CATEGORIES, SHAPE_PATHS, type ShapeMeta } from '../../core/render/shape'

/** 居中插入一个元素 */
function insertCentered(factory: (x: number, y: number) => PPTElement): void {
  const store = useEditorStore.getState()
  const slideW = store.presentation.width
  const slideH = Math.round(slideW / store.presentation.viewportRatio)
  store.pushHistory()
  store.addElement(factory(slideW / 2 - 150, slideH / 2 - 120))
}

/** 按预设尺寸居中插入线条 */
function insertLinePreset(presetId: LinePresetId): void {
  const store = useEditorStore.getState()
  const slideW = store.presentation.width
  const slideH = Math.round(slideW / store.presentation.viewportRatio)
  const preset = LINE_PRESETS[presetId]
  store.pushHistory()
  store.addElement(createLineElementFromPreset(
    (slideW - preset.w) / 2,
    (slideH - Math.max(preset.h, 24)) / 2,
    presetId,
  ))
}

function ShapePreview({ meta }: { meta: ShapeMeta }) {
  return (
    <svg viewBox="0 0 100 100" className="size-8">
      <path d={meta.path} fill="#4b5563" fillRule={meta.evenodd ? 'evenodd' : undefined} />
    </svg>
  )
}

const DROPDOWN_KEYS = new Set(['shape', 'line', 'chart'])

export function InsertMenu() {
  const [open, setOpen] = useState<string | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(null)
    }
    window.addEventListener('mousedown', onDown)
    return () => window.removeEventListener('mousedown', onDown)
  }, [])

  const toggleDropdown = (key: string) => {
    setOpen(open === key ? null : key)
  }

  const insertImage = () => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = 'image/*'
    input.onchange = () => {
      const file = input.files?.[0]
      if (file) void insertImageFile(file)
      input.value = ''
    }
    input.click()
  }

  const insertMedia = (kind: 'video' | 'audio') => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = kind === 'video' ? 'video/*' : 'audio/*'
    input.onchange = () => {
      const file = input.files?.[0]
      if (!file) return
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
    }
    input.click()
  }

  const items: Array<{ key: string; label: string; icon: React.ReactNode; onClick: () => void }> = [
    { key: 'text', label: '文本框', icon: <Type size={15} />, onClick: () => insertCentered((x, y) => createTextElement(x, y)) },
    { key: 'shape', label: '形状', icon: <SquareIcon />, onClick: () => toggleDropdown('shape') },
    { key: 'line', label: '线条', icon: <Spline size={15} />, onClick: () => toggleDropdown('line') },
    { key: 'table', label: '表格', icon: <Table2 size={15} />, onClick: () => insertCentered((x, y) => createTableElement(x, y)) },
    { key: 'chart', label: '图表', icon: <BarChart3 size={15} />, onClick: () => toggleDropdown('chart') },
    { key: 'image', label: '图片', icon: <ImageIcon />, onClick: insertImage },
    { key: 'video', label: '视频', icon: <Video size={15} />, onClick: () => insertMedia('video') },
    { key: 'audio', label: '音频', icon: <Music size={15} />, onClick: () => insertMedia('audio') },
    { key: 'formula', label: '公式', icon: <Sigma size={15} />, onClick: () => insertCentered((x, y) => createFormulaElement(x, y)) },
  ]

  return (
    <div ref={rootRef} className="relative flex items-center" data-testid="insert-menu">
      {items.map((item) => (
        <div key={item.key} className="relative">
          <button
            className={`flex items-center gap-1 rounded-md px-2 py-1.5 text-xs ${open === item.key ? 'bg-gray-100 text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'}`}
            title={item.label}
            onClick={() => {
              if (DROPDOWN_KEYS.has(item.key)) {
                toggleDropdown(item.key)
              } else {
                setOpen(null)
                item.onClick()
              }
            }}
            data-testid={`insert-${item.key}`}
          >
            {item.icon}
            {DROPDOWN_KEYS.has(item.key) && <ChevronDown size={12} />}
          </button>

          {item.key === 'shape' && open === 'shape' && (
            <div className="absolute left-0 top-9 z-50 max-h-[70vh] w-[420px] overflow-y-auto rounded-lg border border-gray-200 bg-white p-3 shadow-xl">
              {SHAPE_CATEGORIES.map((cat) => (
                <div key={cat.id} className="mb-3">
                  <div className="mb-1.5 text-xs font-medium text-gray-400">{cat.name}</div>
                  <div className="grid grid-cols-8 gap-1.5">
                    {Object.entries(SHAPE_PATHS)
                      .filter(([, meta]) => meta.category === cat.id)
                      .map(([key, meta]) => (
                        <button
                          key={key}
                          title={meta.name}
                          className="flex items-center justify-center rounded-md p-1.5 hover:bg-[#fbeae5]"
                          onClick={() => {
                            insertCentered((x, y) => createShapeElement(x, y, key))
                            setOpen(null)
                          }}
                        >
                          <ShapePreview meta={meta} />
                        </button>
                      ))}
                  </div>
                </div>
              ))}
            </div>
          )}

          {item.key === 'line' && open === 'line' && (
            <div className="absolute left-0 top-9 z-50 w-[320px] rounded-lg border border-gray-200 bg-white p-3 shadow-xl" data-testid="line-preset-menu">
              {PRESETS_BY_GROUP.map((group) => (
                <div key={group.id} className="mb-3 last:mb-0">
                  <div className="mb-2 flex items-center gap-1.5 text-xs font-medium text-gray-500">
                    <span className="h-3 w-0.5 rounded bg-gray-300" />
                    {group.name}
                  </div>
                  <div className={`grid gap-1.5 ${group.id === 'straight' ? 'grid-cols-5' : 'grid-cols-4'}`}>
                    {group.presets.map((preset) => (
                      <button
                        key={preset.id}
                        title={preset.id}
                        className="flex items-center justify-center rounded-md border border-transparent p-1.5 text-gray-600 hover:border-[#d14424]/30 hover:bg-[#fbeae5] hover:text-[#d14424]"
                        onClick={() => {
                          insertLinePreset(preset.id)
                          setOpen(null)
                        }}
                      >
                        <LinePresetIcon presetId={preset.id} />
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}

          {item.key === 'chart' && open === 'chart' && (
            <ChartPickerPanel
              onPick={(chartType: ChartType) => {
                setOpen(null)
                insertCentered((x, y) => createChartElement(x, y, chartType))
              }}
              onClose={() => setOpen(null)}
            />
          )}
        </div>
      ))}
    </div>
  )
}

function SquareIcon() {
  return (
    <svg viewBox="0 0 24 24" className="size-[15px]" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="4" y="4" width="16" height="16" rx="2" />
    </svg>
  )
}

function ImageIcon() {
  return (
    <svg viewBox="0 0 24 24" className="size-[15px]" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <circle cx="8.5" cy="8.5" r="1.5" />
      <path d="M21 15l-5-5L5 21" />
    </svg>
  )
}
