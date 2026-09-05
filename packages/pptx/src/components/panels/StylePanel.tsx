// 右侧样式面板：卡片式分组布局，紧凑排列
import { useState } from 'react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import type { ChartElement, PPTElement, ShapeElement, TableCell } from '../../types/slides'
import { type AlignMode } from '../../core/utils/align'
import { SHAPE_PATHS } from '../../core/render/shape'
import type { TransitionPreset } from '../../types/slides'
import { chartTypeGroups, effectOptions, transitionPresets } from './options'
import { DEFAULT_CHART_COLORS } from '../canvas/ChartRenderer'
import { CHART_PALETTES, DEFAULT_CHART_ELEMENTS, type ChartElements } from '../../core/chart/chartOptions'
import {
  AlignCenter, AlignHorizontalDistributeCenter, AlignLeft, AlignRight,
  AlignVerticalDistributeCenter, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, AlignVerticalJustifyStart,
  ChevronDown, ChevronRight, Copy, FlipHorizontal, FlipVertical,
  Group, Link2, Lock, Pencil, RotateCcw, Trash2, Ungroup,
} from 'lucide-react'
import { IconLayerBackward, IconLayerForward, IconLayerToBack, IconLayerToFront } from './LayerIcons'
import { LineStyleSection } from './LineStyleControls'
import { FormulaEditorDialog } from '../formula/FormulaEditorDialog'
import { ColorInput } from '../color/ColorInput'

/* ---------- 基础控件 ---------- */

function NumberInput({ value, onChange, min = 0, max = 1000, step = 1, suffix }: {
  value: number | undefined
  onChange: (v: number) => void
  min?: number
  max?: number
  step?: number
  suffix?: string
}) {
  return (
    <span className="flex items-center gap-1">
      <input
        type="number"
        value={value ?? 0}
        min={min}
        max={max}
        step={step}
        onChange={(e) => onChange(Number(e.target.value))}
        className="w-16 rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
      />
      {suffix && <span className="text-xs text-gray-400">{suffix}</span>}
    </span>
  )
}

function SelectInput({ value, onChange, options, testId }: {
  value: string | undefined
  onChange: (v: string) => void
  options: Array<{ value: string; label: string }>
  testId?: string
}) {
  return (
    <select
      value={value ?? ''}
      onChange={(e) => onChange(e.target.value)}
      data-testid={testId}
      className="w-full rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
    >
      {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  )
}

/** 分组下拉（optgroup），用于按类别展示图表类型 */
function GroupedSelectInput({ value, onChange, groups, testId }: {
  value: string | undefined
  onChange: (v: string) => void
  groups: Array<{ group: string; types: Array<{ id: string; label: string }> }>
  testId?: string
}) {
  return (
    <select
      value={value ?? ''}
      onChange={(e) => onChange(e.target.value)}
      data-testid={testId}
      className="w-full rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
    >
      {groups.map((g) => (
        <optgroup key={g.group} label={g.group}>
          {g.types.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
        </optgroup>
      ))}
    </select>
  )
}

function ActionButton({ label, onClick, danger, disabled, active }: { label: string; onClick: () => void; danger?: boolean; disabled?: boolean; active?: boolean }) {
  return (
    <button
      className={`rounded border px-2 py-1 text-xs hover:bg-gray-50 disabled:opacity-40 ${
        active ? 'border-[#d14424] bg-[#fbeae5] text-[#d14424]' :
        danger ? 'border-gray-200 text-red-500' : 'border-gray-200 text-gray-600'
      }`}
      onClick={onClick}
      disabled={disabled}
    >
      {label}
    </button>
  )
}

/** 图标工具按钮（与富文本工具条 TB 一致） */
function IconButton({ icon, title, onClick, danger, disabled, active }: {
  icon: React.ReactNode
  title: string
  onClick: () => void
  danger?: boolean
  disabled?: boolean
  active?: boolean
}) {
  return (
    <button
      type="button"
      title={title}
      aria-label={title}
      disabled={disabled}
      onClick={onClick}
      className={`flex size-7 items-center justify-center rounded transition-colors disabled:opacity-40 ${
        active ? 'bg-[#fbeae5] text-[#d14424]' :
        danger ? 'text-red-500 hover:bg-red-50' : 'text-gray-600 hover:bg-gray-100'
      }`}
    >
      {icon}
    </button>
  )
}

/** 图标按钮组（带小标题与分隔） */
function ToolGroup({ label, children }: { label?: string; children: React.ReactNode }) {
  return (
    <div className="mb-2 last:mb-0">
      {label && <div className="mb-1 text-[10px] font-medium uppercase tracking-wide text-gray-400">{label}</div>}
      <div className="flex flex-wrap items-center gap-0.5">{children}</div>
    </div>
  )
}

function ToolDivider() {
  return <div className="mx-0.5 h-5 w-px shrink-0 bg-gray-200" />
}

/* ---------- 可折叠分组卡片 ---------- */

function Section({ title, defaultOpen = true, children }: { title: string; defaultOpen?: boolean; children: React.ReactNode }) {
  const [open, setOpen] = useState(defaultOpen)
  return (
    <div className="mb-2 overflow-hidden rounded-lg border border-gray-100 bg-gray-50/50">
      <button
        type="button"
        className="flex w-full items-center gap-1.5 px-3 py-2 text-left hover:bg-gray-50/80"
        onClick={() => setOpen(!open)}
      >
        {open ? <ChevronDown size={13} className="text-gray-400" /> : <ChevronRight size={13} className="text-gray-400" />}
        <span className="text-xs font-semibold text-gray-700">{title}</span>
      </button>
      {open && <div className="space-y-1.5 border-t border-gray-100 px-3 py-2">{children}</div>}
    </div>
  )
}

/* ---------- 设置行（紧凑布局） ---------- */

function Field({ label, children, className }: { label: string; children: React.ReactNode; className?: string }) {
  return (
    <div className={`flex items-center gap-2 ${className ?? ''}`}>
      <span className="w-16 shrink-0 text-xs text-gray-500">{label}</span>
      <div className="min-w-0 flex-1">{children}</div>
    </div>
  )
}

function ToggleField({ label, checked, onChange, icon }: { label: string; checked: boolean; onChange: (v: boolean) => void; icon?: React.ReactNode }) {
  return (
    <label className="flex cursor-pointer items-center justify-between gap-2 py-1">
      <span className="flex items-center gap-1.5 text-xs text-gray-600">
        {icon && <span className="text-gray-400">{icon}</span>}
        {label}
      </span>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative h-4 w-7 shrink-0 rounded-full transition-colors ${checked ? 'bg-[#d14424]' : 'bg-gray-200'}`}
      >
        <span className={`absolute top-0.5 left-0.5 size-3 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-3' : 'translate-x-0'}`} />
      </button>
    </label>
  )
}

/* ---------- 面板主体 ---------- */

export function StylePanel() {
  const tab = useUIStore((s) => s.rightPanelTab)
  const setTab = useUIStore((s) => s.setRightPanelTab)

  return (
    <div className="flex h-full w-[260px] shrink-0 flex-col border-l border-gray-200 bg-white" data-testid="style-panel">
      <div className="flex border-b border-gray-100 text-xs">
        {([['style', '样式'], ['animation', '动画'], ['comment', '批注']] as const).map(([key, label]) => (
          <button
            key={key}
            className={`flex-1 py-2 text-center ${tab === key ? 'border-b-2 border-[#d14424] font-medium text-[#d14424]' : 'text-gray-500 hover:text-gray-700'}`}
            onClick={() => setTab(key)}
          >
            {label}
          </button>
        ))}
      </div>
      <div className="flex-1 overflow-y-auto p-2">
        {tab === 'style' && <StyleTab />}
        {tab === 'animation' && <AnimationTab />}
        {tab === 'comment' && <CommentTab />}
      </div>
    </div>
  )
}

/* ---------- 样式页签 ---------- */

function StyleTab() {
  const selectedIds = useEditorStore((s) => s.selectedIds)
  const slide = useEditorStore((s) => s.presentation.slides[s.slideIndex])

  if (!slide) return null
  const selected = slide.elements.filter((el) => selectedIds.includes(el.id))

  if (!selected.length) {
    return (
      <div>
        <PageSection />
        <ArrangeSection selectedIds={[]} />
      </div>
    )
  }

  if (selected.length === 1) {
    const el = selected[0]
    return (
      <div>
        <div className="mb-2 rounded-lg bg-[#fbeae5] px-3 py-1.5 text-xs font-medium text-[#d14424]">
          {typeLabel(el)}
        </div>
        <ElementSection el={el} />
        <ArrangeSection selectedIds={selectedIds} />
      </div>
    )
  }

  return (
    <div>
      <div className="mb-2 rounded-lg bg-[#fbeae5] px-3 py-1.5 text-xs font-medium text-[#d14424]">
        已选 {selected.length} 个元素
      </div>
      <Section title="通用">
        <Field label="不透明度">
          <div className="flex items-center gap-2">
            <input
              type="range"
              min={5}
              max={100}
              value={Math.round((selected[0].opacity ?? 1) * 100)}
              onChange={(e) => {
                const store = useEditorStore.getState()
                store.pushHistory()
                store.updateElements(selectedIds, (d) => { d.opacity = Number(e.target.value) / 100 })
              }}
              className="h-1 flex-1 cursor-pointer accent-[#d14424]"
            />
            <NumberInput
              value={Math.round((selected[0].opacity ?? 1) * 100)}
              min={5}
              max={100}
              onChange={(v) => {
                const store = useEditorStore.getState()
                store.pushHistory()
                store.updateElements(selectedIds, (d) => { d.opacity = v / 100 })
              }}
              suffix="%"
            />
          </div>
        </Field>
      </Section>
      <ArrangeSection selectedIds={selectedIds} />
    </div>
  )
}

function typeLabel(el: PPTElement): string {
  const map: Record<PPTElement['type'], string> = {
    text: '文本', image: '图片', shape: '形状', line: '线条', table: '表格',
    chart: '图表', video: '视频', audio: '音频', formula: '公式',
  }
  return map[el.type]
}

/* 页面设置（未选中元素时） */
function PageSection() {
  const slide = useEditorStore((s) => s.presentation.slides[s.slideIndex])
  if (!slide) return null
  const bg = slide.background ?? { type: 'solid' as const, color: '#ffffff' }

  return (
    <div>
      <Section title="页面背景">
        <Field label="类型">
          <SelectInput
            testId="bg-type"
            value={bg.type}
            onChange={(v) => {
              const store = useEditorStore.getState()
              store.pushHistory()
              if (v === 'solid') store.setSlideBackground({ type: 'solid', color: '#ffffff' })
              else if (v === 'gradient') store.setSlideBackground({ type: 'gradient', gradient: { type: 'linear', colors: [{ pos: 0, color: '#d14424' }, { pos: 100, color: '#f5c56b' }], rotate: 90 } })
              else store.setSlideBackground({ type: 'image', image: { src: '', size: 'cover' } })
            }}
            options={[
              { value: 'solid', label: '纯色' },
              { value: 'gradient', label: '渐变' },
              { value: 'image', label: '图片' },
            ]}
          />
        </Field>
        {bg.type === 'solid' && (
          <Field label="颜色">
            <ColorInput value={bg.color} onChange={(v) => useEditorStore.getState().setSlideBackground({ color: v })} />
          </Field>
        )}
        {bg.type === 'gradient' && bg.gradient && (
          <>
            <Field label="起始色">
              <ColorInput value={bg.gradient.colors[0]?.color} onChange={(v) => {
                const colors = [...bg.gradient!.colors]
                colors[0] = { ...colors[0], color: v }
                useEditorStore.getState().setSlideBackground({ gradient: { ...bg.gradient!, colors } })
              }} />
            </Field>
            <Field label="结束色">
              <ColorInput value={bg.gradient.colors[1]?.color} onChange={(v) => {
                const colors = [...bg.gradient!.colors]
                colors[1] = { ...colors[1] ?? { pos: 100 }, color: v }
                useEditorStore.getState().setSlideBackground({ gradient: { ...bg.gradient!, colors } })
              }} />
            </Field>
            <Field label="角度">
              <NumberInput value={bg.gradient.rotate ?? 90} onChange={(v) => useEditorStore.getState().setSlideBackground({ gradient: { ...bg.gradient!, rotate: v } })} suffix="°" />
            </Field>
          </>
        )}
        {bg.type === 'image' && (
          <Field label="图片">
            <label className="cursor-pointer rounded border border-gray-200 px-2 py-1 text-xs hover:bg-gray-50">
              选择文件
              <input
                type="file"
                accept="image/*"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0]
                  if (!file) return
                  const reader = new FileReader()
                  reader.onload = () => useEditorStore.getState().setSlideBackground({ image: { src: reader.result as string, size: 'cover' } })
                  reader.readAsDataURL(file)
                }}
              />
            </label>
          </Field>
        )}
      </Section>
      <Section title="画布比例">
        <Field label="比例">
          <SelectInput
            testId="canvas-ratio"
            value={String(useEditorStore.getState().presentation.viewportRatio)}
            onChange={(v) => {
              const store = useEditorStore.getState()
              store.pushHistory()
              store.setViewportRatio(Number(v))
            }}
            options={[
              { value: String(16 / 9), label: '16 : 9' },
              { value: String(16 / 10), label: '16 : 10' },
              { value: String(4 / 3), label: '4 : 3' },
              { value: '1', label: '1 : 1' },
            ]}
          />
        </Field>
      </Section>
      <Section title="切换动画">
        <Field label="效果">
          <SelectInput
            testId="transition-effect"
            value={slide.transition?.preset ?? 'none'}
            onChange={(v) => useEditorStore.getState().setSlideTransition({ preset: v as TransitionPreset })}
            options={transitionPresets}
          />
        </Field>
        <Field label="时长">
          <NumberInput value={slide.transition?.duration ?? 500} onChange={(v) => useEditorStore.getState().setSlideTransition({ duration: v })} suffix="ms" step={100} />
        </Field>
      </Section>
    </div>
  )
}

/* 元素样式分区（按类型） */
function ElementSection({ el }: { el: PPTElement }) {
  const [formulaDialogOpen, setFormulaDialogOpen] = useState(false)
  const pushThen = (fn: () => void) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    fn()
  }
  const upd = (fn: (draft: PPTElement) => void) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    store.updateElements([el.id], fn)
  }

  return (
    <div className="mb-2">
      {/* 文本元素 */}
      {el.type === 'text' && (
        <>
          <Field label="文字色">
            <ColorInput value={el.defaultColor} onChange={(v) => upd((d) => { if (d.type === 'text') d.defaultColor = v })} />
          </Field>
          <div className="grid grid-cols-2 gap-x-2 gap-y-1">
            <Field label="行高">
              <NumberInput value={el.lineHeight ?? 1.5} step={0.1} min={0.5} max={4} onChange={(v) => upd((d) => { if (d.type === 'text') d.lineHeight = v })} />
            </Field>
            <Field label="字间距">
              <NumberInput value={el.charSpace ?? 0} onChange={(v) => upd((d) => { if (d.type === 'text') d.charSpace = v })} suffix="px" />
            </Field>
            <Field label="内边距">
              <NumberInput value={el.padding ?? 8} onChange={(v) => upd((d) => { if (d.type === 'text') d.padding = v })} suffix="px" />
            </Field>
          </div>
          <div className="mt-1 space-y-0.5 rounded-md bg-white px-1">
            <ToggleField label="竖排" checked={el.vertical ?? false} onChange={(v) => upd((d) => { if (d.type === 'text') d.vertical = v })} />
            <ToggleField label="自动收缩" checked={el.autoSize ?? false} onChange={(v) => upd((d) => { if (d.type === 'text') d.autoSize = v })} />
          </div>
        </>
      )}

      {/* 文本特效 */}
      {el.type === 'text' && (
        <Section title="文本特效" defaultOpen={false}>
          <ToggleField
            label="描边"
            checked={Boolean(el.textStroke)}
            onChange={(v) => upd((d) => {
              if (d.type !== 'text') return
              d.textStroke = v ? { color: '#ffffff', width: 2 } : undefined
            })}
          />
          {el.textStroke && (
            <div className="mt-1 flex items-center gap-2">
              <ColorInput value={el.textStroke.color} onChange={(v) => upd((d) => { if (d.type === 'text' && d.textStroke) d.textStroke.color = v })} small />
              <NumberInput value={el.textStroke.width} min={1} max={10} onChange={(v) => upd((d) => { if (d.type === 'text' && d.textStroke) d.textStroke.width = v })} suffix="px" />
            </div>
          )}
          <ToggleField
            label="文字阴影"
            checked={Boolean(el.shadow)}
            onChange={(e) => upd((d) => {
              if (d.type !== 'text') return
              d.shadow = e ? { h: 2, v: 2, blur: 6, color: 'rgba(0,0,0,0.35)' } : undefined
            })}
          />
        </Section>
      )}

      {/* 图片元素 */}
      {el.type === 'image' && (
        <Section title="图片样式" defaultOpen>
          <div className="grid grid-cols-2 gap-2">
            <Field label="圆角">
              <NumberInput value={el.radius ?? 0} onChange={(v) => upd((d) => { if (d.type === 'image') d.radius = v })} suffix="px" max={200} />
            </Field>
            <Field label="裁剪形状">
              <SelectInput
                value={el.clipShape ?? ''}
                onChange={(v) => upd((d) => { if (d.type === 'image') d.clipShape = v || undefined })}
                options={[
                  { value: '', label: '无' },
                  { value: 'ellipse', label: '圆形' },
                  { value: 'roundRect', label: '圆角矩形' },
                  { value: 'triangle', label: '三角形' },
                  { value: 'diamond', label: '菱形' },
                  { value: 'hexagon', label: '六边形' },
                  { value: 'star5', label: '五角星' },
                  { value: 'heart', label: '心形' },
                  { value: 'cloud', label: '云朵' },
                ]}
              />
            </Field>
          </div>
          <Field label="翻转">
            <div className="flex gap-0.5">
              <IconButton icon={<FlipHorizontal size={14} />} title="水平翻转" onClick={() => upd((d) => { if (d.type === 'image') d.flipH = !d.flipH })} active={el.flipH} />
              <IconButton icon={<FlipVertical size={14} />} title="垂直翻转" onClick={() => upd((d) => { if (d.type === 'image') d.flipV = !d.flipV })} active={el.flipV} />
            </div>
          </Field>
          <ToggleField
            label="阴影"
            checked={Boolean(el.shadow)}
            onChange={(e) => upd((d) => {
              if (d.type !== 'image') return
              d.shadow = e ? { h: 4, v: 4, blur: 10, color: 'rgba(0,0,0,0.3)' } : undefined
            })}
          />
          <Field label="边框">
            <div className="flex items-center gap-1">
              <ColorInput allowEmpty value={el.outline?.color} onChange={(v) => upd((d) => { if (d.type === 'image') d.outline = { color: v, width: 2, style: 'solid' } })} />
              <NumberInput value={el.outline?.width ?? 0} min={0} max={20} onChange={(v) => upd((d) => { if (d.type === 'image') d.outline = { color: el.outline?.color ?? '#333', width: v, style: 'solid' } })} suffix="px" />
            </div>
          </Field>
          <ImageCropSection el={el} />
        </Section>
      )}

      {/* 形状元素 */}
      {el.type === 'shape' && (
        <>
          <Section title="填充与边框" defaultOpen>
            <Field label="填充">
              <ColorInput allowEmpty value={typeof el.fill === 'string' ? el.fill : undefined} onChange={(v) => upd((d) => { if (d.type === 'shape') d.fill = v })} />
            </Field>
            <ToggleField
              label="渐变填充"
              checked={typeof el.fill === 'object'}
              onChange={(e) => upd((d) => {
                if (d.type !== 'shape') return
                d.fill = e
                  ? { type: 'linear', colors: [{ pos: 0, color: '#d14424' }, { pos: 100, color: '#f5c56b' }], rotate: 90 }
                  : '#d14424'
              })}
            />
            {typeof el.fill === 'object' && el.fill !== null && (
              <>
                <div className="grid grid-cols-2 gap-2">
                  <Field label="起色">
                    <ColorInput value={el.fill.colors[0]?.color} onChange={(v) => upd((d) => {
                      if (d.type !== 'shape' || typeof d.fill !== 'object' || d.fill === null) return
                      const colors = [...d.fill.colors]
                      colors[0] = { ...colors[0], color: v }
                      d.fill = { ...d.fill, colors }
                    })} />
                  </Field>
                  <Field label="止色">
                    <ColorInput value={el.fill.colors[1]?.color} onChange={(v) => upd((d) => {
                      if (d.type !== 'shape' || typeof d.fill !== 'object' || d.fill === null) return
                      const colors = [...d.fill.colors]
                      colors[1] = { ...colors[1] ?? { pos: 100 }, color: v }
                      d.fill = { ...d.fill, colors }
                    })} />
                  </Field>
                </div>
                <Field label="渐变角">
                  <NumberInput value={typeof el.fill === 'object' && el.fill !== null ? el.fill.rotate ?? 90 : 90} onChange={(v) => upd((d) => {
                    if (d.type !== 'shape' || typeof d.fill !== 'object' || d.fill === null) return
                    d.fill = { ...d.fill, rotate: v }
                  })} suffix="°" />
                </Field>
              </>
            )}
            <div className="grid grid-cols-2 gap-2">
              <Field label="边框色">
                <ColorInput allowEmpty value={el.outline?.color} onChange={(v) => upd((d) => { if (d.type === 'shape') d.outline = { ...(d.outline ?? { width: 1, style: 'solid' }), color: v } })} />
              </Field>
              <Field label="边框宽">
                <NumberInput value={el.outline?.width ?? 0} min={0} max={20} onChange={(v) => upd((d) => { if (d.type === 'shape') d.outline = { ...(d.outline ?? { color: '#333', style: 'solid' }), width: v } })} suffix="px" />
              </Field>
            </div>
            <ToggleField
              label="阴影"
              checked={Boolean(el.shadow)}
              onChange={(e) => upd((d) => {
                if (d.type !== 'shape') return
                d.shadow = e ? { h: 3, v: 3, blur: 8, color: 'rgba(0,0,0,0.3)' } : undefined
              })}
            />
          </Section>
          <Section title="内嵌文本" defaultOpen={false}>
            <Field label="文字">
              <input
                type="text"
                value={el.text ?? ''}
                onChange={(e) => upd((d) => { if (d.type === 'shape') d.text = e.target.value })}
                className="w-full rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
              />
            </Field>
            <div className="grid grid-cols-2 gap-2">
              <Field label="字号">
                <NumberInput value={el.fontSize ?? 18} min={10} max={96} onChange={(v) => upd((d) => { if (d.type === 'shape') d.fontSize = v })} />
              </Field>
              <Field label="文字色">
                <ColorInput value={el.defaultColor} onChange={(v) => upd((d) => { if (d.type === 'shape') d.defaultColor = v })} />
              </Field>
            </div>
            <Field label="对齐">
              <SelectInput
                value={el.align ?? 'center'}
                onChange={(v) => upd((d) => { if (d.type === 'shape') d.align = v as ShapeElement['align'] })}
                options={[{ value: 'left', label: '左' }, { value: 'center', label: '中' }, { value: 'right', label: '右' }]}
              />
            </Field>
          </Section>
        </>
      )}

      {/* 线条元素 */}
      {el.type === 'line' && (
        <Section title="线条样式" defaultOpen>
          <div className="grid grid-cols-2 gap-2">
            <Field label="颜色">
              <ColorInput value={el.color} onChange={(v) => upd((d) => { if (d.type === 'line') d.color = v })} />
            </Field>
            <Field label="线宽">
              <NumberInput value={el.lineWidth} min={1} max={20} onChange={(v) => upd((d) => { if (d.type === 'line') d.lineWidth = v })} suffix="px" />
            </Field>
          </div>
          <LineStyleSection el={el} upd={upd} />
        </Section>
      )}

      {/* 表格元素 */}
      {el.type === 'table' && <TableSection el={el} />}

      {/* 图表元素 */}
      {el.type === 'chart' && <ChartStyleSection el={el} upd={upd} />}

      {/* 公式元素 */}
      {el.type === 'formula' && (
        <Section title="公式" defaultOpen>
          <Field label="LaTeX">
            <button
              className="flex w-full items-center gap-2 rounded border border-gray-200 bg-gray-50 px-3 py-2 text-left text-sm hover:border-[#d14424] hover:bg-white"
              onClick={() => setFormulaDialogOpen(true)}
            >
              <Pencil size={14} className="shrink-0 text-gray-400" />
              <span className="flex-1 truncate font-mono text-xs text-gray-600">{el.latex}</span>
              <span className="shrink-0 text-xs text-[#d14424]">编辑</span>
            </button>
          </Field>
          {formulaDialogOpen && (
            <FormulaEditorDialog
              value={el.latex}
              onConfirm={(latex) => { upd((d) => { if (d.type === 'formula') d.latex = latex }); setFormulaDialogOpen(false) }}
              onClose={() => setFormulaDialogOpen(false)}
            />
          )}
          <Field label="字号">
            <NumberInput value={el.fontSize ?? 32} min={12} max={120} onChange={(v) => upd((d) => { if (d.type === 'formula') d.fontSize = v })} />
          </Field>
          <Field label="颜色">
            <ColorInput value={el.color} onChange={(v) => upd((d) => { if (d.type === 'formula') d.color = v })} />
          </Field>
        </Section>
      )}

      {/* 音视频元素 */}
      {(el.type === 'video' || el.type === 'audio') && (
        <Section title="媒体" defaultOpen>
          <ToggleField label="循环" checked={el.loop} onChange={(e) => upd((d) => { if (d.type === 'video' || d.type === 'audio') d.loop = e })} />
          <ToggleField label="自动播放" checked={el.autoPlay} onChange={(e) => upd((d) => { if (d.type === 'video' || d.type === 'audio') d.autoPlay = e })} />
        </Section>
      )}

      {/* 通用属性 */}
      <Section title="通用属性" defaultOpen={false}>
        <Field label="不透明度">
          <div className="flex items-center gap-2">
            <input
              type="range"
              min={5}
              max={100}
              value={Math.round((el.opacity ?? 1) * 100)}
              onChange={(e) => upd((d) => { d.opacity = Number(e.target.value) / 100 })}
              className="h-1 flex-1 cursor-pointer accent-[#d14424]"
            />
            <NumberInput value={Math.round((el.opacity ?? 1) * 100)} min={5} max={100} onChange={(v) => upd((d) => { d.opacity = v / 100 })} suffix="%" />
          </div>
        </Field>
        <ToggleField label="锁定" icon={<Lock size={12} />} checked={el.lock ?? false} onChange={(v) => upd((d) => { d.lock = v })} />
        <Field label="超链接">
          <div className="flex items-center gap-1.5">
            <Link2 size={13} className="shrink-0 text-gray-400" />
            <input
              type="text"
              placeholder="https:// 或页面id"
              value={el.link?.target ?? ''}
              onChange={(e) => upd((d) => {
                const v = e.target.value
                d.link = v ? { type: v.startsWith('http') || v.startsWith('/') ? 'web' : 'slide', target: v } : undefined
              })}
              className="min-w-0 flex-1 rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
            />
          </div>
        </Field>
        {el.type === 'shape' && (
          <Field label="形状">
            <SelectInput
              value={el.shapeKey}
              onChange={(v) => pushThen(() => useEditorStore.getState().updateElements([el.id], (d) => { if (d.type === 'shape') d.shapeKey = v }))}
              options={Object.entries(SHAPE_PATHS).map(([key, meta]) => ({ value: key, label: meta.name }))}
            />
          </Field>
        )}
      </Section>
    </div>
  )
}

/** 表格编辑：行列管理 + 单元格选择与样式 */
function TableSection({ el }: { el: PPTElement }) {
  const [sel, setSel] = useState<{ r: number; c: number }>({ r: 0, c: 0 })
  if (el.type !== 'table') return null
  const rows = el.cells.length
  const cols = Math.max(...el.cells.map((row) => row?.length ?? 0), 1)
  const r = Math.min(sel.r, rows - 1)
  const c = Math.min(sel.c, cols - 1)
  const cell = el.cells[r]?.[c] ?? null

  const upd = (fn: (draft: PPTElement) => void) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    store.updateElements([el.id], fn)
  }

  const setCell = (patch: Partial<TableCell>) => upd((d) => {
    if (d.type !== 'table') return
    const target = d.cells[r]?.[c]
    if (!target) return
    if (patch.text !== undefined) target.text = patch.text
    if (patch.colspan !== undefined) target.colspan = patch.colspan
    if (patch.rowspan !== undefined) target.rowspan = patch.rowspan
  })

  const setCellStyle = (patch: Partial<NonNullable<TableCell['style']>>) => upd((d) => {
    if (d.type !== 'table') return
    const target = d.cells[r]?.[c]
    if (!target) return
    if (!target.style) target.style = { color: '#333333', align: 'left', valign: 'middle' }
    if (patch.bold !== undefined) target.style.bold = patch.bold
    if (patch.italic !== undefined) target.style.italic = patch.italic
    if (patch.underline !== undefined) target.style.underline = patch.underline
    if (patch.color !== undefined) target.style.color = patch.color
    if (patch.backcolor !== undefined) target.style.backcolor = patch.backcolor
    if (patch.fontsize !== undefined) target.style.fontsize = patch.fontsize
    if (patch.fontface !== undefined) target.style.fontface = patch.fontface
    if (patch.align !== undefined) target.style.align = patch.align
    if (patch.valign !== undefined) target.style.valign = patch.valign
  })

  return (
    <div className="mb-2">
      <Section title="表格结构" defaultOpen>
        <div className="flex flex-wrap gap-1">
          <ActionButton label="加行" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            d.cells.push(Array.from({ length: d.colSizes.length }, () => ({ text: '' })))
            d.rowSizes = d.rowSizes.map((v) => v * (d.rowSizes.length / (d.rowSizes.length + 1)))
            d.rowSizes.push(1 / (d.rowSizes.length + 1))
          })} />
          <ActionButton label="加列" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) row.push({ text: '' })
            d.colSizes.push(1 / (d.colSizes.length + 1))
            d.colSizes = d.colSizes.map((v) => v * (d.colSizes.length / (d.colSizes.length + 1)))
          })} />
          <ActionButton label="删行" danger onClick={() => upd((d) => {
            if (d.type !== 'table' || d.cells.length <= 1) return
            d.cells.pop()
            d.rowSizes.pop()
          })} />
          <ActionButton label="删列" danger onClick={() => upd((d) => {
            if (d.type !== 'table' || d.colSizes.length <= 1) return
            for (const row of d.cells) row.pop()
            d.colSizes.pop()
          })} />
        </div>
        <div className="mt-2 flex items-center gap-2">
          <Field label="表头色" className="!w-auto !mb-0">
            <ColorInput
              value={el.theme?.headColor ?? '#d14424'}
              onChange={(v) => upd((d) => {
                if (d.type !== 'table') return
                if (!d.theme) d.theme = { color: [], headColor: v }
                else d.theme.headColor = v
              })}
              small
            />
          </Field>
          <ActionButton label="应用表头" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            const headColor = d.theme?.headColor ?? '#d14424'
            if (!d.theme) {
              d.theme = { color: [], headColor }
            } else {
              d.theme.headColor = headColor
            }
            const headRow = d.cells[0]
            if (headRow) {
              for (const tc of headRow) {
                if (tc) {
                  if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                  tc.style.backcolor = headColor
                  tc.style.color = '#ffffff'
                  tc.style.bold = true
                  tc.style.align = 'center'
                }
              }
            }
          })} />
        </div>
      </Section>

      <Section title="表格边框" defaultOpen>
        <div className="grid grid-cols-2 gap-2">
          <Field label="边框色">
            <ColorInput
              value={el.outline?.color ?? '#cccccc'}
              onChange={(v) => upd((d) => {
                if (d.type !== 'table') return
                d.outline = { ...(d.outline ?? { width: 1, style: 'solid' }), color: v }
              })}
              small
            />
          </Field>
          <Field label="边框宽">
            <NumberInput
              value={el.outline?.width ?? 1}
              min={0}
              max={10}
              onChange={(v) => upd((d) => {
                if (d.type !== 'table') return
                d.outline = { ...(d.outline ?? { color: '#cccccc', style: 'solid' }), width: v }
              })}
              suffix="px"
            />
          </Field>
        </div>
        <Field label="边框样式">
          <SelectInput
            value={el.outline?.style ?? 'solid'}
            onChange={(v) => upd((d) => {
              if (d.type !== 'table') return
              d.outline = { ...(d.outline ?? { color: '#cccccc', width: 1 }), style: v }
            })}
            options={[
              { value: 'solid', label: '实线' },
              { value: 'dashed', label: '虚线' },
              { value: 'dotted', label: '点线' },
            ]}
          />
        </Field>
      </Section>

      <Section title="批量字体设置" defaultOpen={false}>
        <div className="grid grid-cols-2 gap-2">
          <Field label="字色">
            <ColorInput value="#333333" onChange={(v) => upd((d) => {
              if (d.type !== 'table') return
              for (const row of d.cells) for (const tc of row) {
                if (tc) {
                  if (!tc.style) tc.style = { color: v, align: 'left', valign: 'middle' }
                  else tc.style.color = v
                }
              }
            })} small />
          </Field>
          <Field label="字号">
            <NumberInput value={14} min={10} max={72} onChange={(v) => upd((d) => {
              if (d.type !== 'table') return
              for (const row of d.cells) for (const tc of row) {
                if (tc) {
                  if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                  tc.style.fontsize = v
                }
              }
            })} />
          </Field>
        </div>
        <Field label="字体">
          <SelectInput value="" onChange={(v) => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc) {
                if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                tc.style.fontface = v
              }
            }
          })} options={[
            { value: '', label: '默认' },
            { value: '微软雅黑', label: '微软雅黑' },
            { value: '宋体', label: '宋体' },
            { value: '黑体', label: '黑体' },
            { value: '楷体', label: '楷体' },
            { value: '仿宋', label: '仿宋' },
            { value: 'Arial', label: 'Arial' },
            { value: 'Times New Roman', label: 'Times New Roman' },
          ]} />
        </Field>
        <div className="grid grid-cols-2 gap-2">
          <Field label="水平">
            <SelectInput value="" onChange={(v) => upd((d) => {
              if (d.type !== 'table') return
              for (const row of d.cells) for (const tc of row) {
                if (tc) {
                  if (!tc.style) tc.style = { color: '#333333', align: v as NonNullable<TableCell['style']>['align'], valign: 'middle' }
                  else tc.style.align = v as NonNullable<TableCell['style']>['align']
                }
              }
            })} options={[
              { value: 'left', label: '左对齐' },
              { value: 'center', label: '居中' },
              { value: 'right', label: '右对齐' },
            ]} />
          </Field>
          <Field label="垂直">
            <SelectInput value="" onChange={(v) => upd((d) => {
              if (d.type !== 'table') return
              for (const row of d.cells) for (const tc of row) {
                if (tc) {
                  if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: v as NonNullable<TableCell['style']>['valign'] }
                  else tc.style.valign = v as NonNullable<TableCell['style']>['valign']
                }
              }
            })} options={[
              { value: 'top', label: '上对齐' },
              { value: 'middle', label: '居中' },
              { value: 'bottom', label: '下对齐' },
            ]} />
          </Field>
        </div>
        <div className="mt-1 flex flex-wrap gap-1">
          <ActionButton label="全部加粗" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc) {
                if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                tc.style.bold = true
              }
            }
          })} />
          <ActionButton label="取消加粗" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc?.style) tc.style.bold = false
            }
          })} />
          <ActionButton label="全部斜体" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc) {
                if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                tc.style.italic = true
              }
            }
          })} />
          <ActionButton label="取消斜体" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc?.style) tc.style.italic = false
            }
          })} />
          <ActionButton label="全部下划线" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc) {
                if (!tc.style) tc.style = { color: '#333333', align: 'left', valign: 'middle' }
                tc.style.underline = true
              }
            }
          })} />
          <ActionButton label="取消下划线" onClick={() => upd((d) => {
            if (d.type !== 'table') return
            for (const row of d.cells) for (const tc of row) {
              if (tc?.style) tc.style.underline = false
            }
          })} />
        </div>
      </Section>

      <Section title="单元格编辑" defaultOpen>
        <div className="mb-2 grid gap-0.5" style={{ gridTemplateColumns: `repeat(${cols}, 1fr)` }}>
          {Array.from({ length: rows }, (_, rr) => (
            Array.from({ length: cols }, (_, cc) => (
              <button
                key={`${rr}-${cc}`}
                className={`h-6 border text-[10px] ${rr === r && cc === c ? 'border-[#d14424] bg-[#fbeae5]' : 'border-gray-200 bg-white hover:bg-gray-50'}`}
                onClick={() => setSel({ r: rr, c: cc })}
              >
                {`${rr + 1},${cc + 1}`}
              </button>
            ))
          ))}
        </div>

        {cell && (
          <>
            <div className="mb-1 text-xs text-gray-400">单元格 {r + 1},{c + 1}</div>
            <textarea
              className="mb-2 h-14 w-full rounded border border-gray-200 p-1.5 text-xs focus:border-[#d14424] focus:outline-none"
              value={cell.text}
              onChange={(e) => setCell({ text: e.target.value })}
            />
            <div className="grid grid-cols-2 gap-2">
              <ToggleField label="加粗" checked={cell.style?.bold ?? false} onChange={(v) => setCellStyle({ bold: v })} />
              <ToggleField label="斜体" checked={cell.style?.italic ?? false} onChange={(v) => setCellStyle({ italic: v })} />
              <ToggleField label="下划线" checked={cell.style?.underline ?? false} onChange={(v) => setCellStyle({ underline: v })} />
              <Field label="字色">
                <ColorInput value={cell.style?.color} onChange={(v) => setCellStyle({ color: v })} small />
              </Field>
              <Field label="底色">
                <ColorInput allowEmpty value={cell.style?.backcolor} onChange={(v) => setCellStyle({ backcolor: v })} small />
              </Field>
              <Field label="字号">
                <NumberInput value={cell.style?.fontsize} min={10} max={72} onChange={(v) => setCellStyle({ fontsize: v })} />
              </Field>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <Field label="水平">
                <SelectInput
                  value={cell.style?.align ?? 'left'}
                  onChange={(v) => setCellStyle({ align: v as NonNullable<TableCell['style']>['align'] })}
                  options={[{ value: 'left', label: '左' }, { value: 'center', label: '中' }, { value: 'right', label: '右' }]}
                />
              </Field>
              <Field label="垂直">
                <SelectInput
                  value={cell.style?.valign ?? 'middle'}
                  onChange={(v) => setCellStyle({ valign: v as NonNullable<TableCell['style']>['valign'] })}
                  options={[{ value: 'top', label: '上' }, { value: 'middle', label: '中' }, { value: 'bottom', label: '下' }]}
                />
              </Field>
            </div>
            <div className="mt-1 flex gap-1">
              <ActionButton label="向右合并" onClick={() => upd((d) => {
                if (d.type !== 'table') return
                const target = d.cells[r]?.[c]
                const right = d.cells[r]?.[c + 1]
                if (!target || !right || target.colspan && target.colspan > 1) return
                target.text = `${target.text} ${right.text}`.trim()
                target.colspan = (target.colspan ?? 1) + 1
                d.cells[r].splice(c + 1, 1)
              })} />
              <ActionButton label="拆分" disabled={!cell.colspan || cell.colspan <= 1} onClick={() => upd((d) => {
                if (d.type !== 'table') return
                const target = d.cells[r]?.[c]
                if (!target || !target.colspan || target.colspan <= 1) return
                target.colspan = undefined
                d.cells[r].splice(c + 1, 0, { text: '' })
              })} />
            </div>
          </>
        )}
      </Section>
    </div>
  )
}

/** 图片裁剪区域编辑（比例为源图的 0-100%） */
function ImageCropSection({ el }: { el: PPTElement }) {
  if (el.type !== 'image') return null
  const clip = el.clip ?? { x: 0, y: 0, w: 1, h: 1 }
  const setClip = (patch: Partial<{ x: number; y: number; w: number; h: number }>) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    store.updateElements([el.id], (d) => {
      if (d.type !== 'image') return
      const next = { ...clip, ...patch }
      next.w = Math.min(Math.max(next.w, 0.05), 1)
      next.h = Math.min(Math.max(next.h, 0.05), 1)
      next.x = Math.min(Math.max(next.x, 0), 1 - next.w)
      next.y = Math.min(Math.max(next.y, 0), 1 - next.h)
      d.clip = next
    })
  }
  const pct = (v: number) => Math.round(v * 100)
  return (
    <div className="mt-2 rounded border border-gray-100 bg-white p-2">
      <div className="mb-1.5 flex items-center justify-between">
        <span className="text-xs font-medium text-gray-600">裁剪区域（%）</span>
        <IconButton icon={<RotateCcw size={13} />} title="重置裁剪" onClick={() => {
          const store = useEditorStore.getState()
          store.pushHistory()
          store.updateElements([el.id], (d) => { if (d.type === 'image') d.clip = undefined })
        }} />
      </div>
      <div className="grid grid-cols-2 gap-1.5">
        {([['x', '左'], ['y', '上'], ['w', '宽'], ['h', '高']] as const).map(([key, label]) => (
          <label key={key} className="flex items-center gap-1 text-xs text-gray-500">
            {label}
            <input
              type="number"
              min={key === 'w' || key === 'h' ? 5 : 0}
              max={key === 'w' || key === 'h' ? 100 : 95}
              value={pct(clip[key])}
              onChange={(e) => setClip({ [key]: Number(e.target.value) / 100 })}
              className="w-full rounded border border-gray-200 px-1 py-0.5 text-xs focus:border-[#d14424] focus:outline-none"
            />
          </label>
        ))}
      </div>
    </div>
  )
}

/** 图表元素开关项（与 eflink-excel 的设置面板一致） */
const CHART_ELEMENT_ITEMS: { key: keyof ChartElements; label: string }[] = [
  { key: 'legend', label: '图例' },
  { key: 'gridLine', label: '网格线' },
  { key: 'axis', label: '坐标轴' },
  { key: 'trendline', label: '趋势线' },
  { key: 'chartTitle', label: '图表标题' },
  { key: 'dataLabel', label: '数据标签' },
  { key: 'axisTitle', label: '轴标题' },
]

/** 图表样式设置区（对齐 eflink-excel：类型/元素开关/标题/切换行列 + 配色方案 + 数据编辑） */
function ChartStyleSection({ el, upd }: { el: PPTElement; upd: (fn: (draft: PPTElement) => void) => void }) {
  if (el.type !== 'chart') return null
  const elements: ChartElements = el.elements ?? DEFAULT_CHART_ELEMENTS

  const updElement = (key: keyof ChartElements, v: boolean) =>
    upd((d) => { if (d.type === 'chart') d.elements = { ...(d.elements ?? DEFAULT_CHART_ELEMENTS), [key]: v } })

  // 切换行列：系列名变类别、类别变系列
  const transpose = () => upd((d) => {
    if (d.type !== 'chart') return
    const { labels, series } = d.data
    d.data.labels = series.map((s) => s.name)
    d.data.series = labels.map((l, i) => ({ name: l || `系列${i + 1}`, values: series.map((s) => s.values[i] ?? 0) }))
  })

  const applyPalette = (palette: string[]) => upd((d) => {
    if (d.type !== 'chart') return
    d.chartColors = palette.slice(0, d.data.series.length)
  })

  return (
    <>
      <Section title="图表设置" defaultOpen>
        <Field label="类型">
          <GroupedSelectInput
            testId="chart-type"
            value={el.chartType}
            onChange={(v) => upd((d) => { if (d.type === 'chart') d.chartType = v as ChartElement['chartType'] })}
            groups={chartTypeGroups}
          />
        </Field>
        <div>
          <div className="mb-1 text-xs text-gray-500">图表元素</div>
          <div className="grid grid-cols-2 gap-x-2">
            {CHART_ELEMENT_ITEMS.map(({ key, label }) => (
              <ToggleField key={key} label={label} checked={elements[key]} onChange={(e) => updElement(key, e)} />
            ))}
          </div>
        </div>
        {elements.chartTitle && (
          <Field label="标题">
            <input
              type="text"
              value={el.title ?? ''}
              onChange={(e) => upd((d) => { if (d.type === 'chart') d.title = e.target.value })}
              className="w-full rounded border border-gray-200 px-1.5 py-1 text-xs focus:border-[#d14424] focus:outline-none"
            />
          </Field>
        )}
        <Field label="行列">
          <ActionButton label="切换行列" onClick={transpose} />
        </Field>
        <Field label="系列">
          <div className="flex gap-1">
            <ActionButton label="添加" onClick={() => upd((d) => {
              if (d.type !== 'chart') return
              const n = d.data.series.length + 1
              d.data.series.push({ name: `系列${n}`, values: d.data.labels.map(() => 0) })
              if (d.chartColors) d.chartColors = [...d.chartColors, DEFAULT_CHART_COLORS[(n - 1) % DEFAULT_CHART_COLORS.length]]
            })} />
            <ActionButton label="减少" danger disabled={el.data.series.length <= 1} onClick={() => upd((d) => {
              if (d.type !== 'chart' || d.data.series.length <= 1) return
              d.data.series.pop()
              if (d.chartColors) d.chartColors = d.chartColors.slice(0, d.data.series.length)
            })} />
          </div>
        </Field>
        <DataEditor el={el} />
      </Section>
      <Section title="配色方案" defaultOpen>
        <div className="flex flex-col gap-2">
          {CHART_PALETTES.map((palette, index) => {
            const active = el.paletteIndex === index
            return (
              <button
                key={index}
                type="button"
                title={`配色方案 ${index + 1}`}
                onClick={() => applyPalette(palette)}
                className={`flex items-center gap-1 rounded-md border px-2 py-1.5 transition-colors ${
                  active ? 'border-[#d14424] ring-1 ring-[#d14424]' : 'border-gray-200 hover:border-[#d14424]/40'
                }`}
              >
                {palette.map((color) => (
                  <span key={color} className="h-4 w-4 rounded-[3px]" style={{ backgroundColor: color }} />
                ))}
              </button>
            )
          })}
        </div>
        <Field label="微调">
          <div className="flex flex-wrap gap-1">
            {el.data.series.map((_s, i) => (
              <ColorInput
                key={i}
                value={(el.chartColors ?? DEFAULT_CHART_COLORS)[i]}
                onChange={(v) => upd((d) => {
                  if (d.type !== 'chart') return
                  const colors = [...(d.chartColors ?? DEFAULT_CHART_COLORS.slice(0, d.data.series.length))]
                  colors[i] = v
                  d.chartColors = colors
                })}
                small
              />
            ))}
          </div>
        </Field>
      </Section>
    </>
  )
}

/** 图表数据表格编辑器 */
function DataEditor({ el }: { el: PPTElement }) {
  if (el.type !== 'chart') return null
  const { labels, series } = el.data

  const upd = (fn: (d: PPTElement) => void) => {
    const store = useEditorStore.getState()
    store.pushHistory()
    store.updateElements([el.id], fn)
  }

  const updateLabel = (i: number, value: string) =>
    upd((d) => { if (d.type === 'chart') d.data.labels[i] = value })

  const updateValue = (si: number, li: number, value: number) =>
    upd((d) => { if (d.type === 'chart') d.data.series[si].values[li] = value })

  const updateSeriesName = (si: number, value: string) =>
    upd((d) => { if (d.type === 'chart') d.data.series[si].name = value })

  const addRow = () =>
    upd((d) => {
      if (d.type !== 'chart') return
      const idx = d.data.labels.length + 1
      d.data.labels.push(`类别${idx}`)
      d.data.series.forEach((s) => s.values.push(0))
    })

  const removeRow = (i: number) =>
    upd((d) => {
      if (d.type !== 'chart' || d.data.labels.length <= 1) return
      d.data.labels.splice(i, 1)
      d.data.series.forEach((s) => s.values.splice(i, 1))
    })

  const addSeries = () =>
    upd((d) => {
      if (d.type !== 'chart') return
      const n = d.data.series.length + 1
      d.data.series.push({ name: `系列${n}`, values: d.data.labels.map(() => 0) })
      if (d.chartColors) d.chartColors = [...d.chartColors, DEFAULT_CHART_COLORS[(n - 1) % DEFAULT_CHART_COLORS.length]]
    })

  const removeSeries = (si: number) =>
    upd((d) => {
      if (d.type !== 'chart' || d.data.series.length <= 1) return
      d.data.series.splice(si, 1)
      if (d.chartColors) d.chartColors = d.chartColors.filter((_, i) => i !== si)
    })

  return (
    <div className="mt-2">
      <div className="mb-1 flex items-center justify-between">
        <span className="text-xs text-gray-500">数据编辑</span>
        <div className="flex gap-1">
          <ActionButton label="+ 行" onClick={addRow} />
          <ActionButton label="+ 列" onClick={addSeries} />
        </div>
      </div>
      <div className="overflow-x-auto rounded border border-gray-200">
        <table className="w-full border-collapse text-xs">
          <thead>
            <tr className="bg-gray-50">
              <th className="border-b border-gray-200 px-1.5 py-1 text-left text-[10px] font-medium text-gray-400">标签</th>
              {series.map((s, si) => (
                <th key={si} className="border-b border-l border-gray-200 px-1 py-1 text-center text-[10px] font-medium text-gray-400">
                  <input
                    type="text"
                    value={s.name}
                    onChange={(e) => updateSeriesName(si, e.target.value)}
                    className="w-full bg-transparent text-center text-[10px] text-gray-600 focus:outline-none"
                  />
                </th>
              ))}
              {series.length > 1 && <th className="w-6 border-b border-gray-200" />}
            </tr>
          </thead>
          <tbody>
            {labels.map((l, li) => (
              <tr key={li} className="hover:bg-gray-50/50">
                <td className="border-b border-gray-100 px-1 py-0.5">
                  <input
                    type="text"
                    value={l}
                    onChange={(e) => updateLabel(li, e.target.value)}
                    className="w-full bg-transparent text-xs text-gray-700 focus:outline-none"
                  />
                </td>
                {series.map((s, si) => (
                  <td key={si} className="border-b border-l border-gray-100 px-1 py-0.5 text-center">
                    <input
                      type="number"
                      value={s.values[li] ?? 0}
                      onChange={(e) => updateValue(si, li, Number(e.target.value) || 0)}
                      className="w-full bg-transparent text-center text-xs text-gray-600 focus:outline-none"
                    />
                  </td>
                ))}
                {series.length > 1 && (
                  <td className="border-b border-gray-100 px-0.5 py-0.5 text-center">
                    <button
                      type="button"
                      className="text-gray-300 hover:text-red-400"
                      title="删除此行"
                      onClick={() => removeRow(li)}
                      disabled={labels.length <= 1}
                    >
                      <Trash2 size={11} />
                    </button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="mt-1 flex items-center justify-between">
        <div className="flex gap-1">
          {series.length > 1 && series.map((_s, si) => (
            <button
              key={si}
              type="button"
              className="text-gray-300 hover:text-red-400"
              title={`删除系列 ${_s.name || si + 1}`}
              onClick={() => removeSeries(si)}
            >
              <Trash2 size={11} />
            </button>
          ))}
        </div>
        {series.length > 1 && (
          <span className="text-[10px] text-gray-400">
            共 {labels.length} 行 × {series.length} 列
          </span>
        )}
      </div>
    </div>
  )
}

/* 排列/对齐/组合/删除 */
function ArrangeSection({ selectedIds }: { selectedIds: string[] }) {
  if (!selectedIds.length) return null
  const store = useEditorStore.getState()
  const slideW = store.presentation.width
  const slideH = Math.round(slideW / store.presentation.viewportRatio)

  const align = (mode: AlignMode) => {
    const s = useEditorStore.getState()
    s.pushHistory()
    s.updateElements(selectedIds, () => {
      const els = s.presentation.slides[s.slideIndex].elements.filter((el) => selectedIds.includes(el.id))
      if (!els.length) return
      const multi = els.length > 1
      const minX = multi ? Math.min(...els.map((e) => e.x)) : 0
      const maxX = multi ? Math.max(...els.map((e) => e.x + e.w)) : slideW
      const minY = multi ? Math.min(...els.map((e) => e.y)) : 0
      const maxY = multi ? Math.max(...els.map((e) => e.y + e.h)) : slideH
      // 通过 immer 代理修改，确保触发更新
      const target = s.presentation.slides[s.slideIndex].elements
      for (const el of target) {
        if (!selectedIds.includes(el.id)) continue
        switch (mode) {
          case 'left': el.x = minX; break
          case 'hcenter': el.x = minX + (maxX - minX) / 2 - el.w / 2; break
          case 'right': el.x = maxX - el.w; break
          case 'top': el.y = minY; break
          case 'vcenter': el.y = minY + (maxY - minY) / 2 - el.h / 2; break
          case 'bottom': el.y = maxY - el.h; break
        }
      }
    })
  }

  const distribute = (mode: 'horizontal' | 'vertical') => {
    const s = useEditorStore.getState()
    s.pushHistory()
    s.updateElements(selectedIds, () => {
      const els = s.presentation.slides[s.slideIndex].elements.filter((el) => selectedIds.includes(el.id))
      if (els.length < 3) return
      const isH = mode === 'horizontal'
      const sorted = [...els].sort((a, b) => (isH ? a.x - b.x : a.y - b.y))
      const first = sorted[0]
      const last = sorted[sorted.length - 1]
      const totalSize = sorted.reduce((acc, el) => acc + (isH ? el.w : el.h), 0)
      const span = isH ? (last.x + last.w) - first.x : (last.y + last.h) - first.y
      const gap = (span - totalSize) / (sorted.length - 1)
      const target = s.presentation.slides[s.slideIndex].elements
      let cursor = isH ? first.x : first.y
      for (const el of sorted) {
        const t = target.find((te) => te.id === el.id)
        if (!t) continue
        if (isH) {
          t.x = cursor
          cursor += el.w + gap
        } else {
          t.y = cursor
          cursor += el.h + gap
        }
      }
    })
  }

  return (
    <Section title="排列与对齐">
      <ToolGroup label="对齐">
        <IconButton icon={<AlignLeft size={14} />} title="左对齐" onClick={() => align('left')} />
        <IconButton icon={<AlignCenter size={14} />} title="水平居中" onClick={() => align('hcenter')} />
        <IconButton icon={<AlignRight size={14} />} title="右对齐" onClick={() => align('right')} />
        <ToolDivider />
        <IconButton icon={<AlignVerticalJustifyStart size={14} />} title="顶对齐" onClick={() => align('top')} />
        <IconButton icon={<AlignVerticalJustifyCenter size={14} />} title="垂直居中" onClick={() => align('vcenter')} />
        <IconButton icon={<AlignVerticalJustifyEnd size={14} />} title="底对齐" onClick={() => align('bottom')} />
      </ToolGroup>

      <ToolGroup label="分布">
        <IconButton icon={<AlignHorizontalDistributeCenter size={14} />} title="横向分布" onClick={() => distribute('horizontal')} />
        <IconButton icon={<AlignVerticalDistributeCenter size={14} />} title="纵向分布" onClick={() => distribute('vertical')} />
      </ToolGroup>

      <ToolGroup label="图层">
        <IconButton icon={<IconLayerForward size={14} />} title="上移一层" onClick={() => { store.pushHistory(); store.setElementLevel(selectedIds, 'up') }} />
        <IconButton icon={<IconLayerBackward size={14} />} title="下移一层" onClick={() => { store.pushHistory(); store.setElementLevel(selectedIds, 'down') }} />
        <IconButton icon={<IconLayerToFront size={14} />} title="置顶" onClick={() => { store.pushHistory(); store.setElementLevel(selectedIds, 'top') }} />
        <IconButton icon={<IconLayerToBack size={14} />} title="置底" onClick={() => { store.pushHistory(); store.setElementLevel(selectedIds, 'bottom') }} />
      </ToolGroup>

      <ToolGroup label="操作">
        {selectedIds.length > 1 && (
          <IconButton icon={<Group size={14} />} title="组合" onClick={() => { store.pushHistory(); store.setGroup(selectedIds, `g-${Date.now().toString(36)}`) }} />
        )}
        {selectedIds.length === 1 && useEditorStore.getState().presentation.slides[useEditorStore.getState().slideIndex].elements.find((el) => el.id === selectedIds[0])?.groupId && (
          <IconButton icon={<Ungroup size={14} />} title="取消组合" onClick={() => { store.pushHistory(); store.setGroup(selectedIds, undefined) }} />
        )}
        <IconButton icon={<Copy size={14} />} title="复制" onClick={() => store.duplicateElements(selectedIds)} />
        <IconButton icon={<Trash2 size={14} />} title="删除" danger onClick={() => { store.pushHistory(); store.removeElements(selectedIds) }} />
      </ToolGroup>
    </Section>
  )
}

/* ---------- 动画页签 ---------- */

function AnimationTab() {
  const selectedIds = useEditorStore((s) => s.selectedIds)
  const slide = useEditorStore((s) => s.presentation.slides[s.slideIndex])
  const store = useEditorStore.getState()

  if (!slide || selectedIds.length !== 1) {
    return <div className="text-xs text-gray-400">选中单个元素后可设置动画</div>
  }
  const el = slide.elements.find((e) => e.id === selectedIds[0])
  if (!el) return null
  const anim = el.anim

  return (
    <div>
      <div className="mb-2 rounded-lg bg-[#fbeae5] px-3 py-1.5 text-xs font-medium text-[#d14424]">
        动画 · {typeLabel(el)}
      </div>
      <Section title="动画设置" defaultOpen>
        <Field label="类别">
          <SelectInput
            testId="anim-type"
            value={anim?.type ?? ''}
            onChange={(v) => {
              store.pushHistory()
              store.updateElements([el.id], (d) => {
                d.anim = v
                  ? { type: v as never, effect: effectOptions[v as 'in' | 'emphasis' | 'out'][0]?.value ?? 'fade', duration: 600, delay: 0, trigger: 'click' }
                  : undefined
              })
            }}
            options={[{ value: '', label: '无' }, { value: 'in', label: '进入' }, { value: 'emphasis', label: '强调' }, { value: 'out', label: '退出' }]}
          />
        </Field>
        {anim && (
          <>
            <Field label="效果">
              <SelectInput
                testId="anim-effect"
                value={anim.effect}
                onChange={(v) => {
                  store.pushHistory()
                  store.updateElements([el.id], (d) => { if (d.anim) d.anim.effect = v })
                }}
                options={effectOptions[anim.type] ?? []}
              />
            </Field>
            <div className="grid grid-cols-2 gap-2">
              <Field label="时长">
                <NumberInput value={anim.duration} step={100} onChange={(v) => {
                  store.updateElements([el.id], (d) => { if (d.anim) d.anim.duration = v })
                }} suffix="ms" />
              </Field>
              <Field label="延迟">
                <NumberInput value={anim.delay} step={100} onChange={(v) => {
                  store.updateElements([el.id], (d) => { if (d.anim) d.anim.delay = v })
                }} suffix="ms" />
              </Field>
            </div>
            <Field label="触发">
              <SelectInput
                value={anim.trigger}
                onChange={(v) => {
                  store.pushHistory()
                  store.updateElements([el.id], (d) => { if (d.anim) d.anim.trigger = v as never })
                }}
                options={[{ value: 'click', label: '单击时' }, { value: 'withPrevious', label: '与上一动画同时' }, { value: 'afterPrevious', label: '上一动画之后' }]}
              />
            </Field>
          </>
        )}
      </Section>
    </div>
  )
}

/* ---------- 批注页签 ---------- */

function CommentTab() {
  const selectedIds = useEditorStore((s) => s.selectedIds)
  const slide = useEditorStore((s) => s.presentation.slides[s.slideIndex])
  const store = useEditorStore.getState()

  if (!slide) return null
  const el = selectedIds.length === 1 ? slide.elements.find((e) => e.id === selectedIds[0]) : undefined

  return (
    <div>
      {el ? (
        <Section title="元素批注" defaultOpen>
          <textarea
            className="h-20 w-full rounded border border-gray-200 p-2 text-xs focus:border-[#d14424] focus:outline-none"
            placeholder="输入批注内容…"
            defaultValue={el.comment ?? ''}
            onBlur={(e) => {
              store.pushHistory()
              store.updateElements([el.id], (d) => { d.comment = e.target.value || undefined })
            }}
          />
        </Section>
      ) : (
        <div className="mb-2 text-xs text-gray-400">选中单个元素后可添加元素批注</div>
      )}
      <Section title="演讲者备注（本页）" defaultOpen={false}>
        <textarea
          className="h-20 w-full rounded border border-gray-200 p-2 text-xs focus:border-[#d14424] focus:outline-none"
          placeholder="本页备注…"
          defaultValue={slide.note ?? ''}
          onBlur={(e) => {
            store.pushHistory()
            store.setSlideNote(e.target.value)
          }}
        />
      </Section>
    </div>
  )
}
