// 线条属性面板：带预览图标的可视化控件
import { useEffect, useRef, useState } from 'react'
import { ChevronDown } from 'lucide-react'
import type { LineArrow, LineElement, LineStyle, LineType, PPTElement } from '../../types/slides'

const VB = '0 0 40 24'

function SvgBase({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <svg viewBox={VB} className={className ?? 'h-5 w-10'} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      {children}
    </svg>
  )
}

function ArrowCap({ x, y, angle, fill }: { x: number; y: number; angle: number; fill?: boolean }) {
  return (
    <polygon
      points="0,-3.5 7,0 0,3.5"
      fill={fill ? 'currentColor' : 'none'}
      stroke={fill ? 'none' : 'currentColor'}
      strokeWidth="1.5"
      transform={`translate(${x} ${y}) rotate(${angle})`}
    />
  )
}

export function LineStylePreview({ style }: { style: LineStyle }) {
  const dash = style === 'dashed' ? '5 3' : style === 'dotted' ? '1.5 3' : undefined
  return (
    <SvgBase>
      <line x1="4" y1="12" x2="36" y2="12" strokeDasharray={dash} />
    </SvgBase>
  )
}

export function LineTypePreview({ type }: { type: LineType }) {
  if (type === 'broken') {
    return (
      <SvgBase>
        <polyline points="4,4 20,4 20,20 36,20" />
      </SvgBase>
    )
  }
  if (type === 'curve') {
    return (
      <SvgBase>
        <path d="M4 20 Q 20 20 36 4" />
      </SvgBase>
    )
  }
  return (
    <SvgBase>
      <line x1="4" y1="20" x2="36" y2="4" />
    </SvgBase>
  )
}

export function LineArrowPreview({ arrow, at }: { arrow: LineArrow; at: 'start' | 'end' }) {
  const v = normalizeArrow(arrow)
  if (v === 'none') {
    return (
      <SvgBase className="h-5 w-10 text-gray-300">
        <line x1="8" y1="12" x2="32" y2="12" strokeDasharray="3 3" />
      </SvgBase>
    )
  }
  if (v === 'dot') {
    return (
      <SvgBase>
        {at === 'start'
          ? <><circle cx="8" cy="12" r="3" fill="currentColor" stroke="none" /><line x1="12" y1="12" x2="36" y2="12" /></>
          : <><line x1="4" y1="12" x2="28" y2="12" /><circle cx="32" cy="12" r="3" fill="currentColor" stroke="none" /></>}
      </SvgBase>
    )
  }
  if (v === 'triangle') {
    return (
      <SvgBase>
        {at === 'start'
          ? <><ArrowCap x={8} y={12} angle={180} fill /><line x1="12" y1="12" x2="36" y2="12" /></>
          : <><line x1="4" y1="12" x2="28" y2="12" /><ArrowCap x={32} y={12} angle={0} fill /></>}
      </SvgBase>
    )
  }
  // arrow
  return (
    <SvgBase>
      {at === 'start'
        ? <><ArrowCap x={8} y={12} angle={180} /><line x1="12" y1="12" x2="36" y2="12" /></>
        : <><line x1="4" y1="12" x2="28" y2="12" /><ArrowCap x={32} y={12} angle={0} /></>}
    </SvgBase>
  )
}

export function normalizeArrow(v: LineArrow | undefined): 'none' | 'arrow' | 'triangle' | 'dot' {
  // 旧数据可能存有空串（toLineArrow 的 none 表示），空串已被 !v 覆盖
  if (!v || v === 'none') return 'none'
  return v
}

export function toLineArrow(v: 'none' | 'arrow' | 'triangle' | 'dot'): LineArrow {
  return v === 'none' ? '' : v
}

const STYLE_OPTIONS: Array<{ value: LineStyle; label: string }> = [
  { value: 'solid', label: '实线' },
  { value: 'dashed', label: '虚线' },
  { value: 'dotted', label: '点线' },
]

const TYPE_OPTIONS: Array<{ value: LineType; label: string }> = [
  { value: 'straight', label: '直线' },
  { value: 'broken', label: '折线' },
  { value: 'curve', label: '曲线' },
]

const ARROW_OPTIONS: Array<{ value: 'none' | 'arrow' | 'triangle' | 'dot'; label: string }> = [
  { value: 'none', label: '无' },
  { value: 'arrow', label: '箭头' },
  { value: 'triangle', label: '三角' },
  { value: 'dot', label: '圆点' },
]

/** 三选一图标分段按钮 */
function SegmentedPicker<T extends string>({
  value,
  options,
  onChange,
  renderIcon,
}: {
  value: T
  options: Array<{ value: T; label: string }>
  onChange: (v: T) => void
  renderIcon: (v: T) => React.ReactNode
}) {
  return (
    <div className="flex rounded-md border border-gray-200 bg-white p-0.5">
      {options.map((o) => (
        <button
          key={o.value}
          type="button"
          title={o.label}
          aria-label={o.label}
          aria-pressed={value === o.value}
          onClick={() => onChange(o.value)}
          className={`flex flex-1 flex-col items-center gap-0.5 rounded px-1 py-1 transition-colors ${
            value === o.value
              ? 'bg-[#fbeae5] text-[#d14424] ring-1 ring-[#d14424]/40'
              : 'text-gray-500 hover:bg-gray-50'
          }`}
        >
          {renderIcon(o.value)}
          <span className="text-[10px] leading-none">{o.label}</span>
        </button>
      ))}
    </div>
  )
}

/** 下拉式可视化选择（用于箭头等多选项） */
function VisualDropdown<T extends string>({
  value,
  options,
  onChange,
  renderIcon,
}: {
  value: T
  options: Array<{ value: T; label: string }>
  onChange: (v: T) => void
  renderIcon: (v: T) => React.ReactNode
}) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)
  const current = options.find((o) => o.value === value) ?? options[0]

  useEffect(() => {
    if (!open) return
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', onDoc)
    return () => document.removeEventListener('mousedown', onDoc)
  }, [open])

  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className="flex w-full items-center gap-2 rounded border border-gray-200 bg-white px-2 py-1.5 text-left hover:border-gray-300 focus:border-[#d14424] focus:outline-none"
      >
        <span className="shrink-0 text-gray-700">{renderIcon(current.value)}</span>
        <span className="min-w-0 flex-1 truncate text-xs text-gray-700">{current.label}</span>
        <ChevronDown size={14} className={`shrink-0 text-gray-400 transition-transform ${open ? 'rotate-180' : ''}`} />
      </button>
      {open && (
        <div className="absolute left-0 right-0 top-[calc(100%+4px)] z-50 overflow-hidden rounded-md border border-gray-200 bg-white py-1 shadow-lg">
          {options.map((o) => (
            <button
              key={o.value}
              type="button"
              onClick={() => { onChange(o.value); setOpen(false) }}
              className={`flex w-full items-center gap-2 px-2 py-1.5 text-left hover:bg-gray-50 ${
                o.value === value ? 'bg-[#fbeae5] text-[#d14424]' : 'text-gray-700'
              }`}
            >
              <span className="shrink-0">{renderIcon(o.value)}</span>
              <span className="text-xs">{o.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

export function LineStyleSection({
  el,
  upd,
}: {
  el: LineElement
  upd: (fn: (d: PPTElement) => void) => void
}) {
  return (
    <>
      <div className="grid grid-cols-2 gap-2">
        <div className="col-span-2">
          <div className="mb-1 text-xs text-gray-500">样式</div>
          <SegmentedPicker
            value={el.lineStyle ?? 'solid'}
            options={STYLE_OPTIONS}
            onChange={(v) => upd((d) => { if (d.type === 'line') d.lineStyle = v })}
            renderIcon={(v) => <LineStylePreview style={v} />}
          />
        </div>
      </div>
      <div>
        <div className="mb-1 text-xs text-gray-500">类型</div>
        <SegmentedPicker
          value={el.lineType ?? 'straight'}
          options={TYPE_OPTIONS}
          onChange={(v) => {
            if (v === 'straight') {
              upd((d) => { if (d.type === 'line') { d.lineType = 'straight'; d.points = undefined } })
            } else if (v === 'broken') {
              upd((d) => { if (d.type === 'line') { d.lineType = 'broken'; d.points = [[d.w * 0.5, d.h * 0.5]] } })
            } else {
              upd((d) => { if (d.type === 'line') { d.lineType = 'curve'; d.points = [[d.w * 0.5, d.h * 0.5]] } })
            }
          }}
          renderIcon={(v) => <LineTypePreview type={v} />}
        />
      </div>
      <div className="grid grid-cols-2 gap-2">
        <div>
          <div className="mb-1 text-xs text-gray-500">起点箭头</div>
          <VisualDropdown
            value={normalizeArrow(el.startArrow)}
            options={ARROW_OPTIONS}
            onChange={(v) => upd((d) => { if (d.type === 'line') d.startArrow = toLineArrow(v) })}
            renderIcon={(v) => <LineArrowPreview arrow={v} at="start" />}
          />
        </div>
        <div>
          <div className="mb-1 text-xs text-gray-500">终点箭头</div>
          <VisualDropdown
            value={normalizeArrow(el.endArrow)}
            options={ARROW_OPTIONS}
            onChange={(v) => upd((d) => { if (d.type === 'line') d.endArrow = toLineArrow(v) })}
            renderIcon={(v) => <LineArrowPreview arrow={v} at="end" />}
          />
        </div>
      </div>
    </>
  )
}
