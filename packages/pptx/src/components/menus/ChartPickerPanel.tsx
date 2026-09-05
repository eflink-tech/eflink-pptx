// 图表类型选择面板（与 eflink-excel 同款）：按类型分组展示 mini ECharts 实时缩略图，
// 所见即所插。点击插入菜单中的图表按钮后弹出。
import { useEffect, useRef } from 'react'
import * as echarts from 'echarts'
import type { ChartType } from '../../types/slides'
import { buildChartOption, CHART_GROUPS, SAMPLE_CHART_DATA } from '../../core/chart/chartOptions'

interface Props {
  onPick: (chartType: ChartType) => void
  onClose: () => void
  /** 相对父容器的偏移 */
  offset?: { left: number; top: number }
}

const THUMB_W = 88
const THUMB_H = 58

function Thumb({ typeId }: { typeId: string }) {
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!ref.current) return
    const chart = echarts.init(ref.current, null, { renderer: 'svg', width: THUMB_W, height: THUMB_H })
    chart.setOption(buildChartOption(SAMPLE_CHART_DATA, typeId, true))
    return () => chart.dispose()
  }, [typeId])

  return <div ref={ref} style={{ width: THUMB_W, height: THUMB_H }} />
}

export function ChartPickerPanel({ onPick, onClose, offset = { left: 0, top: 36 } }: Props) {
  const panelRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!panelRef.current?.contains(e.target as Node)) onClose()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('keydown', onKey)
    }
  }, [onClose])

  return (
    <div
      ref={panelRef}
      className="absolute left-0 top-9 z-50 w-[360px] rounded-lg border border-gray-200 bg-white p-3 shadow-xl"
      data-testid="chart-picker-panel"
      style={{ marginLeft: offset.left, marginTop: offset.top - 36 }}
    >
      <div className="max-h-[380px] overflow-y-auto pr-1">
        {CHART_GROUPS.map(({ group, types }) => (
          <div key={group} className="mb-2">
            <div className="mb-1.5 text-[12px] font-medium text-gray-500">{group}</div>
            <div className="grid grid-cols-3 gap-2">
              {types.map(({ id, label }) => (
                <button
                  key={id}
                  type="button"
                  title={label}
                  className="flex flex-col items-center gap-1 rounded-lg border border-transparent p-1.5 transition-colors hover:border-[#d14424]/30 hover:bg-[#fbeae5]"
                  onClick={() => onPick(id as ChartType)}
                >
                  <Thumb typeId={id} />
                  <span className="max-w-full truncate text-[11px] text-gray-500">{label}</span>
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}
