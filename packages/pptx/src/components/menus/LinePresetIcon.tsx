// 线条预设预览图标（插入菜单下拉用）
import type { LinePresetId } from '../../core/schema/linePresets'

const VB = '0 0 48 48'

function Base({ children }: { children: React.ReactNode }) {
  return (
    <svg viewBox={VB} className="size-10" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      {children}
    </svg>
  )
}

function ArrowHead({ x, y, angle }: { x: number; y: number; angle: number }) {
  return (
    <polygon
      points="0,-4 8,0 0,4"
      fill="currentColor"
      stroke="none"
      transform={`translate(${x} ${y}) rotate(${angle})`}
    />
  )
}

function DotCap({ x, y }: { x: number; y: number }) {
  return <circle cx={x} cy={y} r="3.5" fill="currentColor" stroke="none" />
}

/** 各预设的缩略图路径（统一从左上到右下） */
const PRESET_PATHS: Record<LinePresetId, React.ReactNode> = {
  'straight-solid': (
    <>
      <line x1="8" y1="40" x2="40" y2="8" />
    </>
  ),
  'straight-dashed': (
    <>
      <line x1="8" y1="40" x2="40" y2="8" strokeDasharray="6 4" />
    </>
  ),
  'straight-arrow': (
    <>
      <line x1="8" y1="40" x2="34" y2="14" />
      <ArrowHead x={38} y={10} angle={-45} />
    </>
  ),
  'straight-dashed-arrow': (
    <>
      <line x1="8" y1="40" x2="34" y2="14" strokeDasharray="6 4" />
      <ArrowHead x={38} y={10} angle={-45} />
    </>
  ),
  'straight-dot': (
    <>
      <line x1="8" y1="40" x2="34" y2="14" />
      <DotCap x={40} y={8} />
    </>
  ),
  'broken-l': (
    <>
      <polyline points="8,8 36,8 36,40" />
      <ArrowHead x={36} y={44} angle={90} />
    </>
  ),
  'broken-z': (
    <>
      <polyline points="8,8 24,8 24,40 40,40" />
      <ArrowHead x={44} y={40} angle={0} />
    </>
  ),
  'curve-simple': (
    <>
      <path d="M8 40 Q 40 40 40 8" />
      <ArrowHead x={40} y={8} angle={-45} />
    </>
  ),
  'curve-s': (
    <>
      <path d="M8 8 C 40 8 8 40 40 40" />
      <ArrowHead x={44} y={40} angle={0} />
    </>
  ),
}

export function LinePresetIcon({ presetId }: { presetId: LinePresetId }) {
  return <Base>{PRESET_PATHS[presetId]}</Base>
}
