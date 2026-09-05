// 线条插入预设（与插入菜单下拉选项一一对应）
import type { LineArrow, LineElement, LineStyle, LineType } from '../../types/slides'
import { genId } from '../utils/id'

export type LinePresetId =
  | 'straight-solid'
  | 'straight-dashed'
  | 'straight-arrow'
  | 'straight-dashed-arrow'
  | 'straight-dot'
  | 'broken-l'
  | 'broken-z'
  | 'curve-simple'
  | 'curve-s'

export interface LinePreset {
  id: LinePresetId
  category: 'straight' | 'connector'
  w: number
  h: number
  start: [number, number]
  end: [number, number]
  points?: Array<[number, number]>
  lineType: LineType
  lineStyle: LineStyle
  startArrow: LineArrow
  endArrow: LineArrow
}

export const LINE_PRESET_GROUPS = [
  { id: 'straight' as const, name: '直线' },
  { id: 'connector' as const, name: '折线、曲线' },
]

export const LINE_PRESETS: Record<LinePresetId, LinePreset> = {
  'straight-solid': {
    id: 'straight-solid', category: 'straight',
    w: 240, h: 0, start: [0, 0], end: [240, 0],
    lineType: 'straight', lineStyle: 'solid', startArrow: '', endArrow: '',
  },
  'straight-dashed': {
    id: 'straight-dashed', category: 'straight',
    w: 240, h: 0, start: [0, 0], end: [240, 0],
    lineType: 'straight', lineStyle: 'dashed', startArrow: '', endArrow: '',
  },
  'straight-arrow': {
    id: 'straight-arrow', category: 'straight',
    w: 240, h: 0, start: [0, 0], end: [240, 0],
    lineType: 'straight', lineStyle: 'solid', startArrow: '', endArrow: 'arrow',
  },
  'straight-dashed-arrow': {
    id: 'straight-dashed-arrow', category: 'straight',
    w: 240, h: 0, start: [0, 0], end: [240, 0],
    lineType: 'straight', lineStyle: 'dashed', startArrow: '', endArrow: 'arrow',
  },
  'straight-dot': {
    id: 'straight-dot', category: 'straight',
    w: 240, h: 0, start: [0, 0], end: [240, 0],
    lineType: 'straight', lineStyle: 'solid', startArrow: '', endArrow: 'dot',
  },
  'broken-l': {
    id: 'broken-l', category: 'connector',
    w: 200, h: 140, start: [0, 0], end: [200, 140], points: [[200, 0]],
    lineType: 'broken', lineStyle: 'solid', startArrow: '', endArrow: 'arrow',
  },
  'broken-z': {
    id: 'broken-z', category: 'connector',
    w: 200, h: 140, start: [0, 0], end: [200, 140], points: [[100, 0], [100, 140]],
    lineType: 'broken', lineStyle: 'solid', startArrow: '', endArrow: 'arrow',
  },
  'curve-simple': {
    id: 'curve-simple', category: 'connector',
    w: 200, h: 120, start: [0, 120], end: [200, 0], points: [[200, 120]],
    lineType: 'curve', lineStyle: 'solid', startArrow: '', endArrow: 'arrow',
  },
  'curve-s': {
    id: 'curve-s', category: 'connector',
    w: 200, h: 120, start: [0, 0], end: [200, 120], points: [[200, 0], [0, 120]],
    lineType: 'curve', lineStyle: 'solid', startArrow: '', endArrow: 'arrow',
  },
}

const PRESETS_BY_GROUP = LINE_PRESET_GROUPS.map((g) => ({
  ...g,
  presets: Object.values(LINE_PRESETS).filter((p) => p.category === g.id),
}))

export { PRESETS_BY_GROUP }

export function createLineElementFromPreset(x: number, y: number, presetId: LinePresetId): LineElement {
  const p = LINE_PRESETS[presetId]
  return {
    id: genId('l-'),
    type: 'line',
    x,
    y,
    w: p.w,
    h: p.h,
    start: [...p.start],
    end: [...p.end],
    points: p.points?.map((pt) => [...pt] as [number, number]),
    lineType: p.lineType,
    lineStyle: p.lineStyle,
    color: '#333333',
    lineWidth: 2,
    startArrow: p.startArrow,
    endArrow: p.endArrow,
    name: '线条',
  }
}
