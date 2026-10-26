// 面板下拉选项常量（标签按当前界面语言解析）
import type { TransitionPreset } from '../../types/slides'
import { CHART_GROUPS } from '../../core/chart/chartOptions'

type Translate = (key: string) => string

const CHART_GROUP_KEYS = ['bar', 'line', 'pie', 'barHorizontal', 'area', 'scatter', 'radar'] as const

const camelCase = (id: string) => id.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase())

/** 图表类型下拉（按组分组的 optgroup 选项） */
export function chartTypeGroups(t: Translate) {
  return CHART_GROUPS.map((group, index) => ({
    group: t(`pptx.chart.group.${CHART_GROUP_KEYS[index] ?? 'bar'}`),
    types: group.types.map((type) => ({ id: type.id, label: t(`pptx.chart.type.${camelCase(type.id)}`) })),
  }))
}

const TRANSITIONS: Array<{ value: TransitionPreset; key: string }> = [
  { value: 'none', key: 'pptx.transition.none' },
  { value: 'fade', key: 'pptx.transition.fade' },
  { value: 'slideLeft', key: 'pptx.transition.slideLeft' },
  { value: 'slideRight', key: 'pptx.transition.slideRight' },
  { value: 'slideUp', key: 'pptx.transition.slideUp' },
  { value: 'slideDown', key: 'pptx.transition.slideDown' },
  { value: 'wipeLeft', key: 'pptx.transition.wipeLeft' },
  { value: 'wipeRight', key: 'pptx.transition.wipeRight' },
  { value: 'zoomIn', key: 'pptx.transition.zoomIn' },
  { value: 'zoomOut', key: 'pptx.transition.zoomOut' },
  { value: 'blinds', key: 'pptx.transition.blinds' },
  { value: 'flip', key: 'pptx.transition.flip' },
]

export function transitionPresets(t: Translate): Array<{ value: TransitionPreset; label: string }> {
  return TRANSITIONS.map((item) => ({ value: item.value, label: t(item.key) }))
}

const EFFECTS: Record<'in' | 'emphasis' | 'out', Array<{ value: string; key: string }>> = {
  in: [
    { value: 'fade', key: 'pptx.effect.fade' },
    { value: 'flyInLeft', key: 'pptx.effect.flyInLeft' },
    { value: 'flyInRight', key: 'pptx.effect.flyInRight' },
    { value: 'flyInTop', key: 'pptx.effect.flyInTop' },
    { value: 'flyInBottom', key: 'pptx.effect.flyInBottom' },
    { value: 'zoomIn', key: 'pptx.effect.zoomIn' },
    { value: 'wipeRight', key: 'pptx.effect.wipeRight' },
    { value: 'bounceIn', key: 'pptx.effect.bounceIn' },
  ],
  emphasis: [
    { value: 'pulse', key: 'pptx.effect.pulse' },
    { value: 'shake', key: 'pptx.effect.shake' },
    { value: 'spin', key: 'pptx.effect.spin' },
    { value: 'grow', key: 'pptx.effect.grow' },
  ],
  out: [
    { value: 'fade', key: 'pptx.effect.fadeOut' },
    { value: 'flyOutRight', key: 'pptx.effect.flyOutRight' },
    { value: 'flyOutLeft', key: 'pptx.effect.flyOutLeft' },
    { value: 'zoomOut', key: 'pptx.effect.zoomOut' },
  ],
}

export function animEffectOptions(t: Translate): Record<'in' | 'emphasis' | 'out', Array<{ value: string; label: string }>> {
  return {
    in: EFFECTS.in.map((item) => ({ value: item.value, label: t(item.key) })),
    emphasis: EFFECTS.emphasis.map((item) => ({ value: item.value, label: t(item.key) })),
    out: EFFECTS.out.map((item) => ({ value: item.value, label: t(item.key) })),
  }
}
