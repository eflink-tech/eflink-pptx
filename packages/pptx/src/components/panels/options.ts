// 面板下拉选项常量
import type { TransitionPreset } from '../../types/slides'
import { CHART_GROUPS } from '../../core/chart/chartOptions'

/** 图表类型下拉（按组分组的 optgroup 选项） */
export const chartTypeGroups = CHART_GROUPS

export const transitionPresets: Array<{ value: TransitionPreset; label: string }> = [
  { value: 'none', label: '无' },
  { value: 'fade', label: '淡入淡出' },
  { value: 'slideLeft', label: '向左滑入' },
  { value: 'slideRight', label: '向右滑入' },
  { value: 'slideUp', label: '向上滑入' },
  { value: 'slideDown', label: '向下滑入' },
  { value: 'wipeLeft', label: '向左擦除' },
  { value: 'wipeRight', label: '向右擦除' },
  { value: 'zoomIn', label: '缩放进入' },
  { value: 'zoomOut', label: '缩放退出' },
  { value: 'blinds', label: '百叶窗' },
  { value: 'flip', label: '翻转' },
]

export const effectOptions: Record<'in' | 'emphasis' | 'out', Array<{ value: string; label: string }>> = {
  in: [
    { value: 'fade', label: '淡入' },
    { value: 'flyInLeft', label: '自左侧飞入' },
    { value: 'flyInRight', label: '自右侧飞入' },
    { value: 'flyInTop', label: '自顶部飞入' },
    { value: 'flyInBottom', label: '自底部飞入' },
    { value: 'zoomIn', label: '缩放进入' },
    { value: 'wipeRight', label: '向右擦除' },
    { value: 'bounceIn', label: '弹跳进入' },
  ],
  emphasis: [
    { value: 'pulse', label: '脉冲' },
    { value: 'shake', label: '摇晃' },
    { value: 'spin', label: '旋转' },
    { value: 'grow', label: '放大' },
  ],
  out: [
    { value: 'fade', label: '淡出' },
    { value: 'flyOutRight', label: '向右飞出' },
    { value: 'flyOutLeft', label: '向左飞出' },
    { value: 'zoomOut', label: '缩小退出' },
  ],
}
