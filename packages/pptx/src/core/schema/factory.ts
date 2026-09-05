// 元素工厂：插入新元素时的默认值
import type {
  AudioElement, ChartElement, FormulaElement, ImageElement, LineElement,
  PPTElement, ShapeElement, TableElement, TextElement, VideoElement,
} from '../../types/slides'
import { genId } from '../utils/id'
import { createLineElementFromPreset } from './linePresets'
import { DEFAULT_CHART_ELEMENTS } from '../chart/chartOptions'

export function createTextElement(x: number, y: number, w = 400, h = 80, content = '默认文本'): TextElement {
  return {
    id: genId('t-'), type: 'text', x, y, w, h,
    content: `<p>${content}</p>`,
    defaultColor: '#333333', lineHeight: 1.5, padding: 8,
    name: '文本框',
  }
}

export function createImageElement(x: number, y: number, src: string, w: number, h: number): ImageElement {
  return { id: genId('i-'), type: 'image', x, y, w, h, src, name: '图片' }
}

export function createShapeElement(x: number, y: number, shapeKey: string): ShapeElement {
  const w = shapeKey === 'line' ? 200 : 260
  return {
    id: genId('s-'), type: 'shape', x, y, w, h: 200, shapeKey,
    fill: '#d14424', outline: { color: '#00000000', width: 1, style: 'solid' },
    align: 'center', valign: 'middle', name: '形状',
  }
}

export function createLineElement(x: number, y: number): LineElement {
  return createLineElementFromPreset(x, y, 'straight-solid')
}

export function createTableElement(x: number, y: number, rows = 4, cols = 4): TableElement {
  const colSizes = Array.from({ length: cols }, () => 1 / cols)
  const rowSizes = Array.from({ length: rows }, () => 1 / rows)
  const cells = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => ({
      text: r === 0 ? `表头 ${c + 1}` : '',
      style: r === 0
        ? { bold: true, color: '#ffffff', backcolor: '#d14424', align: 'center', valign: 'middle' } as const
        : { color: '#333333', align: 'left', valign: 'middle' } as const,
    })),
  )
  return {
    id: genId('tb-'), type: 'table', x, y, w: 640, h: rows * 48,
    rowSizes, colSizes, cells, outline: { color: '#d0d0d0', width: 1, style: 'solid' },
    theme: { color: ['#ffffff', '#f7f7f7'], headColor: '#d14424' }, name: '表格',
  }
}

export function createChartElement(x: number, y: number, chartType: ChartElement['chartType'] = 'bar-cluster'): ChartElement {
  return {
    id: genId('c-'), type: 'chart', x, y, w: 560, h: 380, chartType,
    data: {
      labels: ['一月', '二月', '三月', '四月'],
      series: [
        { name: '系列1', values: [80, 65, 90, 72] },
        { name: '系列2', values: [55, 78, 60, 85] },
      ],
    },
    elements: { ...DEFAULT_CHART_ELEMENTS }, paletteIndex: 0, fontSize: 14, name: '图表',
  }
}

export function createVideoElement(x: number, y: number, src: string): VideoElement {
  return { id: genId('v-'), type: 'video', x, y, w: 640, h: 360, src, loop: false, autoPlay: false, name: '视频' }
}

export function createAudioElement(x: number, y: number, src: string): AudioElement {
  return { id: genId('a-'), type: 'audio', x, y, w: 64, h: 64, src, loop: false, autoPlay: false, name: '音频' }
}

export function createFormulaElement(x: number, y: number, latex = 'a^2+b^2=c^2'): FormulaElement {
  return { id: genId('f-'), type: 'formula', x, y, w: 320, h: 96, latex, color: '#333333', fontSize: 32, name: '公式' }
}

/** 按类型创建元素（菜单/粘贴/AI 通用入口） */
export function createElementByType(type: PPTElement['type'], x: number, y: number): PPTElement {
  switch (type) {
    case 'text': return createTextElement(x, y)
    case 'shape': return createShapeElement(x, y, 'roundRect')
    case 'line': return createLineElement(x, y)
    case 'table': return createTableElement(x, y)
    case 'chart': return createChartElement(x, y)
    case 'formula': return createFormulaElement(x, y)
    default: return createTextElement(x, y)
  }
}
