// 内置成套模板（封面/内容/结束页）
import type { Background, Slide, Theme } from '../../types/slides'
import { genId } from '../utils/id'
import type { ShapeElement, TextElement } from '../../types/slides'

function text(x: number, y: number, w: number, h: number, content: string, opts: Partial<TextElement> = {}): TextElement {
  return {
    id: genId('t-'), type: 'text', x, y, w, h,
    content, lineHeight: 1.5, padding: 8, name: '文本', ...opts,
  }
}

function shape(shapeKey: string, x: number, y: number, w: number, h: number, fill: string, opts: Partial<ShapeElement> = {}): ShapeElement {
  return {
    id: genId('s-'), type: 'shape', x, y, w, h, shapeKey, fill,
    outline: { color: '#00000000', width: 1, style: 'solid' },
    align: 'center', valign: 'middle', name: '装饰', ...opts,
  }
}

export interface PPTTemplate {
  id: string
  name: string
  colors: string[]
  background: Background
  build: () => Slide[]
}

interface Palette {
  id: string; name: string
  colors: string[]
  bg: string
  accent: string
  text: string
  sub: string
}

const PALETTES: Palette[] = [
  { id: 'brand', name: '商务红', colors: ['#d14424', '#e6935c', '#f5c56b', '#7cb342', '#42a5f5', '#455a64'], bg: '#ffffff', accent: '#d14424', text: '#333333', sub: '#8a8a8a' },
  { id: 'tech', name: '深色科技', colors: ['#00c2ff', '#4a7dff', '#8e6bbf', '#00d6a1', '#f5c56b', '#202124'], bg: '#181c25', accent: '#00c2ff', text: '#f0f2f5', sub: '#8f97a3' },
  { id: 'fresh', name: '清新绿', colors: ['#2e9e6b', '#7cb342', '#c0d860', '#42a5f5', '#f5c56b', '#26523f'], bg: '#f7fbf7', accent: '#2e9e6b', text: '#2d3a31', sub: '#7c8b80' },
  { id: 'warm', name: '暖橙', colors: ['#e8730c', '#f5a623', '#f5c56b', '#d14424', '#8e6bbf', '#5d4037'], bg: '#fffbf5', accent: '#e8730c', text: '#4a3728', sub: '#a08c78' },
  { id: 'blue', name: '沉稳蓝', colors: ['#1e5eff', '#42a5f5', '#26c6da', '#5e6db8', '#f5c56b', '#0d1b3e'], bg: '#ffffff', accent: '#1e5eff', text: '#1f2733', sub: '#7a8699' },
]

function buildSlides(p: Palette): Slide[] {
  const cover: Slide = {
    id: genId('slide-'),
    background: { type: 'solid', color: p.bg },
    elements: [
      shape('rect', 0, 0, 1280, 720, p.accent, { opacity: 0.08 }),
      shape('rect', 0, 660, 1280, 60, p.accent),
      text(120, 280, 1040, 110, `<p style="text-align:center"><span style="font-size:56px;color:${p.text}">演示文稿标题</span></p>`, { name: '封面标题' }),
      text(240, 410, 800, 60, `<p style="text-align:center"><span style="font-size:22px;color:${p.sub}">副标题 · 汇报人 · 2026</span></p>`, { name: '封面副标题' }),
      shape('rect', 560, 396, 160, 4, p.accent),
    ],
  }
  const content: Slide = {
    id: genId('slide-'),
    background: { type: 'solid', color: p.bg },
    elements: [
      shape('rect', 100, 92, 8, 44, p.accent),
      text(128, 84, 900, 60, `<p><span style="font-size:32px;color:${p.text}">内容页标题</span></p>`, { name: '页标题' }),
      text(128, 200, 480, 320, `<p><span style="font-size:20px;color:${p.sub}">左栏要点：描述核心观点，支持分行排版，保持简洁。</span></p>`, { name: '正文' }),
      shape('roundRect', 660, 200, 520, 320, p.colors[1], { opacity: 0.16 }),
      text(700, 330, 440, 60, `<p style="text-align:center"><span style="font-size:20px;color:${p.text}">图示 / 数据占位</span></p>`, { name: '占位' }),
    ],
  }
  const end: Slide = {
    id: genId('slide-'),
    background: { type: 'solid', color: p.accent },
    elements: [
      text(240, 300, 800, 100, `<p style="text-align:center"><span style="font-size:48px;color:#ffffff">感谢观看</span></p>`, { name: '结束' }),
      text(340, 420, 600, 50, `<p style="text-align:center"><span style="font-size:20px;color:#ffffffcc">THANKS FOR WATCHING</span></p>`, { name: '结束副标题' }),
    ],
  }
  return [cover, content, end]
}

export const PPT_TEMPLATES: PPTTemplate[] = PALETTES.map((p) => ({
  id: `tpl-${p.id}`,
  name: p.name,
  colors: p.colors,
  background: { type: 'solid', color: p.bg },
  build: () => buildSlides(p),
}))

export function templateTheme(tpl: PPTTemplate): Partial<Theme> {
  return { colors: tpl.colors }
}
