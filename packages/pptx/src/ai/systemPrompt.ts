// AI 提示词与结果解析
import type { AIGenElement, AIGenSlide, OutlineResult } from './types'

export const SLIDES_SCHEMA_PROMPT = `输出 JSON（不要 markdown 代码块包裹之外的任何文字），结构如下：
{
  "slides": [
    {
      "title": "本页标题",
      "elements": [
        { "type": "shape", "role": "accent", "x":0, "y":0, "w":1280, "h":720, "text":"", "fill":"#d14424" },
        { "type": "text", "role": "title", "x":100, "y":240, "w":1080, "h":90, "text":"标题文字", "fontSize":44 },
        { "type": "text", "role": "body", "x":100, "y":360, "w":1080, "h":240, "text":"正文要点，每行一条", "fontSize":22 }
      ]
    }
  ]
}
坐标系为 1280x720 画布（16:9），x/y 为左上角，w/h 为宽高。每页必须包含 role=title 的标题文本；正文用 role=body。可用 role=accent 的形状做装饰色块（建议细长条或半屏色块）。文本内容支持 \\n 换行。`

export function buildOutlinePrompt(topic: string, count: number): string {
  return `请围绕主题「${topic}」生成一份演示文稿大纲，共 ${count} 页：第一页为封面（标题+副标题），中间页每页一个论点（含 2-4 个要点），最后一页为总结/致谢。
只输出 JSON：{"topic":"主题","slides":[{"title":"页标题","points":["要点1","要点2"]}]}
不要输出 JSON 之外的任何文字。`
}

export function buildDeckPrompt(outline: OutlineResult, themeColors: string[]): string {
  return `你是一位专业的 PPT 版式设计师。请根据下面的大纲生成完整演示文稿的页面数据。
大纲：${JSON.stringify(outline)}
主色板（可从中选取 fill 颜色）：${themeColors.join(', ')}
${SLIDES_SCHEMA_PROMPT}
要求：版式有变化（左右分栏、上下层次、色块装饰等），标题醒目，正文简洁。第一页为封面页（大标题居中），最后一页为结束页（如「感谢观看」）。`
}

export function parseJSONFromText(text: string): unknown {
  // 去掉可能的 markdown 代码块包裹
  const cleaned = text.replace(/```(?:json)?/g, '').trim()
  try {
    return JSON.parse(cleaned)
  } catch {
    // 尝试截取第一个 { 到最后一个 }
    const start = cleaned.indexOf('{')
    const end = cleaned.lastIndexOf('}')
    if (start >= 0 && end > start) {
      return JSON.parse(cleaned.slice(start, end + 1))
    }
    throw new Error('AI 返回的内容不是有效的 JSON')
  }
}

export function parseOutline(text: string): OutlineResult {
  const data = parseJSONFromText(text) as Partial<OutlineResult>
  if (!data || !Array.isArray(data.slides)) throw new Error('大纲 JSON 结构不正确')
  return {
    topic: typeof data.topic === 'string' ? data.topic : '未命名主题',
    slides: data.slides.map((s) => ({
      title: String(s?.title ?? '未命名页'),
      points: Array.isArray(s?.points) ? s.points.map(String) : [],
    })),
  }
}

export function parseDeck(text: string): AIGenSlide[] {
  const data = parseJSONFromText(text) as { slides?: AIGenSlide[] }
  if (!data || !Array.isArray(data.slides)) throw new Error('页面 JSON 结构不正确')
  return data.slides
}

/** 校验并裁剪 AI 生成的元素，防止越界与非法值 */
export function sanitizeElements(elements: AIGenElement[]): AIGenElement[] {
  return elements
    .filter((el) => el && (el.type === 'text' || el.type === 'shape') && typeof el.text === 'string')
    .map((el) => ({
      ...el,
      x: clamp(Number(el.x) || 0, 0, 1280),
      y: clamp(Number(el.y) || 0, 0, 720),
      w: clamp(Number(el.w) || 200, 40, 1280),
      h: clamp(Number(el.h) || 60, 20, 720),
      fontSize: clamp(Number(el.fontSize) || 20, 12, 96),
      fill: typeof el.fill === 'string' && /^#[0-9a-f]{6}$/i.test(el.fill) ? el.fill : undefined,
    }))
}

function clamp(v: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, v))
}
