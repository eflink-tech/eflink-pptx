import { describe, expect, it } from 'vitest'
import { parseDeck, parseJSONFromText, parseOutline, sanitizeElements } from './systemPrompt'

describe('parseJSONFromText', () => {
  it('解析裸 JSON', () => {
    expect(parseJSONFromText('{"a":1}')).toEqual({ a: 1 })
  })
  it('剥掉 markdown 代码块', () => {
    expect(parseJSONFromText('```json\n{"a":1}\n```')).toEqual({ a: 1 })
  })
  it('截取首尾大括号', () => {
    expect(parseJSONFromText('好的，以下是结果：{"a":1} 请查收')).toEqual({ a: 1 })
  })
  it('非法输入抛错', () => {
    expect(() => parseJSONFromText('不是JSON')).toThrow()
  })
})

describe('parseOutline', () => {
  it('解析大纲', () => {
    const outline = parseOutline('{"topic":"AI","slides":[{"title":"封面","points":["副标题"]},{"title":"总结"}]}')
    expect(outline.topic).toBe('AI')
    expect(outline.slides).toHaveLength(2)
    expect(outline.slides[0].points).toEqual(['副标题'])
    expect(outline.slides[1].points).toEqual([])
  })
  it('结构错误抛错', () => {
    expect(() => parseOutline('{"a":1}')).toThrow()
  })
})

describe('parseDeck', () => {
  it('解析页面数据', () => {
    const slides = parseDeck('{"slides":[{"title":"t","elements":[{"type":"text","x":0,"y":0,"w":100,"h":50,"text":"hi"}]}]}')
    expect(slides).toHaveLength(1)
    expect(slides[0].elements[0].text).toBe('hi')
  })
})

describe('sanitizeElements', () => {
  it('裁剪越界坐标与非法值', () => {
    const els = sanitizeElements([
      { type: 'text', x: -50, y: 2000, w: 9999, h: -10, text: 'hi', fontSize: 500 },
      { type: 'image', x: 0, y: 0, w: 100, h: 100, text: 'x' },
      { type: 'shape', x: 0, y: 0, w: 100, h: 100, text: '', fill: 'not-a-color' },
    ] as never)
    expect(els[0].x).toBe(0)
    expect(els[0].y).toBe(720)
    expect(els[0].w).toBe(1280)
    expect(els[0].h).toBe(20)
    expect(els[0].fontSize).toBe(96)
    expect(els).toHaveLength(2) // image 被过滤，空 text 的 shape 被过滤
  })
})
