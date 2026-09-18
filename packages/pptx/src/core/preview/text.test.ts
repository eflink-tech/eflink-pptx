import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './test-utils'
import { renderText } from './text'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const txBodyOf = (inner: string) =>
  parseXML(`<root ${NS}><p:txBody>${inner}</p:txBody></root>`).documentElement.firstElementChild as Element

const BOX = { x: 100, y: 100, w: 500, h: 100, rot: 0, flipH: false, flipV: false }

describe('preview/renderText', () => {
  it('foreignObject + 复用 txBodyToHTML 富文本 + bodyPr 内边距默认值', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN" sz="2400"/><a:t>标题</a:t></a:r></a:p>')
    const fo = (await renderText(txBody, ctx, BOX))!
    expect(fo.tagName).toBe('foreignObject')
    expect(fo.getAttribute('x')).toBe('100')
    expect(fo.getAttribute('y')).toBe('100')
    expect(fo.getAttribute('width')).toBe('500')
    expect(fo.getAttribute('height')).toBe('100')
    const div = fo.firstElementChild as HTMLElement
    expect(div.getAttribute('xmlns')).toBe('http://www.w3.org/1999/xhtml')
    // 默认内边距：lIns/rIns 91440 → 9.6px，tIns/bIns 45720 → 4.8px
    expect(div.getAttribute('style')).toContain('padding:4.8px 9.6px 4.8px 9.6px')
    expect(div.innerHTML).toContain('标题')
    expect(div.innerHTML).toContain('font-size:32px') // sz 2400 → 32px
  })

  it('anchor=t → 顶部对齐 flex；wrap=none → nowrap', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr anchor="t" wrap="none"/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    const div = ((await renderText(txBody, ctx, { ...BOX, x: 0, y: 0, w: 10, h: 10 }))!
      .firstElementChild as HTMLElement)
    expect(div.getAttribute('style')).toContain('justify-content:flex-start')
    expect(div.getAttribute('style')).toContain('white-space:nowrap')
  })

  it('vert=eaVert → writing-mode:vertical-rl', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr vert="eaVert"/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    const div = ((await renderText(txBody, ctx, { ...BOX, x: 0, y: 0, w: 10, h: 10 }))!
      .firstElementChild as HTMLElement)
    expect(div.getAttribute('style')).toContain('writing-mode:vertical-rl')
  })
})
