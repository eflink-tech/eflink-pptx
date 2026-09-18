/** txBody → 富文本 HTML（项目符号/编号、超链接、autofit、竖排） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import type { PptxTheme } from '../theme'
import type { PptxPackage } from '../package'

export interface TextBodyResult {
  html: string
  autoSize: boolean
  vertical: boolean
}

function escapeHTML(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

interface ParaInfo {
  align: string
  kind: 'none' | 'bullet' | 'number'
  inner: string
}

function bulletKind(p: Element): 'none' | 'bullet' | 'number' {
  const pPr = directChild(p, 'a:pPr')
  if (!pPr) return 'none'
  if (directChild(pPr, 'a:buNone')) return 'none'
  if (directChild(pPr, 'a:buChar')) return 'bullet'
  if (directChild(pPr, 'a:buAutoNum')) return 'number'
  return 'none'
}

async function paragraphToInner(p: Element, theme: PptxTheme, pkg: PptxPackage, partPath: string): Promise<string> {
  let inner = ''
  for (const node of Array.from(p.children)) {
    if (node.nodeName === 'a:br') {
      inner += '<br>'
      continue
    }
    if (node.nodeName !== 'a:r') continue
    const t = firstDescendant(node, 'a:t')
    const text = t?.textContent ?? ''
    if (!text) continue
    const rPr = firstDescendant(node, 'a:rPr')
    const styles: string[] = []
    if (rPr) {
      const sz = attr(rPr, 'sz')
      if (sz) styles.push(`font-size:${Math.round(parseInt(sz, 10) / 100 / 0.75)}px`)
      if (attr(rPr, 'b') === '1') styles.push('font-weight:bold')
      if (attr(rPr, 'i') === '1') styles.push('font-style:italic')
      if (attr(rPr, 'u') === 'sng') styles.push('text-decoration:underline')
      if (attr(rPr, 'strike') === 'sng') styles.push('text-decoration:line-through')
      const color = resolveColor(directChild(rPr, 'a:solidFill'), theme)
      if (color) styles.push(`color:${color}`)
      const latin = directChild(rPr, 'a:latin')
      const typeface = attr(latin, 'typeface')
      if (typeface && !typeface.startsWith('+')) styles.push(`font-family:${typeface}`)
    }
    const styleAttr = styles.length ? ` style="${styles.join(';')}"` : ''
    let run = `<span${styleAttr}>${escapeHTML(text)}</span>`
    // 超链接：a:hlinkClick@r:id → rels（External）；无包上下文时跳过
    const hlink = rPr ? directChild(rPr, 'a:hlinkClick') : null
    const hlinkId = attr(hlink, 'r:id')
    if (hlinkId && pkg && typeof pkg.rels === 'function') {
      const rel = (await pkg.rels(partPath)).get(hlinkId)
      if (rel && rel.mode === 'External') {
        run = `<a href="${escapeHTML(rel.target)}"${styleAttr}>${escapeHTML(text)}</a>`
      }
    }
    inner += run
  }
  return inner
}

/** txBody → 富文本 HTML；autoSize 来自 normAutofit；vertical 来自 bodyPr@vert */
export async function txBodyToHTML(
  txBody: Element,
  theme: PptxTheme,
  pkg?: PptxPackage,
  partPath = '',
): Promise<TextBodyResult> {
  const bodyPr = directChild(txBody, 'a:bodyPr')
  const autoSize = Boolean(bodyPr && directChild(bodyPr, 'a:normAutofit'))
  const vert = attr(bodyPr, 'vert')
  const vertical = vert === 'eaVert' || vert === 'vert' || vert === 'mongolianVert'

  const paragraphs = directChildren(txBody, 'a:p')
  const paras: ParaInfo[] = []
  for (const p of paragraphs) {
    const pPr = directChild(p, 'a:pPr')
    const algn = attr(pPr, 'algn')
    const align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left'
    const inner = await paragraphToInner(p, theme, pkg ?? ({} as PptxPackage), partPath)
    paras.push({ align, kind: bulletKind(p), inner })
  }

  const html: string[] = []
  let i = 0
  while (i < paras.length) {
    const para = paras[i]
    if (para.kind === 'none') {
      html.push(`<p style="text-align:${para.align}">${para.inner || '&nbsp;'}</p>`)
      i += 1
      continue
    }
    const tag = para.kind === 'bullet' ? 'ul' : 'ol'
    const items: string[] = []
    while (i < paras.length && paras[i].kind === para.kind) {
      const cur = paras[i]
      items.push(`<li style="text-align:${cur.align}">${cur.inner || '&nbsp;'}</li>`)
      i += 1
    }
    html.push(`<${tag}>${items.join('')}</${tag}>`)
  }
  return { html: html.join('') || '<p></p>', autoSize, vertical }
}
