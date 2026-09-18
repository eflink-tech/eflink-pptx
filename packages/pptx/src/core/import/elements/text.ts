/** txBody → 富文本 HTML（项目符号/编号、超链接、autofit、竖排、字体栈） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { fontStackOf } from '../fonts'
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
  /** 行距样式（line-height:...），无行距时为空串 */
  spacing: string
}

function bulletKind(p: Element): 'none' | 'bullet' | 'number' {
  const pPr = directChild(p, 'a:pPr')
  if (!pPr) return 'none'
  if (directChild(pPr, 'a:buNone')) return 'none'
  if (directChild(pPr, 'a:buChar')) return 'bullet'
  if (directChild(pPr, 'a:buAutoNum')) return 'number'
  return 'none'
}

/** run 字符样式提取：rPr 显式属性优先，缺失时回退 lstStyle defRPr（同级默认字符样式）。
 * 属性分两类：字号/粗斜下划线/字距在元素属性上，颜色/字体在子元素上，两类均按 rPr → defRPr 回退。 */
function runStyles(rPr: Element | null, defRPr: Element | null, theme: PptxTheme): string[] {
  const styles: string[] = []
  const attrOf = (name: string): string | null =>
    (rPr ? attr(rPr, name) : null) ?? (defRPr ? attr(defRPr, name) : null)
  const sz = attrOf('sz')
  if (sz) styles.push(`font-size:${Math.round(parseInt(sz, 10) / 100 / 0.75)}px`)
  if (attrOf('b') === '1') styles.push('font-weight:bold')
  if (attrOf('i') === '1') styles.push('font-style:italic')
  if (attrOf('u') === 'sng') styles.push('text-decoration:underline')
  if (attrOf('strike') === 'sng') styles.push('text-decoration:line-through')
  const spc = attrOf('spc')
  if (spc) {
    // 字距 1/100 pt → px（与 font-size 同源换算），保留两位小数防浮点噪声
    const ls = Math.round((parseInt(spc, 10) / 100 / 0.75) * 100) / 100
    styles.push(`letter-spacing:${ls}px`)
  }
  const fill = (rPr ? directChild(rPr, 'a:solidFill') : null) ?? (defRPr ? directChild(defRPr, 'a:solidFill') : null)
  const color = resolveColor(fill, theme)
  if (color) styles.push(`color:${color}`)
  // 字体栈：latin + ea（含 +mj/+mn 主题引用），中文回退栈兜底；字体名来自不可信属性，fontStackOf 内已剔除引号防注入
  const latin = ((rPr ? directChild(rPr, 'a:latin') : null) ?? (defRPr ? directChild(defRPr, 'a:latin') : null))
  const ea = ((rPr ? directChild(rPr, 'a:ea') : null) ?? (defRPr ? directChild(defRPr, 'a:ea') : null))
  if (attr(latin, 'typeface') || attr(ea, 'typeface')) {
    styles.push(`font-family:${fontStackOf(attr(latin, 'typeface'), attr(ea, 'typeface'), theme)}`)
  }
  return styles
}

/** lstStyle → 段落级别对应的默认字符样式（a:lvl{N}pPr > a:defRPr，未命中级别时回退 a:defPPr） */
export function defRPrOf(lst: Element | null, lvl: number): Element | null {
  if (!lst) return null
  const lvlPr = directChild(lst, `a:lvl${lvl + 1}pPr`) ?? directChild(lst, 'a:defPPr')
  return lvlPr ? directChild(lvlPr, 'a:defRPr') : null
}

async function paragraphToInner(p: Element, theme: PptxTheme, pkg: PptxPackage | undefined, partPath: string, lst: Element | null): Promise<string> {
  const pPr = directChild(p, 'a:pPr')
  const lvlRaw = parseInt(attr(pPr, 'lvl') ?? '0', 10)
  const defRPr = defRPrOf(lst, Number.isFinite(lvlRaw) ? lvlRaw : 0)
  let inner = ''
  for (const node of Array.from(p.children)) {
    if (node.nodeName === 'a:br') {
      inner += '<br>'
      continue
    }
    // a:r 普通 run；a:fld 字段（页码/日期等）结构相同（a:rPr + a:t），按缓存文本渲染为静态 run
    if (node.nodeName !== 'a:r' && node.nodeName !== 'a:fld') continue
    const t = firstDescendant(node, 'a:t')
    const text = t?.textContent ?? ''
    if (!text) continue
    const rPr = firstDescendant(node, 'a:rPr')
    const styles = runStyles(rPr, defRPr, theme)
    const styleAttr = styles.length ? ` style="${styles.join(';')}"` : ''
    let run = `<span${styleAttr}>${escapeHTML(text)}</span>`
    // 超链接：a:hlinkClick@r:id → rels（External，仅放行 http/https/mailto 白名单协议）；无包上下文时跳过
    const hlink = rPr ? directChild(rPr, 'a:hlinkClick') : null
    const hlinkId = attr(hlink, 'r:id')
    if (hlinkId && pkg) {
      const rel = (await pkg.rels(partPath)).get(hlinkId)
      const safeTarget = rel?.mode === 'External' && /^(https?:|mailto:)/i.test(rel.target) ? rel.target : null
      if (safeTarget) {
        run = `<a href="${escapeHTML(safeTarget)}"${styleAttr}>${escapeHTML(text)}</a>`
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
    const inner = await paragraphToInner(p, theme, pkg, partPath, directChild(txBody, 'a:lstStyle'))
    // 行距：spcPct（1/100000 → 倍数）优先，spcPts（1/100 pt → px）覆盖；val 非法/非正数时跳过，避免产出 line-height:0 压扁文字
    const lnSpc = pPr ? directChild(pPr, 'a:lnSpc') : null
    let spacing = ''
    const pct = lnSpc ? firstDescendant(lnSpc, 'a:spcPct') : null
    const pctVal = pct ? parseInt(attr(pct, 'val') ?? '', 10) : NaN
    if (Number.isFinite(pctVal) && pctVal > 0) spacing = `line-height:${pctVal / 100000}`
    const pts = lnSpc ? firstDescendant(lnSpc, 'a:spcPts') : null
    const ptsVal = pts ? parseInt(attr(pts, 'val') ?? '', 10) : NaN
    if (Number.isFinite(ptsVal) && ptsVal > 0) spacing = `line-height:${Math.round(ptsVal / 100 / 0.75)}px`
    paras.push({ align, kind: bulletKind(p), inner, spacing })
  }

  const html: string[] = []
  let i = 0
  while (i < paras.length) {
    const para = paras[i]
    if (para.kind === 'none') {
      html.push(`<p style="text-align:${para.align}${para.spacing ? `;${para.spacing}` : ''}">${para.inner || '&nbsp;'}</p>`)
      i += 1
      continue
    }
    const tag = para.kind === 'bullet' ? 'ul' : 'ol'
    const items: string[] = []
    while (i < paras.length && paras[i].kind === para.kind) {
      const cur = paras[i]
      items.push(`<li style="text-align:${cur.align}${cur.spacing ? `;${cur.spacing}` : ''}">${cur.inner || '&nbsp;'}</li>`)
      i += 1
    }
    html.push(`<${tag}>${items.join('')}</${tag}>`)
  }
  return { html: html.join('') || '<p></p>', autoSize, vertical }
}
