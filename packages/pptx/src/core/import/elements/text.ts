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
  /** bodyPr 上下 insets 折算的内边距（px）；无 bodyPr 时为 null（调用方回退默认值）。
   * TextElement.padding 为单值，取上下平均——垂直方向是内容溢出/裁剪的主因 */
  padding: number | null
  /** bodyPr 含 a:spAutoFit：框随文本自动增高（PowerPoint 中此类文本框内容超高时溢出显示而非裁剪） */
  autoFitShape: boolean
  /** bodyPr@anchor 垂直对齐（t/ctr/b）；未声明为 null（OOXML 默认顶对齐） */
  anchor: 't' | 'ctr' | 'b' | null
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
  /** 段前/段后/缩进/RTL 等附加样式（margin-top/margin-bottom/margin-left/direction） */
  extra: string[]
  /** 项目符号层级（0 基） */
  level: number
  /** 自动编号起始值（a:buAutoNum@startAt） */
  startAt?: number
  /** 项目符号字符（a:buChar@char，已按符号字体映射为 Unicode）；扁平列表模式下作为行首字符渲染 */
  bulletChar?: string
  /** 悬挂缩进宽度（含单位）：符号占位宽度，使正文对齐到 marL；非悬挂缩进时为空 */
  hangPx?: string
}

/** 符号字体私有编码 → Unicode。Wingdings/Symbol 的 'l' 等项目符号字符直接渲染会显示字母本身 */
const SYMBOL_BULLETS: Record<string, Record<string, string>> = {
  wingdings: {
    l: '●', n: '■', o: '□', q: '❑', u: '◆', v: '❖', w: '◆',
    '§': '▪', Ø: '➢', ü: '✓', '¨': '❏', 'þ': '❖',
  },
  symbol: { l: '●', n: '■', u: '◆', '·': '•', '§': '▪', Ø: '➢' },
}

function bulletGlyph(char: string, buFont: Element | null): string {
  const font = attr(buFont, 'typeface')?.toLowerCase() ?? ''
  const table = font.includes('wingdings') ? SYMBOL_BULLETS.wingdings : font.includes('symbol') ? SYMBOL_BULLETS.symbol : null
  return (table && table[char]) || char
}

/** 段落间距（spcBef/spcAft 的 spcPts 1/100pt 或 spcPct 百分比）→ px / em */
function paraMargin(pPr: Element | null, tag: 'a:spcBef' | 'a:spcAft', scale: number): string | null {
  const spc = pPr ? directChild(pPr, tag) : null
  if (!spc) return null
  const pts = firstDescendant(spc, 'a:spcPts')
  const ptsVal = pts ? parseInt(attr(pts, 'val') ?? '', 10) : NaN
  if (Number.isFinite(ptsVal) && ptsVal > 0) return `${Math.round((ptsVal / 100 / 0.75) * scale * 100) / 100}px`
  const pct = firstDescendant(spc, 'a:spcPct')
  const pctVal = pct ? parseInt(attr(pct, 'val') ?? '', 10) : NaN
  if (Number.isFinite(pctVal) && pctVal > 0) return `${pctVal / 100000}em`
  return null
}

function bulletKind(p: Element): 'none' | 'bullet' | 'number' {
  const pPr = directChild(p, 'a:pPr')
  if (!pPr) return 'none'
  if (directChild(pPr, 'a:buNone')) return 'none'
  if (directChild(pPr, 'a:buChar')) return 'bullet'
  if (directChild(pPr, 'a:buAutoNum')) return 'number'
  return 'none'
}

/** run 字符样式提取：rPr 显式属性优先，缺失时按段落 defRPr → lstStyle defRPr 逐级回退。
 * 属性分两类：字号/粗斜下划线/字距在元素属性上，颜色/字体在子元素上，两类走同一回退链。
 *  fontScale：normAutofit 的缩字系数（PowerPoint 把文本压进固定框时会缩小字号与字距）。 */
function runStyles(rPr: Element | null, defRPrs: (Element | null)[], theme: PptxTheme, scale: number, defaultSz: number, fontScale = 1): string[] {
  const styles: string[] = []
  const attrOf = (name: string): string | null =>
    (rPr ? attr(rPr, name) : null) ?? defRPrs.map(d => (d ? attr(d, name) : null)).find(v => v !== null) ?? null
  const childOf = (tag: string): Element | null =>
    (rPr ? directChild(rPr, tag) : null) ?? defRPrs.map(d => (d ? directChild(d, tag) : null)).find(v => v !== null) ?? null
  const sz = attrOf('sz')
  // 未声明字号时显式回退段落默认字号（presentation defaultTextStyle lvl1pPr），
  // 否则画布继承浏览器默认 16px 导致文字偏小
  // 保留两位小数：取整会让导出往返放大误差（20pt→26.67→27px→20.25pt，居中长文本错位累积）
  const szVal = sz ? parseInt(sz, 10) : defaultSz
  const textScale = scale * fontScale
  styles.push(`font-size:${Math.round((szVal / 100 / 0.75) * textScale * 100) / 100}px`)
  if (attrOf('b') === '1') styles.push('font-weight:bold')
  if (attrOf('i') === '1') styles.push('font-style:italic')
  // 下划线/删除线：sng/dbl 共存时合并为一条 text-decoration，dbl 另加 style
  const u = attrOf('u')
  const strike = attrOf('strike')
  const decos: string[] = []
  if (u === 'sng' || u === 'dbl') decos.push('underline')
  if (strike === 'sng' || strike === 'dbl') decos.push('line-through')
  if (decos.length) {
    styles.push(`text-decoration:${decos.join(' ')}`)
    if (u === 'dbl' || strike === 'dbl') styles.push('text-decoration-style:double')
  }
  // 上/下标：baseline 为百分比（正上负下），PowerPoint 只抬基线不缩小字号
  const baseline = parseInt(attrOf('baseline') ?? '0', 10)
  if (baseline > 0) styles.push('vertical-align:super')
  else if (baseline < 0) styles.push('vertical-align:sub')
  // 大小写
  const cap = attrOf('cap')
  if (cap === 'all') styles.push('text-transform:uppercase')
  else if (cap === 'small') styles.push('font-variant:small-caps')
  const spc = attrOf('spc')
  if (spc) {
    // 字距 1/100 pt → px（与 font-size 同源换算），保留两位小数防浮点噪声
    const ls = Math.round((parseInt(spc, 10) / 100 / 0.75) * textScale * 100) / 100
    styles.push(`letter-spacing:${ls}px`)
  }
  // 高亮底色
  const highlight = childOf('a:highlight')
  const highlightColor = resolveColor(highlight, theme)
  if (highlightColor) styles.push(`background-color:${highlightColor}`)
  const fill = childOf('a:solidFill')
  let color = resolveColor(fill, theme)
  // 空心字（run 显式 noFill + a:ln 描边）：编辑器不支持文字描边，取描边色近似。
  // 优先于 defRPr 继承色——源文件 lstStyle 常给 accent1 等默认色，回退会让绿描边字串成蓝
  if (childOf('a:noFill')) {
    const ln = childOf('a:ln')
    const lnColor = ln ? resolveColor(directChild(ln, 'a:solidFill'), theme) : null
    if (lnColor) color = lnColor
  }
  if (color) styles.push(`color:${color}`)
  // 文字描边（a:ln）：HTML 端用 -webkit-text-stroke 近似
  const runLn = childOf('a:ln')
  const strokeColor = runLn ? resolveColor(directChild(runLn, 'a:solidFill'), theme) : null
  if (strokeColor) {
    const w = Math.max(1, Math.round((parseInt(attr(runLn, 'w') ?? '0', 10) / 9525) * scale))
    styles.push(`-webkit-text-stroke:${w}px ${strokeColor}`)
  }
  // 文字阴影（rPr/defRPr effectLst/outerShdw）
  const effectLst = childOf('a:effectLst')
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (shdw) {
    const dist = (parseInt(attr(shdw, 'dist') ?? '0', 10) / 9525) * scale
    const dirRad = (parseInt(attr(shdw, 'dir') ?? '0', 10) / 60000) * (Math.PI / 180)
    const blur = Math.round((parseInt(attr(shdw, 'blurRad') ?? '0', 10) / 9525) * scale)
    const sc = resolveColor(shdw, theme) ?? '#00000066'
    styles.push(`text-shadow:${Math.round(dist * Math.cos(dirRad))}px ${Math.round(dist * Math.sin(dirRad))}px ${blur}px ${sc}`)
  }
  // 字体栈：latin + ea（含 +mj/+mn 主题引用）；run 未声明字体时回退主题字体——多数源文件
  // 字体继承自主题（run 无 a:latin/a:ea），完全不写字体会让导出 PPTX 在 PowerPoint 中
  // 回退到默认宋体。字体名来自不可信属性，fontStackOf 内已剔除引号防注入
  const latin = childOf('a:latin')
  const ea = childOf('a:ea')
  styles.push(`font-family:${fontStackOf(attr(latin, 'typeface') ?? '+mn-lt', attr(ea, 'typeface') ?? '+mn-ea', theme)}`)
  return styles
}

/** a:normAutofit@fontScale（1/1000 百分比）→ 字号缩放系数。
 *  PowerPoint 用它把文本压进固定框（如单行标题），忽略会让 CJK 折行溢出。 */
export function autofitFontScale(bodyPr: Element | null): number {
  const norm = bodyPr ? directChild(bodyPr, 'a:normAutofit') : null
  const raw = norm ? parseInt(attr(norm, 'fontScale') ?? '', 10) : NaN
  if (!Number.isFinite(raw) || raw <= 0) return 1
  return Math.min(1, Math.max(0.1, raw / 100000))
}

/** lstStyle → 段落级别对应的属性节点（a:lvl{N}pPr，未命中级别时回退 a:defPPr） */
function lstLvlPrOf(lst: Element | null, lvl: number): Element | null {
  if (!lst) return null
  return directChild(lst, `a:lvl${lvl + 1}pPr`) ?? directChild(lst, 'a:defPPr')
}

/** lstStyle → 段落级别对应的默认字符样式（a:lvl{N}pPr > a:defRPr） */
export function defRPrOf(lst: Element | null, lvl: number): Element | null {
  const lvlPr = lstLvlPrOf(lst, lvl)
  return lvlPr ? directChild(lvlPr, 'a:defRPr') : null
}

/** lstStyle → 段落默认对齐（a:lvl{N}pPr@algn，未命中回退 a:defPPr）。
 * 大量源文件段落不写显式 a:pPr，对齐只声明在 lstStyle（如叠加在形状上的居中文本框），
 * 缺此回退会把居中文字渲染成左对齐 */
export function lstAlgnOf(lst: Element | null, lvl: number): string | null {
  const lvlPr = lstLvlPrOf(lst, lvl)
  return lvlPr ? attr(lvlPr, 'algn') : null
}

async function paragraphToInner(p: Element, theme: PptxTheme, pkg: PptxPackage | undefined, partPath: string, defRPrs: (Element | null)[], scale: number, defaultSz: number, fontScale = 1): Promise<string> {
  let inner = ''
  for (const node of Array.from(p.children)) {
    if (node.localName === 'br') {
      inner += '<br>'
      continue
    }
    // a:r 普通 run；a:fld 字段（页码/日期等）结构相同（a:rPr + a:t），按缓存文本渲染为静态 run
    if (node.localName !== 'r' && node.localName !== 'fld') continue
    const t = firstDescendant(node, 'a:t')
    const text = t?.textContent ?? ''
    if (!text) continue
    const rPr = firstDescendant(node, 'a:rPr')
    const styles = runStyles(rPr, defRPrs, theme, scale, defaultSz, fontScale)
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

/** 占位符继承的文字属性（版式/母版原型 lstStyle 与 bodyPr） */
export interface TextInherit {
  lstStyle?: Element
  bodyPr?: Element
}

/** 悬挂缩进（marL/indent）EMU → px */
function emuPx(emu: number, scale: number): string {
  return `${Math.round((emu / 9525) * scale * 100) / 100}px`
}

/** txBody → 富文本 HTML；autoSize 来自 normAutofit；vertical 来自 bodyPr@vert。
 *  inherit：占位符未声明的 lstStyle/bodyPr 从版式/母版原型继承。
 *  flatBullets：项目符号展开为行首字符 + 悬挂缩进（画布形状文本不支持 ul 列表样式时使用）。
 *  scale：目标画布相对 96dpi 的缩放（导入画布 1280px 时源页 960px → 1.333），字号/间距/缩进同比例。
 *  defaultSz：run 未声明字号时的段落默认字号（1/100 pt，来自 presentation defaultTextStyle lvl1pPr）。 */
export async function txBodyToHTML(
  txBody: Element,
  theme: PptxTheme,
  pkg?: PptxPackage,
  partPath = '',
  inherit?: TextInherit,
  opts?: { flatBullets?: boolean; scale?: number; defaultSz?: number },
): Promise<TextBodyResult> {
  const scale = opts?.scale ?? 1
  const defaultSz = opts?.defaultSz ?? 1800
  const bodyPr = directChild(txBody, 'a:bodyPr') ?? inherit?.bodyPr ?? null
  const autoSize = Boolean(bodyPr && directChild(bodyPr, 'a:normAutofit'))
  // 字号按 normAutofit 缓存比例折算：与 PowerPoint 排版一致，避免固定框内折行/溢出
  const fontScale = autofitFontScale(bodyPr)
  const vert = attr(bodyPr, 'vert')
  const vertical = vert === 'eaVert' || vert === 'vert' || vert === 'mongolianVert'
  // insets（EMU→px）：OOXML 默认 tIns/bIns=45720；padding 单值取上下平均
  let padding: number | null = null
  if (bodyPr) {
    const tIns = parseInt(attr(bodyPr, 'tIns') ?? '45720', 10) || 0
    const bIns = parseInt(attr(bodyPr, 'bIns') ?? '45720', 10) || 0
    padding = Math.round(((tIns + bIns) / 2 / 9525) * scale)
  }
  const autoFitShape = Boolean(bodyPr && directChild(bodyPr, 'a:spAutoFit'))
  const anchorRaw = attr(bodyPr, 'anchor')
  const anchor: 't' | 'ctr' | 'b' | null =
    anchorRaw === 'ctr' || anchorRaw === 'b' || anchorRaw === 't' ? anchorRaw : null

  const paragraphs = directChildren(txBody, 'a:p')
  const lst = directChild(txBody, 'a:lstStyle') ?? inherit?.lstStyle ?? null
  const paras: ParaInfo[] = []
  for (const p of paragraphs) {
    const pPr = directChild(p, 'a:pPr')
    const lvlRaw = parseInt(attr(pPr, 'lvl') ?? '0', 10)
    const lvl = Number.isFinite(lvlRaw) ? lvlRaw : 0
    // 对齐回退链：显式 a:pPr@algn → lstStyle lvl{N}pPr@algn（OOXML 缺省左对齐）
    const algn = attr(pPr, 'algn') ?? lstAlgnOf(lst, lvl)
    const align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left'
    // 段落默认字符样式回退链：a:pPr/a:defRPr → lstStyle lvl{N}pPr/a:defRPr → presentation 默认字号。
    // 源文件常把字号/颜色/行距只写在段落 defRPr（目录标题、封面主标题等），缺失会导致字号颜色整体丢失
    const pDefRPr = pPr ? directChild(pPr, 'a:defRPr') : null
    const lstLvlPr = lstLvlPrOf(lst, lvl)
    const inner = await paragraphToInner(p, theme, pkg, partPath, [pDefRPr, defRPrOf(lst, lvl)], scale, defaultSz, fontScale)
    // 行距：spcPct（1/100000）是相对 PowerPoint 单倍行距的比例，单倍行距≈1.2em（CJK 字体
    // ascent/descent 约为 1.2），直接当 CSS line-height 会让全文行距偏紧约 17%；
    // spcPts（1/100 pt → px）覆盖；val 非法/非正数时跳过，避免产出 line-height:0 压扁文字
    const lnSpc = (pPr ? directChild(pPr, 'a:lnSpc') : null) ?? (lstLvlPr ? directChild(lstLvlPr, 'a:lnSpc') : null)
    let spacing = ''
    const pct = lnSpc ? firstDescendant(lnSpc, 'a:spcPct') : null
    const pctVal = pct ? parseInt(attr(pct, 'val') ?? '', 10) : NaN
    if (Number.isFinite(pctVal) && pctVal > 0) spacing = `line-height:${Math.round((pctVal / 100000) * 1.2 * 1000) / 1000}`
    const pts = lnSpc ? firstDescendant(lnSpc, 'a:spcPts') : null
    const ptsVal = pts ? parseInt(attr(pts, 'val') ?? '', 10) : NaN
    if (Number.isFinite(ptsVal) && ptsVal > 0) spacing = `line-height:${Math.round((ptsVal / 100 / 0.75) * scale * fontScale * 100) / 100}px`
    // 段前/段后 + 层级缩进 + 阅读方向（RTL）
    const extra: string[] = []
    const mt = paraMargin(pPr, 'a:spcBef', scale)
    if (mt) extra.push(`margin-top:${mt}`)
    const mb = paraMargin(pPr, 'a:spcAft', scale)
    if (mb) extra.push(`margin-bottom:${mb}`)
    // 段落左缩进/首行缩进（EMU）：悬挂缩进 = marL 正 + indent 负，决定项目符号与折行对齐
    const marL = parseInt(attr(pPr, 'marL') ?? '', 10)
    const indent = parseInt(attr(pPr, 'indent') ?? '', 10)
    if (Number.isFinite(marL) && marL > 0) extra.push(`margin-left:${emuPx(marL, scale)}`)
    else if (lvl > 0) extra.push(`margin-left:${lvl * 1.5}em`)
    if (Number.isFinite(indent) && indent !== 0) extra.push(`text-indent:${emuPx(indent, scale)}`)
    if (attr(pPr, 'rtl') === '1') extra.push('direction:rtl')
    const buAuto = pPr ? directChild(pPr, 'a:buAutoNum') : null
    const startAtRaw = buAuto ? parseInt(attr(buAuto, 'startAt') ?? '', 10) : NaN
    const startAt = Number.isFinite(startAtRaw) && startAtRaw > 1 ? startAtRaw : undefined
    const buChar = pPr ? directChild(pPr, 'a:buChar') : null
    const rawChar = buChar ? attr(buChar, 'char') : null
    const bulletChar = rawChar ? bulletGlyph(rawChar, pPr ? directChild(pPr, 'a:buFont') : null) : undefined
    // 悬挂缩进：正文折行对齐到 marL，符号占位宽度取 marL（首行 indent 负值拉到行首）
    const hangPx = Number.isFinite(marL) && marL > 0 && Number.isFinite(indent) && indent < 0 ? emuPx(marL, scale) : undefined
    paras.push({ align, kind: bulletKind(p), inner, spacing, extra, level: lvl, startAt, bulletChar, hangPx })
  }

  const styleOf = (para: ParaInfo): string =>
    [`text-align:${para.align}`, para.spacing, ...para.extra].filter(Boolean).join(';')

  const html: string[] = []
  let i = 0
  while (i < paras.length) {
    const para = paras[i]
    if (para.kind === 'none') {
      html.push(`<p style="${styleOf(para)}">${para.inner || '&nbsp;'}</p>`)
      i += 1
      continue
    }
    if (opts?.flatBullets) {
      // 画布形状文本层无列表样式：符号作为行首字符写入段落，配合 marL/indent 形成悬挂缩进
      let n = para.kind === 'number' ? para.startAt ?? 1 : 0
      while (i < paras.length && paras[i].kind === para.kind) {
        const cur = paras[i]
        const glyph = cur.kind === 'bullet' ? cur.bulletChar ?? '•' : `${n}.`
        if (cur.kind === 'number') n += 1
        i += 1
        // 源文件常留空的项目符号段落：PowerPoint 不显示空符号（编号继续计数），跳过避免画布出现孤立序号
        if (!cur.inner) continue
        const marker = cur.hangPx
          ? `<span style="display:inline-block;width:${cur.hangPx};text-indent:0">${escapeHTML(glyph)}</span>`
          : `${escapeHTML(glyph)}&nbsp;`
        html.push(`<p style="${styleOf(cur)}">${marker}${cur.inner}</p>`)
      }
      continue
    }
    const tag = para.kind === 'bullet' ? 'ul' : 'ol'
    const start = tag === 'ol' && para.startAt ? ` start="${para.startAt}"` : ''
    const items: string[] = []
    while (i < paras.length && paras[i].kind === para.kind) {
      const cur = paras[i]
      items.push(`<li style="${styleOf(cur)}">${cur.inner || '&nbsp;'}</li>`)
      i += 1
    }
    html.push(`<${tag}${start}>${items.join('')}</${tag}>`)
  }
  return { html: html.join('') || '<p></p>', autoSize, vertical, padding, autoFitShape, anchor }
}
