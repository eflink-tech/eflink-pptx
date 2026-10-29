// src/core/import/master.ts
/** slide ancestry：layout/master 链、继承背景、装饰元素、占位符位置表 */
import { attr, directChild, directChildren, firstDescendant, parseXML } from './xml'
import { refColor, resolveColor, styleMatrixEntry } from './styles'
import { parseThemeForMaster } from './theme'
import type { PptxTheme } from './theme'
import { relTypeIs, type PptxPackage } from './package'
import type { ParseContext } from './context'
import { IDENTITY_XFORM, addSkipped } from './context'
import type { Background, Gradient, ImportReport, PPTElement } from '../../types/slides'
import { parseSpTreeNode } from './elements'

export interface SlideAncestry {
  layoutPath: string | null
  masterPath: string | null
  theme: PptxTheme
  /** 版式/母版继承背景（slide 自身 bg 优先，由调用方处理） */
  background?: Background
  /** 母版 + 版式的非占位符装饰元素（master 在下、layout 在上，已按 scale 缩放到目标画布 px） */
  decorations: PPTElement[]
  /** 占位符位置表（EMU；key = idx 与 type 均可查） */
  placeholders: Map<string, { x: number; y: number; w: number; h: number }>
  /** 占位符文字样式原型（a:lstStyle / a:bodyPr，key = idx 或 type） */
  placeholderText: Map<string, { lstStyle?: Element; bodyPr?: Element }>
  /** 占位符外观原型（spPr / p:style，key = idx 或 type）：几何（custGeom/prstGeom）、填充、边框继承 */
  placeholderShapes: Map<string, { spPr?: Element; style?: Element }>
  /** 母版 p:txStyles（title/body/other） */
  masterTextStyles: { title?: Element; body?: Element; other?: Element }
}

/** slide → layout / layout → master 关系查找（Type 按末段匹配，兼容 strict OOXML） */
async function relTargetByType(pkg: PptxPackage, partPath: string, relName: string): Promise<string | null> {
  const rels = await pkg.rels(partPath)
  for (const rel of rels.values()) {
    if (rel.mode !== 'External' && relTypeIs(rel.type, relName)) return rel.target
  }
  return null
}

/** slide → layout → master 部件链 */
export async function findAncestry(pkg: PptxPackage, slidePath: string): Promise<{ layoutPath: string | null; masterPath: string | null }> {
  const layoutPath = await relTargetByType(pkg, slidePath, 'slideLayout')
  const masterPath = layoutPath ? await relTargetByType(pkg, layoutPath, 'slideMaster') : null
  return { layoutPath, masterPath }
}

/** 收集 spTree 中带 p:ph 且有 xfrm 的占位符位置（EMU）；idx 与 type 双键登记便于 slide 侧回查 */
export function collectPlaceholders(spTree: Element): Map<string, { x: number; y: number; w: number; h: number }> {
  const map = new Map<string, { x: number; y: number; w: number; h: number }>()
  for (const sp of Array.from(spTree.children)) {
    if (sp.localName !== 'sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (!ph) continue
    const xfrm = firstDescendant(sp, 'a:xfrm')
    const off = xfrm ? directChild(xfrm, 'a:off') : null
    const ext = xfrm ? directChild(xfrm, 'a:ext') : null
    if (!off || !ext) continue
    const box = {
      x: parseInt(attr(off, 'x') ?? '0', 10),
      y: parseInt(attr(off, 'y') ?? '0', 10),
      w: parseInt(attr(ext, 'cx') ?? '0', 10),
      h: parseInt(attr(ext, 'cy') ?? '0', 10),
    }
    const idx = attr(ph, 'idx')
    const type = attr(ph, 'type')
    if (idx && !map.has(idx)) map.set(idx, box)
    if (type && !map.has(type)) map.set(type, box)
  }
  return map
}

/** 收集占位符原型的外观（spPr 几何/填充/边框 + p:style）；idx 与 type 双键登记 */
export function collectPlaceholderShapes(spTree: Element): Map<string, { spPr?: Element; style?: Element }> {
  const map = new Map<string, { spPr?: Element; style?: Element }>()
  for (const sp of Array.from(spTree.children)) {
    if (sp.localName !== 'sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (!ph) continue
    const spPr = directChild(sp, 'p:spPr') ?? directChild(sp, 'a:spPr') ?? undefined
    const style = directChild(sp, 'p:style') ?? undefined
    if (!spPr && !style) continue
    const entry = { spPr, style }
    const idx = attr(ph, 'idx')
    const type = attr(ph, 'type')
    if (idx && !map.has(idx)) map.set(idx, entry)
    if (type && !map.has(type)) map.set(type, entry)
  }
  return map
}

/** 收集占位符原型的文字样式（版式/母版 sp 的 lstStyle/bodyPr）；idx 与 type 双键登记 */
export function collectPlaceholderText(
  spTree: Element,
): Map<string, { lstStyle?: Element; bodyPr?: Element }> {
  const map = new Map<string, { lstStyle?: Element; bodyPr?: Element }>()
  for (const sp of Array.from(spTree.children)) {
    if (sp.localName !== 'sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (!ph) continue
    const txBody = firstDescendant(sp, 'p:txBody')
    if (!txBody) continue
    const lstStyle = directChild(txBody, 'a:lstStyle') ?? undefined
    const bodyPr = directChild(txBody, 'a:bodyPr') ?? undefined
    if (!lstStyle && !bodyPr) continue
    const entry = { lstStyle, bodyPr }
    const idx = attr(ph, 'idx')
    const type = attr(ph, 'type')
    if (idx && !map.has(idx)) map.set(idx, entry)
    if (type && !map.has(type)) map.set(type, entry)
  }
  return map
}

/** 母版 p:txStyles：titleStyle/bodyStyle/otherStyle（层级默认字符样式） */
function parseMasterTextStyles(masterRoot: Element): { title?: Element; body?: Element; other?: Element } {
  const txStyles = firstDescendant(masterRoot, 'p:txStyles')
  if (!txStyles) return {}
  return {
    title: directChild(txStyles, 'p:titleStyle') ?? undefined,
    body: directChild(txStyles, 'p:bodyStyle') ?? undefined,
    other: directChild(txStyles, 'p:otherStyle') ?? undefined,
  }
}

/** a:gradFill → 渐变背景（停站色 + 角度），供 Background.gradient 使用 */
function gradToGradient(gradFill: Element, theme: PptxTheme): Gradient | undefined {
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildren(gsLst, 'a:gs') : []
  if (!stops.length) return undefined
  const lin = directChild(gradFill, 'a:lin')
  return {
    type: 'linear',
    colors: stops.map((gs) => ({
      pos: parseInt(attr(gs, 'pos') ?? '0', 10) / 100000,
      color: resolveColor(gs, theme) ?? '#000000',
    })),
    rotate: lin ? Math.round(parseInt(attr(lin, 'ang') ?? '0', 10) / 60000) : 0,
  }
}

/** a:pattFill（lgGrid 网格）→ SVG 平铺图背景；设计稿常用网格纹理底，其余图案不识别返回 undefined */
function pattFillToBackground(patt: Element, theme: PptxTheme): Background | undefined {
  if (attr(patt, 'prst') !== 'lgGrid') return undefined
  const fg = resolveColor(directChild(patt, 'a:fgClr'), theme) ?? '#F2F2F2'
  const bg = resolveColor(directChild(patt, 'a:bgClr'), theme) ?? '#FFFFFF'
  // OOXML 预设图案 tile 逻辑尺寸为 8×8 单位，主流渲染器（LibreOffice 实测）按 0.1 cm
  // 渲染 tile → 网格距 ≈ 3.78px@96dpi；导出侧 tile 固化为位图，按此密度合成（线宽 0.5px、
  // 置于 0.25 偏移抗锯齿），密度取整会使长程相位漂移、网格大面积错位
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="3.78" height="3.78"><rect width="3.78" height="3.78" fill="${bg}"/><path d="M0.25 0V3.78M0 0.25H3.78" stroke="${fg}" stroke-width="0.5" fill="none"/></svg>`
  return { type: 'image', image: { src: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`, size: 'repeat' } }
}

/** p:bg → Background（solid/grad/blip；bgRef 引用主题 fillStyleLst；lgGrid 图案转平铺图）。
 *  图片背景返回 src 为空的 image（size 标记平铺/覆盖），src 由调用方按 rels 补充。 */
export function parseBackgroundFill(bg: Element | null, theme: PptxTheme, report?: ImportReport): Background | undefined {
  if (!bg) return undefined
  const bgPr = directChild(bg, 'p:bgPr')
  const bgRef = directChild(bg, 'p:bgRef')
  if (bgPr) {
    const solid = directChild(bgPr, 'a:solidFill')
    if (solid) {
      const color = resolveColor(solid, theme)
      if (color) return { type: 'solid', color }
    }
    const grad = directChild(bgPr, 'a:gradFill')
    if (grad) {
      const gradient = gradToGradient(grad, theme)
      if (gradient) return { type: 'gradient', gradient }
    }
    const patt = directChild(bgPr, 'a:pattFill')
    if (patt) {
      const pattern = pattFillToBackground(patt, theme)
      if (pattern) return pattern
      if (report) addSkipped(report, 'patternFill')
      return undefined
    }
    const blip = directChild(bgPr, 'a:blipFill')
    const tile = blip ? directChild(blip, 'a:tile') : null
    return { type: 'image', image: { src: '', size: tile ? 'repeat' : 'cover' } }
  }
  if (bgRef) {
    // bgRef idx：1001+ 指向 fillStyleLst[idx-1000]，0 = 纯色子节点
    const idx = parseInt(attr(bgRef, 'idx') ?? '0', 10)
    const ph = refColor(bgRef)
    const matrix = idx >= 1001 ? styleMatrixEntry(theme, 'fill', idx - 1000) : null
    if (matrix?.localName === 'solidFill') {
      const color = resolveColor(matrix, theme, ph)
      if (color) return { type: 'solid', color }
    }
    if (matrix?.localName === 'gradFill') {
      const gradient = gradToGradient(matrix, theme)
      if (gradient) return { type: 'gradient', gradient }
    }
    if (matrix?.localName === 'pattFill' && report) addSkipped(report, 'patternFill')
    const color = resolveColor(bgRef, theme)
    if (color) return { type: 'solid', color }
  }
  return undefined
}

/** 解析 slide 的 ancestry（layout/master 背景与装饰；装饰按 partPath 解析 rels，用传入 scale 缩放） */
export async function parseSlideAncestry(
  pkg: PptxPackage,
  slidePath: string,
  report: ImportReport,
  scale: { x: number; y: number },
): Promise<SlideAncestry> {
  const { layoutPath, masterPath } = await findAncestry(pkg, slidePath)
  const theme = await parseThemeForMaster(pkg, masterPath)

  const placeholders = new Map<string, { x: number; y: number; w: number; h: number }>()
  const placeholderText = new Map<string, { lstStyle?: Element; bodyPr?: Element }>()
  const placeholderShapes = new Map<string, { spPr?: Element; style?: Element }>()
  const decorations: PPTElement[] = []
  let background: Background | undefined

  // 部件文档按需解析：单个版式/母版畸形时降级跳过该部件，不拖垮整页
  const docs = new Map<string, Document>()
  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath) continue
    const xml = await pkg.text(partPath)
    if (!xml) continue
    try {
      docs.set(partPath, parseXML(xml))
    } catch (e) {
      console.warn('[pptx-import] 版式/母版解析失败:', partPath, e)
      addSkipped(report, 'slideParseFailed')
    }
  }

  const masterRoot = masterPath ? docs.get(masterPath)?.documentElement : undefined
  const layoutRoot = layoutPath ? docs.get(layoutPath)?.documentElement : undefined
  const masterTextStyles = masterRoot ? parseMasterTextStyles(masterRoot) : {}
  // showMasterSp=0：布局隐藏母版形状；母版自身声明 0 时母版装饰不显示
  const layoutShowsMaster = !layoutRoot || attr(layoutRoot, 'showMasterSp') !== '0'
  const masterShowsOwn = !masterRoot || attr(masterRoot, 'showMasterSp') !== '0'

  // 背景回退：版式 → 母版（slide 自身 bg 由调用方优先处理）
  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath || background) continue
    const doc = docs.get(partPath)
    if (!doc) continue
    const bg = firstDescendant(doc.documentElement, 'p:bg')
    const partBg = parseBackgroundFill(bg, theme, report)
    if (!partBg || !bg) continue
    if (partBg.type === 'image') {
      // 图片背景：src 解析成功才赋值，否则保持 undefined 由后续部件/lt1 兜底（避免悬空 image 背景）
      const blip = firstDescendant(bg, 'a:blip')
      const embedId = attr(blip, 'r:embed')
      const target = embedId ? await pkg.relTarget(partPath, embedId) : null
      const src = target ? await pkg.mediaDataUrl(target) : undefined
      if (src) background = { type: 'image', image: { src, size: partBg.image?.size ?? 'cover' } }
    } else {
      background = partBg
    }
  }

  // 占位符位置/文字样式：版式优先于母版（不覆盖已存在键）
  for (const partPath of [layoutPath, masterPath]) {
    const doc = partPath ? docs.get(partPath) : undefined
    if (!doc) continue
    const spTree = firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')
    if (!spTree) continue
    for (const [k, v] of collectPlaceholders(spTree)) {
      if (!placeholders.has(k)) placeholders.set(k, v)
    }
    for (const [k, v] of collectPlaceholderText(spTree)) {
      if (!placeholderText.has(k)) placeholderText.set(k, v)
    }
    for (const [k, v] of collectPlaceholderShapes(spTree)) {
      if (!placeholderShapes.has(k)) placeholderShapes.set(k, v)
    }
  }

  // 装饰层序：母版在下、版式在上（幻灯片内容再叠于其上）；showMasterSp=0 时隐藏对应层
  for (const partPath of [masterPath, layoutPath]) {
    if (!partPath) continue
    if (partPath === masterPath && (!layoutShowsMaster || !masterShowsOwn)) continue
    const doc = docs.get(partPath)
    if (!doc) continue
    const spTree = firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')
    if (!spTree) continue
    const ctx: ParseContext = { pkg, partPath, theme, report, scale, placeholders, placeholderText, placeholderShapes, masterTextStyles }
    for (const child of Array.from(spTree.children)) {
      if (!['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp', 'AlternateContent'].includes(child.localName)) continue
      if (firstDescendant(child, 'p:ph')) continue
      const els = await parseSpTreeNode(child, ctx, IDENTITY_XFORM, pkg)
      decorations.push(...els)
    }
  }

  if (!background) {
    const lt1 = theme.schemeColors.lt1 ?? '#FFFFFF'
    background = { type: 'solid', color: lt1 }
  }
  return { layoutPath, masterPath, theme, background, decorations, placeholders, placeholderText, placeholderShapes, masterTextStyles }
}
