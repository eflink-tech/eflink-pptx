// src/core/import/master.ts
/** slide ancestry：layout/master 链、继承背景、装饰元素、占位符位置表 */
import { attr, directChild, directChildren, firstDescendant, parseXML } from './xml'
import { resolveColor } from './styles'
import { parseThemeForMaster } from './theme'
import type { PptxTheme } from './theme'
import type { PptxPackage } from './package'
import type { ParseContext } from './context'
import { IDENTITY_XFORM } from './context'
import type { Background, Gradient, ImportReport, PPTElement } from '../../types/slides'
import { parseSpTreeNode } from './elements'

export interface SlideAncestry {
  layoutPath: string | null
  masterPath: string | null
  theme: PptxTheme
  /** 版式/母版继承背景（slide 自身 bg 优先，由调用方处理） */
  background?: Background
  /** 版式 + 母版的非占位符装饰元素（已按 scale 缩放到目标画布 px） */
  decorations: PPTElement[]
  /** 占位符位置表（EMU；key = idx ?? type） */
  placeholders: Map<string, { x: number; y: number; w: number; h: number }>
}

/** slide → layout / layout → master 关系 Type URI（与 theme.ts THEME_REL_TYPE 同模式） */
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'

async function relTargetByType(pkg: PptxPackage, partPath: string, relType: string): Promise<string | null> {
  const rels = await pkg.rels(partPath)
  for (const rel of rels.values()) {
    if (rel.mode !== 'External' && rel.type === relType) return rel.target
  }
  return null
}

/** slide → layout → master 部件链 */
export async function findAncestry(pkg: PptxPackage, slidePath: string): Promise<{ layoutPath: string | null; masterPath: string | null }> {
  const layoutPath = await relTargetByType(pkg, slidePath, SLIDE_LAYOUT_REL)
  const masterPath = layoutPath ? await relTargetByType(pkg, layoutPath, SLIDE_MASTER_REL) : null
  return { layoutPath, masterPath }
}

/** 收集 spTree 中带 p:ph 且有 xfrm 的占位符位置（EMU） */
export function collectPlaceholders(spTree: Element): Map<string, { x: number; y: number; w: number; h: number }> {
  const map = new Map<string, { x: number; y: number; w: number; h: number }>()
  for (const sp of Array.from(spTree.children)) {
    if (sp.nodeName !== 'p:sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (!ph) continue
    const xfrm = firstDescendant(sp, 'a:xfrm')
    const off = xfrm ? directChild(xfrm, 'a:off') : null
    const ext = xfrm ? directChild(xfrm, 'a:ext') : null
    if (!off || !ext) continue
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    if (!key || map.has(key)) continue
    map.set(key, {
      x: parseInt(attr(off, 'x') ?? '0', 10),
      y: parseInt(attr(off, 'y') ?? '0', 10),
      w: parseInt(attr(ext, 'cx') ?? '0', 10),
      h: parseInt(attr(ext, 'cy') ?? '0', 10),
    })
  }
  return map
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

/** p:bg → Background（solid/grad/blip；bgRef 以颜色子节点走 solid） */
export function parseBackgroundFill(bg: Element | null, theme: PptxTheme): Background | undefined {
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
    // 其余（blipFill 等）：图片背景，src 由调用方按 rels 补充
    return { type: 'image' }
  }
  if (bgRef) {
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
  const decorations: PPTElement[] = []
  let background: Background | undefined

  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath) continue
    const xml = await pkg.text(partPath)
    if (!xml) continue
    const doc = parseXML(xml)
    const spTree = firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')
    if (!spTree) continue
    for (const [k, v] of collectPlaceholders(spTree)) {
      if (!placeholders.has(k)) placeholders.set(k, v)
    }
    const bg = firstDescendant(doc.documentElement, 'p:bg')
    const partBg = parseBackgroundFill(bg, theme)
    if (partBg && bg && !background) {
      if (partBg.type === 'image') {
        // 图片背景：src 解析成功才赋值，否则保持 undefined 由后续部件/lt1 兜底（避免悬空 image 背景）
        const blip = firstDescendant(bg, 'a:blip')
        const embedId = attr(blip, 'r:embed')
        const target = embedId ? await pkg.relTarget(partPath, embedId) : null
        const src = target ? await pkg.mediaDataUrl(target) : undefined
        if (src) background = { type: 'image', image: { src, size: 'cover' } }
      } else {
        background = partBg
      }
    }
    // 非占位符装饰元素
    const ctx: ParseContext = { pkg, partPath, theme, report, scale, placeholders }
    for (const child of Array.from(spTree.children)) {
      if (!['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp'].includes(child.nodeName)) continue
      if (firstDescendant(child, 'p:ph')) continue
      const els = await parseSpTreeNode(child, ctx, IDENTITY_XFORM, pkg)
      decorations.push(...els)
    }
  }

  if (!background) {
    const lt1 = theme.schemeColors.lt1 ?? '#FFFFFF'
    background = { type: 'solid', color: lt1 }
  }
  return { layoutPath, masterPath, theme, background, decorations, placeholders }
}
