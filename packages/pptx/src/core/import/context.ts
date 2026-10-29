// src/core/import/context.ts
/** 解析上下文：贯穿整个导入管线的共享状态 */
import type { ElementLink, ImportReport } from '../../types/slides'
import type { PptxPackage } from './package'
import type { PptxTheme } from './theme'
import { attr, directChild } from './xml'

export interface ParseContext {
  pkg: PptxPackage
  /** 当前部件路径（rels 查询基准，如 'ppt/slides/slide1.xml'） */
  partPath: string
  theme: PptxTheme
  report: ImportReport
  /** EMU→px 缩放（画布适配） */
  scale: { x: number; y: number }
  /** 版式/母版占位符位置表（EMU 空间，key = idx 或 type） */
  placeholders: Map<string, { x: number; y: number; w: number; h: number }>
  /** 占位符原型文字样式（key = idx/type → a:lstStyle / a:bodyPr）；slide 占位符继承版式/母版声明 */
  placeholderText?: Map<string, { lstStyle?: Element; bodyPr?: Element }>
  /** 占位符原型外观（key = idx/type → spPr / p:style）；几何、填充、边框按属性逐项继承 */
  placeholderShapes?: Map<string, { spPr?: Element; style?: Element }>
  /** 母版 p:txStyles（title/body/other）；占位符自身未声明 lstStyle 时的兜底 */
  masterTextStyles?: { title?: Element; body?: Element; other?: Element }
  /** 幻灯片尺寸（EMU），占位符几何缺失时用于兜底定位 */
  slideSize?: { w: number; h: number }
  /** 包内 slide 部件路径 → 导入后的 slide id（内部超链接解析） */
  slideIdByPath?: Map<string, string>
}

/** 记录一类降级/丢弃项 */
export function addSkipped(report: ImportReport, kind: string): void {
  report.skipped[kind] = (report.skipped[kind] ?? 0) + 1
}

/** 组合变换：子空间 EMU → 父空间 EMU 的完整仿射矩阵（含旋转/翻转），
 *  x' = a·x + c·y + e，y' = b·x + d·y + f；rot 为累计旋转角（度），flip 为累计翻转 */
export interface GroupXform {
  a: number
  b: number
  c: number
  d: number
  e: number
  f: number
  rot: number
  flipH: boolean
  flipV: boolean
}

export const IDENTITY_XFORM: Readonly<GroupXform> = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, rot: 0, flipH: false, flipV: false,
})

export function mapPoint(t: Readonly<GroupXform>, x: number, y: number): [number, number] {
  return [t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f]
}

/** 轴向缩放（旋转后 x 轴/y 轴的长度比例） */
export function xformScale(t: Readonly<GroupXform>): [number, number] {
  return [Math.hypot(t.a, t.b), Math.hypot(t.c, t.d)]
}

export function mapX(t: Readonly<GroupXform>, v: number): number {
  return t.a * v + t.e
}

export function mapY(t: Readonly<GroupXform>, v: number): number {
  return t.d * v + t.f
}

/** 复合变换 outer ∘ inner（先 inner 后 outer）；累计旋转在外层含奇数次翻转时反向（F·R(θ)·F = R(-θ)） */
export function composeXform(outer: Readonly<GroupXform>, inner: Readonly<GroupXform>): GroupXform {
  const reflected = outer.flipH !== outer.flipV
  return {
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
    rot: outer.rot + (reflected ? -inner.rot : inner.rot),
    flipH: outer.flipH !== inner.flipH,
    flipV: outer.flipV !== inner.flipV,
  }
}

/** 占位符位置：先按 idx 再按 type 查版式/母版原型（slide 占位符常只有 idx，原型只有 type） */
export function resolvePlaceholderGeometry(
  ctx: ParseContext,
  ph: Element,
): { x: number; y: number; w: number; h: number } | null {
  const idx = attr(ph, 'idx')
  const type = attr(ph, 'type')
  return (idx ? ctx.placeholders.get(idx) : undefined)
    ?? (type ? ctx.placeholders.get(type) : undefined)
    ?? null
}

/** 占位符继承的外观原型（spPr / p:style）：idx 优先、type 兜底 */
export function inheritedPlaceholderShape(
  ctx: ParseContext,
  ph: Element,
): { spPr?: Element; style?: Element } | undefined {
  const idx = attr(ph, 'idx')
  const type = attr(ph, 'type')
  return (idx ? ctx.placeholderShapes?.get(idx) : undefined)
    ?? (type ? ctx.placeholderShapes?.get(type) : undefined)
}

/** 占位符继承的文字样式：版式/母版原型 lstStyle 优先，其次母版 p:txStyles 对应槽位 */
export function inheritedPlaceholderText(
  ctx: ParseContext,
  ph: Element,
): { lstStyle?: Element; bodyPr?: Element } {
  const idx = attr(ph, 'idx')
  const type = attr(ph, 'type')
  const proto = (idx ? ctx.placeholderText?.get(idx) : undefined)
    ?? (type ? ctx.placeholderText?.get(type) : undefined)
  if (proto?.lstStyle) return proto
  const slot = type === 'title' || type === 'ctrTitle'
    ? 'title'
    : type === 'body' || type === 'subTitle' || type === 'obj'
      ? 'body'
      : 'other'
  const masterLst = ctx.masterTextStyles?.[slot]
  return { lstStyle: masterLst ?? proto?.lstStyle, bodyPr: proto?.bodyPr }
}

/** 元素级超链接（cNvPr/a:hlinkClick）：外部 http(s)/mailto → web，包内 slide 关系 → 内部页链接 */
export async function parseElementLink(
  cNvPr: Element | null,
  ctx: ParseContext,
): Promise<ElementLink | undefined> {
  const hlink = cNvPr ? directChild(cNvPr, 'a:hlinkClick') : null
  const rId = attr(hlink, 'r:id')
  if (!rId) return undefined
  const rel = (await ctx.pkg.rels(ctx.partPath)).get(rId)
  if (!rel) return undefined
  if (rel.mode === 'External') {
    return /^(https?:|mailto:)/i.test(rel.target) ? { type: 'web', target: rel.target } : undefined
  }
  const slideId = ctx.slideIdByPath?.get(rel.target)
  return slideId ? { type: 'slide', target: slideId } : undefined
}

/** 占位符几何缺失时的兜底位置（EMU）：保持内容可见并计入报告 */
export function fallbackPlaceholderBox(ph: Element, size: { w: number; h: number }): { x: number; y: number; w: number; h: number } {
  const type = attr(ph, 'type') ?? ''
  if (type === 'title' || type === 'ctrTitle') {
    return { x: size.w * 0.05, y: size.h * 0.03, w: size.w * 0.9, h: size.h * 0.15 }
  }
  if (type === 'dt' || type === 'ftr' || type === 'sldNum') {
    return { x: size.w * 0.4, y: size.h * 0.92, w: size.w * 0.2, h: size.h * 0.05 }
  }
  return { x: size.w * 0.08, y: size.h * 0.22, w: size.w * 0.84, h: size.h * 0.68 }
}
