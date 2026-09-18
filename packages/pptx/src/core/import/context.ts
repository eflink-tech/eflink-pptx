// src/core/import/context.ts
/** 解析上下文：贯穿整个导入管线的共享状态 */
import type { ImportReport } from '../../types/slides'
import type { PptxPackage } from './package'
import type { PptxTheme } from './theme'

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
}

/** 记录一类降级/丢弃项 */
export function addSkipped(report: ImportReport, kind: string): void {
  report.skipped[kind] = (report.skipped[kind] ?? 0) + 1
}

/** 组合变换：EMU 空间的平移 + 轴缩放仿射（rot 为需要叠加到子元素的角度 1/60000 deg） */
export interface GroupXform {
  ox: number
  oy: number
  sx: number
  sy: number
  rot: number
}

export const IDENTITY_XFORM: GroupXform = { ox: 0, oy: 0, sx: 1, sy: 1, rot: 0 }

export function mapX(t: GroupXform, v: number): number {
  return t.ox + v * t.sx
}

export function mapY(t: GroupXform, v: number): number {
  return t.oy + v * t.sy
}
