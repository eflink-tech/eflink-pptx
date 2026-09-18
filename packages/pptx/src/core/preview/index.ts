// src/core/preview/index.ts
/** 预览入口：pptx → 每页一个 <svg>（OOXML 语义直渲，不做编辑器模型映射取舍） */
import { PptxPackage } from '../import/package'
import { listSlidePaths } from '../import'
import { addSkipped } from '../import/context'
import type { ImportReport } from '../../types/slides'
import { renderSlide } from './slide'

export interface PreviewResult {
  pages: SVGSVGElement[]
  report: ImportReport
}

/** pptx → 页面 SVG 列表 + 兼容性报告（单页失败计入报告不中断） */
export async function previewPPTXDetailed(file: File): Promise<PreviewResult> {
  const pkg = await PptxPackage.load(file)
  const { paths, srcW, srcH } = await listSlidePaths(pkg)
  const report: ImportReport = { skipped: {} }
  const pages: SVGSVGElement[] = []
  for (let i = 0; i < paths.length; i += 1) {
    try {
      const page = await renderSlide(pkg, paths[i], srcW, srcH, i, report)
      if (page) pages.push(page)
      else {
        console.warn('[pptx-preview] 页面部件缺失:', paths[i])
        addSkipped(report, 'slideParseFailed')
      }
    } catch (e) {
      // 单页失败不拖垮整个预览：计入报告后继续下一页
      console.warn('[pptx-preview] 页面渲染失败:', paths[i], e)
      addSkipped(report, 'slideParseFailed')
    }
  }
  if (!pages.length) throw new Error('PPTX 中没有可预览的幻灯片')
  return { pages, report }
}

/** 兼容设计文档 API：仅返回页面列表 */
export async function previewPPTX(file: File): Promise<SVGSVGElement[]> {
  return (await previewPPTXDetailed(file)).pages
}
