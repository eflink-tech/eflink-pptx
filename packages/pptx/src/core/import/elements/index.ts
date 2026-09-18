// src/core/import/elements/index.ts
/** spTree 子节点分发：p:sp / p:cxnSp / p:pic / p:graphicFrame / p:grpSp */
import { addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { PPTElement } from '../../../types/slides'
import { parseShapeEl } from './shape'
import { parsePictureEl } from './picture'
import { parseTableEl } from './table'
import { parseChartEl } from './chart'
import { parseGroupEl } from './group'

export async function parseSpTreeNode(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  try {
    switch (node.nodeName) {
      case 'p:sp':
      case 'p:cxnSp': {
        const el = await parseShapeEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'p:pic': {
        const el = await parsePictureEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'p:graphicFrame': {
        // graphicFrame 内既有 a:tbl 也可能挂 chart rels：表格优先，找不到 a:tbl 落到图表
        const el = (await parseTableEl(node, ctx, xf, pkg)) ?? (await parseChartEl(node, ctx, xf, pkg))
        return el ? [el] : []
      }
      case 'p:grpSp':
        return await parseGroupEl(node, ctx, xf, pkg)
      default:
        return []
    }
  } catch {
    addSkipped(ctx.report, 'elementParseFailed')
    return []
  }
}
