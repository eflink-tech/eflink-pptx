/** spTree 单节点分发器：p:sp/p:cxnSp/p:pic/p:graphicFrame/p:grpSp → 对应渲染器；未知节点跳过。
 * 独立成文件打破 slide → group ↔ dispatch 的模块循环（函数级互调在 ESM 下经由 live binding 正常工作）。 */
import { firstDescendant } from '../import/xml'
import { addSkipped } from '../import/context'
import type { PreviewCtx } from './svg'
import { renderShape } from './shape'
import { renderPicture } from './picture'
import { renderTable } from './table'
import { renderChart } from './chart'
import { renderGroup } from './group'

export async function renderSpTreeNode(node: Element, ctx: PreviewCtx): Promise<SVGElement[]> {
  try {
    switch (node.nodeName) {
      case 'p:sp':
      case 'p:cxnSp': {
        const el = await renderShape(node, ctx)
        return el ? [el] : []
      }
      case 'p:pic': {
        const el = await renderPicture(node, ctx)
        return el ? [el] : []
      }
      case 'p:graphicFrame': {
        // 表格（a:tbl）与图表（c:chart）共用 graphicFrame 容器；SmartArt/OLE 等未知类型计入报告
        const isTable = Boolean(firstDescendant(node, 'a:tbl'))
        const isChart = Boolean(firstDescendant(node, 'c:chart'))
        if (!isTable && !isChart) addSkipped(ctx.report, 'graphicFrameUnknown')
        const el = isTable ? await renderTable(node, ctx) : await renderChart(node, ctx)
        return el ? [el] : []
      }
      case 'p:grpSp': {
        const el = await renderGroup(node, ctx)
        return el ? [el] : []
      }
      default:
        return []
    }
  } catch (e) {
    // 单元素失败不拖垮整页（与一期 import/index.ts 的页面级告警同风格）
    console.warn('[pptx-preview] 元素渲染失败:', node.nodeName, e)
    addSkipped(ctx.report, 'elementParseFailed')
    return []
  }
}
