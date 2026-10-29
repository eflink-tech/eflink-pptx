/** spTree 单节点分发器：p:sp/p:cxnSp/p:pic/p:graphicFrame/p:grpSp → 对应渲染器；未知节点跳过。
 * 独立成文件打破 slide → group ↔ dispatch 的模块循环（函数级互调在 ESM 下经由 live binding 正常工作）。 */
import { alternateContentChildren, alternateContentRequires, firstDescendant, graphicDataKind } from '../import/xml'
import { addSkipped } from '../import/context'
import { olePreviewPic } from '../import/ole'
import type { PreviewCtx } from './svg'
import { renderShape } from './shape'
import { renderPicture } from './picture'
import { renderTable } from './table'
import { renderChart } from './chart'
import { renderGroup } from './group'
import { renderSmartArt } from './smartart'

/** OLE：有预览图则按图片渲染，否则计入报告 */
async function renderOle(container: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const pic = olePreviewPic(container)
  const el = pic ? await renderPicture(pic, ctx) : null
  if (el) {
    addSkipped(ctx.report, 'olePreviewOnly')
    return el
  }
  addSkipped(ctx.report, 'oleObjectUnsupported')
  return null
}

export async function renderSpTreeNode(node: Element, ctx: PreviewCtx): Promise<SVGElement[]> {
  try {
    switch (node.localName) {
      case 'sp':
      case 'cxnSp': {
        const el = await renderShape(node, ctx)
        return el ? [el] : []
      }
      case 'pic': {
        const el = await renderPicture(node, ctx)
        return el ? [el] : []
      }
      case 'graphicFrame': {
        // 表格（a:tbl）与图表（c:chart）共用 graphicFrame 容器；SmartArt 按缓存绘制，OLE 按预览图
        const isTable = Boolean(firstDescendant(node, 'a:tbl'))
        const isChart = Boolean(firstDescendant(node, 'c:chart'))
        if (!isTable && !isChart) {
          const kind = graphicDataKind(node)
          if (kind === 'diagram') {
            const el = await renderSmartArt(node, ctx)
            addSkipped(ctx.report, el ? 'smartartDrawing' : 'smartartUnsupported')
            return el ? [el] : []
          }
          if (kind === 'ole') {
            const el = await renderOle(node, ctx)
            return el ? [el] : []
          }
          addSkipped(ctx.report, 'graphicFrameUnknown')
          return []
        }
        const el = isTable ? await renderTable(node, ctx) : await renderChart(node, ctx)
        return el ? [el] : []
      }
      case 'grpSp': {
        const el = await renderGroup(node, ctx)
        return el ? [el] : []
      }
      case 'AlternateContent': {
        const requires = alternateContentRequires(node)
        const out: SVGElement[] = []
        for (const child of alternateContentChildren(node)) {
          out.push(...(await renderSpTreeNode(child, ctx)))
        }
        if (!out.length) addSkipped(ctx.report, 'alternateContentUnsupported')
        else if (requires.includes('am3d')) addSkipped(ctx.report, 'model3dFallback')
        else if (requires.includes('dgm')) addSkipped(ctx.report, 'smartartFallback')
        return out
      }
      case 'oleObj': {
        const el = await renderOle(node, ctx)
        return el ? [el] : []
      }
      case 'contentPart':
        addSkipped(ctx.report, 'contentPartUnsupported')
        return []
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
