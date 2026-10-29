// src/core/import/elements/index.ts
/** spTree 子节点分发：p:sp / p:cxnSp / p:pic / p:graphicFrame / p:grpSp / mc:AlternateContent */
import { addSkipped, composeXform } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { PPTElement } from '../../../types/slides'
import { alternateContentChildren, alternateContentRequires, graphicDataKind, isPresentationml } from '../xml'
import { diagramBackgroundColor, diagramFrameXform, diagramIndex, drawingSpTree, findDiagramDrawing } from '../smartart'
import { olePreviewPic } from '../ole'
import { parseFrameGeom } from './frame-geom'
import { genId } from '../../utils/id'
import { parseShapeEl } from './shape'
import { parsePictureEl } from './picture'
import { parseTableEl } from './table'
import { parseChartEl } from './chart'
import { parseGroupEl } from './group'

/** spTree 结构节点（不属于页面内容，命中默认分支不报告）；按 localName 匹配以兼容 strict OOXML/非常规前缀 */
const STRUCTURAL_LOCALS = ['nvGrpSpPr', 'grpSpPr', 'extLst']

/** SmartArt：按幻灯片缓存的 dsp:drawing 静态绘制展开（绘制坐标为框架本地 EMU，整体叠加框架仿射） */
async function parseSmartArt(
  frame: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  const idx = diagramIndex(frame)
  const drawing = await findDiagramDrawing(pkg, ctx.partPath, idx)
  const tree = drawing ? drawingSpTree(drawing.doc) : null
  if (!drawing || !tree) return []
  const inner = composeXform(xf, diagramFrameXform(frame))
  // 绘制部件有自己的 rels（内嵌图片等），切换解析基准
  const dctx: ParseContext = { ...ctx, partPath: drawing.path }
  const out: PPTElement[] = []
  // SmartArt 背景色不在缓存绘制里（dgm:bg 只存在于数据部件），先铺一层矩形
  const bg = await diagramBackgroundColor(pkg, ctx.partPath, idx, ctx.theme)
  const frameBox = parseFrameGeom(frame, xf, ctx)
  if (bg && frameBox) {
    out.push({ id: genId('s-'), type: 'shape', ...frameBox, shapeKey: 'rect', fill: bg, name: 'SmartArt' })
  }
  for (const child of Array.from(tree.children)) {
    out.push(...(await parseSpTreeNode(child, dctx, inner, pkg)))
  }
  return out
}

/** OLE 对象：有预览图则按图片导入，否则计入报告 */
async function parseOle(
  container: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  const pic = olePreviewPic(container)
  const el = pic ? await parsePictureEl(pic, ctx, xf, pkg) : null
  if (el) {
    addSkipped(ctx.report, 'olePreviewOnly')
    return [el]
  }
  addSkipped(ctx.report, 'oleObjectUnsupported')
  return []
}

/** AlternateContent 降级报告：按 Requires 区分 3D/ SmartArt / OLE */
function reportAlternateContent(requires: string, ctx: ParseContext, rendered: boolean): void {
  if (!rendered) {
    addSkipped(ctx.report, 'alternateContentUnsupported')
  } else if (requires.includes('am3d')) {
    addSkipped(ctx.report, 'model3dFallback')
  } else if (requires.includes('dgm')) {
    addSkipped(ctx.report, 'smartartFallback')
  }
}

export async function parseSpTreeNode(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  try {
    switch (node.localName) {
      case 'sp':
      case 'cxnSp': {
        const el = await parseShapeEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'pic': {
        const el = await parsePictureEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'graphicFrame': {
        // graphicFrame 内既有 a:tbl 也可能挂 chart rels：表格优先，找不到 a:tbl 落到图表
        const el = (await parseTableEl(node, ctx, xf, pkg)) ?? (await parseChartEl(node, ctx, xf, pkg))
        if (el) return [el]
        // SmartArt 优先取缓存的静态绘制；OLE 取预览图。其余未知容器如实报告
        const kind = graphicDataKind(node)
        if (kind === 'diagram') {
          const els = await parseSmartArt(node, ctx, xf, pkg)
          if (els.length) {
            addSkipped(ctx.report, 'smartartDrawing')
            return els
          }
          addSkipped(ctx.report, 'smartartUnsupported')
          return []
        }
        if (kind === 'ole') return await parseOle(node, ctx, xf, pkg)
        addSkipped(ctx.report, 'graphicFrameUnknown')
        return []
      }
      case 'grpSp':
        return await parseGroupEl(node, ctx, xf, pkg)
      case 'AlternateContent': {
        // 优先 Fallback（3D/SmartArt/OLE 的平面位图），逐个解析；子节点可再嵌套 AlternateContent
        const requires = alternateContentRequires(node)
        const out: PPTElement[] = []
        for (const child of alternateContentChildren(node)) {
          out.push(...(await parseSpTreeNode(child, ctx, xf, pkg)))
        }
        reportAlternateContent(requires, ctx, out.length > 0)
        return out
      }
      case 'oleObj':
        return await parseOle(node, ctx, xf, pkg)
      case 'contentPart':
        addSkipped(ctx.report, 'contentPartUnsupported')
        return []
      default:
        if (!STRUCTURAL_LOCALS.includes(node.localName) && (isPresentationml(node) || node.nodeName.startsWith('p:'))) {
          addSkipped(ctx.report, 'unknownElement')
        }
        return []
    }
  } catch {
    addSkipped(ctx.report, 'elementParseFailed')
    return []
  }
}
