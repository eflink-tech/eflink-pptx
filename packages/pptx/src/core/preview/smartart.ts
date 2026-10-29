/** SmartArt 静态回退渲染：读取幻灯片缓存的 dsp:drawing，按框架仿射渲染到 <g>。
 * 坐标与导入一致（绘制部件为框架本地 EMU），整体由 matrix 映射到页面 px 空间。 */
import { diagramBackgroundColor, diagramFrameXform, diagramIndex, drawingSpTree, findDiagramDrawing } from '../import/smartart'
import { emu2pxF, svgEl, geomOf, type PreviewCtx } from './svg'
import { renderSpTreeNode } from './dispatch'

export async function renderSmartArt(frame: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const idx = diagramIndex(frame)
  const drawing = await findDiagramDrawing(ctx.pkg, ctx.partPath, idx)
  const tree = drawing ? drawingSpTree(drawing.doc) : null
  if (!drawing || !tree) return null
  const m = diagramFrameXform(frame)
  const g = svgEl('g', {
    transform: `matrix(${m.a},${m.b},${m.c},${m.d},${emu2pxF(m.e)},${emu2pxF(m.f)})`,
  })
  // SmartArt 背景（dgm:bg）；绘制部件本地坐标原点即框架左上角
  const bg = await diagramBackgroundColor(ctx.pkg, ctx.partPath, idx, ctx.theme)
  const box = geomOf(frame, ctx)
  if (bg && box) g.appendChild(svgEl('rect', { x: 0, y: 0, width: box.w, height: box.h, fill: bg }))
  const dctx: PreviewCtx = { ...ctx, partPath: drawing.path }
  for (const child of Array.from(tree.children)) {
    g.append(...(await renderSpTreeNode(child, dctx)))
  }
  return g.children.length ? g : null
}
