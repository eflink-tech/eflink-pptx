/** p:pic → <image>：srcRect 裁剪（裁剪子区映射满框 + clipPath 裁回）、阴影、边框；音视频渲染海报帧 + 播放标记。
 * 已知取舍：prstGeom 形状裁剪（圆形头像等）忽略，一律按矩形框渲染。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import { addSkipped } from '../import/context'
import { registerShadowFilter, lnOf } from './shape'
import { svgEl, geomOf, boxTransform, type PreviewCtx } from './svg'

export async function renderPicture(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const box = geomOf(node, ctx)
  if (!box) return null
  const blipFill = firstDescendant(node, 'p:blipFill') ?? firstDescendant(node, 'a:blipFill')
  const blip = blipFill ? directChild(blipFill, 'a:blip') : null
  const embedId = attr(blip, 'r:embed')
  // pkg 异常（测试桩缺失 relTarget、畸形包）一律按图片缺失降级，不中断整页渲染
  let src: string | undefined
  try {
    const target = embedId ? await ctx.pkg.relTarget(ctx.partPath, embedId) : null
    src = target ? await ctx.pkg.mediaDataUrl(target) : undefined
  } catch {
    src = undefined
  }
  if (!src) {
    addSkipped(ctx.report, 'missingImage')
    return null
  }

  // srcRect（十万分之一比例）：裁剪子区 [l, l+w] 映射满框 → 全图按比例放大后 clip 回框
  const srcRect = blipFill ? directChild(blipFill, 'a:srcRect') : null
  let imgX = box.x
  let imgY = box.y
  let imgW = box.w
  let imgH = box.h
  let clipId: string | undefined
  if (srcRect) {
    const l = parseInt(attr(srcRect, 'l') ?? '0', 10) / 100000
    const t = parseInt(attr(srcRect, 't') ?? '0', 10) / 100000
    const r = parseInt(attr(srcRect, 'r') ?? '0', 10) / 100000
    const b = parseInt(attr(srcRect, 'b') ?? '0', 10) / 100000
    if (l || t || r || b) {
      const sw = Math.max(0.00001, 1 - l - r)
      const sh = Math.max(0.00001, 1 - t - b)
      imgW = box.w / sw
      imgH = box.h / sh
      imgX = box.x - l * imgW
      imgY = box.y - t * imgH
      clipId = ctx.uid('clip')
      const clipPath = svgEl('clipPath', { id: clipId })
      clipPath.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h }))
      ctx.defs.appendChild(clipPath)
    }
  }

  const g = svgEl('g', { transform: boxTransform(box) })
  const spPr = firstDescendant(node, 'p:spPr') ?? firstDescendant(node, 'a:spPr')
  const filter = registerShadowFilter(spPr ? directChild(spPr, 'a:effectLst') : null, ctx)
  g.appendChild(svgEl('image', {
    x: imgX, y: imgY, width: imgW, height: imgH, href: src,
    preserveAspectRatio: 'none', filter,
    'clip-path': clipId ? `url(#${clipId})` : undefined,
  }))

  // 边框：仅 a:ln 存在且非 noFill 时绘制（与 shape.ts 线条缺省可见性语义区分：图片缺 ln 即无边框）
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  if (ln && !directChild(ln, 'a:noFill')) {
    g.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h, fill: 'none', ...lnOf(ln, ctx) }))
  }

  // 音视频：海报帧 + 播放标记（预览不播放）
  const nvPr = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:nvPr')
  const isMedia = Boolean(nvPr && (directChild(nvPr, 'a:videoFile') || directChild(nvPr, 'a:audioFile')))
  if (isMedia) {
    const r = Math.min(box.w, box.h) * 0.15
    const cx = box.x + box.w / 2
    const cy = box.y + box.h / 2
    const s = r * 0.7
    g.appendChild(svgEl('circle', { cx, cy, r, fill: '#00000066' }))
    g.appendChild(svgEl('path', {
      d: `M ${cx - s * 0.4} ${cy - s * 0.6} L ${cx + s * 0.7} ${cy} L ${cx - s * 0.4} ${cy + s * 0.6} Z`,
      fill: '#FFFFFF',
    }))
  }
  return g
}
