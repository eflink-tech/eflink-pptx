// src/core/import/elements/picture.ts
/** p:pic → image / video / audio */
import { attr, directChild, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { parseShadow } from './shape'
import { parseFrameGeom } from './frame-geom'
import { genId } from '../../utils/id'
import { addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { AudioElement, ImageElement, PPTElement, VideoElement } from '../../../types/slides'

/** a:srcRect（十万分之一百分比）→ 内部 clip 比例矩形 */
export function parseSrcRect(srcRect: Element | null): ImageElement['clip'] | undefined {
  if (!srcRect) return undefined
  const l = parseInt(attr(srcRect, 'l') ?? '0', 10) / 100000
  const t = parseInt(attr(srcRect, 't') ?? '0', 10) / 100000
  const r = parseInt(attr(srcRect, 'r') ?? '0', 10) / 100000
  const b = parseInt(attr(srcRect, 'b') ?? '0', 10) / 100000
  if (!l && !t && !r && !b) return undefined
  // 十万分之一精度取整，规避浮点误差（如 1-0.2-0.2 = 0.6000000000000001）
  const round5 = (v: number) => Math.round(v * 100000) / 100000
  return { x: round5(l), y: round5(t), w: round5(Math.max(0, 1 - l - r)), h: round5(Math.max(0, 1 - t - b)) }
}

export async function parsePictureEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement | null> {
  const geom = parseFrameGeom(node, xf, ctx)
  if (!geom) return null

  // 海报帧：blipFill/a:blip@r:embed → 媒体数据 URL
  const blipFill = firstDescendant(node, 'p:blipFill') ?? firstDescendant(node, 'a:blipFill')
  const blip = blipFill ? directChild(blipFill, 'a:blip') : null
  const embedId = attr(blip, 'r:embed')
  const poster = embedId ? await pkg.mediaDataUrl(await pkg.relTarget(ctx.partPath, embedId) ?? '') : undefined

  // 音视频：nvPr 下的 a:videoFile / a:audioFile
  const nvPr = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:nvPr')
  const videoFile = nvPr ? directChild(nvPr, 'a:videoFile') : null
  const audioFile = nvPr ? directChild(nvPr, 'a:audioFile') : null
  const mediaLinkId = attr(videoFile, 'r:link') ?? attr(audioFile, 'r:link')
  const mediaPath = mediaLinkId ? await pkg.relTarget(ctx.partPath, mediaLinkId) : null
  const mediaSrc = mediaPath ? await pkg.mediaDataUrl(mediaPath) : undefined

  if ((videoFile || audioFile) && mediaSrc) {
    const base = { ...geom, loop: false, autoPlay: false }
    if (videoFile) {
      const video: VideoElement = { id: genId('v-'), type: 'video', src: mediaSrc, ...base, name: '视频' }
      if (poster) video.poster = poster
      return video
    }
    const audio: AudioElement = { id: genId('a-'), type: 'audio', src: mediaSrc, ...base, name: '音频' }
    return audio
  }

  if (!poster) {
    addSkipped(ctx.report, 'missingImage')
    return null
  }

  const spPr = firstDescendant(node, 'p:spPr') ?? firstDescendant(node, 'a:spPr')
  const xfrm = firstDescendant(node, 'a:xfrm')
  // 旋转：xfrm rot（1/60000 deg）+ 组合变换 rot，与 shape.ts 同模式
  const rot = (xfrm ? parseInt(attr(xfrm, 'rot') ?? '0', 10) : 0) / 60000 + (xf.rot ? xf.rot / 60000 : 0)
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const image: ImageElement = {
    id: genId('i-'), type: 'image', src: poster, ...geom,
    name: '图片',
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: attr(xfrm, 'flipV') === '1' || undefined,
  }
  if (rot) image.rotate = Math.round(rot)
  const clip = parseSrcRect(blipFill ? directChild(blipFill, 'a:srcRect') : null)
  if (clip) image.clip = clip
  const shadow = parseShadow(spPr ? directChild(spPr, 'a:effectLst') : null, ctx)
  if (shadow) image.shadow = shadow
  if (ln) {
    const color = resolveColor(directChild(ln, 'a:solidFill'), ctx.theme)
    const width = Math.max(1, Math.round(parseInt(attr(ln, 'w') ?? '0', 10) / 9525))
    if (color) image.outline = { color, width, style: 'solid' }
  }
  return image
}
