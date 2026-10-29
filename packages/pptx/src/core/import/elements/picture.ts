// src/core/import/elements/picture.ts
/** p:pic → image / video / audio */
import { attr, directChild, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { getShapeKey } from '../geometry'
import { parseShadow } from './shape'
import { frameRotOf, mapEmuBox, parseFrameGeom } from './frame-geom'
import { genId } from '../../utils/id'
import {
  addSkipped, fallbackPlaceholderBox, inheritedPlaceholderShape, parseElementLink,
  resolvePlaceholderGeometry,
} from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { AudioElement, ImageElement, PPTElement, VideoElement } from '../../../types/slides'

/** 浏览器可直接渲染的媒体扩展名；EMF/WMF/JXR/WDP 等需矢量/私有解码，按不支持上报 */
const RENDERABLE_IMAGE = new Set(['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp', 'bmp', 'tiff', 'tif', 'avif', 'ico', 'apng'])
const RENDERABLE_VIDEO = new Set(['mp4', 'webm', 'mov', 'm4v', 'ogv'])
const RENDERABLE_AUDIO = new Set(['mp3', 'wav', 'm4a', 'wma', 'flac', 'ogg', 'oga', 'opus', 'aac'])

function extOf(path: string): string {
  return path.split('.').pop()?.toLowerCase() ?? ''
}

/** a:blip 内 SVG 矢量替代（asvg:svgBlip，PowerPoint 同时存 PNG 后备）：优先矢量避免放大模糊 */
function svgBlipOf(blip: Element | null): Element | null {
  if (!blip) return null
  return firstDescendant(blip, 'asvg:svgBlip') ?? firstDescendant(blip, 'svgBlip')
}

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
  const ph = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:ph')
  // 图片占位符位置继承：自身无 xfrm 时按版式/母版原型定位（缺失时回退类型默认位置）
  let geom = parseFrameGeom(node, xf, ctx)
  if (!geom && ph) {
    let pos = resolvePlaceholderGeometry(ctx, ph)
    if (!pos && ctx.slideSize) {
      addSkipped(ctx.report, 'missingPlaceholder')
      pos = fallbackPlaceholderBox(ph, ctx.slideSize)
    }
    if (!pos) return null
    geom = mapEmuBox(xf, pos, ctx.scale)
  }
  if (!geom) return null

  // 海报帧：blipFill/a:blip（SVG 矢量替代优先）→ 媒体数据 URL
  const blipFill = firstDescendant(node, 'p:blipFill') ?? firstDescendant(node, 'a:blipFill')
  const blip = blipFill ? directChild(blipFill, 'a:blip') : null
  const svgBlip = svgBlipOf(blip)
  const embedId = attr(svgBlip, 'r:embed') ?? attr(blip, 'r:embed')
  const posterPath = embedId ? await pkg.relTarget(ctx.partPath, embedId) : null
  let poster = posterPath ? await pkg.mediaDataUrl(posterPath) : undefined
  // 非位图/矢量图（如内嵌 PDF/OLE 图标）不可渲染：报 unsupportedMedia，避免再重复报 missingImage
  let posterUnsupported = false
  if (poster && posterPath && !RENDERABLE_IMAGE.has(extOf(posterPath))) {
    addSkipped(ctx.report, 'unsupportedMedia')
    poster = undefined
    posterUnsupported = true
  }

  // 音视频：nvPr 下的 a:videoFile / a:audioFile（源 rel 可挂在 r:link，也可挂在 p14:media@r:embed）
  const nvPr = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:nvPr')
  const videoFile = nvPr ? directChild(nvPr, 'a:videoFile') : null
  const audioFile = nvPr ? directChild(nvPr, 'a:audioFile') : null
  const p14media = firstDescendant(node, 'p14:media')
  const mediaRelId = attr(videoFile, 'r:link') || attr(audioFile, 'r:link') || attr(p14media, 'r:embed')
  const mediaPath = mediaRelId ? await pkg.relTarget(ctx.partPath, mediaRelId) : null
  let mediaSrc: string | undefined
  if (mediaPath) {
    const ext = extOf(mediaPath)
    const allowed = videoFile ? RENDERABLE_VIDEO : RENDERABLE_AUDIO
    if (!allowed.has(ext)) addSkipped(ctx.report, 'unsupportedMedia')
    else mediaSrc = await pkg.mediaDataUrl(mediaPath)
  }

  const cNvPr = firstDescendant(node, 'p:cNvPr')
  const link = await parseElementLink(cNvPr, ctx)
  const altText = attr(cNvPr, 'descr')?.trim() || undefined

  if ((videoFile || audioFile) && mediaSrc) {
    const base = { ...geom, loop: false, autoPlay: false }
    if (videoFile) {
      const video: VideoElement = { id: genId('v-'), type: 'video', src: mediaSrc, ...base, name: altText ?? '视频' }
      if (poster) video.poster = poster
      if (link) video.link = link
      return video
    }
    const audio: AudioElement = { id: genId('a-'), type: 'audio', src: mediaSrc, ...base, name: altText ?? '音频' }
    if (link) audio.link = link
    return audio
  }

  if (!poster) {
    if (!posterUnsupported) addSkipped(ctx.report, 'missingImage')
    return null
  }

  const spPr = firstDescendant(node, 'p:spPr') ?? firstDescendant(node, 'a:spPr')
  const phShape = ph ? inheritedPlaceholderShape(ctx, ph) : undefined
  const protoSpPr = phShape?.spPr ?? null
  // 属性组继承（同 shape.ts）：自身声明即用自身，否则回退版式/母版原型
  const groupEl = (parent: Element | null, ...names: string[]): Element | null => {
    if (!parent) return null
    for (const name of names) {
      const el = directChild(parent, name)
      if (el) return el
    }
    return null
  }
  const geomEl = groupEl(spPr, 'a:prstGeom', 'a:custGeom') ?? groupEl(protoSpPr, 'a:prstGeom', 'a:custGeom')
  const xfrm = firstDescendant(node, 'a:xfrm')
  const { rot, flipH, flipV } = frameRotOf(xfrm, xf)
  const ln = groupEl(spPr, 'a:ln') ?? groupEl(protoSpPr, 'a:ln')
  const effectLst = groupEl(spPr, 'a:effectLst') ?? groupEl(protoSpPr, 'a:effectLst')
  const image: ImageElement = {
    id: genId('i-'), type: 'image', src: poster, ...geom,
    name: altText ?? '图片',
    flipH: flipH || undefined,
    flipV: flipV || undefined,
  }
  if (link) image.link = link
  if (rot) image.rotate = Math.round(rot)
  const clip = parseSrcRect(blipFill ? directChild(blipFill, 'a:srcRect') : null)
  if (clip) image.clip = clip
  // 图片透明度（a:blip > a:alphaModFix amt，十万分之一）
  const alphaFix = blip ? directChild(blip, 'a:alphaModFix') : null
  if (alphaFix) {
    const amt = parseInt(attr(alphaFix, 'amt') ?? '100000', 10) / 100000
    if (Number.isFinite(amt) && amt < 1) image.opacity = Math.max(0, amt)
  }
  // 形状裁剪：prstGeom（圆角矩形映射为 radius，其余映射为 clipShape）；图片模型无自定义路径裁剪
  const prst = geomEl?.localName === 'prstGeom' ? attr(geomEl, 'prst') ?? '' : ''
  if (geomEl?.localName === 'custGeom') addSkipped(ctx.report, 'imageClipGeom')
  else if (prst === 'roundRect') {
    const avLst = geomEl ? directChild(geomEl, 'a:avLst') : null
    const gd = avLst ? firstDescendant(avLst, 'a:gd') : null
    const fmla = attr(gd, 'fmla') ?? ''
    const val = parseInt(fmla.replace(/^val\s+/, ''), 10)
    const adj = Number.isFinite(val) ? val / 100000 : 0.16667
    image.radius = Math.round(Math.min(geom.w, geom.h) * adj)
  } else if (prst && prst !== 'rect') {
    image.clipShape = getShapeKey(prst)
  }
  const shadow = parseShadow(effectLst, ctx)
  if (shadow) image.shadow = shadow
  if (ln) {
    const color = resolveColor(directChild(ln, 'a:solidFill'), ctx.theme)
    const width = Math.max(1, Math.round(parseInt(attr(ln, 'w') ?? '0', 10) / 9525))
    if (color) image.outline = { color, width, style: 'solid' }
  }
  return image
}
