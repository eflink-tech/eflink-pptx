/** a:txBody → foreignObject：复用 txBodyToHTML（已 XSS 加固：escapeHTML + 链接协议白名单）。
 * 不做 normAutofit fontScale 缩放（已知取舍），ins/anchor/wrap/vert 按 bodyPr 直渲。 */
import { attr, directChild } from '../import/xml'
import { txBodyToHTML } from '../import/elements/text'
import { emu2pxF, svgEl, type PreviewCtx, type Box } from './svg'

export async function renderText(txBody: Element, ctx: PreviewCtx, box: Box): Promise<SVGElement | null> {
  const body = await txBodyToHTML(txBody, ctx.theme, ctx.pkg, ctx.partPath)
  const bodyPr = directChild(txBody, 'a:bodyPr')
  // 内边距 EMU 默认值：lIns/rIns 91440、tIns/bIns 45720
  const lIns = emu2pxF(parseInt(attr(bodyPr, 'lIns') ?? '91440', 10))
  const tIns = emu2pxF(parseInt(attr(bodyPr, 'tIns') ?? '45720', 10))
  const rIns = emu2pxF(parseInt(attr(bodyPr, 'rIns') ?? '91440', 10))
  const bIns = emu2pxF(parseInt(attr(bodyPr, 'bIns') ?? '45720', 10))
  const anchor = attr(bodyPr, 'anchor')
  const justify = anchor === 't' ? 'flex-start' : anchor === 'b' ? 'flex-end' : anchor ? 'center' : undefined
  const vert = attr(bodyPr, 'vert')

  const fo = svgEl('foreignObject', { x: box.x, y: box.y, width: box.w, height: box.h })
  const div = document.createElement('div')
  div.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml')
  div.setAttribute('style', [
    'width:100%', 'height:100%', 'box-sizing:border-box', 'overflow:hidden',
    `padding:${tIns}px ${rIns}px ${bIns}px ${lIns}px`,
    justify ? `display:flex;flex-direction:column;justify-content:${justify}` : '',
    vert === 'vert' || vert === 'eaVert' || vert === 'mongolianVert' ? 'writing-mode:vertical-rl' : '',
    attr(bodyPr, 'wrap') === 'none' ? 'white-space:nowrap' : '',
  ].filter(Boolean).join(';'))
  div.innerHTML = body.html
  fo.appendChild(div)
  return fo
}
