// 图片导出：离屏渲染幻灯片 → html-to-image 截图
import { createRoot } from 'react-dom/client'
import { toPng, toJpeg } from 'html-to-image'
import { createElement } from 'react'
import JSZip from 'jszip'
import type { Presentation, Slide } from '../../types/slides'
import { SlideRenderer } from '../../components/canvas/SlideRenderer'
import { downloadBlob } from './json'

/** 离屏渲染单页并导出 PNG/JPEG */
export async function renderSlideToBlob(
  slide: Slide,
  presentation: Presentation,
  format: 'png' | 'jpeg' = 'png',
): Promise<Blob> {
  const w = presentation.width
  const h = Math.round(presentation.width / presentation.viewportRatio)

  const host = document.createElement('div')
  host.style.cssText = `position:fixed;left:-99999px;top:0;width:${w}px;height:${h}px;z-index:-1;`
  document.body.appendChild(host)
  const root = createRoot(host)
  root.render(createElement(SlideRenderer, { slide, width: w, height: h, staticMode: true }))

  // 等待渲染与图片解码
  await new Promise((resolve) => setTimeout(resolve, 120))
  const images = host.querySelectorAll('img')
  await Promise.all(Array.from(images).map((img) => img.complete
    ? Promise.resolve()
    : new Promise((resolve) => { img.onload = resolve; img.onerror = resolve })))

  const capture = format === 'png' ? toPng : toJpeg
  const dataUrl = await capture(host, { width: w, height: h, pixelRatio: 2, backgroundColor: '#ffffff' })
  root.unmount()
  host.remove()

  const res = await fetch(dataUrl)
  return res.blob()
}

/** 导出单页或全部页面（打包 zip 时动态引入 JSZip） */
export async function exportImages(
  slides: Slide[],
  presentation: Presentation,
  format: 'png' | 'jpeg',
  docName: string,
  onProgress?: (done: number, total: number) => void,
): Promise<void> {
  if (slides.length === 1) {
    const blob = await renderSlideToBlob(slides[0], presentation, format)
    downloadBlob(blob, `${docName || 'slide'}.${format === 'png' ? 'png' : 'jpg'}`)
    return
  }
  const zip = new JSZip()
  for (let i = 0; i < slides.length; i++) {
    const blob = await renderSlideToBlob(slides[i], presentation, format)
    zip.file(`slide-${String(i + 1).padStart(2, '0')}.${format === 'png' ? 'png' : 'jpg'}`, blob)
    onProgress?.(i + 1, slides.length)
  }
  const out = await zip.generateAsync({ type: 'blob' })
  downloadBlob(out, `${docName || 'slides'}-images.zip`)
}
