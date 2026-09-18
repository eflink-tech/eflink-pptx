// 图片导出：离屏渲染幻灯片 → modern-screenshot 截图
// （不用 html-to-image：其 foreignObject 克隆对内嵌 SVG 渲染整体失效，形状兜底图全白）
import { createRoot } from 'react-dom/client'
import { domToPng, domToJpeg } from 'modern-screenshot'
import { createElement } from 'react'
import JSZip from 'jszip'
import type { Presentation, Slide } from '../../types/slides'
import { SlideRenderer } from '../../components/canvas/SlideRenderer'
import { downloadBlob } from './json'

/** 离屏渲染单页并导出 PNG/JPEG；transparent 时背景透明（元素兜底图用——
 * 白底裁剪图会作为不透明矩形盖住下层相邻元素） */
export async function renderSlideToBlob(
  slide: Slide,
  presentation: Presentation,
  format: 'png' | 'jpeg' = 'png',
  opts?: { transparent?: boolean },
): Promise<Blob> {
  const w = presentation.width
  const h = Math.round(presentation.width / presentation.viewportRatio)

  const host = document.createElement('div')
  host.style.cssText = `position:fixed;left:-99999px;top:0;width:${w}px;height:${h}px;z-index:-1;`
  document.body.appendChild(host)
  const root = createRoot(host)
  root.render(createElement(SlideRenderer, {
    slide, width: w, height: h, staticMode: true,
    // 元素兜底渲染：背景透明（SlideRenderer 默认白底会盖住下层相邻元素）
    transparentBg: opts?.transparent,
  }))

  // 等待 React 提交（并发渲染下固定延时不可靠：拍早了得到空白 foreignObject）：
  // 双 rAF 确保至少一帧完成绘制，再等元素实际挂载；随后等待图片解码
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  for (let i = 0; i < 50 && !host.firstElementChild; i++) {
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
  const images = host.querySelectorAll('img')
  await Promise.all(Array.from(images).map((img) => img.complete
    ? Promise.resolve()
    : new Promise((resolve) => { img.onload = resolve; img.onerror = resolve })))

  const capture = format === 'png' ? domToPng : domToJpeg
  const dataUrl = await capture(host, {
    width: w, height: h, scale: 2,
    backgroundColor: opts?.transparent ? undefined : '#ffffff',
    // JPEG 质量 1：默认 0.8 的量化会把纯白背景压成 249（全页约 2% 的灰蒙色偏）
    ...(format === 'jpeg' ? { quality: 1 } : {}),
  })
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
