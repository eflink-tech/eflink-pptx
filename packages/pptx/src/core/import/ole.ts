// src/core/import/ole.ts
/** OLE 对象预览图定位：p:oleObj 内嵌的 p:pic 是 PowerPoint 显示给用户的静态预览帧。
 * 无预览图（纯 p:embed/p:link）才视为无法导入。 */
import { descendants, directChild } from './xml'

/** 容器（graphicFrame 或 oleObj）内的预览图 pic；多个 oleObj 时取首个带 pic 的（mc:Fallback 分支） */
export function olePreviewPic(container: Element): Element | null {
  const objs = container.localName === 'oleObj' ? [container] : descendants(container, 'p:oleObj')
  for (const obj of objs) {
    const pic = directChild(obj, 'p:pic')
    if (pic) return pic
  }
  return null
}
