// 图片文件读取与插入（顶栏 / 插入菜单 / 粘贴 / 拖拽共用）
import { useEditorStore } from '../../store/editorStore'
import { createImageElement } from '../schema/factory'

export function readImageFile(file: File): Promise<{ src: string; width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error('读取图片失败'))
    reader.onload = () => {
      const src = reader.result as string
      const img = new Image()
      img.onload = () => resolve({ src, width: img.naturalWidth, height: img.naturalHeight })
      img.onerror = () => reject(new Error('图片解码失败'))
      img.src = src
    }
    reader.readAsDataURL(file)
  })
}

/** 将图片按画布 60% 上限等比插入画布中央 */
export async function insertImageFile(file: File): Promise<void> {
  const { src, width, height } = await readImageFile(file)
  const store = useEditorStore.getState()
  const slideW = store.presentation.width
  const slideH = Math.round(slideW / store.presentation.viewportRatio)
  const ratio = Math.min(slideW * 0.6 / width, slideH * 0.6 / height, 1)
  const w = Math.round(width * ratio)
  const h = Math.round(height * ratio)
  store.pushHistory()
  store.addElement(createImageElement((slideW - w) / 2, (slideH - h) / 2, src, w, h))
}

/** 判断文件是否为可插入的图片 */
export function isImageFile(file: File): boolean {
  return file.type.startsWith('image/')
}
