// JSON 导入导出
import type { Presentation } from '../../types/slides'

export function exportJSON(presentation: Presentation, name: string): void {
  const blob = new Blob([JSON.stringify(presentation, null, 2)], { type: 'application/json' })
  downloadBlob(blob, `${name || '未命名'}.efppt.json`)
}

export function parseJSONFile(text: string): Presentation {
  const data = JSON.parse(text) as Presentation
  if (!data || !Array.isArray(data.slides)) throw new Error('JSON 结构不正确：缺少 slides')
  return data
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}
