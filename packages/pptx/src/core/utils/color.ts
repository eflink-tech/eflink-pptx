/** 颜色工具：hex/rgba 转换与渐变序列化 */

/** '#rrggbb' -> {r,g,b} */
export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  let h = hex.replace('#', '')
  if (h.length === 3) h = h.split('').map((c) => c + c).join('')
  const n = parseInt(h, 16)
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 }
}

/** 透明度混合到白色背景，返回 hex（导出 pptx 无 alpha 时用） */
export function blendToWhite(hex: string, alpha: number): string {
  const { r, g, b } = hexToRgb(hex)
  const mix = (c: number) => Math.round(c * alpha + 255 * (1 - alpha))
  const to2 = (n: number) => n.toString(16).padStart(2, '0')
  return `#${to2(mix(r))}${to2(mix(g))}${to2(mix(b))}`
}

/** '#rrggbb' -> 'r,g,b'（pptxgenjs 透明度方案用） */
export function hexToRgbStr(hex: string): string {
  const { r, g, b } = hexToRgb(hex)
  return `${r},${g},${b}`
}

/** 是否为合法 hex 颜色 */
export function isHexColor(v: string): boolean {
  return /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(v)
}
