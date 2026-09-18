// src/core/import/styles.ts
/** 颜色求值：OOXML 颜色节点 → 内部 hex（含 alpha 8 位形式） */
import { resolveSchemeColor } from './theme'
import type { PptxTheme } from './theme'

const BYTE = 255

function clampByte(n: number): number {
  return Math.min(BYTE, Math.max(0, Math.round(n)))
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((v) => clampByte(v).toString(16).padStart(2, '0').toUpperCase()).join('')}`
}

function hexToRgb(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ]
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / BYTE, gn = g / BYTE, bn = b / BYTE
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6
  else if (max === gn) h = ((bn - rn) / d + 2) / 6
  else h = ((rn - gn) / d + 4) / 6
  return [h, s, l]
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * BYTE, l * BYTE, l * BYTE]
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const hue = (t: number): number => {
    let tt = t
    if (tt < 0) tt += 1
    if (tt > 1) tt -= 1
    if (tt < 1 / 6) return p + (q - p) * 6 * tt
    if (tt < 1 / 2) return q
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6
    return p
  }
  return [hue(h + 1 / 3) * BYTE, hue(h) * BYTE, hue(h - 1 / 3) * BYTE]
}

/** lumMod/lumOff：HSL 亮度缩放/偏移（mod 与 off 以 1 为满值） */
function applyLum(r: number, g: number, b: number, mod: number, off: number): [number, number, number] {
  const [h, s, l] = rgbToHsl(r, g, b)
  return hslToRgb(h, s, Math.min(1, Math.max(0, l * mod + off)))
}

/** 单个颜色节点（a:srgbClr / a:sysClr / a:schemeClr / a:prstClr）求值 */
export function resolveColorOf(clrEl: Element, theme: PptxTheme): string | undefined {
  const name = clrEl.nodeName
  let hex: string | undefined
  if (name === 'a:srgbClr') hex = `#${(clrEl.getAttribute('val') ?? '000000').toUpperCase()}`
  else if (name === 'a:sysClr') hex = `#${(clrEl.getAttribute('lastClr') ?? '000000').toUpperCase()}`
  else if (name === 'a:schemeClr') hex = resolveSchemeColor(clrEl.getAttribute('val') ?? 'tx1', theme)
  else if (name === 'a:prstClr') hex = `#${(clrEl.getAttribute('lastClr') ?? '000000').toUpperCase()}`
  else return undefined
  if (!hex) return undefined
  // 非法 hex（如 'GGHHII'）直接判无效，避免 NaN 传播
  if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return undefined

  let [r, g, b] = hexToRgb(hex)
  let alpha: number | null = null
  let mod = 1
  let off = 0
  for (const modEl of Array.from(clrEl.children)) {
    const val = parseInt(modEl.getAttribute('val') ?? '0', 10) / 100000
    switch (modEl.nodeName) {
      case 'a:alpha': alpha = val; break
      case 'a:lumMod': mod = val; break
      case 'a:lumOff': off = val; break
      case 'a:shade': r *= val; g *= val; b *= val; break
      case 'a:tint': r += (BYTE - r) * (1 - val); g += (BYTE - g) * (1 - val); b += (BYTE - b) * (1 - val); break
      default: break
    }
  }
  // 修饰符文档顺序不保证：先应用亮度变换，最后拼接 alpha
  if (mod !== 1 || off !== 0) [r, g, b] = applyLum(r, g, b, mod, off)
  const rgb = rgbToHex(r, g, b)
  if (alpha === null) return rgb
  const a = Math.round(alpha * BYTE).toString(16).padStart(2, '0').toUpperCase()
  return rgb + a
}

/** 容器节点（a:solidFill / a:ln / a:bgPr 等）取第一个颜色子节点求值 */
export function resolveColor(container: Element | null, theme: PptxTheme): string | undefined {
  if (!container) return undefined
  for (const child of Array.from(container.children)) {
    const c = resolveColorOf(child, theme)
    if (c) return c
  }
  return undefined
}
