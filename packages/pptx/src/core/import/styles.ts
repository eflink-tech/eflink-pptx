// src/core/import/styles.ts
/** 颜色求值：OOXML 颜色节点 → 内部 hex（含 alpha 8 位形式） */
import { resolveSchemeColor } from './theme'
import type { PptxTheme } from './theme'

const BYTE = 255

/** ST_PresetColorVal 全量（CSS/X11 命名色，含 dk/lt/med 短别名）。
 *  紧凑 'name:HEX' 串只为源码体积；未知值仍按无效处理（避免错色）。 */
const PRESET_LIST = [
  'aliceBlue:F0F8FF antiqueWhite:FAEBD7 aqua:00FFFF aquamarine:7FFFD4 azure:F0FFFF beige:F5F5DC',
  'bisque:FFE4C4 black:000000 blanchedAlmond:FFEBCD blue:0000FF blueViolet:8A2BE2 brown:A52A2A',
  'burlyWood:DEB887 cadetBlue:5F9EA0 chartreuse:7FFF00 chocolate:D2691E coral:FF7F50',
  'cornflowerBlue:6495ED cornsilk:FFF8DC crimson:DC143C cyan:00FFFF darkBlue:00008B darkCyan:008B8B',
  'darkGoldenrod:B8860B darkGray:A9A9A9 darkGreen:006400 darkGrey:A9A9A9 darkKhaki:BDB76B',
  'darkMagenta:8B008B darkOliveGreen:556B2F darkOrange:FF8C00 darkOrchid:9932CC darkRed:8B0000',
  'darkSalmon:E9967A darkSeaGreen:8FBC8F darkSlateBlue:483D8B darkSlateGray:2F4F4F darkSlateGrey:2F4F4F',
  'darkTurquoise:00CED1 darkViolet:9400D3 deepPink:FF1493 deepSkyBlue:00BFFF dimGray:696969 dimGrey:696969',
  'dkBlue:00008B dkCyan:008B8B dkGoldenrod:B8860B dkGray:A9A9A9 dkGreen:006400 dkGrey:A9A9A9 dkKhaki:BDB76B',
  'dkMagenta:8B008B dkOliveGreen:556B2F dkOrange:FF8C00 dkOrchid:9932CC dkRed:8B0000 dkSalmon:E9967A',
  'dkSeaGreen:8FBC8F dkSlateBlue:483D8B dkSlateGray:2F4F4F dkSlateGrey:2F4F4F dkTurquoise:00CED1 dkViolet:9400D3',
  'dodgerBlue:1E90FF firebrick:B22222 floralWhite:FFFAF0 forestGreen:228B22 fuchsia:FF00FF gainsboro:DCDCDC',
  'ghostWhite:F8F8FF gold:FFD700 goldenrod:DAA520 gray:808080 green:008000 greenYellow:ADFF2F grey:808080',
  'honeydew:F0FFF0 hotPink:FF69B4 indianRed:CD5C5C indigo:4B0082 ivory:FFFFF0 khaki:F0E68C lavender:E6E6FA',
  'lavenderBlush:FFF0F5 lawnGreen:7CFC00 lemonChiffon:FFFACD lightBlue:ADD8E6 lightCoral:F08080 lightCyan:E0FFFF',
  'lightGoldenrodYellow:FAFAD2 lightGray:D3D3D3 lightGreen:90EE90 lightGrey:D3D3D3 lightPink:FFB6C1',
  'lightSalmon:FFA07A lightSeaGreen:20B2AA lightSkyBlue:87CEFA lightSlateGray:778899 lightSlateGrey:778899',
  'lightSteelBlue:B0C4DE lightYellow:FFFFE0 lime:00FF00 limeGreen:32CD32 linen:FAF0E6 ltBlue:ADD8E6 ltCoral:F08080',
  'ltCyan:E0FFFF ltGoldenrodYellow:FAFAD2 ltGray:D3D3D3 ltGreen:90EE90 ltGrey:D3D3D3 ltPink:FFB6C1 ltSalmon:FFA07A',
  'ltSeaGreen:20B2AA ltSkyBlue:87CEFA ltSlateGray:778899 ltSlateGrey:778899 ltSteelBlue:B0C4DE ltYellow:FFFFE0',
  'magenta:FF00FF maroon:800000 medAquamarine:66CDAA medBlue:0000CD medOrchid:BA55D3 medPurple:9370DB',
  'medSeaGreen:3CB371 medSlateBlue:7B68EE medSpringGreen:00FA9A medTurquoise:48D1CC medVioletRed:C71585',
  'mediumAquamarine:66CDAA mediumBlue:0000CD mediumOrchid:BA55D3 mediumPurple:9370DB mediumSeaGreen:3CB371',
  'mediumSlateBlue:7B68EE mediumSpringGreen:00FA9A mediumTurquoise:48D1CC mediumVioletRed:C71585 midnightBlue:191970',
  'mintCream:F5FFFA mistyRose:FFE4E1 moccasin:FFE4B5 navajoWhite:FFDEAD navy:000080 oldLace:FDF5E6 olive:808000',
  'oliveDrab:6B8E23 orange:FFA500 orangeRed:FF4500 orchid:DA70D6 paleGoldenrod:EEE8AA paleGreen:98FB98',
  'paleTurquoise:AFEEEE paleVioletRed:DB7093 papayaWhip:FFEFD5 peachPuff:FFDAB9 peru:CD853F pink:FFC0CB',
  'plum:DDA0DD powderBlue:B0E0E6 purple:800080 red:FF0000 rosyBrown:BC8F8F royalBlue:4169E1 saddleBrown:8B4513',
  'salmon:FA8072 sandyBrown:F4A460 seaGreen:2E8B57 seashell:FFF5EE sienna:A0522D silver:C0C0C0 skyBlue:87CEEB',
  'slateBlue:6A5ACD slateGray:708090 slateGrey:708090 snow:FFFAFA springGreen:00FF7F steelBlue:4682B4 tan:D2B48C',
  'teal:008080 thistle:D8BFD8 tomato:FF6347 turquoise:40E0D0 violet:EE82EE violetRed:D02090 wheat:F5DEB3',
  'white:FFFFFF whiteSmoke:F5F5F5 yellow:FFFF00 yellowGreen:9ACD32',
].join(' ')

const PRESET_COLORS: Record<string, string> = Object.fromEntries(
  PRESET_LIST.split(' ').map((pair) => {
    const [name, hex] = pair.split(':')
    return [name, `#${hex}`]
  }),
)

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

/** 基础色解析（不含修饰符）：srgb/sys/prst/scheme/scrgb/hsl；phClr 用引用处颜色替换 */
function baseHexOf(clrEl: Element, theme: PptxTheme, phColor: Element | null | undefined): string | undefined {
  const name = clrEl.localName
  if (name === 'srgbClr') return `#${(clrEl.getAttribute('val') ?? '000000').toUpperCase()}`
  if (name === 'sysClr') return `#${(clrEl.getAttribute('lastClr') ?? '000000').toUpperCase()}`
  if (name === 'prstClr') return PRESET_COLORS[clrEl.getAttribute('val') ?? '']
  if (name === 'schemeClr') {
    const val = clrEl.getAttribute('val') ?? 'tx1'
    if (val === 'phClr') return phColor ? resolveColorOf(phColor, theme) : undefined
    return resolveSchemeColor(val, theme)
  }
  if (name === 'scrgbClr') {
    // ST_Percentage 0–100000 → 0–255（线性 RGB）
    const ch = (attrName: string): number => (parseInt(clrEl.getAttribute(attrName) ?? '0', 10) / 100000) * 255
    return rgbToHex(ch('r'), ch('g'), ch('b'))
  }
  if (name === 'hslClr') {
    const h = (parseInt(clrEl.getAttribute('hue') ?? '0', 10) / 60000) / 360
    const s = parseInt(clrEl.getAttribute('sat') ?? '0', 10) / 100000
    const l = parseInt(clrEl.getAttribute('lum') ?? '0', 10) / 100000
    const [r, g, b] = hslToRgb(h, s, l)
    return rgbToHex(r, g, b)
  }
  return undefined
}

/** 单个颜色节点（a:srgbClr / a:sysClr / a:schemeClr / a:prstClr / a:scrgbClr / a:hslClr）求值。
 *  phColor 用于解析 a:schemeClr val="phClr"（主题样式矩阵引用场景）。 */
export function resolveColorOf(clrEl: Element, theme: PptxTheme, phColor?: Element | null): string | undefined {
  const raw = baseHexOf(clrEl, theme, phColor)
  if (!raw) return undefined
  const baseAlpha = /^#[0-9A-Fa-f]{8}$/.test(raw) ? parseInt(raw.slice(7, 9), 16) : null
  // 非法 hex（如 'GGHHII'）直接判无效，避免 NaN 传播
  if (!/^#[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/.test(raw)) return undefined

  let [r, g, b] = hexToRgb(raw)
  let alpha = baseAlpha === null ? null : baseAlpha / BYTE
  // 修饰符按文档顺序依次作用；alpha 最后叠加（顺序无关，与 OOXML 语义一致）
  for (const modEl of Array.from(clrEl.children)) {
    const val = parseInt(modEl.getAttribute('val') ?? '0', 10) / 100000
    switch (modEl.localName) {
      case 'alpha': alpha = val; break
      case 'alphaMod': alpha = (alpha ?? 1) * val; break
      case 'alphaOff': alpha = (alpha ?? 1) + val; break
      case 'lumMod': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb(h, s, Math.min(1, Math.max(0, l * val)))
        break
      }
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb(h, s, Math.min(1, Math.max(0, l + val)))
        break
      }
      case 'satMod': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb(h, Math.min(1, s * val), l)
        break
      }
      case 'satOff': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb(h, Math.min(1, Math.max(0, s + val)), l)
        break
      }
      case 'hueMod': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb((h * val) % 1, s, l)
        break
      }
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb((((h + val) % 1) + 1) % 1, s, l)
        break
      }
      case 'red': r = val * BYTE; break
      case 'redMod': r *= val; break
      case 'redOff': r += val * BYTE; break
      case 'green': g = val * BYTE; break
      case 'greenMod': g *= val; break
      case 'greenOff': g += val * BYTE; break
      case 'blue': b = val * BYTE; break
      case 'blueMod': b *= val; break
      case 'blueOff': b += val * BYTE; break
      case 'shade': r *= val; g *= val; b *= val; break
      case 'tint': r += (BYTE - r) * (1 - val); g += (BYTE - g) * (1 - val); b += (BYTE - b) * (1 - val); break
      case 'comp': {
        const [h, s, l] = rgbToHsl(r, g, b)
        ;[r, g, b] = hslToRgb((h + 0.5) % 1, s, l)
        break
      }
      case 'inv': r = BYTE - r; g = BYTE - g; b = BYTE - b; break
      case 'gray': {
        const y = 0.299 * r + 0.587 * g + 0.114 * b
        r = r + (y - r) * val; g = g + (y - g) * val; b = b + (y - b) * val
        break
      }
      case 'gamma': r = Math.pow(r / BYTE, val) * BYTE; g = Math.pow(g / BYTE, val) * BYTE; b = Math.pow(b / BYTE, val) * BYTE; break
      default: break
    }
  }
  const rgb = rgbToHex(r, g, b)
  if (alpha === null) return rgb
  const a = Math.round(Math.min(1, Math.max(0, alpha)) * BYTE).toString(16).padStart(2, '0').toUpperCase()
  return rgb + a
}

/** 容器节点（a:solidFill / a:ln / a:bgPr 等）取第一个颜色子节点求值 */
export function resolveColor(container: Element | null, theme: PptxTheme, phColor?: Element | null): string | undefined {
  if (!container) return undefined
  for (const child of Array.from(container.children)) {
    const c = resolveColorOf(child, theme, phColor)
    if (c) return c
  }
  return undefined
}

/** 样式矩阵条目：idx 从 1 起（0 = 显式无该项）；越界取最后一个（与 PowerPoint 语义一致） */
export function styleMatrixEntry(
  theme: PptxTheme,
  kind: 'fill' | 'line' | 'effect',
  idx: number,
): Element | null {
  if (!theme.styles || idx < 1) return null
  const list = kind === 'fill' ? theme.styles.fills : kind === 'line' ? theme.styles.lines : theme.styles.effects
  if (!list.length) return null
  return list[Math.min(idx, list.length) - 1] ?? null
}

/** 样式引用（fillRef/lnRef/effectRef）的直接颜色子节点（phClr 的替换源） */
export function refColor(refEl: Element | null): Element | null {
  if (!refEl) return null
  for (const child of Array.from(refEl.children)) {
    if (['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'].includes(child.localName)) return child
  }
  return null
}
