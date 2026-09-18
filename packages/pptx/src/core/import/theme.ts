// src/core/import/theme.ts
/** 主题解析：clrScheme / fontScheme / 母版 clrMap */
import { attr, firstDescendant, directChild, parseXML } from './xml'
import type { PptxPackage } from './package'

export interface PptxTheme {
  /** dk1 lt1 dk2 lt2 accent1-6 hlink folHlink → '#RRGGBB' */
  schemeColors: Record<string, string>
  majorFont: string
  minorFont: string
  /** 母版 p:clrMap：bg1→lt1 等；未指定的键与 key 相同 */
  colorMap: Record<string, string>
}

export const DEFAULT_SCHEME: Record<string, string> = {
  dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6',
  accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000',
  accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72',
}

const CLR_MAP_KEYS = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function readClrNode(clr: Element | null): string | undefined {
  if (!clr) return undefined
  const srgb = firstDescendant(clr, 'a:srgbClr')
  if (srgb) return `#${(attr(srgb, 'val') ?? '000000').toUpperCase()}`
  const sys = firstDescendant(clr, 'a:sysClr')
  if (sys) return `#${(attr(sys, 'lastClr') ?? '000000').toUpperCase()}`
  return undefined
}

function fontOf(fontScheme: Element | null, tag: string): string | undefined {
  const node = fontScheme ? directChild(fontScheme, `a:${tag}`) : null
  const latin = node ? directChild(node, 'a:latin') : null
  const typeface = attr(latin, 'typeface')
  return typeface && !typeface.startsWith('+') ? typeface : undefined
}

/** 解析 master 关联的主题部件 + master 自身 clrMap（master 缺失时返回默认主题） */
export async function parseThemeForMaster(pkg: PptxPackage, masterPath: string | null): Promise<PptxTheme> {
  const colorMap: Record<string, string> = {}
  if (masterPath) {
    const masterXml = await pkg.text(masterPath)
    if (masterXml) {
      const clrMap = firstDescendant(parseXML(masterXml).documentElement, 'p:clrMap')
      for (const k of CLR_MAP_KEYS) {
        const v = attr(clrMap, k)
        if (v) colorMap[k] = v
      }
    }
  }

  const schemeColors: Record<string, string> = { ...DEFAULT_SCHEME }
  let majorFont = 'Calibri'
  let minorFont = 'Calibri'

  if (masterPath) {
    const rels = await pkg.rels(masterPath)
    let themePath: string | null = null
    for (const rel of rels.values()) {
      if (rel.mode !== 'External' && rel.target.includes('/theme/')) {
        themePath = rel.target
        break
      }
    }
    if (themePath) {
      const xml = await pkg.text(themePath)
      if (xml) {
        const root = parseXML(xml).documentElement
        const scheme = firstDescendant(root, 'a:clrScheme')
        if (scheme) {
          for (const child of Array.from(scheme.children)) {
            const name = child.nodeName.replace(/^a:/, '')
            const hex = readClrNode(child)
            if (hex) schemeColors[name] = hex
          }
        }
        const fontScheme = firstDescendant(root, 'a:fontScheme')
        majorFont = fontOf(fontScheme, 'majorFont') ?? majorFont
        minorFont = fontOf(fontScheme, 'minorFont') ?? minorFont
      }
    }
  }

  return { schemeColors, majorFont, minorFont, colorMap }
}

/** schemeClr 名称求值：先过母版 clrMap（bg1→lt1 等），再查主题色表 */
export function resolveSchemeColor(name: string, theme: PptxTheme): string {
  const mapped = theme.colorMap[name] ?? name
  return theme.schemeColors[mapped] ?? theme.schemeColors[name] ?? '#000000'
}
