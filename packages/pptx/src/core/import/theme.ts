// src/core/import/theme.ts
/** 主题解析：clrScheme / fontScheme / 母版 clrMap */
import { attr, firstDescendant, directChild, directChildren, parseXML } from './xml'
import { relTypeIs, type PptxPackage } from './package'

/** theme1.xml 样式矩阵：p:style 的 fillRef/lnRef/effectRef 按 idx（1 起）索引，
 *  矩阵元素内的 phClr 由引用处的颜色替换 */
export interface PptxStyleMatrix {
  fills: Element[]
  lines: Element[]
  effects: Element[]
}

export interface PptxTheme {
  /** dk1 lt1 dk2 lt2 accent1-6 hlink folHlink → '#RRGGBB' */
  schemeColors: Record<string, string>
  majorFont: string
  minorFont: string
  /** 东亚字体（a:fontScheme 的 a:ea），中文模板的实际渲染字体；缺失为 undefined */
  majorEaFont?: string
  minorEaFont?: string
  /** 母版 p:clrMap：bg1→lt1 等；未指定的键与 key 相同 */
  colorMap: Record<string, string>
  /** 主题样式矩阵（fillStyleLst/lnStyleLst/effectStyleLst）；部件缺失时 undefined */
  styles?: PptxStyleMatrix
}

export const DEFAULT_SCHEME: Record<string, string> = {
  dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6',
  accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000',
  accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72',
}

const CLR_MAP_KEYS = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

/** OOXML theme 关系 Type URI */

function readClrNode(clr: Element | null): string | undefined {
  if (!clr) return undefined
  const srgb = firstDescendant(clr, 'a:srgbClr')
  if (srgb) return `#${(attr(srgb, 'val') ?? '000000').toUpperCase()}`
  const sys = firstDescendant(clr, 'a:sysClr')
  if (sys) return `#${(attr(sys, 'lastClr') ?? '000000').toUpperCase()}`
  return undefined
}

function fontOf(fontScheme: Element | null, tag: string, ea: boolean): string | undefined {
  const node = fontScheme ? directChild(fontScheme, `a:${tag}`) : null
  const face = node ? directChild(node, ea ? 'a:ea' : 'a:latin') : null
  const typeface = attr(face, 'typeface')
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
  let majorEaFont: string | undefined
  let minorEaFont: string | undefined
  let styles: PptxStyleMatrix | undefined

  if (masterPath) {
    const rels = await pkg.rels(masterPath)
    let themePath: string | null = null
    for (const rel of rels.values()) {
      if (rel.mode !== 'External' && relTypeIs(rel.type, 'theme')) {
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
        majorFont = fontOf(fontScheme, 'majorFont', false) ?? majorFont
        minorFont = fontOf(fontScheme, 'minorFont', false) ?? minorFont
        majorEaFont = fontOf(fontScheme, 'majorFont', true)
        minorEaFont = fontOf(fontScheme, 'minorFont', true)

        const fillLst = firstDescendant(root, 'a:fillStyleLst')
        const lnLst = firstDescendant(root, 'a:lnStyleLst')
        const effectLst = firstDescendant(root, 'a:effectStyleLst')
        styles = {
          fills: fillLst ? Array.from(fillLst.children) : [],
          lines: lnLst ? directChildren(lnLst, 'a:ln') : [],
          // effectStyle 包裹一层：取内部 a:effectLst（缺失给空壳，保持 idx 对齐）
          effects: effectLst
            ? Array.from(effectLst.children).map((es) => directChild(es, 'a:effectLst') ?? es)
            : [],
        }
      }
    }
  }

  return { schemeColors, majorFont, minorFont, majorEaFont, minorEaFont, colorMap, styles }
}

/** schemeClr 名称求值：先过母版 clrMap（bg1→lt1 等），再查主题色表 */
export function resolveSchemeColor(name: string, theme: PptxTheme): string {
  const mapped = theme.colorMap[name] ?? name
  return theme.schemeColors[mapped] ?? theme.schemeColors[name] ?? '#000000'
}
