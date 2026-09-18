// src/core/import/fonts.ts
/** 字体栈生成：OOXML typeface（含 +mj/+mn 主题引用）→ CSS font-family。
 * 栈序：latin → ea → 中文系统回退（CSS 按字符逐字回退：西文命中 latin，中文落至 ea 或系统字体）。 */
import type { PptxTheme } from './theme'

/** 中文回退栈：macOS PingFang SC / Windows Microsoft YaHei，末位通用族 */
const CN_FALLBACK = ['PingFang SC', 'Microsoft YaHei']

/** 主题字体引用 → 主题字体名（latin 位与 ea 位各自解析） */
function resolveThemeRef(typeface: string, theme: PptxTheme, ea: boolean): string | undefined {
  if (!typeface.startsWith('+')) return typeface
  const major = typeface.includes('+mj')
  if (ea) return major ? theme.majorEaFont : theme.minorEaFont
  return major ? theme.majorFont : theme.minorFont
}

/** 字体名转义：单引号包裹并剔除引号/反斜杠/分号，防 style 属性逃逸（字体名来自不可信文件属性） */
function quote(name: string): string {
  return `'${name.replace(/['"\\;]/g, '')}'`
}

/** 组装字体栈：latin/ea 依序去重，附中文回退栈；全部缺失时仅回退栈 */
export function fontStackOf(
  latin: string | null | undefined,
  ea: string | null | undefined,
  theme: PptxTheme,
): string {
  const resolved = [
    latin ? resolveThemeRef(latin, theme, false) : undefined,
    ea ? resolveThemeRef(ea, theme, true) : undefined,
  ].filter((v): v is string => Boolean(v))
  const seen = new Set<string>()
  const faces: string[] = []
  for (const name of [...resolved, ...CN_FALLBACK]) {
    if (seen.has(name)) continue
    seen.add(name)
    faces.push(quote(name))
  }
  return [...faces, 'sans-serif'].join(', ')
}
