// src/core/import/theme.ts —— Task 5 会整体替换
export interface PptxTheme {
  schemeColors: Record<string, string>
  majorFont: string
  minorFont: string
  colorMap: Record<string, string>
}
