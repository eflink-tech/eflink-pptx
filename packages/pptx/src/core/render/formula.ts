// 公式静态渲染：与 MathLive 编辑器共用同一套排版引擎，保证所见即所得
import { convertLatexToMarkup } from 'mathlive'

/** MathLive 默认占位符符号（U+25A2 ▢），与编辑器 placeholderSymbol 一致 */
const PLACEHOLDER_SYMBOL = '\u25A2'

/**
 * 将 MathLive 的 \\placeholder{} 转为可显示的 ▢ 符号。
 * convertLatexToMarkup 对 placeholder 只渲染空白 Box，需预处理。
 */
function latexForDisplay(latex: string): string {
  return latex.replace(
    /\\placeholder(?:\[[^\]]*\])?\{\}/g,
    `\\class{ML__placeholder}{${PLACEHOLDER_SYMBOL}}`,
  )
}

/** 将 LaTeX 转为 MathLive 静态 HTML markup */
export function renderFormulaMarkup(latex: string): string {
  try {
    return convertLatexToMarkup(latexForDisplay(latex), { defaultMode: 'math' })
  } catch {
    return '<span style="color:#f00">公式错误</span>'
  }
}
