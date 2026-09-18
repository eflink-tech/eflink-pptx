/** 预览层跨文件测试工厂：makeCtx 供各 preview/*.test.ts 共用。
 * 独立成模块以避免从 *.test.ts 导入导致 vitest 重复注册用例（导入方的 describe/it 会在导入方上下文再跑一遍）。 */
import type { ImportReport } from '../../types/slides'
import { SVG_NS, type PreviewCtx } from './svg'

/** 构造最小预览上下文（结构兼容 ParseContext，pkg 仅占位） */
export function makeCtx(overrides?: Partial<PreviewCtx>): PreviewCtx {
  const defs = document.createElementNS(SVG_NS, 'defs')
  let n = 0
  return {
    pkg: {} as never,
    partPath: 'ppt/slides/slide1.xml',
    theme: { schemeColors: {}, majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {} },
    report: { skipped: {} } as ImportReport,
    scale: { x: 1, y: 1 },
    placeholders: new Map(),
    defs,
    uid: (prefix: string) => `${prefix}-${n++}`,
    ...overrides,
  }
}
