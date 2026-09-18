// src/core/import/index.ts
/** 导入入口：解压 → ancestry → 逐页解析 → Presentation（分层容错，失败项计入 report） */
import type { ImportReport, PPTElement, Presentation, Slide, Theme } from '../../types/slides'
import { createDefaultTheme } from '../../types/slides'
import { PptxPackage, type PartRel } from './package'
import type { PptxTheme } from './theme'
import { parseSlideAncestry, parseBackgroundFill } from './master'
import { parseSpTreeNode } from './elements'
import { IDENTITY_XFORM, addSkipped } from './context'
import { attr, firstDescendant, parseXML } from './xml'
import { genId } from '../utils/id'

export { resolveTarget } from './package'

export interface ImportResult {
  presentation: Presentation
  report: ImportReport
}

/** 装饰元素模板克隆：深拷贝并重编 id，避免多页共享同一对象 */
function cloneDecoration(el: PPTElement): PPTElement {
  const clone = JSON.parse(JSON.stringify(el)) as PPTElement
  const dash = el.id.indexOf('-')
  clone.id = genId(dash > 0 ? `${el.id.slice(0, dash + 1)}` : '')
  return clone
}

/** PptxTheme → 输出 Theme（色板按 pptx 语义顺序映射，缺失项回退 Office 默认色） */
function buildOutputTheme(theme: PptxTheme | undefined): Theme {
  if (!theme) return createDefaultTheme()
  const sc = theme.schemeColors
  return {
    colors: [
      sc.accent1 ?? '#4472C4', sc.accent2 ?? '#ED7D31', sc.accent3 ?? '#A5A5A5',
      sc.accent4 ?? '#FFC000', sc.accent5 ?? '#5B9BD5', sc.accent6 ?? '#70AD47',
      sc.dk2 ?? '#44546A', sc.lt2 ?? '#E7E6E6', sc.dk1 ?? '#000000', sc.lt1 ?? '#FFFFFF',
    ],
    background: { type: 'solid', color: sc.lt1 ?? '#FFFFFF' },
    fontName: theme.minorFont,
    fontColor: sc.dk1 ?? '#333333',
    // 保留 fontScheme 原始字体：导出还原 theme1.xml，run 与 theme 字体一致时用 +mn 引用
    ooxmlFonts: {
      major: theme.majorFont, majorEa: theme.majorEaFont,
      minor: theme.minorFont, minorEa: theme.minorEaFont,
    },
  }
}

/** 页面路径与源画布尺寸：sldIdLst → rels；缺失时按 slideN.xml 编号兜底。
 * 导入与预览两条链路共享此入口，保证页面顺序/尺寸/兜底口径一致。 */
export async function listSlidePaths(
  pkg: PptxPackage,
): Promise<{ paths: string[]; srcW: number; srcH: number; ratio: number }> {
  const presXml = await pkg.text('ppt/presentation.xml')
  if (!presXml) throw new Error('不是有效的 PPTX 文件（缺少 presentation.xml）')
  const root = parseXML(presXml).documentElement

  // 幻灯片尺寸（EMU）→ 源画布尺寸与宽高比
  const sldSz = firstDescendant(root, 'p:sldSz')
  const srcW = parseInt(attr(sldSz, 'cx') ?? '12192000', 10)
  const srcH = parseInt(attr(sldSz, 'cy') ?? '6858000', 10)
  const ratio = srcW / srcH || 16 / 9

  // 页面顺序：sldIdLst → rels；缺失时按 slideN.xml 编号兜底
  const sldIdLst = firstDescendant(root, 'p:sldIdLst')
  const slideIds = sldIdLst ? Array.from(sldIdLst.getElementsByTagName('p:sldId')) : []
  const rels = await pkg.rels('ppt/presentation.xml')
  const slidePaths = slideIds
    .map((id) => attr(id, 'r:id'))
    .map((rid) => (rid ? rels.get(rid) : undefined))
    .filter((rel): rel is PartRel => Boolean(rel && rel.mode !== 'External'))
    .map((rel) => (rel.target.startsWith('ppt/') ? rel.target : `ppt/${rel.target}`))
  const paths = slidePaths.length
    ? slidePaths
    : pkg.paths()
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => (parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10) - parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10)))
  if (!paths.length) throw new Error('PPTX 中没有幻灯片')
  return { paths, srcW, srcH, ratio }
}

/** 解析 pptx → Presentation + 兼容性报告 */
export async function importPPTXDetailed(file: File): Promise<ImportResult> {
  const pkg = await PptxPackage.load(file)
  const { paths, srcW, srcH, ratio } = await listSlidePaths(pkg)
  const scale = { x: 1280 / (srcW / 9525), y: (1280 / ratio) / (srcH / 9525) }

  const report: ImportReport = { skipped: {} }
  const slides: Slide[] = []
  let firstTheme: PptxTheme | undefined

  for (const slidePath of paths) {
    const xml = await pkg.text(slidePath)
    if (!xml) continue
    try {
      const doc = parseXML(xml)
      const cSld = firstDescendant(doc.documentElement, 'p:cSld')
      const spTree = cSld ? firstDescendant(cSld, 'p:spTree') : null

      // ancestry 内部已解析主题/母版链，这里直接复用其结果（rels 有缓存，无重复 IO）
      const ancestry = await parseSlideAncestry(pkg, slidePath, report, scale)
      const ctx = {
        pkg,
        partPath: slidePath,
        theme: ancestry.theme,
        report,
        scale,
        placeholders: ancestry.placeholders,
      }

      const elements: PPTElement[] = ancestry.decorations.map(cloneDecoration)
      if (spTree) {
        for (const child of Array.from(spTree.children)) {
          elements.push(...(await parseSpTreeNode(child, ctx, IDENTITY_XFORM, pkg)))
        }
      }

      // 背景：slide 自身 bg 优先 → 版式/母版（ancestry.background 已兜底 lt1/白）→ 白色
      const bgEl = firstDescendant(doc.documentElement, 'p:bg')
      let background = parseBackgroundFill(bgEl, ancestry.theme)
      if (background?.type === 'image' && bgEl) {
        const blip = firstDescendant(bgEl, 'a:blip')
        const embedId = attr(blip, 'r:embed')
        if (embedId) {
          const target = await pkg.relTarget(slidePath, embedId)
          const src = target ? await pkg.mediaDataUrl(target) : undefined
          background = src ? { type: 'image', image: { src, size: 'cover' } } : undefined
        } else {
          background = undefined
        }
      }

      if (!firstTheme) firstTheme = ancestry.theme

      slides.push({
        id: genId('slide-'),
        elements,
        background: background ?? ancestry.background ?? { type: 'solid', color: '#ffffff' },
      })
    } catch (e) {
      // 单页失败不拖垮整个导入：计入报告后继续下一页
      console.warn('[pptx-import] 页面解析失败:', slidePath, e)
      addSkipped(report, 'slideParseFailed')
    }
  }
  if (!slides.length) throw new Error('PPTX 中没有可解析的幻灯片')

  return {
    presentation: {
      slides,
      theme: buildOutputTheme(firstTheme),
      width: 1280,
      viewportRatio: ratio,
    },
    report,
  }
}

/** 兼容既有调用方：仅返回 Presentation */
export async function importPPTX(file: File): Promise<Presentation> {
  return (await importPPTXDetailed(file)).presentation
}
