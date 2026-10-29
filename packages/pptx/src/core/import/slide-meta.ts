// src/core/import/slide-meta.ts
/** 页面级元数据：切换效果、备注、分节、隐藏页 */
import { attr, directChild, directChildren, firstDescendant, isPresentationml, parseXML } from './xml'
import { relTypeIs, type PptxPackage } from './package'
import type { SlideTransition, TransitionPreset } from '../../types/slides'

/** slide → notesSlide 的正文纯文本；缺失/空返回 undefined */
export async function parseNotesText(pkg: PptxPackage, slidePath: string): Promise<string | undefined> {
  const rels = await pkg.rels(slidePath)
  let notesPath: string | null = null
  for (const rel of rels.values()) {
    if (rel.mode !== 'External' && relTypeIs(rel.type, 'notesSlide')) {
      notesPath = rel.target
      break
    }
  }
  if (!notesPath) return undefined
  const xml = await pkg.text(notesPath)
  if (!xml) return undefined
  const root = parseXML(xml).documentElement
  const spTree = firstDescendant(firstDescendant(root, 'p:cSld') ?? root, 'p:spTree')
  if (!spTree) return undefined
  // 备注正文占位符：ph type="body"（编号占位符等噪音不计入）
  let bodySp: Element | null = null
  for (const sp of Array.from(spTree.children)) {
    if (sp.localName !== 'sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (ph && attr(ph, 'type') === 'body') { bodySp = sp; break }
  }
  const txBody = bodySp ? firstDescendant(bodySp, 'p:txBody') : null
  if (!txBody) return undefined
  const lines: string[] = []
  for (const p of directChildren(txBody, 'a:p')) {
    let line = ''
    for (const node of Array.from(p.children)) {
      if (node.localName === 'br') line += '\n'
      else if (node.localName === 'r' || node.localName === 'fld') line += firstDescendant(node, 'a:t')?.textContent ?? ''
    }
    lines.push(line)
  }
  const text = lines.join('\n').trim()
  return text || undefined
}

/** p:transition → 内部 SlideTransition；未识别预设返回 { transition: undefined, unsupported: true } */
export function parseTransition(root: Element): { transition?: SlideTransition; unsupported: boolean } {
  const trans = firstDescendant(root, 'p:transition')
  if (!trans) return { unsupported: false }
  let child: Element | null = null
  for (const node of Array.from(trans.children)) {
    if ((isPresentationml(node) || node.nodeName.startsWith('p:')) && node.localName !== 'sndAc' && node.localName !== 'extLst') {
      child = node
      break
    }
  }
  if (!child) return { unsupported: false }
  const dir = attr(child, 'dir') ?? ''
  const preset = mapPreset(child.nodeName.replace(/^p:/, ''), dir)
  if (!preset) return { unsupported: true }
  const spd = attr(trans, 'spd')
  const p14dur = parseInt(attr(firstDescendant(trans, 'p14:dur'), 'val') ?? '', 10)
  const duration = Number.isFinite(p14dur) && p14dur > 0
    ? p14dur
    : spd === 'slow' ? 1000 : spd === 'fast' ? 500 : 750
  return { transition: { preset, duration, easing: 'ease' }, unsupported: false }
}

function mapPreset(kind: string, dir: string): TransitionPreset | undefined {
  switch (kind) {
    case 'fade':
      return 'fade'
    case 'push':
      return dir === 'u' ? 'slideUp' : dir === 'd' ? 'slideDown' : dir === 'r' ? 'slideRight' : 'slideLeft'
    case 'wipe':
      return dir === 'u' || dir === 'd' ? 'wipeRight' : 'wipeLeft'
    case 'zoom':
      return 'zoomIn'
    case 'blinds':
      return 'blinds'
    case 'flip':
      return 'flip'
    default:
      return undefined
  }
}

/** p:timing 中存在动画节点时报告 animation（仅切换时长等空壳不算） */
export function hasAnimation(root: Element): boolean {
  const timing = firstDescendant(root, 'p:timing')
  if (!timing) return false
  for (const tag of ['p:par', 'p:anim', 'p:animEffect', 'p:animMotion', 'p:animRot', 'p:animScale', 'p:set', 'p:cmd']) {
    if (firstDescendant(timing, tag)) return true
  }
  return false
}

/** p14:sectionLst → sldId@id → 分节名（p14 前缀缺失时按本地名兜底） */
export function parseSectionMap(root: Element): Map<string, string> {
  const map = new Map<string, string>()
  const sectionLst = firstDescendant(root, 'p14:sectionLst')
  if (!sectionLst) return map
  for (const section of Array.from(sectionLst.children)) {
    if (section.localName !== 'section') continue
    const name = attr(section, 'name') ?? ''
    const idLst = directChild(section, 'p14:sldIdLst') ?? directChild(section, 'sldIdLst')
    if (!idLst || !name) continue
    for (const sldId of Array.from(idLst.children)) {
      const id = attr(sldId, 'id')
      if (id) map.set(id, name)
    }
  }
  return map
}
