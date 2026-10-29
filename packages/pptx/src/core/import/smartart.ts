// src/core/import/smartart.ts
/** SmartArt 静态回退：定位幻灯片缓存的 dsp:drawing（Microsoft 2007 diagramDrawing 关系）并换算框架坐标。
 * 导入与预览共用；不依赖元素解析，避免与分发器循环引用。 */
import { attr, directChild, firstDescendant, graphicDataKind, parseXML } from './xml'
import { relTypeIs, type PptxPackage } from './package'
import { resolveColor } from './styles'
import type { PptxTheme } from './theme'
import type { GroupXform } from './context'

export interface DiagramDrawing {
  /** 绘制部件包内路径（rels 解析基准） */
  path: string
  doc: Document
}

/** 部件文档缓存：同一包内多次访问（如母版装饰复用于每页）只解析一次 */
const docCache = new WeakMap<PptxPackage, Map<string, Document | null>>()

async function partDoc(pkg: PptxPackage, path: string): Promise<Document | null> {
  let cache = docCache.get(pkg)
  if (!cache) {
    cache = new Map()
    docCache.set(pkg, cache)
  }
  const cached = cache.get(path)
  if (cached !== undefined) return cached
  let doc: Document | null = null
  const xml = await pkg.text(path)
  if (xml) {
    try {
      doc = parseXML(xml)
    } catch {
      doc = null
    }
  }
  cache.set(path, doc)
  return doc
}

/** 按部件序号取关系目标（drawing1/data1… 与幻灯片中 SmartArt 出现顺序一致） */
async function relTargetsByIndex(pkg: PptxPackage, partPath: string, kind: string): Promise<string[]> {
  const rels = await pkg.rels(partPath)
  return [...rels.values()]
    .filter((rel) => rel.mode !== 'External' && relTypeIs(rel.type, kind))
    .map((rel) => rel.target)
    .sort((a, b) => (parseInt(a.match(/(\d+)\.xml$/i)?.[1] ?? '0', 10)) - (parseInt(b.match(/(\d+)\.xml$/i)?.[1] ?? '0', 10)))
}

/** 当前 graphicFrame 是本部件中第几个 SmartArt（按 preceding diagram 兄弟节点计数） */
export function diagramIndex(frame: Element): number {
  const parent = frame.parentElement
  if (!parent) return 0
  let n = 0
  for (const sib of Array.from(parent.children)) {
    if (sib === frame) break
    if (sib.localName === 'graphicFrame' && graphicDataKind(sib) === 'diagram') n += 1
  }
  return n
}

/** 定位与第 index 个 SmartArt 对应的绘制部件；缺失返回 null */
export async function findDiagramDrawing(
  pkg: PptxPackage,
  partPath: string,
  index: number,
): Promise<DiagramDrawing | null> {
  const targets = await relTargetsByIndex(pkg, partPath, 'diagramDrawing')
  const valid: DiagramDrawing[] = []
  for (const path of targets) {
    const doc = await partDoc(pkg, path)
    if (doc) valid.push({ path, doc })
  }
  return valid[index] ?? valid[0] ?? null
}

/** SmartArt 自身背景色（dgm:dataModel/dgm:bg；缓存绘制部件不含它，需从数据部件取） */
export async function diagramBackgroundColor(
  pkg: PptxPackage,
  partPath: string,
  index: number,
  theme: PptxTheme,
): Promise<string | undefined> {
  const targets = await relTargetsByIndex(pkg, partPath, 'diagramData')
  const path = targets[index] ?? targets[0]
  if (!path) return undefined
  const doc = await partDoc(pkg, path)
  const bg = doc ? firstDescendant(doc.documentElement, 'dgm:bg') : null
  if (!bg) return undefined
  return resolveColor(directChild(bg, 'a:solidFill'), theme)
}

/** dsp:drawing > dsp:spTree */
export function drawingSpTree(doc: Document): Element | null {
  return firstDescendant(doc.documentElement, 'dsp:spTree')
}

/** 框架 p:xfrm → 绘制子空间（框架本地 EMU）到父空间的仿射：p' = 中心 + R·F·(p + off - 中心) */
export function diagramFrameXform(frame: Element): GroupXform {
  const xfrm = firstDescendant(frame, 'p:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const ox = off ? parseInt(attr(off, 'x') ?? '0', 10) : 0
  const oy = off ? parseInt(attr(off, 'y') ?? '0', 10) : 0
  const cx = ext ? parseInt(attr(ext, 'cx') ?? '0', 10) : 0
  const cy = ext ? parseInt(attr(ext, 'cy') ?? '0', 10) : 0
  const deg = xfrm ? parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000 : 0
  const fx = xfrm && attr(xfrm, 'flipH') === '1' ? -1 : 1
  const fy = xfrm && attr(xfrm, 'flipV') === '1' ? -1 : 1
  const theta = (deg * Math.PI) / 180
  const cos = Math.cos(theta)
  const sin = Math.sin(theta)
  const ccx = ox + cx / 2
  const ccy = oy + cy / 2
  const m00 = cos * fx
  const m01 = -sin * fy
  const m10 = sin * fx
  const m11 = cos * fy
  return {
    a: m00,
    b: m10,
    c: m01,
    d: m11,
    e: ccx + m00 * (ox - ccx) + m01 * (oy - ccy),
    f: ccy + m10 * (ox - ccx) + m11 * (oy - ccy),
    rot: deg,
    flipH: fx < 0,
    flipV: fy < 0,
  }
}
