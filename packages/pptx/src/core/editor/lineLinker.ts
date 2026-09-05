// 线条端点：锚点计算、吸附、附着跟随
import { useEditorStore } from '../../store/editorStore'
import type { LineAnchorIndex, LineAttach, LineElement, PPTElement } from '../../types/slides'

const SNAP_RADIUS = 14
const BBOX_PAD = 6

type Point = [number, number]

/** 可连接图形（除线条自身外） */
export function isConnectable(el: PPTElement): boolean {
  return el.type !== 'line'
}

/** 元素局部锚点（相对左上角，未旋转） */
export function getLocalAnchorPoint(w: number, h: number, anchor: LineAnchorIndex): Point {
  switch (anchor) {
    case 0: return [w / 2, 0]
    case 1: return [w, h / 2]
    case 2: return [w / 2, h]
    default: return [0, h / 2]
  }
}

/** 元素局部坐标 → 画布世界坐标（含旋转） */
export function elementLocalToWorld(el: PPTElement, localX: number, localY: number): Point {
  const rot = ((el.rotate ?? 0) * Math.PI) / 180
  const cx = el.x + el.w / 2
  const cy = el.y + el.h / 2
  const dx = localX - el.w / 2
  const dy = localY - el.h / 2
  if (!rot) return [el.x + localX, el.y + localY]
  const cos = Math.cos(rot)
  const sin = Math.sin(rot)
  return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos]
}

/** 线条局部坐标 → 画布世界坐标（含旋转） */
export function lineLocalToWorld(line: LineElement, local: Point): Point {
  const rot = ((line.rotate ?? 0) * Math.PI) / 180
  if (!rot) return [line.x + local[0], line.y + local[1]]
  const cx = line.x + line.w / 2
  const cy = line.y + line.h / 2
  const lx = local[0] - line.w / 2
  const ly = local[1] - line.h / 2
  const cos = Math.cos(rot)
  const sin = Math.sin(rot)
  return [cx + lx * cos - ly * sin, cy + lx * sin + ly * cos]
}

/** 画布世界坐标 → 线条局部坐标（重算 bbox 后 rotation 清零） */
export function worldToLineLocal(line: LineElement, world: Point): Point {
  return [world[0] - line.x, world[1] - line.y]
}

/** 获取图形全部锚点的世界坐标 */
export function getElementAnchors(el: PPTElement): Array<{ anchor: LineAnchorIndex; x: number; y: number }> {
  if (!isConnectable(el)) return []
  return ([0, 1, 2, 3] as LineAnchorIndex[]).map((anchor) => {
    const [lx, ly] = getLocalAnchorPoint(el.w, el.h, anchor)
    const [x, y] = elementLocalToWorld(el, lx, ly)
    return { anchor, x, y }
  })
}

/** 解析线条端点世界坐标（优先附着锚点） */
export function resolveLineEndpoint(
  line: LineElement,
  endpoint: 'start' | 'end',
  elements: PPTElement[],
): Point {
  const attach = endpoint === 'start' ? line.startAttach : line.endAttach
  if (attach) {
    const target = elements.find((e) => e.id === attach.elementId)
    if (target && isConnectable(target)) {
      const [lx, ly] = getLocalAnchorPoint(target.w, target.h, attach.anchor)
      return elementLocalToWorld(target, lx, ly)
    }
  }
  return lineLocalToWorld(line, endpoint === 'start' ? line.start : line.end)
}

/** 根据世界坐标点重算线条 bbox 并重映射局部坐标 */
export function recomputeLineFromWorld(
  line: LineElement,
  worldStart: Point,
  worldPoints: Point[],
  worldEnd: Point,
): void {
  const all = [worldStart, ...worldPoints, worldEnd]
  const minX = Math.min(...all.map((p) => p[0]))
  const minY = Math.min(...all.map((p) => p[1]))
  const maxX = Math.max(...all.map((p) => p[0]))
  const maxY = Math.max(...all.map((p) => p[1]))
  line.x = Math.floor(minX - BBOX_PAD)
  line.y = Math.floor(minY - BBOX_PAD)
  line.w = Math.max(8, Math.ceil(maxX - minX + BBOX_PAD * 2))
  line.h = Math.max(8, Math.ceil(maxY - minY + BBOX_PAD * 2))
  line.rotate = undefined
  line.start = worldToLineLocal(line, worldStart)
  line.end = worldToLineLocal(line, worldEnd)
  line.points = worldPoints.length ? worldPoints.map((p) => worldToLineLocal(line, p)) : undefined
}

export interface SnapResult {
  x: number
  y: number
  attach?: LineAttach
}

/** 吸附到最近图形锚点 */
export function snapToAnchor(
  x: number,
  y: number,
  elements: PPTElement[],
  excludeIds: string[],
): SnapResult {
  let best: SnapResult = { x, y }
  let bestDist = SNAP_RADIUS
  for (const el of elements) {
    if (!isConnectable(el) || excludeIds.includes(el.id)) continue
    for (const a of getElementAnchors(el)) {
      const d = Math.hypot(a.x - x, a.y - y)
      if (d < bestDist) {
        bestDist = d
        best = { x: a.x, y: a.y, attach: { elementId: el.id, anchor: a.anchor } }
      }
    }
  }
  return best
}

/** 线条端点是否附着到指定图形 */
export function lineAttachedTo(line: LineElement, elementId: string): boolean {
  return line.startAttach?.elementId === elementId || line.endAttach?.elementId === elementId
}

/** 图形移动/缩放后，更新附着线条端点（需在 updateElements 内调用） */
export function syncLinesAttachedTo(elements: PPTElement[], movedIds: string[]): void {
  if (!movedIds.length) return
  for (const el of elements) {
    if (el.type !== 'line') continue
    const attached = (el.startAttach && movedIds.includes(el.startAttach.elementId))
      || (el.endAttach && movedIds.includes(el.endAttach.elementId))
    if (!attached) continue
    const ws = resolveLineEndpoint(el, 'start', elements)
    const we = resolveLineEndpoint(el, 'end', elements)
    const wps = (el.points ?? []).map((p) => lineLocalToWorld(el, p))
    recomputeLineFromWorld(el, ws, wps, we)
  }
}

/** 图形变换后刷新附着线条（通过 store 触发响应式更新） */
export function refreshAttachedLines(movedIds: string[]): void {
  if (!movedIds.length) return
  const { presentation, slideIndex, updateElements } = useEditorStore.getState()
  const elements = presentation.slides[slideIndex].elements
  const lineIds = elements
    .filter((el): el is LineElement => el.type === 'line' && movedIds.some((id) => lineAttachedTo(el, id)))
    .map((el) => el.id)
  if (!lineIds.length) return
  updateElements(lineIds, (line) => {
    if (line.type !== 'line') return
    const els = useEditorStore.getState().presentation.slides[useEditorStore.getState().slideIndex].elements
    const ws = resolveLineEndpoint(line, 'start', els)
    const we = resolveLineEndpoint(line, 'end', els)
    const wps = (line.points ?? []).map((p) => lineLocalToWorld(line, p))
    recomputeLineFromWorld(line, ws, wps, we)
  }, { history: false })
}

/** 设置线条某一端点（世界坐标 + 可选附着） */
export function setLineEndpointWorld(
  line: LineElement,
  endpoint: 'start' | 'end',
  world: Point,
  attach: LineAttach | undefined,
  elements: PPTElement[],
): void {
  if (endpoint === 'start') {
    line.startAttach = attach
  } else {
    line.endAttach = attach
  }
  const ws = resolveLineEndpoint(line, 'start', elements)
  const we = resolveLineEndpoint(line, 'end', elements)
  const wps = (line.points ?? []).map((p) => lineLocalToWorld(line, p))
  if (endpoint === 'start') {
    recomputeLineFromWorld(line, world, wps, we)
  } else {
    recomputeLineFromWorld(line, ws, wps, world)
  }
}

/** 设置中间控制点（世界坐标） */
export function setLineMidPointWorld(
  line: LineElement,
  index: number,
  world: Point,
  elements: PPTElement[],
): void {
  const ws = resolveLineEndpoint(line, 'start', elements)
  const we = resolveLineEndpoint(line, 'end', elements)
  const wps = (line.points ?? []).map((p) => lineLocalToWorld(line, p))
  while (wps.length <= index) wps.push([...wps[wps.length - 1] ?? ws] as Point)
  wps[index] = world
  recomputeLineFromWorld(line, ws, wps, we)
}

/** 收集画布上全部可显示锚点（去重） */
export function collectAllAnchors(
  elements: PPTElement[],
  excludeLineId?: string,
): Array<{ elementId: string; anchor: LineAnchorIndex; x: number; y: number }> {
  const out: Array<{ elementId: string; anchor: LineAnchorIndex; x: number; y: number }> = []
  for (const el of elements) {
    if (!isConnectable(el)) continue
    if (el.id === excludeLineId) continue
    for (const a of getElementAnchors(el)) {
      out.push({ elementId: el.id, anchor: a.anchor, x: a.x, y: a.y })
    }
  }
  return out
}
