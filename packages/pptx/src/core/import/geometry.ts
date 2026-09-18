/** 几何解析：全量 OOXML prstGeom 预设映射 + custGeom → SVG path（0-100 视口空间） */
import { attr, directChildren, firstDescendant } from './xml'

// PRST_MAP 仅列真实存在的 OOXML 预设名（不编造预设）；未列出的预设由 getShapeKey 兜底 'rect'。
// connector 类预设映射为 'rect' 仅为占位——p:cxnSp 节点在 shape.ts 中先按线条处理，不会走到该映射。
/** OOXML prstGeom → 内部 shapeKey（无视觉等价 → 'rect'） */
export const PRST_MAP: Record<string, string> = {
  rect: 'rect', roundRect: 'roundRect', round1Rect: 'roundRect', round2SameRect: 'roundRect',
  round2DiagRect: 'roundRect', snip1Rect: 'rect', snip2SameRect: 'rect', snip2DiagRect: 'rect',
  snipRoundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle',
  parallelogram: 'parallelogram', trapezoid: 'trapezoid', nonIsoscelesTrapezoid: 'trapezoid',
  pentagon: 'pentagon', hexagon: 'hexagon', heptagon: 'heptagon', octagon: 'octagon',
  decagon: 'octagon', dodecagon: 'octagon',
  star4: 'star4', star5: 'star5', star6: 'star6', star7: 'star5', star8: 'star8',
  star10: 'star5', star12: 'star8', star16: 'star16', star24: 'star16', star32: 'star16',
  donut: 'donut', noSmoking: 'donut', blockArc: 'arc', pie: 'pie', chord: 'chord', arc: 'arc',
  frame: 'frame', halfFrame: 'halfFrame', corner: 'rtTriangle', diagStripe: 'parallelogram',
  can: 'can', cube: 'cube', bevel: 'cube', plaque: 'plaque', plaqueTab: 'plaque',
  foldedCorner: 'frame', teardrop: 'ellipse',
  cross: 'cross', plus: 'mathPlus',
  heart: 'heart', lightningBolt: 'lightning', sun: 'sun', moon: 'moon', cloud: 'cloud',
  smileyFace: 'smile', irregularSeal1: 'cloud', irregularSeal2: 'cloud',
  rightArrow: 'arrowRight', leftArrow: 'arrowLeft', upArrow: 'arrowUp', downArrow: 'arrowDown',
  leftRightArrow: 'arrowLeftRight', upDownArrow: 'arrowUpDown', quadArrow: 'arrowQuad',
  leftUpArrow: 'arrowLeft', bentArrow: 'bentArrow', bentUpArrow: 'bentArrow', uturnArrow: 'bentArrow',
  circularArrow: 'bentArrow', curvedLeftArrow: 'curvedRightArrow', curvedRightArrow: 'curvedRightArrow',
  curvedUpArrow: 'bentArrow', curvedDownArrow: 'bentArrow',
  stripedRightArrow: 'arrowRight', notchedRightArrow: 'arrowRight', homePlate: 'homePlate',
  chevron: 'chevron',
  wedgeRectCallout: 'callout1', wedgeRoundRectCallout: 'callout2', wedgeEllipseCallout: 'callout3',
  cloudCallout: 'callout3', borderRectCallout: 'callout1', borderRoundRectCallout: 'callout2',
  borderEllipseCallout: 'callout3',
  mathPlus: 'mathPlus', mathMinus: 'mathMinus', mathMultiply: 'mathMultiply',
  mathDivide: 'mathDivide', mathEqual: 'mathEqual', mathNotEqual: 'mathNotEqual',
  flowChartProcess: 'flowProcess', flowChartAlternateProcess: 'flowProcess',
  flowChartDecision: 'flowDecision', flowChartInputOutput: 'flowData', flowChartData: 'flowData',
  flowChartDocument: 'flowDocument', flowChartMultidocument: 'flowDocument',
  flowChartPredefinedProcess: 'flowPredefined', flowChartInternalStorage: 'flowPredefined',
  flowChartTerminator: 'flowTerminal', flowChartPreparation: 'flowProcess',
  flowChartManualInput: 'flowManualInput', flowChartManualOperation: 'flowManualOperation',
  flowChartConnector: 'flowConnector', flowChartOffpageConnector: 'flowOffpage',
  flowChartCard: 'flowDocument', flowChartPunchedTape: 'flowData',
  flowChartSummingJunction: 'flowOr', flowChartOr: 'flowOr', flowChartCollate: 'flowMerge',
  flowChartSort: 'flowMerge', flowChartExtract: 'flowExtract', flowChartMerge: 'flowMerge',
  flowChartOnlineStorage: 'flowData', flowChartMagneticTape: 'flowData',
  flowChartMagneticDisk: 'flowData', flowChartMagneticDrum: 'flowData',
  flowChartDirectAccessStorage: 'flowData', flowChartDisplay: 'flowDocument',
  flowChartDelay: 'flowDelay', flowChartStoredData: 'flowData',
  line: 'rect', straightConnector1: 'rect', bentConnector2: 'rect', bentConnector3: 'rect',
  bentConnector4: 'rect', bentConnector5: 'rect', curvedConnector2: 'rect', curvedConnector3: 'rect',
  curvedConnector4: 'rect', curvedConnector5: 'rect',
  verticalScroll: 'frame', horizontalScroll: 'frame', ribbon: 'plaque', ribbon2: 'plaque',
  ellipseRibbon: 'plaque', ellipseRibbon2: 'plaque', seal: 'star8', seal4: 'star8',
}

export function getShapeKey(prst: string): string {
  return PRST_MAP[prst] ?? 'rect'
}

/** custGeom → SVG path 字符串（归一化 0-100 视口；返回 null 表示无可解析路径） */
export function custGeomToPath(cust: Element, w: number, h: number): string | null {
  const pathLst = firstDescendant(cust, 'a:pathLst')
  if (!pathLst) return null
  const emuW = Math.max(1, w * 9525)
  const emuH = Math.max(1, h * 9525)
  const cmds: string[] = []
  const fmt = (n: number) => Math.round(n * 100) / 100
  for (const pathEl of directChildren(pathLst, 'a:path')) {
    const pw = parseInt(attr(pathEl, 'w') ?? '0', 10) || emuW
    const ph = parseInt(attr(pathEl, 'h') ?? '0', 10) || emuH
    // 纯归一化到 0-100 视口空间：坐标超出路径尺寸合法，不做有损裁剪（渲染端 viewBox 负责裁剪显示）
    const fx = (x: number) => (x / pw) * 100
    const fy = (y: number) => (y / ph) * 100
    for (const seg of Array.from(pathEl.children)) {
      const pts = directChildren(seg, 'a:pt').map((pt) => [
        parseInt(attr(pt, 'x') ?? '0', 10),
        parseInt(attr(pt, 'y') ?? '0', 10),
      ] as const)
      switch (seg.nodeName) {
        case 'a:moveTo':
          if (pts[0]) cmds.push(`M${fmt(fx(pts[0][0]))},${fmt(fy(pts[0][1]))}`)
          break
        case 'a:lnTo':
          if (pts[0]) cmds.push(`L${fmt(fx(pts[0][0]))},${fmt(fy(pts[0][1]))}`)
          break
        case 'a:cubicBezTo':
          if (pts.length >= 3) {
            cmds.push(`C${fmt(fx(pts[0][0]))},${fmt(fy(pts[0][1]))} ${fmt(fx(pts[1][0]))},${fmt(fy(pts[1][1]))} ${fmt(fx(pts[2][0]))},${fmt(fy(pts[2][1]))}`)
          }
          break
        case 'a:quadBezTo':
          if (pts.length >= 2) {
            cmds.push(`Q${fmt(fx(pts[0][0]))},${fmt(fy(pts[0][1]))} ${fmt(fx(pts[1][0]))},${fmt(fy(pts[1][1]))}`)
          }
          break
        case 'a:close':
          cmds.push('Z')
          break
        case 'a:arcTo':
          // a:arcTo 无 a:pt 子元素，暂不支持圆弧参数：跳过该段（弦线近似，由前后段连线闭合）
          break
        default:
          break
      }
    }
  }
  return cmds.length ? cmds.join(' ') : null
}
