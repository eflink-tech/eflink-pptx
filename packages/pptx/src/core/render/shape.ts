// 预设形状库：归一化 viewBox 0 0 100 100 的 SVG path
// key 与 PPTX preset geometry 名称尽量对齐，便于导入导出双向映射

export interface ShapeMeta {
  /** 分类（插入菜单分组用） */
  category: 'basic' | 'arrow' | 'star' | 'flow' | 'callout' | 'math'
  name: string
  /** SVG path d 串（可含多段）；evenodd 形状需以 evenodd 渲染 */
  path: string
  /** 是否需要 fill-rule evenodd（框/环等镂空形状） */
  evenodd?: boolean
}

const R = 8 // 圆角矩形默认圆角

export const SHAPE_PATHS: Record<string, ShapeMeta> = {
  /* ---------- basic ---------- */
  rect: { category: 'basic', name: '矩形', path: 'M0 0H100V100H0Z' },
  roundRect: { category: 'basic', name: '圆角矩形', path: `M${R} 0H${100 - R}Q100 0 100 ${R}V${100 - R}Q100 100 ${100 - R} 100H${R}Q0 100 0 ${100 - R}V${R}Q0 0 ${R} 0Z` },
  ellipse: { category: 'basic', name: '圆形', path: 'M50 0A50 50 0 1 1 49.9 0Z' },
  triangle: { category: 'basic', name: '等腰三角形', path: 'M50 0L100 100H0Z' },
  rtTriangle: { category: 'basic', name: '直角三角形', path: 'M0 0H100L0 100Z' },
  parallelogram: { category: 'basic', name: '平行四边形', path: 'M25 0H100L75 100H0Z' },
  trapezoid: { category: 'basic', name: '梯形', path: 'M20 0H80L100 100H0Z' },
  diamond: { category: 'basic', name: '菱形', path: 'M50 0L100 50L50 100L0 50Z' },
  pentagon: { category: 'basic', name: '正五边形', path: 'M50 0L98.5 35.3L80.2 92.7H19.8L1.5 35.3Z' },
  hexagon: { category: 'basic', name: '正六边形', path: 'M25 0H75L100 50L75 100H25L0 50Z' },
  heptagon: { category: 'basic', name: '正七边形', path: 'M50 0L90.1 22.2L99.6 68.3L69.4 103.1L22 97.6L0 57.6L15.3 14.2Z' },
  octagon: { category: 'basic', name: '正八边形', path: 'M30 0H70L100 30V70L70 100H30L0 70V30Z' },
  donut: { category: 'basic', name: '圆环', path: 'M50 0A50 50 0 1 1 49.9 0ZM50 25A25 25 0 1 0 50.1 25Z', evenodd: true },
  frame: { category: 'basic', name: '框架', path: 'M0 0H100V100H0ZM15 15V85H85V15Z', evenodd: true },
  halfFrame: { category: 'basic', name: 'L形', path: 'M0 0H50V50H100V100H0Z' },
  pie: { category: 'basic', name: '饼形', path: 'M50 50L50 0A50 50 0 0 1 100 50Z' },
  chord: { category: 'basic', name: '弦形', path: 'M0 0A75 75 0 0 0 100 0Z' },
  arc: { category: 'basic', name: '弧形', path: 'M0 0A75 75 0 0 0 100 0' },
  cross: { category: 'basic', name: '十字形', path: 'M35 0H65V35H100V65H65V100H35V65H0V35H35Z' },
  plus: { category: 'basic', name: '加号', path: 'M40 0H60V40H100V60H60V100H40V60H0V40H40Z' },
  plaque: { category: 'basic', name: '匾形', path: 'M0 10Q25 0 50 0Q75 0 100 10V90Q75 100 50 100Q25 100 0 90Z' },
  can: { category: 'basic', name: '圆柱', path: 'M0 15A50 15 0 0 1 100 15V85A50 15 0 0 1 0 85ZM0 15A50 15 0 0 0 100 15' },
  cube: { category: 'basic', name: '立方体', path: 'M0 20L20 0H100V80L80 100H0ZM20 0V20H100M0 20H80V100' },

  /* ---------- arrow ---------- */
  arrowRight: { category: 'arrow', name: '右箭头', path: 'M0 30H60V0L100 50L60 100V70H0Z' },
  arrowLeft: { category: 'arrow', name: '左箭头', path: 'M100 30H40V0L0 50L40 100V70H100Z' },
  arrowUp: { category: 'arrow', name: '上箭头', path: 'M30 100V40H0L50 0L100 40H70V100Z' },
  arrowDown: { category: 'arrow', name: '下箭头', path: 'M30 0V60H0L50 100L100 60H70V0Z' },
  arrowLeftRight: { category: 'arrow', name: '左右箭头', path: 'M25 30V10L0 35L25 60V40H75V60L100 35L75 10V30Z' },
  arrowUpDown: { category: 'arrow', name: '上下箭头', path: 'M30 25H10L35 0L60 25H40V75H60L35 100L10 75H30Z' },
  arrowQuad: { category: 'arrow', name: '四向箭头', path: 'M50 0L65 20H55V42H78V32L100 50L78 68V58H55V80H65L50 100L35 80H45V58H22V68L0 50L22 32V42H45V20H35Z' },
  bentArrow: { category: 'arrow', name: '直角箭头', path: 'M0 100V30Q0 0 30 0H60V-0L100 25L60 50V25H40Q40 25 40 45V100Z' },
  curvedRightArrow: { category: 'arrow', name: '弧形箭头', path: 'M10 90Q10 30 70 30V0L100 40L70 80V50Q50 50 50 90Z' },
  chevron: { category: 'arrow', name: 'V形箭头', path: 'M0 0H60L100 50L60 100H0L40 50Z' },
  homePlate: { category: 'arrow', name: '五边形箭头', path: 'M0 0H65L100 50L65 100H0Z' },

  /* ---------- star ---------- */
  star5: { category: 'star', name: '五角星', path: 'M50 0L61.8 35.3L98.5 36.4L69.4 58.9L80.2 94.2L50 72.7L19.8 94.2L30.6 58.9L1.5 36.4L38.2 35.3Z' },
  star4: { category: 'star', name: '四角星', path: 'M50 0L62 38L100 50L62 62L50 100L38 62L0 50L38 38Z' },
  star6: { category: 'star', name: '六角星', path: 'M50 0L61.3 27.7L90.1 22.2L80.5 50L99.6 68.3L69.4 68.7L61.3 96.9L50 72.7L38.7 96.9L30.6 68.7L0.4 68.3L19.5 50L9.9 22.2L38.7 27.7Z' },
  star8: { category: 'star', name: '八角星', path: 'M50 0L61.8 22.7L86.6 13.4L77.3 38.2L100 50L77.3 61.8L86.6 86.6L61.8 77.3L50 100L38.2 77.3L13.4 86.6L22.7 61.8L0 50L22.7 38.2L13.4 13.4L38.2 22.7Z' },
  star16: { category: 'star', name: '十六角星', path: 'M50 0L56 12.4L62.5 0.4L66 13.5L74.6 3L75.2 16.7L85.4 8.1L83.3 21.6L94.3 15.5L89.3 28.3L100 25L92.4 35.4L103.4 35.4L93.1 43.5L104.4 46.4L92.1 51.6L103.5 54.7L92.2 57.6L102.5 65.7L91.5 65.7L99.1 76.1L88.1 72.9L93.1 85.7L82.1 79.6L84.2 93.1L74 84.5L73.4 98.2L64.8 87.7L61.3 100.8L54.8 88.8L48.8 100L44 88L38.7 99.9L35.2 86.8L26.6 97.3L26 83.6L15.8 92.2L17.9 78.7L6.9 84.8L11.9 72L1.2 75.3L8.8 64.9L-2.2 64.9L8.1 56.8L-3.2 53.9L9.1 48.7L-2.3 45.6L9 42.7L-1.3 34.6L9.7 34.6L2.1 24.2L13.1 27.4L8.1 14.6L19.1 20.7L17 7.2L27.2 15.8L27.8 2.1L36.4 12.6L39.9 -0.5L46.4 11.5Z' },
  burst: { category: 'star', name: '爆炸形', path: 'M50 0L58 18L70 4L74 23L90 14L88 33L100 34L90 45L100 55L88 56L90 75L74 66L70 85L58 71L50 89L42 71L30 85L26 66L10 75L12 56L0 55L10 45L0 34L12 33L10 14L26 23L30 4L42 18Z' },

  /* ---------- flow ---------- */
  flowProcess: { category: 'flow', name: '流程', path: 'M0 0H100V100H0Z' },
  flowDecision: { category: 'flow', name: '决策', path: 'M50 0L100 50L50 100L0 50Z' },
  flowData: { category: 'flow', name: '数据', path: 'M15 0H100L85 100H0Z' },
  flowDocument: { category: 'flow', name: '文档', path: 'M0 0H100V85Q50 110 0 85Z' },
  flowPredefined: { category: 'flow', name: '预定义', path: 'M8 0H92Q100 0 100 8V92Q100 100 92 100H8Q0 100 0 92V8Q0 0 8 0Z' },
  flowTerminal: { category: 'flow', name: '终止', path: 'M50 0H50A50 50 0 0 1 50 100H50A50 50 0 0 1 50 0Z M28 0H72Q100 0 100 50Q100 100 72 100H28Q0 100 0 50Q0 0 28 0Z' },
  flowConnector: { category: 'flow', name: '连接', path: 'M50 0A50 50 0 1 1 49.9 0Z' },
  flowOffpage: { category: 'flow', name: '离页', path: 'M0 0H100V70L50 100L0 70Z' },
  flowDelay: { category: 'flow', name: '延迟', path: 'M0 0H50A50 50 0 0 1 50 100H0Z' },
  flowManualInput: { category: 'flow', name: '手动输入', path: 'M0 20L100 0V100H0Z' },
  flowManualOperation: { category: 'flow', name: '手动操作', path: 'M0 0H100L85 100H15Z' },
  flowMerge: { category: 'flow', name: '合并', path: 'M0 0H100L50 100Z' },
  flowOr: { category: 'flow', name: '或', path: 'M50 0A50 50 0 1 1 49.9 0ZM50 25A25 25 0 1 0 50.1 25Z', evenodd: true },
  flowExtract: { category: 'flow', name: '抽取', path: 'M50 0L100 100H0Z' },

  /* ---------- callout ---------- */
  callout1: { category: 'callout', name: '矩形标注', path: 'M0 0H100V75H55L30 100V75H0Z' },
  callout2: { category: 'callout', name: '圆角标注', path: `M${R} 0H${100 - R}Q100 0 100 ${R}V60Q100 75 ${100 - R} 75H55L30 100V75H${R}Q0 75 0 60V${R}Q0 0 ${R} 0Z` },
  callout3: { category: 'callout', name: '椭圆形标注', path: 'M50 0A50 37.5 0 1 1 49.9 0ZM40 66L25 100L58 71Z' },
  calloutCloud: { category: 'callout', name: '云朵标注', path: 'M28 55A18 18 0 0 1 30 20A22 22 0 0 1 72 22A16 16 0 0 1 75 54L45 100L42 55Z' },

  /* ---------- math ---------- */
  mathPlus: { category: 'math', name: '加', path: 'M40 0H60V40H100V60H60V100H40V60H0V40H40Z' },
  mathMinus: { category: 'math', name: '减', path: 'M0 40H100V60H0Z' },
  mathMultiply: { category: 'math', name: '乘', path: 'M42 0H58V42H100V58H58V100H42V58H0V42H42Z' },
  mathDivide: { category: 'math', name: '除', path: 'M0 42H100V58H0ZM50 22A11 11 0 1 1 50.1 22ZM50 67A11 11 0 1 1 50.1 67Z', evenodd: true },
  mathEqual: { category: 'math', name: '等号', path: 'M0 28H100V44H0ZM0 56H100V72H0Z' },
  mathNotEqual: { category: 'math', name: '不等号', path: 'M18 20H36V37L82 25V41L36 53V59L82 47V63L36 75V100H18V79L0 84V68L18 63V57L0 62V46L18 41Z' },
  heart: { category: 'math', name: '心形', path: 'M50 90C20 65 0 45 0 26C0 10 12 0 25 0C35 0 45 6 50 16C55 6 65 0 75 0C88 0 100 10 100 26C100 45 80 65 50 90Z' },
  lightning: { category: 'math', name: '闪电', path: 'M55 0L20 55H45L35 100L80 40H52L65 0Z' },
  sun: { category: 'math', name: '太阳', path: 'M50 22A28 28 0 1 1 49.9 22ZM50 0L54 14H46ZM50 100L46 86H54ZM0 50L14 46V54ZM100 50L86 54V46ZM14 14L25 23L18 28ZM86 86L75 77L82 72ZM86 14L82 28L75 23ZM14 86L18 72L25 77Z' },
  moon: { category: 'math', name: '月形', path: 'M63 0A55 55 0 0 0 63 100A55 55 0 1 1 63 0Z' },
  cloud: { category: 'math', name: '云形', path: 'M25 75A20 20 0 0 1 27 35A25 25 0 0 1 75 30A18 18 0 0 1 78 75Z' },
  smile: { category: 'math', name: '笑脸', path: 'M50 0A50 50 0 1 1 49.9 0ZM30 38A8 8 0 1 0 30.1 38ZM70 38A8 8 0 1 0 70.1 38ZM28 60Q50 80 72 60L78 66Q50 90 22 66Z', evenodd: true },
  umbrella: { category: 'math', name: '伞形', path: 'M50 0A50 50 0 0 1 100 50L90 42A10 10 0 0 0 70 42L60 42A10 10 0 0 0 40 42L30 42A10 10 0 0 0 10 42L0 50A50 50 0 0 1 50 0ZM47 45H53V85A8 8 0 0 1 37 85H43A2 2 0 0 0 47 83Z' },
}

/** 形状分类顺序（插入菜单） */
export const SHAPE_CATEGORIES: Array<{ id: ShapeMeta['category']; name: string }> = [
  { id: 'basic', name: '基本形状' },
  { id: 'arrow', name: '箭头' },
  { id: 'star', name: '星形' },
  { id: 'flow', name: '流程图' },
  { id: 'callout', name: '标注' },
  { id: 'math', name: '图标' },
]

/** 获取形状 path（未知 key 回退矩形） */
export function getShapePath(key: string): ShapeMeta {
  return SHAPE_PATHS[key] ?? SHAPE_PATHS.rect
}
