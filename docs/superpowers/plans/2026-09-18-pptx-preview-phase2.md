# PPTX 第二期：OOXML→SVG 高保真预览 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增 `src/core/preview/` OOXML→SVG 渲染器：每页 slide → 一个 `<svg>`，按 OOXML 语义直接渲染（不做编辑器模型映射取舍），提供 `previewPPTX(file): Promise<SVGSVGElement[]>` 入口，并接入 ImportDialog「先预览再导入」。

**Architecture:** 复用一期解析层（theme/styles/geometry/text/table/chart/master），预览侧只做「OOXML → SVG 节点」直渲。预览坐标系 = 源画布 px（viewBox = 源宽高 EMU/9525，浮点不取整），因此 `ctx.scale = {x:1,y:1}`；组合用 SVG transform 表达（旋转/翻转无需一期的一期降级）。渐变/阴影/裁剪登记到页级 `<defs>`，文本用 foreignObject 复用已 XSS 加固的 `txBodyToHTML`，形状路径复用 `SHAPE_PATHS`（嵌套 svg viewBox 0 0 100 100 + preserveAspectRatio="none"）与 `custGeomToPath`。

**Tech Stack:** TypeScript + DOMParser/createElementNS（vitest jsdom 可跑）、JSZip（测试内动态构造最小 pptx，沿用 `core/import/index.test.ts` 的 `buildMinimalPptx` 模式）。

**红线（延续一期）:** 禁止移植/复制 PPTist（AGPL-3.0）代码；全局中文注释；提交信息不加 Co-Authored-By；命令一律在 `packages/pptx` 目录执行（`npx tsc -b` / `npx vitest run` / oxlint）。

---

## 文件结构

```
packages/pptx/src/core/preview/          # 全部新建
├── svg.ts           # SVG 助手：PreviewCtx / svgEl / emu2pxF / sz2px / Box / geomOf / boxTransform
├── svg.test.ts
├── shape.ts         # renderShape / renderLine / registerLinearGradient / registerShadowFilter / lnOf / registerMarker
├── shape.test.ts
├── text.ts          # renderText（foreignObject）
├── picture.ts       # renderPicture（srcRect 裁剪 / 阴影 / 边框 / 音视频海报+播放标记）
├── picture.test.ts
├── table.ts         # renderTable（合并格 / 富文本单元格）
├── table.test.ts
├── chart.ts         # renderChart（bar/line/pie/area/scatter 直渲；radar 占位）
├── chart.test.ts
├── group.ts         # renderGroup（SVG transform，含旋转）
├── group.test.ts
├── dispatch.ts      # renderSpTreeNode 分发器（避免 slide↔group 循环引用）
├── slide.ts         # renderSlide（背景/装饰/占位符/组装）
├── slide.test.ts
└── index.ts         # previewPPTXDetailed / previewPPTX 入口 + index.test.ts

修改（一期文件，均为非破坏性增量）：
- core/import/elements/text.ts      # rPr@spc 字距、pPr a:lnSpc 行距（编辑导入同步受益）
- core/import/elements/table.ts     # buildCellMatrix 增加返回 tcEls 元素矩阵（预览富文本单元格用）
- core/import/index.ts              # 提取共享 listSlidePaths(pkg)（页面顺序 + 源画布尺寸）
- components/dialogs/ImportDialog.tsx  # 先预览再导入 UI
```

依赖方向：`index → slide → dispatch → {shape,text,picture,table,chart,group}`，`group → dispatch`（无环）；`shape → text`（形状内文本）。

---

### Task 1: 新建分支 + preview/svg.ts 基础助手

**Files:**
- Create: `packages/pptx/src/core/preview/svg.ts`
- Test: `packages/pptx/src/core/preview/svg.test.ts`

- [ ] **Step 1: 新建功能分支**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx
git checkout main && git pull
git checkout -b feat/pptx-preview-phase2
```

- [ ] **Step 2: 写失败测试** `packages/pptx/src/core/preview/svg.test.ts`

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import type { ImportReport } from '../../types/slides'
import { SVG_NS, svgEl, emu2pxF, sz2px, geomOf, boxTransform, type PreviewCtx } from './svg'

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

describe('preview/svg 助手', () => {
  it('emu2pxF：EMU→px 浮点不取整', () => {
    expect(emu2pxF(9525)).toBe(1)
    expect(emu2pxF(12192000)).toBe(1280)
    expect(emu2pxF(9524)).toBeCloseTo(0.9996, 4)
  })

  it('sz2px：字号（1/100 pt）→ px', () => {
    expect(sz2px(1800)).toBe(24)
    expect(sz2px(2400)).toBe(32)
  })

  it('svgEl：构建命名空间元素，undefined/null 属性跳过', () => {
    const el = svgEl('rect', { x: 1, fill: undefined, stroke: null, width: '10' })
    expect(el.namespaceURI).toBe('http://www.w3.org/2000/svg')
    expect(el.getAttribute('x')).toBe('1')
    expect(el.getAttribute('width')).toBe('10')
    expect(el.hasAttribute('fill')).toBe(false)
    expect(el.hasAttribute('stroke')).toBe(false)
  })

  it('geomOf：a:xfrm → Box（px 浮点 + rot/flip）', () => {
    const xml = `<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:sp><p:spPr>
        <a:xfrm rot="1800000" flipH="1"><a:off x="952500" y="190500"/><a:ext cx="4762500" cy="952500"/></a:xfrm>
      </p:spPr></p:sp></root>`
    const sp = parseXML(xml).documentElement.firstElementChild as Element
    const box = geomOf(sp, makeCtx())
    expect(box).not.toBeNull()
    expect(box!.x).toBe(100)
    expect(box!.y).toBe(20)
    expect(box!.w).toBe(500)
    expect(box!.h).toBe(100)
    expect(box!.rot).toBe(30)
    expect(box!.flipH).toBe(true)
    expect(box!.flipV).toBe(false)
  })

  it('geomOf：无 xfrm 时回退占位符位置表（EMU→px）', () => {
    const xml = `<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:sp><p:nvSpPr><p:ph type="title"/></p:nvSpPr><p:spPr/></p:sp></root>`
    const sp = parseXML(xml).documentElement.firstElementChild as Element
    const placeholders = new Map([['title', { x: 952500, y: 952500, w: 9525000, h: 952500 }]])
    const box = geomOf(sp, makeCtx({ placeholders }))
    expect(box).toEqual({ x: 100, y: 100, w: 1000, h: 100, rot: 0, flipH: false, flipV: false })
  })

  it('boxTransform：rot 绕自身中心、flip 平移翻转', () => {
    const t1 = boxTransform({ x: 100, y: 100, w: 500, h: 100, rot: 30, flipH: false, flipV: false })
    expect(t1).toBe('rotate(30,350,150)')
    const t2 = boxTransform({ x: 100, y: 100, w: 500, h: 100, rot: 0, flipH: true, flipV: false })
    expect(t2).toBe('translate(600,100) scale(-1,1) translate(-100,-100)')
    expect(boxTransform({ x: 0, y: 0, w: 10, h: 10, rot: 0, flipH: false, flipV: false })).toBeUndefined()
  })
})
```

- [ ] **Step 3: 运行确认失败**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx/packages/pptx
npx vitest run src/core/preview/svg.test.ts
```

预期：FAIL（模块 `./svg` 不存在）。

- [ ] **Step 4: 实现** `packages/pptx/src/core/preview/svg.ts`

```typescript
/** SVG 预览助手：命名空间元素构建、单位换算、元素几何 Box。
 * 预览坐标系 = 源画布 px（EMU/9525 浮点），viewBox 即源画布尺寸；因此 ctx.scale 恒为 {1,1}。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import type { ParseContext } from '../import/context'

export const SVG_NS = 'http://www.w3.org/2000/svg'

/** 预览上下文：结构兼容 ParseContext（可直接传给一期解析函数），额外持有 defs 与 id 生成器 */
export interface PreviewCtx extends ParseContext {
  /** 当前页 SVG 的 <defs>：渐变 / 滤镜 / 裁剪路径登记处 */
  defs: SVGElement
  /** 唯一 id 生成器（实现方保证跨页唯一，如 `${prefix}-p${pageIdx}-${n++}`） */
  uid: (prefix: string) => string
}

/** 构建 SVG 命名空间元素；值为 undefined/null 的属性跳过 */
export function svgEl(tag: string, attrs?: Record<string, string | number | undefined | null>): SVGElement {
  const el = document.createElementNS(SVG_NS, tag)
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null) continue
      el.setAttribute(k, String(v))
    }
  }
  return el
}

/** EMU → px（96dpi，浮点不取整：viewBox 用浮点保证布局精确） */
export function emu2pxF(emu: number): number {
  return emu / 9525
}

/** 字号/字距（1/100 pt）→ px */
export function sz2px(hundredthsPt: number): number {
  return hundredthsPt / 100 / 0.75
}

export interface Box {
  x: number
  y: number
  w: number
  h: number
  /** 度 */
  rot: number
  flipH: boolean
  flipV: boolean
}

/** 元素几何：a:xfrm（sp/pic 用）或 p:xfrm（graphicFrame 用）→ 源画布 px 浮点 Box。
 * 无 xfrm 时回退占位符位置表（EMU → px），与一期 parseShapeEl 的占位符继承语义一致。 */
export function geomOf(node: Element, ctx: PreviewCtx): Box | null {
  const xfrm = firstDescendant(node, 'a:xfrm') ?? firstDescendant(node, 'p:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (xfrm && off && ext) {
    return {
      x: emu2pxF(parseInt(attr(off, 'x') ?? '0', 10)),
      y: emu2pxF(parseInt(attr(off, 'y') ?? '0', 10)),
      w: emu2pxF(parseInt(attr(ext, 'cx') ?? '0', 10)),
      h: emu2pxF(parseInt(attr(ext, 'cy') ?? '0', 10)),
      rot: parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000,
      flipH: attr(xfrm, 'flipH') === '1',
      flipV: attr(xfrm, 'flipV') === '1',
    }
  }
  const ph = firstDescendant(firstDescendant(node, 'p:nvSpPr') ?? node, 'p:ph')
  if (ph) {
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    const pos = ctx.placeholders.get(key)
    if (pos) {
      return {
        x: emu2pxF(pos.x), y: emu2pxF(pos.y), w: emu2pxF(pos.w), h: emu2pxF(pos.h),
        rot: 0, flipH: false, flipV: false,
      }
    }
  }
  return null
}

/** 元素变换属性串：flip（绕自身包围盒翻转）+ rot（绕自身中心）。无变换返回 undefined。 */
export function boxTransform(box: Box): string | undefined {
  const parts: string[] = []
  if (box.flipH || box.flipV) {
    const fx = box.flipH ? -1 : 1
    const fy = box.flipV ? -1 : 1
    parts.push(
      `translate(${box.x + (fx < 0 ? box.w : 0)},${box.y + (fy < 0 ? box.h : 0)}) scale(${fx},${fy}) translate(${-box.x},${-box.y})`,
    )
  }
  if (box.rot) {
    parts.push(`rotate(${box.rot},${box.x + box.w / 2},${box.y + box.h / 2})`)
  }
  return parts.length ? parts.join(' ') : undefined
}
```

- [ ] **Step 5: 运行确认通过**

```bash
npx vitest run src/core/preview/svg.test.ts
```

预期：PASS（6 用例）。

- [ ] **Step 6: 提交**

```bash
git add packages/pptx/src/core/preview/svg.ts packages/pptx/src/core/preview/svg.test.ts
git commit -m "feat(preview): SVG 预览基础助手（PreviewCtx/svgEl/几何换算）"
```

---

### Task 2: text.ts 扩展 —— rPr@spc 字距 + pPr a:lnSpc 行距

设计文档要求「精确字距」按 OOXML 语义渲染；该能力在共享层实现，编辑导入与预览同时受益。`spc` 单位 1/100 pt，`a:spcPct@val` 单位 1/100000（行高倍数），`a:spcPts@val` 单位 1/100 pt。

**Files:**
- Modify: `packages/pptx/src/core/import/elements/text.ts`
- Test: `packages/pptx/src/core/import/elements/text.test.ts`（追加用例）

- [ ] **Step 1: 追加失败测试**（在 `text.test.ts` 末尾追加；沿用该文件既有的 XML 构造方式，若已有 helper 直接复用）

```typescript
describe('txBodyToHTML 字距与行距', () => {
  it('rPr@spc → letter-spacing（1/100 pt → px）', async () => {
    const { txBodyOf } = await import('./text.test-helpers').catch(() => ({ txBodyOf: null }) as never) as never
  })
})
```

> 注意：上面仅是占位说明——实际写法按 `text.test.ts` 现有构造模式（`new DOMParser().parseFromString('<p:txBody xmlns:p="…" xmlns:a="…">…</p:txBody>', 'text/xml').documentElement`）。以下三个用例为最终形态：

```typescript
describe('txBodyToHTML 字距与行距', () => {
  const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  const theme = { schemeColors: {}, majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {} }

  const txBodyOf = (inner: string) =>
    parseXML(`<p:txBody ${NS}>${inner}</p:txBody>`).documentElement

  it('rPr@spc → letter-spacing', async () => {
    const body = txBodyOf('<a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN" spc="300"/><a:t>AB</a:t></a:r></a:p>')
    const { html } = await txBodyToHTML(body, theme)
    expect(html).toContain('letter-spacing:4px') // 300/100/0.75 = 4
  })

  it('a:lnSpc a:spcPct → line-height 倍数', async () => {
    const body = txBodyOf('<a:bodyPr/><a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN"/><a:t>x</a:t></a:r></a:p>')
    const { html } = await txBodyToHTML(body, theme)
    expect(html).toContain('line-height:1.5')
  })

  it('a:lnSpc a:spcPts → line-height px', async () => {
    const body = txBodyOf('<a:bodyPr/><a:p><a:pPr><a:lnSpc><a:spcPts val="2000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="zh-CN"/><a:t>x</a:t></a:r></a:p>')
    const { html } = await txBodyToHTML(body, theme)
    expect(html).toContain('line-height:27px') // 2000/100/0.75 = 26.67 → 四舍五入 27
  })
})
```

（`parseXML` 从 `../xml` 导入；若 `text.test.ts` 里已有同用途局部 helper，用既有 helper 替换 `txBodyOf`/`NS`，断言不变。）

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/import/elements/text.test.ts
```

预期：新增 3 例 FAIL（无 letter-spacing / line-height 输出）。

- [ ] **Step 3: 实现**

`text.ts` 三处修改：

1. 顶部 import 增加 `sz2px`（从 `../../preview/svg` 导入会形成 import→preview 反向依赖，**禁止**；在本文件内用既有公式实现，与 font-size 同源）：

```typescript
// paragraphToInner 内，typeface 分支之后追加：
      const spc = attr(rPr, 'spc')
      if (spc) {
        // 字距 1/100 pt → px（与 font-size 同源换算），保留两位小数防浮点噪声
        const ls = Math.round((parseInt(spc, 10) / 100 / 0.75) * 100) / 100
        styles.push(`letter-spacing:${ls}px`)
      }
```

2. `ParaInfo` 增加字段：

```typescript
interface ParaInfo {
  align: string
  kind: 'none' | 'bullet' | 'number'
  inner: string
  /** 行距样式（line-height:...），无行距时为空串 */
  spacing: string
}
```

3. `txBodyToHTML` 主循环内，`algn` 读取之后追加行距解析，并改 html 生成：

```typescript
    // 行距：spcPct（1/100000 → 倍数）优先，spcPts（1/100 pt → px）覆盖
    const lnSpc = pPr ? directChild(pPr, 'a:lnSpc') : null
    let spacing = ''
    const pct = lnSpc ? firstDescendant(lnSpc, 'a:spcPct') : null
    if (pct) spacing = `line-height:${(parseInt(attr(pct, 'val') ?? '100000', 10)) / 100000}`
    const pts = lnSpc ? firstDescendant(lnSpc, 'a:spcPts') : null
    if (pts) spacing = `line-height:${Math.round(parseInt(attr(pts, 'val') ?? '0', 10) / 100 / 0.75)}px`
    paras.push({ align, kind: bulletKind(p), inner, spacing })
```

```typescript
  while (i < paras.length) {
    const para = paras[i]
    if (para.kind === 'none') {
      html.push(`<p style="text-align:${para.align}${para.spacing ? `;${para.spacing}` : ''}">${para.inner || '&nbsp;'}</p>`)
      i += 1
      continue
    }
    const tag = para.kind === 'bullet' ? 'ul' : 'ol'
    const items: string[] = []
    while (i < paras.length && paras[i].kind === para.kind) {
      const cur = paras[i]
      items.push(`<li style="text-align:${cur.align}${cur.spacing ? `;${cur.spacing}` : ''}">${cur.inner || '&nbsp;'}</li>`)
      i += 1
    }
    html.push(`<${tag}>${items.join('')}</${tag}>`)
  }
```

- [ ] **Step 4: 运行确认通过（全量 text + 基线）**

```bash
npx vitest run src/core/import/elements/text.test.ts
npx vitest run src/core/import
```

预期：全部 PASS（既有 141 基线用例不得回归）。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/import/elements/text.ts packages/pptx/src/core/import/elements/text.test.ts
git commit -m "feat(import): 文本支持 OOXML 字距(spc)与行距(lnSpc)"
```

---

### Task 3: preview/shape.ts —— 形状/线条渲染 + 渐变/阴影/描边登记

**Files:**
- Create: `packages/pptx/src/core/preview/shape.ts`
- Test: `packages/pptx/src/core/preview/shape.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { SVG_NS, type PreviewCtx } from './svg'
import { renderShape, renderLine, registerLinearGradient } from './shape'
import { makeCtx } from './svg.test'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

describe('preview/renderShape', () => {
  it('预设形状：嵌套 svg + SHAPE_PATHS 路径 + solidFill 填充 + 描边', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="4762500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
      <a:ln w="25400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.tagName).toBe('g')
    const nested = g.querySelector('svg')!
    expect(nested.getAttribute('viewBox')).toBe('0 0 100 100')
    expect(nested.getAttribute('preserveAspectRatio')).toBe('none')
    expect(nested.getAttribute('x')).toBe('100')
    expect(nested.getAttribute('width')).toBe('500')
    const path = nested.querySelector('path')!
    expect(path.getAttribute('fill')).toBe('#42A5F5')
    expect(Number(path.getAttribute('stroke-width'))).toBeCloseTo(2.6667, 3) // 25400/9525
  })

  it('渐变填充：defs 登记 linearGradient 全停站，fill 引用 url(#id)', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill><a:gsLst>
        <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
      </a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    const path = g.querySelector('path')!
    const url = path.getAttribute('fill')!
    expect(url.startsWith('url(#grad-')).toBe(true)
    const grad = ctx.defs.querySelector('linearGradient')!
    expect(grad.querySelectorAll('stop')).toHaveLength(3)
    // ang=90°：渐变方向向下（y1 < y2），中间站保留
    expect(grad.querySelectorAll('stop')[1].getAttribute('stop-color')).toBe('#00FF00')
  })

  it('阴影：defs 登记 feDropShadow（blur/2 为 stdDeviation）', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.querySelector('path')!.getAttribute('filter')!.startsWith('url(#shadow-')).toBe(true)
    const drop = ctx.defs.querySelector('feDropShadow')!
    expect(Number(drop.getAttribute('stdDeviation'))).toBeCloseTo(50800 / 9525 / 2, 3)
    expect(Number(drop.getAttribute('dy'))).toBeCloseTo(38100 / 9525, 3) // dir=45° → dx≈dy
  })

  it('noFill → fill:none；rot/flip 上到 g 的 transform', async () => {
    const ctx = makeCtx()
    const sp = wrap(`<p:sp><p:spPr>
      <a:xfrm rot="1800000" flipH="1"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
    </p:spPr></p:sp>`)
    const g = (await renderShape(sp, ctx))!
    expect(g.querySelector('path')!.getAttribute('fill')).toBe('none')
    expect(g.getAttribute('transform')).toContain('scale(-1,1)')
    expect(g.getAttribute('transform')).toContain('rotate(30,')
  })

  it('p:cxnSp → line + 箭头 marker', () => {
    const ctx = makeCtx()
    const node = wrap(`<p:cxnSp><p:spPr>
      <a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        <a:tailEnd type="triangle"/></a:ln>
    </p:spPr></p:cxnSp>`)
    const line = renderLine(node, ctx)!
    expect(line.getAttribute('x1')).toBe('0')
    expect(line.getAttribute('y1')).toBe('100') // flipV
    expect(line.getAttribute('stroke')).toBe('#FF0000')
    expect(line.getAttribute('marker-end')!.startsWith('url(#marker-')).toBe(true)
    expect(ctx.defs.querySelector('marker')).not.toBeNull()
  })
})
```

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/shape.test.ts
```

预期：FAIL（模块不存在；`makeCtx` 需从 svg.test 导出——Task 1 的测试文件里已含该函数，补 `export`）。

- [ ] **Step 3: 实现** `packages/pptx/src/core/preview/shape.ts`

```typescript
/** p:sp / p:cxnSp → SVG：形状路径（嵌套 svg 0-100 视口）/ 线条；渐变、阴影、箭头登记到 defs。
 * 已知取舍：嵌套 svg 的 preserveAspectRatio="none" 会非均匀缩放描边（极端宽高比下描边粗细略有失真）。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import { resolveColor } from '../import/styles'
import { custGeomToPath, getShapeKey } from '../import/geometry'
import { getShapePath } from '../render/shape'
import { emu2pxF, svgEl, geomOf, boxTransform, type PreviewCtx, type Box } from './svg'
import { renderText } from './text'

/** a:gradFill → defs 登记 linearGradient（全部停站保留；a:lin@ang 从 3 点钟方向顺时针，SVG y 向下三角函数直接适用） */
export function registerLinearGradient(gradFill: Element, ctx: PreviewCtx): string {
  const id = ctx.uid('grad')
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildrenOf(gsLst, 'a:gs') : []
  const lin = directChild(gradFill, 'a:lin')
  const ang = ((parseInt(attr(lin, 'ang') ?? '0', 10)) / 60000) * (Math.PI / 180)
  const grad = svgEl('linearGradient', {
    id,
    x1: 0.5 - Math.cos(ang) / 2,
    y1: 0.5 - Math.sin(ang) / 2,
    x2: 0.5 + Math.cos(ang) / 2,
    y2: 0.5 + Math.sin(ang) / 2,
    gradientUnits: 'objectBoundingBox',
  })
  for (const gs of stops) {
    grad.appendChild(svgEl('stop', {
      offset: parseInt(attr(gs, 'pos') ?? '0', 10) / 100000,
      // resolveColor 输出 #RRGGBB（或含 alpha 的 8 位形式），浏览器 stop-color 均支持
      'stop-color': resolveColor(gs, ctx.theme) ?? '#000000',
    }))
  }
  ctx.defs.appendChild(grad)
  return `url(#${id})`
}

/** a:outerShdw → defs 登记 feDropShadow filter（stdDeviation = blur/2），无阴影返回 undefined */
export function registerShadowFilter(effectLst: Element | null, ctx: PreviewCtx): string | undefined {
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (!shdw) return undefined
  const id = ctx.uid('shadow')
  const blur = emu2pxF(parseInt(attr(shdw, 'blurRad') ?? '0', 10))
  const dist = emu2pxF(parseInt(attr(shdw, 'dist') ?? '0', 10))
  const dirRad = ((parseInt(attr(shdw, 'dir') ?? '0', 10)) / 60000) * (Math.PI / 180)
  const filter = svgEl('filter', { id, x: '-50%', y: '-50%', width: '200%', height: '200%' })
  filter.appendChild(svgEl('feDropShadow', {
    dx: dist * Math.cos(dirRad),
    dy: dist * Math.sin(dirRad),
    stdDeviation: blur / 2,
    'flood-color': resolveColor(shdw, ctx.theme) ?? '#000000',
  }))
  ctx.defs.appendChild(filter)
  return `url(#${id})`
}

/** a:ln → 描边属性；缺失或 a:noFill 返回 { 'stroke-width': 0 }（无边框） */
export function lnOf(ln: Element | null, ctx: PreviewCtx): Record<string, string | number> {
  if (!ln || directChild(ln, 'a:noFill')) return { 'stroke-width': 0 }
  const dash = attr(directChild(ln, 'a:prstDash'), 'val') ?? ''
  const dasharray = dash === 'dash' || dash === 'sysDash' || dash === 'lgDash'
    ? '12,6'
    : dash === 'dot' || dash === 'sysDot' || dash === 'lgDot' ? '2,4' : undefined
  const out: Record<string, string | number> = {
    stroke: resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? '#000000',
    'stroke-width': emu2pxF(parseInt(attr(ln, 'w') ?? '12700', 10)), // 缺省 12700 EMU = 1pt
  }
  if (dasharray) out['stroke-dasharray'] = dasharray
  return out
}

/** 端点箭头 → defs 登记 marker（triangle/stealth 实心三角、arrow 开放箭头、oval 圆点） */
function registerMarker(ctx: PreviewCtx, type: string, color: string): string {
  const id = ctx.uid('marker')
  let marker: SVGElement
  if (type === 'oval') {
    marker = svgEl('marker', { id, markerWidth: 8, markerHeight: 8, refX: 4, refY: 4, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('circle', { cx: 4, cy: 4, r: 3, fill: color }))
  } else if (type === 'arrow') {
    marker = svgEl('marker', { id, markerWidth: 10, markerHeight: 10, refX: 9, refY: 5, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('path', { d: 'M1,1 L9,5 L1,9', fill: 'none', stroke: color, 'stroke-width': 1.5 }))
  } else {
    marker = svgEl('marker', { id, markerWidth: 10, markerHeight: 10, refX: 9, refY: 5, orient: 'auto', markerUnits: 'strokeWidth' })
    marker.appendChild(svgEl('path', { d: 'M0,0 L10,5 L0,10 Z', fill: color }))
  }
  ctx.defs.appendChild(marker)
  return `url(#${id})`
}

/** p:cxnSp / 直线 prst → <line>（flip 决定起止点，rot 上 transform） */
export function renderLine(node: Element, ctx: PreviewCtx): SVGElement | null {
  const box = geomOf(node, ctx)
  if (!box) return null
  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
  const color = resolveColor(ln ? directChild(ln, 'a:solidFill') : solidFill, ctx.theme) ?? '#333333'
  const line = svgEl('line', {
    x1: box.flipH ? box.x + box.w : box.x,
    y1: box.flipV ? box.y + box.h : box.y,
    x2: box.flipH ? box.x : box.x + box.w,
    y2: box.flipV ? box.y : box.y + box.h,
    stroke: color,
    'stroke-width': ln ? emu2pxF(parseInt(attr(ln, 'w') ?? '12700', 10)) : 1,
    transform: boxTransform({ ...box, flipH: false, flipV: false }),
  })
  for (const [k, v] of Object.entries(lnOf(ln, ctx))) {
    if (k === 'stroke' || v === undefined) continue
    line.setAttribute(k, String(v))
  }
  const head = attr(ln ? directChild(ln, 'a:headEnd') : null, 'type')
  const tail = attr(ln ? directChild(ln, 'a:tailEnd') : null, 'type')
  if (head) line.setAttribute('marker-start', registerMarker(ctx, head, color))
  if (tail) line.setAttribute('marker-end', registerMarker(ctx, tail, color))
  return line
}

/** p:sp → <g>：路径（custGeom 优先，否则预设形状库）+ 渐变/阴影 + 形状内文本（renderText） */
export async function renderShape(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const prst = attr(firstDescendant(node, 'a:prstGeom'), 'prst') ?? ''
  if (node.nodeName === 'p:cxnSp' || prst === 'line' || prst === 'straightConnector1') {
    return renderLine(node, ctx)
  }
  const box = geomOf(node, ctx)
  if (!box) return null
  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const custGeom = spPr ? directChild(spPr, 'a:custGeom') : null
  const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
  const noFill = spPr ? directChild(spPr, 'a:noFill') : null
  const gradFill = spPr ? directChild(spPr, 'a:gradFill') : null
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const effectLst = spPr ? directChild(spPr, 'a:effectLst') : null

  const g = svgEl('g', { transform: boxTransform(box) })
  // 路径：custGeomToPath 产出 0-100 视口空间；预设形状库同为 0-100（未知 key 回退 rect）
  const meta = custGeom ? null : getShapePath(prst ? getShapeKey(prst) : 'rect')
  const d = custGeom ? custGeomToPath(custGeom, box.w, box.h) : meta?.path ?? null
  if (d) {
    const fill = gradFill
      ? registerLinearGradient(gradFill, ctx)
      : noFill ? 'none'
        : resolveColor(solidFill, ctx.theme) ?? 'none'
    const nested = svgEl('svg', {
      x: box.x, y: box.y, width: box.w, height: box.h,
      viewBox: '0 0 100 100', preserveAspectRatio: 'none', overflow: 'visible',
    })
    nested.appendChild(svgEl('path', {
      d, fill, filter: registerShadowFilter(effectLst, ctx),
      ...lnOf(ln, ctx),
      'fill-rule': meta?.evenodd ? 'evenodd' : undefined,
    }))
    g.appendChild(nested)
  }

  const txBody = firstDescendant(node, 'p:txBody') ?? firstDescendant(node, 'a:txBody')
  if (txBody) {
    const text = await renderText(txBody, ctx, box)
    if (text) g.appendChild(text)
  }
  return g.children.length ? g : null
}

/** directChildren 别名（xml.ts 未导出 directChildren 时的本文件内实现） */
function directChildrenOf(parent: Element, name: string): Element[] {
  return Array.from(parent.children).filter((c) => c.nodeName === name)
}
```

> 实现注记：`xml.ts` 已导出 `directChildren`，实现时**直接 import 使用**，删掉 `directChildrenOf` 别名（上方保留仅为展示完整逻辑）。`renderText` 在 Task 4 实现——本任务先创建 `preview/text.ts` 的**空壳**（`export async function renderText(...): Promise<SVGElement | null> { return null }`），保证 Task 3 可独立提交；Task 4 填充实现。

- [ ] **Step 4: 运行确认通过**

```bash
npx vitest run src/core/preview/shape.test.ts
```

预期：PASS（5 用例）。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/preview/
git commit -m "feat(preview): 形状/线条 SVG 渲染（渐变/阴影/描边/箭头登记 defs）"
```

---

### Task 4: preview/text.ts —— foreignObject 富文本

**Files:**
- Create: `packages/pptx/src/core/preview/text.ts`（替换 Task 3 空壳）
- Test: `packages/pptx/src/core/preview/text.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './svg.test'
import { renderText } from './text'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const txBodyOf = (inner: string) =>
  parseXML(`<root ${NS}><p:txBody>${inner}</p:txBody></root>`).documentElement.firstElementChild as Element

describe('preview/renderText', () => {
  it('foreignObject + 复用 txBodyToHTML 富文本 + bodyPr 内边距默认值', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN" sz="2400"/><a:t>标题</a:t></a:r></a:p>')
    const fo = (await renderText(txBody, ctx, { x: 100, y: 100, w: 500, h: 100, rot: 0, flipH: false, flipV: false }))!
    expect(fo.tagName).toBe('foreignObject')
    expect(fo.getAttribute('x')).toBe('100')
    expect(fo.getAttribute('width')).toBe('500')
    const div = fo.firstElementChild as HTMLElement
    expect(div.getAttribute('xmlns')).toBe('http://www.w3.org/1999/xhtml')
    // 默认内边距：lIns/rIns 91440 → 9.6px，tIns/bIns 45720 → 4.8px
    expect(div.getAttribute('style')).toContain('padding:4.8px 9.6px 4.8px 9.6px')
    expect(div.innerHTML).toContain('标题')
    expect(div.innerHTML).toContain('font-size:32px') // sz 2400 → 32px
  })

  it('anchor=t → 顶部对齐 flex；wrap=none → nowrap', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr anchor="t" wrap="none"/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    const div = ((await renderText(txBody, ctx, { x: 0, y: 0, w: 10, h: 10, rot: 0, flipH: false, flipV: false }))!
      .firstElementChild as HTMLElement)
    expect(div.getAttribute('style')).toContain('justify-content:flex-start')
    expect(div.getAttribute('style')).toContain('white-space:nowrap')
  })

  it('vert=eaVert → writing-mode:vertical-rl', async () => {
    const ctx = makeCtx()
    const txBody = txBodyOf('<a:bodyPr vert="eaVert"/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    const div = ((await renderText(txBody, ctx, { x: 0, y: 0, w: 10, h: 10, rot: 0, flipH: false, flipV: false }))!
      .firstElementChild as HTMLElement)
    expect(div.getAttribute('style')).toContain('writing-mode:vertical-rl')
  })
})
```

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/text.test.ts
```

预期：FAIL（空壳返回 null）。

- [ ] **Step 3: 实现**（整体替换 `preview/text.ts` 空壳）

```typescript
/** a:txBody → foreignObject：复用 txBodyToHTML（已 XSS 加固：escapeHTML + 链接协议白名单）。
 * 不做 normAutofit fontScale 缩放（记录于已知取舍），ins/anchor/wrap/vert 按 bodyPr 直渲。 */
import { attr, directChild } from '../import/xml'
import { txBodyToHTML } from '../import/elements/text'
import { emu2pxF, svgEl, type PreviewCtx, type Box } from './svg'

export async function renderText(txBody: Element, ctx: PreviewCtx, box: Box): Promise<SVGElement | null> {
  const body = await txBodyToHTML(txBody, ctx.theme, ctx.pkg, ctx.partPath)
  const bodyPr = directChild(txBody, 'a:bodyPr')
  // 内边距 EMU 默认值：lIns/rIns 91440、tIns/bIns 45720
  const lIns = emu2pxF(parseInt(attr(bodyPr, 'lIns') ?? '91440', 10))
  const tIns = emu2pxF(parseInt(attr(bodyPr, 'tIns') ?? '45720', 10))
  const rIns = emu2pxF(parseInt(attr(bodyPr, 'rIns') ?? '91440', 10))
  const bIns = emu2pxF(parseInt(attr(bodyPr, 'bIns') ?? '45720', 10))
  const anchor = attr(bodyPr, 'anchor')
  const justify = anchor === 't' ? 'flex-start' : anchor === 'b' ? 'flex-end' : anchor ? 'center' : undefined
  const vert = attr(bodyPr, 'vert')

  const fo = svgEl('foreignObject', { x: box.x, y: box.y, width: box.w, height: box.h })
  const div = document.createElement('div')
  div.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml')
  div.setAttribute('style', [
    'width:100%', 'height:100%', 'box-sizing:border-box', 'overflow:hidden',
    `padding:${tIns}px ${rIns}px ${bIns}px ${lIns}px`,
    justify ? `display:flex;flex-direction:column;justify-content:${justify}` : '',
    vert === 'vert' || vert === 'eaVert' || vert === 'mongolianVert' ? 'writing-mode:vertical-rl' : '',
    attr(bodyPr, 'wrap') === 'none' ? 'white-space:nowrap' : '',
  ].filter(Boolean).join(';'))
  div.innerHTML = body.html
  fo.appendChild(div)
  return fo
}
```

- [ ] **Step 4: 运行确认通过 + Task 3 回归**

```bash
npx vitest run src/core/preview/text.test.ts src/core/preview/shape.test.ts
```

预期：PASS。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/preview/text.ts packages/pptx/src/core/preview/text.test.ts
git commit -m "feat(preview): foreignObject 富文本渲染（内边距/对齐/竖排/不换行）"
```

---

### Task 5: preview/picture.ts —— 图片/音视频

**Files:**
- Create: `packages/pptx/src/core/preview/picture.ts`
- Test: `packages/pptx/src/core/preview/picture.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from '../import/package'
import { parseXML } from '../import/xml'
import { makeCtx } from './svg.test'
import { renderPicture } from './picture'
import { SVG_NS } from './svg'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

/** 构造含 1x1 PNG（红色）媒体 + slide1 rels 的包，返回挂好 pkg/partPath 的 ctx */
async function makePictureCtx(): Promise<{ ctx: ReturnType<typeof makeCtx>; cleanup: () => void }> {
  const zip = new JSZip()
  // 1x1 红色 PNG
  const pngB64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  zip.file('ppt/media/a.png', pngB64, { base64: true })
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/a.png"/>
</Relationships>`)
  const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
  const ctx = makeCtx({ pkg: pkg as never })
  return { ctx, cleanup: () => { /* zip 由 GC 回收 */ } }
}

const PIC = (extra = '') => `<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${extra}</p:spPr></p:pic>`

describe('preview/renderPicture', () => {
  it('基本图片：image + preserveAspectRatio=none', async () => {
    const { ctx } = await makePictureCtx()
    const g = (await renderPicture(wrap(PIC()), ctx))!
    const img = g.querySelector('image')!
    expect(img.getAttribute('href')).toContain('data:image/png')
    expect(img.getAttribute('preserveAspectRatio')).toBe('none')
    expect(img.getAttribute('width')).toBe('100')
  })

  it('srcRect 裁剪：图片放大映射 + clipPath 裁回框内', async () => {
    const { ctx } = await makePictureCtx()
    const g = (await renderPicture(wrap(PIC('<a:effectLst/>')), ctx))! // 无关分支兜底
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/><a:srcRect l="20000" t="10000" r="20000" b="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="6000" cy="3000"/></a:xfrm></p:spPr></p:pic>`)
    const g2 = (await renderPicture(node, ctx))!
    const img = g2.querySelector('image')!
    // 裁剪区宽 60% → 全图宽 1000/0.6 ≈ 1666.7；x = 1000/9525*9525... 直接断言关键关系
    expect(Number(img.getAttribute('width'))).toBeGreaterThan(9525 / 9525 * 600 / 0.6 - 1)
    expect(img.getAttribute('clip-path')!.startsWith('url(#clip-')).toBe(true)
    expect(ctx.defs.querySelector('clipPath')).not.toBeNull()
  })

  it('缺失图片 → null 且计入报告', async () => {
    const ctx = makeCtx()
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdX"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic>`)
    expect(await renderPicture(node, ctx)).toBeNull()
    expect(ctx.report.skipped.missingImage).toBe(1)
  })

  it('视频：海报帧 + 居中播放标记', async () => {
    const { ctx } = await makePictureCtx()
    const node = wrap(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="v"/><p:nvPr><a:videoFile r:link="rId1"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr></p:pic>`)
    const g = (await renderPicture(node, ctx))!
    expect(g.querySelector('image')).not.toBeNull() // 海报帧照常渲染
    expect(g.querySelector('circle')).not.toBeNull() // 播放标记
    expect(g.querySelector('path')).not.toBeNull()
  })
})
```

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/picture.test.ts
```

- [ ] **Step 3: 实现** `packages/pptx/src/core/preview/picture.ts`

```typescript
/** p:pic → <image>：srcRect 裁剪（裁剪子区映射满框 + clipPath 裁回）、阴影、边框；音视频渲染海报帧 + 播放标记。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import { addSkipped } from '../import/context'
import { registerShadowFilter, lnOf } from './shape'
import { svgEl, geomOf, boxTransform, type PreviewCtx } from './svg'

export async function renderPicture(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const box = geomOf(node, ctx)
  if (!box) return null
  const blipFill = firstDescendant(node, 'p:blipFill') ?? firstDescendant(node, 'a:blipFill')
  const blip = blipFill ? directChild(blipFill, 'a:blip') : null
  const embedId = attr(blip, 'r:embed')
  const target = embedId ? await ctx.pkg.relTarget(ctx.partPath, embedId) : null
  const src = target ? await ctx.pkg.mediaDataUrl(target) : undefined
  if (!src) {
    addSkipped(ctx.report, 'missingImage')
    return null
  }

  // srcRect（十万分之一比例）：裁剪子区 [l, l+w] 映射满框 → 全图按比例放大后 clip 回框
  const srcRect = blipFill ? directChild(blipFill, 'a:srcRect') : null
  let imgX = box.x
  let imgY = box.y
  let imgW = box.w
  let imgH = box.h
  let clipId: string | undefined
  if (srcRect) {
    const l = parseInt(attr(srcRect, 'l') ?? '0', 10) / 100000
    const t = parseInt(attr(srcRect, 't') ?? '0', 10) / 100000
    const r = parseInt(attr(srcRect, 'r') ?? '0', 10) / 100000
    const b = parseInt(attr(srcRect, 'b') ?? '0', 10) / 100000
    if (l || t || r || b) {
      const sw = Math.max(0.00001, 1 - l - r)
      const sh = Math.max(0.00001, 1 - t - b)
      imgW = box.w / sw
      imgH = box.h / sh
      imgX = box.x - l * imgW
      imgY = box.y - t * imgH
      clipId = ctx.uid('clip')
      const clipPath = svgEl('clipPath', { id: clipId })
      clipPath.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h }))
      ctx.defs.appendChild(clipPath)
    }
  }

  const g = svgEl('g', { transform: boxTransform(box) })
  const spPr = firstDescendant(node, 'p:spPr') ?? firstDescendant(node, 'a:spPr')
  const filter = registerShadowFilter(spPr ? directChild(spPr, 'a:effectLst') : null, ctx)
  g.appendChild(svgEl('image', {
    x: imgX, y: imgY, width: imgW, height: imgH, href: src,
    preserveAspectRatio: 'none', filter,
    'clip-path': clipId ? `url(#${clipId})` : undefined,
  }))

  const ln = spPr ? directChild(spPr, 'a:ln') : null
  if (ln && !directChild(ln, 'a:noFill')) {
    g.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h, fill: 'none', ...lnOf(ln, ctx) }))
  }

  // 音视频：海报帧 + 播放标记（预览不播放）
  const nvPr = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:nvPr')
  const isMedia = Boolean(nvPr && (directChild(nvPr, 'a:videoFile') || directChild(nvPr, 'a:audioFile')))
  if (isMedia) {
    const r = Math.min(box.w, box.h) * 0.15
    const cx = box.x + box.w / 2
    const cy = box.y + box.h / 2
    const s = r * 0.7
    g.appendChild(svgEl('circle', { cx, cy, r, fill: '#00000066' }))
    g.appendChild(svgEl('path', {
      d: `M ${cx - s * 0.4} ${cy - s * 0.6} L ${cx + s * 0.7} ${cy} L ${cx - s * 0.4} ${cy + s * 0.6} Z`,
      fill: '#FFFFFF',
    }))
  }
  return g
}
```

- [ ] **Step 4: 运行确认通过**

```bash
npx vitest run src/core/preview/picture.test.ts
```

预期：PASS（4 用例）。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/preview/picture.ts packages/pptx/src/core/preview/picture.test.ts
git commit -m "feat(preview): 图片/音视频 SVG 渲染（srcRect 裁剪/阴影/边框/播放标记）"
```

---

### Task 6: buildCellMatrix 增量 + preview/table.ts

**Files:**
- Modify: `packages/pptx/src/core/import/elements/table.ts`（`buildCellMatrix` 返回值增加非破坏性字段 `tcEls`）
- Create: `packages/pptx/src/core/preview/table.ts`
- Test: `packages/pptx/src/core/preview/table.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './svg.test'
import { renderTable } from './table'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

const TBL = `<p:graphicFrame>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
    <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
    <a:tr h="1000000">
      <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>表头</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:tcPr></a:tc>
    </a:tr>
    <a:tr h="1000000">
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
    </a:tr>
  </a:tbl></a:graphicData></a:graphic>
</p:graphicFrame>`

describe('preview/renderTable', () => {
  it('合并格 rect 跨格 + 单元格富文本 foreignObject + 默认边框', async () => {
    const ctx = makeCtx()
    const g = (await renderTable(wrap(TBL), ctx))!
    const rects = Array.from(g.querySelectorAll('rect'))
    // 3 个单元格（合并格跨 2 列算 1 个 rect）
    expect(rects).toHaveLength(3)
    // 合并格：x=0, w=200（4000000/9525/2*2 = 420.2 总宽的一半…精确值见断言）
    expect(Number(rects[0].getAttribute('width'))).toBeCloseTo((2000000 / 9525) * 2, 1)
    expect(Number(rects[0].getAttribute('height'))).toBeCloseTo(1000000 / 9525, 1)
    expect(rects[0].getAttribute('fill')).toBe('#EEEEEE')
    expect(rects[0].getAttribute('stroke')).toBe('#BFBFBF')
    // 富文本进入 foreignObject（复用 txBodyToHTML：加粗/字号/颜色生效）
    const fo = g.querySelector('foreignObject')!
    expect(fo.firstElementChild!.innerHTML).toContain('<strong>') // b=1 → bold（txBodyToHTML 输出 bold 加粗形式）
    expect(fo.firstElementChild!.innerHTML).toContain('表头')
  })
})
```

> 断言注记：`txBodyToHTML` 的加粗输出形式以一期实现为准（`<span style="font-weight:bold">`）——若断言 `<strong>` 不符，改为断言 `toContain('font-weight:bold')`。宽高精确值：总宽 4000000/9525 ≈ 419.95px，单列 ≈ 209.97px，合并宽 ≈ 419.95px；`toBeCloseTo(..., 1)` 按此核对。

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/table.test.ts
```

- [ ] **Step 3a: buildCellMatrix 增量**（`core/import/elements/table.ts`）

返回类型与实现各改一处（保持既有调用方 `parseTableEl` 不变）：

```typescript
/** a:tbl → cells 矩阵（合并原点带 colspan/rowspan，被并格为 null）；tcEls 为对应源 a:tc 元素矩阵（预览富文本用） */
export function buildCellMatrix(tbl: Element, theme: PptxTheme): {
  cells: Array<Array<TableCell | null>>
  tcEls: Array<Array<Element | null>>
} {
  const trs = directChildren(tbl, 'a:tr')
  const cells: Array<Array<TableCell | null>> = []
  const tcEls: Array<Array<Element | null>> = []
  const occupied: Array<Array<boolean>> = []
  for (let r = 0; r < trs.length; r += 1) {
    cells.push([])
    tcEls.push([])
    occupied.push([])
  }
  // ...原逻辑不变，仅在以下两处同步维护 tcEls：
  // 1) hMerge/vMerge 分支：tcEls[r][c] = null
  // 2) 普通格：tcEls[r][c] = tc；覆盖区循环内 tcEls[rr][cc] = null（与 cells 同步）
  ...
  return { cells, tcEls }
}
```

- [ ] **Step 3b: 实现** `packages/pptx/src/core/preview/table.ts`

```typescript
/** p:graphicFrame(a:tbl) → 表格 SVG：合并格 rect + 复用 txBodyToHTML 的富文本单元格。
 * 边框简化：统一细边框（#BFBFBF 0.75px），OOXML 逐边 lnL/lnR/lnT/lnB 定制不还原（记录于已知取舍）。 */
import { attr, directChildren, firstDescendant } from '../import/xml'
import { buildCellMatrix } from '../import/elements/table'
import { renderText } from './text'
import { svgEl, geomOf, type PreviewCtx } from './svg'

export async function renderTable(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const box = geomOf(node, ctx)
  if (!box) return null
  const tbl = firstDescendant(node, 'a:tbl')
  if (!tbl) return null

  const gridCols = directChildren(firstDescendant(node, 'a:tblGrid') ?? tbl, 'a:gridCol')
  const colSizesRaw = gridCols.map((c) => parseInt(attr(c, 'w') ?? '0', 10))
  const colTotal = colSizesRaw.reduce((a, b) => a + b, 0) || 1
  const trs = directChildren(tbl, 'a:tr')
  const rowSizesRaw = trs.map((tr) => parseInt(attr(tr, 'h') ?? '0', 10))
  const rowTotal = rowSizesRaw.reduce((a, b) => a + b, 0) || 1

  // 每列/行相对偏移（px）
  const colX: number[] = []
  let acc = 0
  for (const w of colSizesRaw) {
    colX.push(acc)
    acc += (w / colTotal) * box.w
  }
  const rowY: number[] = []
  acc = 0
  for (const h of rowSizesRaw) {
    rowY.push(acc)
    acc += (h / rowTotal) * box.h
  }

  const { cells, tcEls } = buildCellMatrix(tbl, ctx.theme)
  const g = svgEl('g')
  for (let r = 0; r < cells.length; r += 1) {
    for (let c = 0; c < cells[r].length; c += 1) {
      const cell = cells[r][c]
      if (!cell) continue
      const colspan = cell.colspan ?? 1
      const rowspan = cell.rowspan ?? 1
      const x = box.x + (colX[c] ?? 0)
      const y = box.y + (rowY[r] ?? 0)
      const w = (colX[c + colspan] ?? box.w) - (colX[c] ?? 0)
      const h = (rowY[r + rowspan] ?? box.h) - (rowY[r] ?? 0)
      g.appendChild(svgEl('rect', {
        x, y, width: w, height: h,
        fill: cell.style?.backcolor ?? 'none',
        stroke: '#BFBFBF', 'stroke-width': 0.75,
      }))
      const tc = tcEls[r][c]
      const txBody = tc ? firstDescendant(tc, 'a:txBody') : null
      if (txBody) {
        const text = await renderText(txBody, ctx, { x, y, w, h, rot: 0, flipH: false, flipV: false })
        if (text) g.appendChild(text)
      }
    }
  }
  return g
}
```

- [ ] **Step 4: 运行确认通过 + 一期表格回归**

```bash
npx vitest run src/core/preview/table.test.ts src/core/import/elements/table.test.ts
```

预期：全部 PASS。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/import/elements/table.ts packages/pptx/src/core/preview/table.ts packages/pptx/src/core/preview/table.test.ts
git commit -m "feat(preview): 表格 SVG 渲染（合并格/富文本单元格；buildCellMatrix 增量返回源元素）"
```

---

### Task 7: preview/chart.ts —— 图表直渲

**Files:**
- Create: `packages/pptx/src/core/preview/chart.ts`
- Test: `packages/pptx/src/core/preview/chart.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './svg.test'
import { renderChart } from './chart'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

/** 2 系列 × 3 类目的柱状图 graphicFrame（缓存数据内联在 chart part） */
function barFrame(): Element {
  const chartPart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:barChart>
    <c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列一</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
    <c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列二</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
  </c:barChart></c:plotArea></c:chart></c:chartSpace>`
  // relTarget 走 ctx.pkg —— 测试用桩包：仅实现 relTarget 返回 null 使 parseChartEl 走 missingChart？
  // 不行——parseChartEl 需要真实 chart part。用 makeCtx({pkg: stubPkg})，stub 实现 text/relTarget/rels：
  ...
}
```

> 实现注记：测试里的包桩用最简对象实现 `PptxPackage` 所需方法（`relTarget(partPath, rId) → 'chart1.xml'`、`text('chart1.xml') → chartPart`、`rels(partPath) → new Map()`），类型断言 `as never` 挂进 `makeCtx({ pkg: stub as never })`。完整测试文件以此桩展开：

```typescript
function stubPkg(chartPart: string): never {
  return {
    relTarget: async () => 'chart1.xml',
    text: async (p: string) => (p === 'chart1.xml' ? chartPart : null),
    rels: async () => new Map(),
    mediaDataUrl: async () => undefined,
  } as never
}

describe('preview/renderChart', () => {
  it('柱状图：每系列×类目一个 rect，标题渲染', async () => {
    const ctx = makeCtx({ pkg: stubPkg(BAR_CHART_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('rect').length).toBeGreaterThanOrEqual(6) // 2 系列 × 3 类目
    const title = g.querySelector('text')
    expect(title?.textContent).toBe('月度报表')
  })

  it('饼图：每数据项一个扇形 path', async () => {
    const ctx = makeCtx({ pkg: stubPkg(PIE_CHART_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    expect(g.querySelectorAll('path').length).toBe(3)
  })

  it('雷达图：降级占位框', async () => {
    const ctx = makeCtx({ pkg: stubPkg(RADAR_CHART_PART) })
    const g = (await renderChart(wrap(FRAME), ctx))!
    const placeholder = g.querySelector('rect[fill="#F2F2F2"]')
    expect(placeholder).not.toBeNull()
  })

  it('缺失图表 part → null 且计入报告', async () => {
    const ctx = makeCtx({ pkg: stubPkg('') })
    expect(await renderChart(wrap(FRAME), ctx)).toBeNull()
    expect(ctx.report.skipped.missingChart).toBe(1)
  })
})
```

`FRAME`（graphicFrame，c:chart 引用 rId1 + 图表标题「月度报表」在 chart part 的 `<c:title>` 里）与三个 chart part 常量（BAR：barChart 2 系列×3 类目；PIE：pieChart 3 数据项；RADAR：radarChart）按 Task 6/一期 `chart.test.ts` 的既有构造模式编写——一期 `core/import/elements/chart.test.ts` 已有同类 fixture，**复用其 XML 片段**。

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/chart.test.ts
```

- [ ] **Step 3: 实现** `packages/pptx/src/core/preview/chart.ts`

```typescript
/** p:graphicFrame(c:chart) → 图表 SVG：复用一期 parseChartEl（scale 恒 {1,1}，geom 即源画布 px），
 * 按 chartType 直渲 SVG 图元；radar 降级占位框（已知取舍）。 */
import { firstDescendant } from '../import/xml'
import { parseChartEl } from '../import/elements/chart'
import { IDENTITY_XFORM } from '../import/context'
import { svgEl, geomOf, type PreviewCtx } from './svg'
import type { ChartElement } from '../../types/slides'

const CHART_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

type BoxLike = { x: number; y: number; w: number; h: number }

export async function renderChart(node: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const box = geomOf(node, ctx)
  if (!box) return null
  const chart = await parseChartEl(node, ctx, IDENTITY_XFORM, ctx.pkg)
  if (!chart) return null
  const g = svgEl('g')
  if (chart.chartType === 'radar') {
    g.appendChild(svgEl('rect', { x: box.x, y: box.y, width: box.w, height: box.h, fill: '#F2F2F2', stroke: '#BFBFBF', 'stroke-width': 1 }))
    const t = svgEl('text', { x: box.x + box.w / 2, y: box.y + box.h / 2, 'text-anchor': 'middle', 'font-size': 12, fill: '#999999' })
    t.textContent = '雷达图暂不支持预览'
    g.appendChild(t)
    return g
  }
  drawAxis(g, box)
  const t = chart.chartType
  if (t === 'pie' || t === 'pie-doughnut') drawPie(g, chart, box, t === 'pie-doughnut')
  else if (t.startsWith('bar')) drawBars(g, chart, box, t.includes('horizontal'), t.includes('stack') || t === 'bar-percent')
  else if (t.startsWith('area')) drawArea(g, chart, box, t === 'area-stack')
  else if (t === 'scatter') drawScatter(g, chart, box)
  else drawLine(g, chart, box, t === 'line-stack', t === 'line-marker')
  if (chart.title) {
    const title = svgEl('text', { x: box.x + box.w / 2, y: box.y + 16, 'text-anchor': 'middle', 'font-size': 14, fill: '#333333' })
    title.textContent = chart.title
    g.appendChild(title)
  }
  return g
}

/** 绘图区：留出轴标签与标题空间 */
function plotRect(box: BoxLike): BoxLike {
  return { x: box.x + 36, y: box.y + 24, w: box.w - 44, h: box.h - 48 }
}

function valueMax(chart: ChartElement): number {
  return Math.max(0, ...chart.data.series.flatMap((s) => s.values)) * 1.1 || 1
}

function drawAxis(g: SVGElement, box: BoxLike): void {
  const p = plotRect(box)
  g.appendChild(svgEl('line', { x1: p.x, y1: p.y, x2: p.x, y2: p.y + p.h, stroke: '#999999', 'stroke-width': 1 }))
  g.appendChild(svgEl('line', { x1: p.x, y1: p.y + p.h, x2: p.x + p.w, y2: p.y + p.h, stroke: '#999999', 'stroke-width': 1 }))
}

function drawBars(g: SVGElement, chart: ChartElement, box: BoxLike, horizontal: boolean, stacked: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  const band = (horizontal ? p.h : p.w) / n
  const groupSize = stacked ? 1 : series.length
  for (let i = 0; i < n; i += 1) {
    let acc = 0
    for (let si = 0; si < series.length; si += 1) {
      const v = series[si].values[i] ?? 0
      const color = CHART_COLORS[si % CHART_COLORS.length]
      if (horizontal) {
        const seg = (v / maxV) * p.w
        const y = stacked
          ? p.y + band * i + (band * acc) / maxV
          : p.y + band * i + (band / groupSize) * si + (band / groupSize) * 0.1
        const h = stacked ? (band * v) / maxV : (band / groupSize) * 0.8
        g.appendChild(svgEl('rect', { x: p.x, y, width: seg, height: h, fill: color }))
      } else {
        const x = stacked
          ? p.x + band * i
          : p.x + band * i + (band / groupSize) * si + (band / groupSize) * 0.1
        const w = stacked ? band : (band / groupSize) * 0.8
        const h = stacked ? (p.h * v) / maxV : (p.h * v) / maxV
        const y = stacked ? p.y + p.h - (p.h * (acc + v)) / maxV : p.y + p.h - h
        g.appendChild(svgEl('rect', { x, y, width: w, height: h, fill: color }))
      }
      acc += v
    }
  }
}

function drawLine(g: SVGElement, chart: ChartElement, box: BoxLike, stacked: boolean, marker: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  for (let si = 0; si < series.length; si += 1) {
    const color = CHART_COLORS[si % CHART_COLORS.length]
    const pts: string[] = []
    for (let i = 0; i < n; i += 1) {
      const v = (series[si].values[i] ?? 0)
        + (stacked ? series.slice(0, si).reduce((s, sr) => s + (sr.values[i] ?? 0), 0) : 0)
      pts.push(`${p.x + (p.w * i) / Math.max(1, n - 1)},${p.y + p.h - (v / maxV) * p.h}`)
    }
    g.appendChild(svgEl('polyline', { points: pts.join(' '), fill: 'none', stroke: color, 'stroke-width': 2 }))
    if (marker) {
      for (const pt of pts) {
        const [x, y] = pt.split(',')
        g.appendChild(svgEl('circle', { cx: x, cy: y, r: 3, fill: color }))
      }
    }
  }
}

function drawArea(g: SVGElement, chart: ChartElement, box: BoxLike, stacked: boolean): void {
  const { labels, series } = chart.data
  const p = plotRect(box)
  const maxV = valueMax(chart)
  const n = Math.max(1, labels.length)
  const cum = new Array(n).fill(0)
  const rows = series.map((sr) => sr.values.map((v, i) => (stacked ? (cum[i] += v ?? 0) : v ?? 0)))
  const order = stacked ? [...rows.keys()].reverse() : [...rows.keys()]
  for (const si of order) {
    const pts = rows[si].map((v, i) => `${p.x + (p.w * i) / Math.max(1, n - 1)},${p.y + p.h - (v / maxV) * p.h}`)
    if (!pts.length) continue
    const d = `M ${pts[0]} L ${pts.join(' L ')} L ${p.x + p.w},${p.y + p.h} L ${p.x},${p.y + p.h} Z`
    g.appendChild(svgEl('path', { d, fill: CHART_COLORS[si % CHART_COLORS.length], 'fill-opacity': stacked ? 1 : 0.5 }))
  }
}

function drawScatter(g: SVGElement, chart: ChartElement, box: BoxLike): void {
  const p = plotRect(box)
  const xs = chart.data.labels.map(Number).filter(Number.isFinite)
  const ys = chart.data.series.flatMap((s) => s.values).filter(Number.isFinite)
  const minX = Math.min(0, ...xs)
  const maxX = Math.max(1, ...xs)
  const minY = Math.min(0, ...ys)
  const maxY = Math.max(1, ...ys) * 1.1
  for (let si = 0; si < chart.data.series.length; si += 1) {
    const color = CHART_COLORS[si % CHART_COLORS.length]
    const vals = chart.data.series[si].values
    xs.forEach((x, i) => {
      const y = vals[i]
      if (!Number.isFinite(y)) return
      g.appendChild(svgEl('circle', {
        cx: p.x + ((x - minX) / (maxX - minX || 1)) * p.w,
        cy: p.y + p.h - ((y - minY) / (maxY - minY || 1)) * p.h,
        r: 4, fill: color,
      }))
    })
  }
}

function drawPie(g: SVGElement, chart: ChartElement, box: BoxLike, doughnut: boolean): void {
  const series = chart.data.series[0]
  if (!series) return
  const total = series.values.reduce((a, b) => a + b, 0) || 1
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const r = Math.min(box.w, box.h) * 0.4
  const rIn = doughnut ? r * 0.6 : 0
  let a0 = -Math.PI / 2
  for (let i = 0; i < series.values.length; i += 1) {
    const a1 = a0 + ((series.values[i] ?? 0) / total) * Math.PI * 2
    g.appendChild(svgEl('path', { d: slicePath(cx, cy, r, rIn, a0, a1), fill: CHART_COLORS[i % CHART_COLORS.length] }))
    a0 = a1
  }
}

/** 扇环路径：外弧顺时针 + 内弧逆时针（rIn=0 时为扇形） */
function slicePath(cx: number, cy: number, r: number, rIn: number, a0: number, a1: number): string {
  const large = a1 - a0 > Math.PI ? 1 : 0
  const x0 = cx + r * Math.cos(a0)
  const y0 = cy + r * Math.sin(a0)
  const x1 = cx + r * Math.cos(a1)
  const y1 = cy + r * Math.sin(a1)
  if (rIn <= 0) return `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`
  const ix0 = cx + rIn * Math.cos(a1)
  const iy0 = cy + rIn * Math.sin(a1)
  const ix1 = cx + rIn * Math.cos(a0)
  const iy1 = cy + rIn * Math.sin(a0)
  return `M ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} L ${ix0} ${iy0} A ${rIn} ${rIn} 0 ${large} 0 ${ix1} ${iy1} Z`
}
```

- [ ] **Step 4: 运行确认通过**

```bash
npx vitest run src/core/preview/chart.test.ts
```

预期：PASS（4 用例）。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/preview/chart.ts packages/pptx/src/core/preview/chart.test.ts
git commit -m "feat(preview): 图表 SVG 直渲（bar/line/pie/doughnut/area/scatter；radar 占位）"
```

---

### Task 8: preview/group.ts + dispatch.ts + slide.ts

**Files:**
- Create: `packages/pptx/src/core/preview/group.ts`、`packages/pptx/src/core/preview/dispatch.ts`、`packages/pptx/src/core/preview/slide.ts`
- Test: `packages/pptx/src/core/preview/group.test.ts`、`packages/pptx/src/core/preview/slide.test.ts`

- [ ] **Step 1: 写失败测试**

`group.test.ts`：

```typescript
import { describe, expect, it } from 'vitest'
import { parseXML } from '../import/xml'
import { makeCtx } from './svg.test'
import { renderGroup } from './group'

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const wrap = (inner: string) => parseXML(`<root ${NS}>${inner}</root>`).documentElement.firstElementChild as Element

const GRP = `<p:grpSp>
  <p:grpSpPr>
    <a:xfrm rot="5400000">
      <a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/>
      <a:chOff x="0" y="0"/><a:chExt cx="100" cy="50"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>
</p:grpSp>`

describe('preview/renderGroup', () => {
  it('SVG transform 表达 child→parent 映射 + 组合旋转（无需降级）', async () => {
    const ctx = makeCtx()
    const g = (await renderGroup(wrap(GRP), ctx))!
    const tf = g.getAttribute('transform')!
    expect(tf).toContain('translate(100,100)')
    expect(tf).toContain('scale(19,19)')      // 1905000/9525/100 与 952500/9525/50 均为 19
    expect(tf).toContain('translate(-0,-0)')
    expect(tf).toContain('rotate(90,')        // 组合旋转在预览中可完整还原
    expect(ctx.report.skipped.groupRotation).toBeUndefined() // 不计入降级
    expect(g.querySelector('path')).not.toBeNull()
  })

  it('嵌套组合递归展开', async () => {
    const ctx = makeCtx()
    const inner = GRP.replace('p:grpSp>', 'p:grpSp>') // 原样
    const outer = wrap(`<p:grpSp><p:grpSpPr><a:xfrm>
      <a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/>
    </a:xfrm></p:grpSpPr>${GRP}</p:grpSp>`)
    const g = (await renderGroup(outer, ctx))!
    expect(g.querySelectorAll('g').length).toBeGreaterThanOrEqual(1)
  })
})
```

`slide.test.ts`：

```typescript
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from '../import/package'
import { renderSlide } from './slide'

/** 最小 pptx：master(红底矩形装饰 + title 占位符) + layout + slide(引用占位符) */
async function buildPptx(): Promise<File> {
  const zip = new JSZip()
  // [Content_Types]/_rels/.rels/presentation.xml 等骨架与 core/import/index.test.ts 的 buildMinimalPptx 相同
  //（实现时复制该 helper 骨架），另加：
  // - ppt/slideMasters/slideMaster1.xml：p:bg solidFill lt2 + 一个非占位符装饰 p:sp（红矩形）+ title 占位符 p:sp（有 xfrm）
  // - ppt/slideMasters/_rels/slideMaster1.xml.rels → slideLayout1 + theme1
  // - ppt/slideLayouts/slideLayout1.xml：空 spTree
  // - ppt/slideLayouts/_rels/slideLayout1.xml.rels → slideMaster1
  // - ppt/theme/theme1.xml：clrScheme（accent1 等）+ fontScheme
  // - ppt/slides/slide1.xml：一个无 xfrm 的 title 占位符 p:sp（继承 master 位置）
  //   rels → slideLayout1
  // （XML 模板按一期 master.test.ts / theme.test.ts 的既有写法拼装）
  const blob = await zip.generateAsync({ type: 'blob' })
  return new File([blob], 'preview.pptx')
}

describe('preview/renderSlide', () => {
  it('页面组装：viewBox 源画布尺寸 / 母版背景与装饰合入 / 占位符位置继承', async () => {
    const pkg = await PptxPackage.load(await buildPptx())
    const svg = (await renderSlide(pkg, 'ppt/slides/slide1.xml', 12192000, 6858000, 0, { skipped: {} }))!
    expect(svg.getAttribute('viewBox')).toBe('0 0 1280 720')
    expect(svg.getAttribute('width')).toBe('1280')
    // 背景 rect（母版 bg）
    expect(svg.querySelector('rect[fill]')).not.toBeNull()
    // 装饰元素（母版红矩形）与占位符文本（继承 master 位置的 title）都在页内
    expect(svg.querySelector('path, rect')).not.toBeNull()
    expect(svg.querySelector('foreignObject')).not.toBeNull()
  })
})
```

> 实现注记：`slide.test.ts` 的 zip 组装代码是本任务的主要测试工作量——**必须完整写出全部 XML**（复制一期 `index.test.ts`/`master.test.ts` 中的 presentation/master/layout/theme XML 模板并按上述清单调整），不得留占位符。断言按实际模板微调，但 viewBox/背景/装饰/foreignObject 四项必须有。

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/group.test.ts src/core/preview/slide.test.ts
```

- [ ] **Step 3a: 实现** `packages/pptx/src/core/preview/group.ts`

```typescript
/** p:grpSp → <g>：SVG transform 完整表达 child→parent 映射与组合旋转（预览无需一期的 groupRotation 降级）。 */
import { attr, directChild, firstDescendant } from '../import/xml'
import { emu2pxF, svgEl, type PreviewCtx } from './svg'
import { renderSpTreeNode } from './dispatch'

export async function renderGroup(grpSp: Element, ctx: PreviewCtx): Promise<SVGElement | null> {
  const xfrm = firstDescendant(grpSp, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const chOff = xfrm ? directChild(xfrm, 'a:chOff') : null
  const chExt = xfrm ? directChild(xfrm, 'a:chExt') : null
  const g = svgEl('g')
  if (off && ext && chOff && chExt) {
    const ox = emu2pxF(parseInt(attr(off, 'x') ?? '0', 10))
    const oy = emu2pxF(parseInt(attr(off, 'y') ?? '0', 10))
    const cx = emu2pxF(parseInt(attr(ext, 'cx') ?? '0', 10))
    const cy = emu2pxF(parseInt(attr(ext, 'cy') ?? '0', 10))
    const chx = emu2pxF(parseInt(attr(chOff, 'x') ?? '0', 10))
    const chy = emu2pxF(parseInt(attr(chOff, 'y') ?? '0', 10))
    const chcx = emu2pxF(parseInt(attr(chExt, 'cx') ?? '0', 10)) || 1
    const chcy = emu2pxF(parseInt(attr(chExt, 'cy') ?? '0', 10)) || 1
    // child→parent：T(off)·S(ext/chExt)·T(-chOff)；旋转置于链尾 = 对已映射的父空间内容绕组合中心旋转
    const parts = [`translate(${ox},${oy})`, `scale(${cx / chcx},${cy / chcy})`, `translate(${-chx},${-chy})`]
    const rot = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000
    if (rot) parts.push(`rotate(${rot},${ox + cx / 2},${oy + cy / 2})`)
    g.setAttribute('transform', parts.join(' '))
  }
  for (const child of Array.from(grpSp.children)) {
    g.append(...(await renderSpTreeNode(child, ctx)))
  }
  return g.children.length ? g : null
}
```

- [ ] **Step 3b: 实现** `packages/pptx/src/core/preview/dispatch.ts`

```typescript
/** spTree 单节点分发器：p:sp/p:cxnSp/p:pic/p:graphicFrame/p:grpSp → 对应渲染器；未知节点跳过。
 * 独立成文件避免 slide ↔ group 循环引用。 */
import { firstDescendant } from '../import/xml'
import { addSkipped } from '../import/context'
import type { PreviewCtx } from './svg'
import { renderShape } from './shape'
import { renderPicture } from './picture'
import { renderTable } from './table'
import { renderChart } from './chart'
import { renderGroup } from './group'

export async function renderSpTreeNode(node: Element, ctx: PreviewCtx): Promise<SVGElement[]> {
  try {
    switch (node.nodeName) {
      case 'p:sp':
      case 'p:cxnSp': {
        const el = await renderShape(node, ctx)
        return el ? [el] : []
      }
      case 'p:pic': {
        const el = await renderPicture(node, ctx)
        return el ? [el] : []
      }
      case 'p:graphicFrame': {
        const el = firstDescendant(node, 'a:tbl')
          ? await renderTable(node, ctx)
          : await renderChart(node, ctx)
        return el ? [el] : []
      }
      case 'p:grpSp': {
        const el = await renderGroup(node, ctx)
        return el ? [el] : []
      }
      default:
        return []
    }
  } catch (e) {
    // 单元素失败不拖垮整页
    console.warn('[pptx-preview] 元素渲染失败:', node.nodeName, e)
    addSkipped(ctx.report, 'elementParseFailed')
    return []
  }
}
```

- [ ] **Step 3c: 实现** `packages/pptx/src/core/preview/slide.ts`

```typescript
/** 单页渲染：背景（slide→layout→master→lt1）+ 版式/母版装饰与占位符表 + slide 元素 → <svg>。 */
import { attr, firstDescendant, parseXML } from '../import/xml'
import { findAncestry, collectPlaceholders, parseBackgroundFill } from '../import/master'
import { parseThemeForMaster } from '../import/theme'
import { addSkipped } from '../import/context'
import type { PptxPackage } from '../import/package'
import type { ImportReport, Gradient } from '../../types/slides'
import { SVG_NS, svgEl, emu2pxF, type PreviewCtx } from './svg'
import { renderSpTreeNode } from './dispatch'

const ELEMENT_NODES = ['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame', 'p:grpSp']

/** 背景绘制节点：solid/grad/blip；当前层级无可用背景返回 undefined（继续向母版回退） */
async function backgroundNode(
  pkg: PptxPackage, partPath: string, theme: Awaited<ReturnType<typeof parseThemeForMaster>>,
  width: number, height: number, uid: (prefix: string) => string,
): Promise<SVGElement | undefined> {
  const xml = await pkg.text(partPath)
  if (!xml) return undefined
  const bg = firstDescendant(parseXML(xml).documentElement, 'p:bg')
  const fill = parseBackgroundFill(bg, theme)
  if (!fill) return undefined
  if (fill.type === 'solid' && fill.color) {
    return svgEl('rect', { x: 0, y: 0, width, height, fill: fill.color })
  }
  if (fill.type === 'gradient' && fill.gradient) {
    const id = uid('bg')
    const grad = svgEl('linearGradient', { id, gradientUnits: 'objectBoundingBox', ...gradientCoords(fill.gradient) })
    for (const stop of fill.gradient.colors) {
      grad.appendChild(svgEl('stop', { offset: stop.pos, 'stop-color': stop.color }))
    }
    return svgEl('g', { 'data-defs': id }) as never // 占位——实际实现见下
  }
  return undefined
}
```

> 上面的 `backgroundNode` 渐变分支需要往 defs 里登记——但该函数拿不到 defs。**最终实现签名**调整为把 defs 与 uid 都传进来（下方完整代码为准）：

```typescript
/** 单页渲染：失败抛出由调用方（previewPPTXDetailed）计入 slideParseFailed */
export async function renderSlide(
  pkg: PptxPackage,
  slidePath: string,
  srcW: number,
  srcH: number,
  pageIdx: number,
  report: ImportReport,
): Promise<SVGSVGElement | null> {
  const xml = await pkg.text(slidePath)
  if (!xml) return null
  const doc = parseXML(xml)
  const spTree = firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')

  const { layoutPath, masterPath } = await findAncestry(pkg, slidePath)
  const theme = await parseThemeForMaster(pkg, masterPath)
  const width = emu2pxF(srcW)
  const height = emu2pxF(srcH)

  const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement
  svg.setAttribute('xmlns', SVG_NS)
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`)
  svg.setAttribute('width', String(width))
  svg.setAttribute('height', String(height))
  let uidN = 0
  const uid = (prefix: string) => `${prefix}-p${pageIdx}-${uidN++}`
  const defs = svgEl('defs')
  svg.appendChild(defs)
  const placeholders = new Map<string, { x: number; y: number; w: number; h: number }>()

  const makeCtx = (partPath: string): PreviewCtx => ({
    pkg, partPath, theme, report,
    scale: { x: 1, y: 1 },
    placeholders,
    defs,
    uid,
  })

  // 1) 背景：slide → layout → master 逐级回退；blip 源缺失继续回退；最终 lt1 兜底
  let bgNode: SVGElement | undefined
  for (const partPath of [slidePath, layoutPath, masterPath]) {
    if (!partPath) continue
    const partXml = await pkg.text(partPath)
    if (!partXml) continue
    const bg = firstDescendant(parseXML(partXml).documentElement, 'p:bg')
    const fill = parseBackgroundFill(bg, theme)
    if (!fill) continue
    if (fill.type === 'solid' && fill.color) {
      bgNode = svgEl('rect', { x: 0, y: 0, width, height, fill: fill.color })
    } else if (fill.type === 'gradient' && fill.gradient) {
      const id = uid('bg')
      const ang = ((fill.gradient.rotate ?? 0) * Math.PI) / 180
      const grad = svgEl('linearGradient', {
        id, gradientUnits: 'objectBoundingBox',
        x1: 0.5 - Math.cos(ang) / 2, y1: 0.5 - Math.sin(ang) / 2,
        x2: 0.5 + Math.cos(ang) / 2, y2: 0.5 + Math.sin(ang) / 2,
      })
      for (const stop of fill.gradient.colors) {
        grad.appendChild(svgEl('stop', { offset: stop.pos, 'stop-color': stop.color }))
      }
      defs.appendChild(grad)
      bgNode = svgEl('rect', { x: 0, y: 0, width, height, fill: `url(#${id})` })
    } else if (fill.type === 'image' && bg) {
      const embedId = attr(firstDescendant(bg, 'a:blip'), 'r:embed')
      const target = embedId ? await pkg.relTarget(partPath, embedId) : null
      const src = target ? await pkg.mediaDataUrl(target) : undefined
      if (src) {
        // cover：xMidYMid slice
        bgNode = svgEl('image', { x: 0, y: 0, width, height, href: src, preserveAspectRatio: 'xMidYMid slice' })
      }
    }
    if (bgNode) break
  }
  svg.appendChild(bgNode ?? svgEl('rect', { x: 0, y: 0, width, height, fill: theme.schemeColors.lt1 ?? '#FFFFFF' }))

  // 2) 版式/母版：占位符位置表 + 非占位符装饰元素（rels 按 partPath 解析）
  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath) continue
    const partXml = await pkg.text(partPath)
    if (!partXml) continue
    const partDoc = parseXML(partXml)
    const partSpTree = firstDescendant(firstDescendant(partDoc.documentElement, 'p:cSld') ?? partDoc.documentElement, 'p:spTree')
    if (!partSpTree) continue
    for (const [k, v] of collectPlaceholders(partSpTree)) {
      if (!placeholders.has(k)) placeholders.set(k, v)
    }
    const partCtx = makeCtx(partPath)
    for (const child of Array.from(partSpTree.children)) {
      if (!ELEMENT_NODES.includes(child.nodeName)) continue
      if (firstDescendant(child, 'p:ph')) continue // 占位符本体不作为装饰渲染
      svg.append(...(await renderSpTreeNode(child, partCtx)))
    }
  }

  // 3) slide 自身元素
  if (spTree) {
    const slideCtx = makeCtx(slidePath)
    for (const child of Array.from(spTree.children)) {
      if (!ELEMENT_NODES.includes(child.nodeName)) continue
      svg.append(...(await renderSpTreeNode(child, slideCtx)))
    }
  }
  return svg
}
```

（若最终不需要 `gradientCoords`/`Gradient` 导入，删除未用项保持 oxlint 干净；`import type { Gradient }` 仅在需要时保留。）

- [ ] **Step 4: 运行确认通过**

```bash
npx vitest run src/core/preview/group.test.ts src/core/preview/slide.test.ts
```

预期：PASS。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/preview/
git commit -m "feat(preview): 组合/分发器/单页渲染（SVG transform 组合旋转、背景回退链、装饰合入）"
```

---

### Task 9: import/index.ts 提取 listSlidePaths + preview/index.ts 入口

**Files:**
- Modify: `packages/pptx/src/core/import/index.ts`（提取共享 `listSlidePaths`，`importPPTXDetailed` 改为调用它）
- Create: `packages/pptx/src/core/preview/index.ts`
- Test: `packages/pptx/src/core/preview/index.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { previewPPTXDetailed, previewPPTX } from './index'

/** 复用 core/import/index.test.ts 的 buildMinimalPptx（文本框 + 圆角矩形两元素，无版式/母版）——复制其实现 */
// （此处粘贴该 helper 全文，改名不改内容）

describe('previewPPTX', () => {
  it('最小包：1 页 SVG，viewBox 源画布 px，元素齐全，报告无跳过', async () => {
    const { pages, report } = await previewPPTXDetailed(await buildMinimalPptx())
    expect(pages).toHaveLength(1)
    expect(pages[0].getAttribute('viewBox')).toBe('0 0 1280 720')
    // 文本框（foreignObject）+ 形状（嵌套 svg path）
    expect(pages[0].querySelector('foreignObject')).not.toBeNull()
    expect(pages[0].querySelector('path')).not.toBeNull()
    expect(report.skipped).toEqual({})
    // previewPPTX 兼容 API
    await expect(previewPPTX(await buildMinimalPptx())).resolves.toHaveLength(1)
  })

  it('坏页容错：畸形第二页计入 slideParseFailed 不中断', async () => {
    const { pages, report } = await previewPPTXDetailed(await buildTwoSlidesOneBroken())
    expect(pages).toHaveLength(1)
    expect(report.skipped.slideParseFailed).toBe(1)
  })

  it('非 pptx → 明确中文错误', async () => {
    const zip = new JSZip()
    const file = new File([await zip.generateAsync({ type: 'blob' })], 'empty.pptx')
    await expect(previewPPTXDetailed(file)).rejects.toThrow('缺少 presentation.xml')
  })
})
```

- [ ] **Step 2: 运行确认失败**

```bash
npx vitest run src/core/preview/index.test.ts
```

- [ ] **Step 3a: 提取 `listSlidePaths`**（`core/import/index.ts`）

把 `importPPTXDetailed` 开头「presentation.xml 读取 + sldSz + 页面顺序」整段（含 `slidePaths` 兜底逻辑）提取为导出函数，并让 `importPPTXDetailed` 调用它：

```typescript
/** 页面路径与源画布尺寸：sldIdLst → rels；缺失时按 slideN.xml 编号兜底 */
export async function listSlidePaths(
  pkg: PptxPackage,
): Promise<{ paths: string[]; srcW: number; srcH: number; ratio: number }> {
  const presXml = await pkg.text('ppt/presentation.xml')
  if (!presXml) throw new Error('不是有效的 PPTX 文件（缺少 presentation.xml）')
  const root = parseXML(presXml).documentElement

  const sldSz = firstDescendant(root, 'p:sldSz')
  const srcW = parseInt(attr(sldSz, 'cx') ?? '12192000', 10)
  const srcH = parseInt(attr(sldSz, 'cy') ?? '6858000', 10)
  const ratio = srcW / srcH || 16 / 9

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
```

`importPPTXDetailed` 改为：

```typescript
export async function importPPTXDetailed(file: File): Promise<ImportResult> {
  const pkg = await PptxPackage.load(file)
  const { paths, srcW, srcH, ratio } = await listSlidePaths(pkg)
  const scale = { x: 1280 / (srcW / 9525), y: (1280 / ratio) / (srcH / 9525) }
  // ……其余保持原实现（report/slides 循环原样）
}
```

- [ ] **Step 3b: 实现** `packages/pptx/src/core/preview/index.ts`

```typescript
/** 预览入口：pptx → 每页一个 <svg>（OOXML 语义直渲，不做编辑器模型映射取舍） */
import { PptxPackage } from '../import/package'
import { listSlidePaths } from '../import'
import { addSkipped } from '../import/context'
import type { ImportReport } from '../../types/slides'
import { renderSlide } from './slide'

export interface PreviewResult {
  pages: SVGSVGElement[]
  report: ImportReport
}

/** pptx → 页面 SVG 列表 + 兼容性报告（单页失败计入报告不中断） */
export async function previewPPTXDetailed(file: File): Promise<PreviewResult> {
  const pkg = await PptxPackage.load(file)
  const { paths, srcW, srcH } = await listSlidePaths(pkg)
  const report: ImportReport = { skipped: {} }
  const pages: SVGSVGElement[] = []
  for (let i = 0; i < paths.length; i += 1) {
    try {
      const page = await renderSlide(pkg, paths[i], srcW, srcH, i, report)
      if (page) pages.push(page)
      else addSkipped(report, 'slideParseFailed')
    } catch (e) {
      console.warn('[pptx-preview] 页面渲染失败:', paths[i], e)
      addSkipped(report, 'slideParseFailed')
    }
  }
  if (!pages.length) throw new Error('PPTX 中没有可预览的幻灯片')
  return { pages, report }
}

/** 设计文档 API：仅返回页面列表 */
export async function previewPPTX(file: File): Promise<SVGSVGElement[]> {
  return (await previewPPTXDetailed(file)).pages
}
```

- [ ] **Step 4: 运行确认通过 + 一期导入回归（listSlidePaths 重构不破坏）**

```bash
npx vitest run src/core/preview/index.test.ts src/core/import
```

预期：全部 PASS。

- [ ] **Step 5: 提交**

```bash
git add packages/pptx/src/core/import/index.ts packages/pptx/src/core/preview/index.ts packages/pptx/src/core/preview/index.test.ts
git commit -m "feat(preview): previewPPTX 入口（listSlidePaths 与导入共享）"
```

---

### Task 10: ImportDialog「先预览再导入」UI

无自动化测试（项目现状：无组件级测试基建），验证 = `npx tsc -b` + `pnpm build` + 手工冒烟（见 Step 4）。

**Files:**
- Modify: `packages/pptx/src/components/dialogs/ImportDialog.tsx`

- [ ] **Step 1: 实现**

```tsx
// 导入对话框：PPTX（先预览再导入）/ JSON
import { useEffect, useRef, useState } from 'react'
import { Modal } from './ModalHost'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { parseJSONFile } from '../../core/export/json'
import { importPPTXDetailed } from '../../core/import'
import { previewPPTXDetailed } from '../../core/preview'

/** 兼容性报告 skip 种类 → 用户可读文案 */
const SKIP_LABELS: Record<string, string> = {
  missingImage: '缺失图片',
  missingChart: '缺失图表数据',
  unknownChart: '未识别图表',
  elementParseFailed: '无法解析的元素',
  slideParseFailed: '无法解析的页面',
  smartartFallback: 'SmartArt 已转图片',
  groupParseFailed: '组合解析失败',
  groupRotation: '组合旋转未还原',
  missingPlaceholder: '占位符缺失',
}

export function ImportDialog() {
  const toast = useToastStore.getState().toast
  const closeModal = useUIStore.getState().closeModal
  const [preview, setPreview] = useState<{ file: File; pages: SVGSVGElement[]; report: { skipped: Record<string, number> } } | null>(null)
  const gridRef = useRef<HTMLDivElement>(null)

  // SVG DOM 节点非 React 元素，直接挂载到网格容器
  useEffect(() => {
    const grid = gridRef.current
    if (!grid || !preview) return
    grid.innerHTML = ''
    preview.pages.forEach((page, i) => {
      const wrap = document.createElement('div')
      wrap.className = 'relative overflow-hidden rounded-lg border border-gray-200 bg-white'
      page.style.width = '100%'
      page.style.height = 'auto'
      page.style.display = 'block'
      wrap.appendChild(page)
      const badge = document.createElement('div')
      badge.className = 'absolute right-1 top-1 rounded bg-black/50 px-1 text-[10px] text-white'
      badge.textContent = String(i + 1)
      wrap.appendChild(badge)
      grid.appendChild(wrap)
    })
  }, [preview])

  const applyImport = async (file: File) => {
    toast('正在解析 PPTX…')
    const { presentation: pres, report } = await importPPTXDetailed(file)
    useEditorStore.getState().pushHistory()
    useEditorStore.getState().replacePresentation(pres)
    const items = Object.entries(report.skipped).map(([k, v]) => `${SKIP_LABELS[k] ?? k} ×${v}`)
    toast(items.length
      ? `已导入 PPTX（${pres.slides.length} 页）；部分内容未完整还原：${items.join('、')}`
      : `已导入 PPTX（${pres.slides.length} 页）`, 'success')
    useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
    closeModal()
  }

  const handleFile = async (file: File) => {
    try {
      if (file.name.toLowerCase().endsWith('.json') || file.name.toLowerCase().endsWith('.efppt.json')) {
        const text = await file.text()
        const pres = parseJSONFile(text)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        toast(`已导入 JSON（${pres.slides.length} 页）`, 'success')
        useEditorStore.setState({ slideIndex: 0, selectedIds: [] })
        closeModal()
      } else if (file.name.toLowerCase().endsWith('.pptx')) {
        // 第二期：先预览再导入
        toast('正在生成预览…')
        const { pages, report } = await previewPPTXDetailed(file)
        setPreview({ file, pages, report })
      } else {
        toast('请选择 .pptx 或 .json 文件', 'error')
        return
      }
    } catch (error) {
      toast(`导入失败：${error instanceof Error ? error.message : '未知错误'}`, 'error')
    }
  }

  if (preview) {
    const items = Object.entries(preview.report.skipped).map(([k, v]) => `${SKIP_LABELS[k] ?? k} ×${v}`)
    return (
      <Modal title="导入预览">
        <div ref={gridRef} className="grid max-h-[60vh] grid-cols-3 gap-3 overflow-auto" />
        {items.length > 0 && (
          <div className="mt-2 text-xs text-gray-500">部分内容预览/还原存在降级：{items.join('、')}</div>
        )}
        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            className="rounded-lg border border-gray-300 px-4 py-2 text-sm text-gray-700 hover:bg-gray-50"
            onClick={() => setPreview(null)}
          >
            返回
          </button>
          <button
            type="button"
            className="rounded-lg bg-[#d14424] px-4 py-2 text-sm text-white hover:opacity-90"
            onClick={() => {
              const file = preview.file
              setPreview(null)
              void applyImport(file).catch((error: unknown) => {
                toast(`导入失败：${error instanceof Error ? error.message : '未知错误'}`, 'error')
              })
            }}
          >
            确认导入（可编辑）
          </button>
        </div>
      </Modal>
    )
  }

  return (
    <Modal title="导入">
      <label className="block cursor-pointer rounded-xl border-2 border-dashed border-gray-300 px-6 py-10 text-center hover:border-[#d14424]">
        <input
          type="file"
          accept=".pptx,.json"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0]
            if (file) void handleFile(file)
            e.target.value = ''
          }}
        />
        <div className="text-sm text-gray-600">点击选择文件</div>
        <div className="mt-1 text-xs text-gray-400">支持 .pptx（PowerPoint 演示文稿）与 .json（本工具工程文件）</div>
        <div className="mt-2 text-[11px] text-gray-400">PPTX 将先展示高保真预览，确认后再导入为可编辑内容</div>
      </label>
    </Modal>
  )
}
```

- [ ] **Step 2: 类型与构建验证**

```bash
npx tsc -b
pnpm build
```

预期：均无错误。

- [ ] **Step 3: 提交**

```bash
git add packages/pptx/src/components/dialogs/ImportDialog.tsx
git commit -m "feat(ui): ImportDialog 先预览再导入（SVG 缩略网格 + 确认导入）"
```

- [ ] **Step 4: 手工冒烟（留给用户，报告中注明）**

`pnpm dev` 后：导入一个真实 pptx → 预览网格出现 → 返回/确认导入两条路径均可走通。

---

### Task 11: 全量验证收尾

**Files:** 无新文件（验证 + 文档注记）

- [ ] **Step 1: 全量测试**

```bash
npx vitest run
```

预期：全部 PASS，含既有 141 基线用例 + 本期新增（预估 ≥ 30）。

- [ ] **Step 2: 类型 / lint / 构建**

```bash
npx tsc -b
npx oxlint --config=../../.oxlintrc.json packages/pptx/src 2>/dev/null || npx oxlint src
pnpm build
```

预期：全部通过（lint 按仓库实际配置执行，见根目录 oxlint 配置）。

- [ ] **Step 3: 已知取舍写入代码注释核查**

确认以下注记存在于对应文件（各任务已含，最后核对一遍）：
- `preview/shape.ts`：嵌套 svg 描边非均匀缩放取舍
- `preview/text.ts`：normAutofit fontScale 不做缩放
- `preview/table.ts`：逐边定制边框不还原（统一细边框）
- `preview/chart.ts`：radar 占位、bar-percent 按堆叠近似

- [ ] **Step 4: 提交（如有零散修正）**

```bash
git add -A
git commit -m "chore(preview): 全量验证收尾"
```

---

## Self-Review 记录（writing-plans）

1. **Spec 覆盖**（设计文档第 4 节逐项）：
   - 每页 slide → 一个 `<svg>`（viewBox 按源画布尺寸）→ Task 8/9 ✓
   - 复用一期解析层（theme/master/styles/geometry）→ Task 8（findAncestry/parseThemeForMaster/collectPlaceholders/parseBackgroundFill）、Task 3（geometry/styles）✓
   - 阴影/渐变中间站/精确字距按 OOXML 语义 → Task 3（feDropShadow、全停站 linearGradient）、Task 2（spc）✓
   - `previewPPTX(file): Promise<SVGSVGElement[]>` → Task 9 ✓
   - ImportDialog 先预览再导入 → Task 10 ✓
   - 字体缺失按主题字体映射 fallback、与编辑同一映射 → 预览复用 `txBodyToHTML`（typeface 直出）+ 一期主题字体映射；浏览器缺字体时 fallback 由 CSS font-family 链自然承接（与编辑器同一 typeface 值）✓
2. **占位符扫描**：Task 7 测试中 chart part fixture 引用一期 `chart.test.ts` 既有模板（明确指示复制），Task 8 slide.test.ts 指示复制一期模板并给出完整清单——两处均已写明「必须完整写出，不得留占位符」。
3. **类型一致性**：`PreviewCtx`（Task 1）在 Task 3-9 全部一致；`geomOf`/`boxTransform`/`emu2pxF`/`sz2px`/`svgEl` 签名前后一致；`buildCellMatrix` 返回 `{cells, tcEls}` 在 Task 6 内一致；`renderSpTreeNode` 定义于 dispatch.ts（Task 8），group.ts 与 slide.ts 均从 dispatch 导入（无循环依赖）。
