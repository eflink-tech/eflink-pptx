# PPTX 导入第一期（可编辑导入分层重写）实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 按规格（`docs/superpowers/specs/2026-09-18-pptx-import-export-design.md`）将 eflink-pptx 的 PPTX 导入重写为分层解析器，覆盖主题色、母版/版式、组合、渐变、阴影、图表、项目符号等，保持对外 API 不变。

**Architecture:** 解压→关系图→主题→母版/版式继承链→逐页递归解析 spTree→内部元素模型。每个解析关注点一个模块（`src/core/import/` 下拆分），单元素解析失败仅跳过并计入 ImportReport。

**Tech Stack:** TypeScript + JSZip + DOMParser（vitest 测试，fixture 用 JSZip 动态构造，不提交二进制）。

**工作目录：** 所有命令均在 `/Users/apple/Documents/myf-project/eflink.tech/eflink-pptx/packages/pptx` 下执行。测试命令：`pnpm vitest run src/core/import`（在包目录内直接 `npx vitest run <file>` 亦可）。

**已知偏差（已同步进 spec）：**
- 表格单元格为纯文本渲染（`ElementRenderer.tsx:364`），第一期提取首 run 样式进 `style` 字段，不做富文本单元格。
- pptxgenjs 4 不支持形状渐变填充，导出侧渐变维持截图兜底，本计划不触碰导出代码。

**模块总览（最终文件结构）：**

```
src/core/import/
├── index.ts        # 入口：importPPTX / importPPTXDetailed，管线组装
├── xml.ts          # XML 安全访问助手（attr/directChild/descendants/firstDescendant）
├── package.ts      # PptxPackage：zip + rels 图 + 媒体 dataURL + resolveTarget
├── theme.ts        # PptxTheme：主题色表/字体/母版 clrMap
├── styles.ts       # resolveColor：srgbClr/sysClr/schemeClr + alpha/lumMod/lumOff
├── geometry.ts     # PRST_MAP 全量预设表 + custGeomToPath
├── context.ts      # ParseContext / ImportReport / addSkipped / GroupXform
├── master.ts       # slide ancestry：layout/master 链、背景、装饰、占位符表
├── elements/
│   ├── index.ts    # spTree 子节点分发 parseSpTreeNode
│   ├── text.ts     # txBody → 富文本 HTML（列表/超链接/autofit/竖排）
│   ├── shape.ts    # p:sp / p:cxnSp → shape/line/text
│   ├── picture.ts  # p:pic → image/video/audio（裁剪/边框/海报帧）
│   ├── table.ts    # graphicFrame a:tbl → table
│   ├── chart.ts    # graphicFrame c:chart → chart
│   └── group.ts    # grpSp 递归展开
└── pptx.test.ts    # 既有测试（迁移 import 路径）+ 新增单测
```

---

### Task 1: 类型增量（ShapeElement.path 与 ImportReport）

**Files:**
- Modify: `src/types/slides.ts`（ShapeElement 定义附近，约 143-162 行）
- Test: 无需新测试文件，随 Task 13 集成测试覆盖；本任务仅类型，用 typecheck 验证

- [ ] **Step 1: 给 ShapeElement 增加 path 字段**

在 `src/types/slides.ts` 的 `ShapeElement` 接口中，`shapeKey` 字段之后加入：

```typescript
  /** 自定义几何 SVG 路径（custGeom 导入产出；设置后渲染优先于 shapeKey 预设路径，坐标为 0-100 视口空间） */
  path?: string
```

- [ ] **Step 2: 新增 ImportReport 类型**

在文件末尾 `createPresentation` 函数之前加入：

```typescript
/* ---------- 导入兼容性报告 ---------- */

/** PPTX 导入兼容性报告：skipped 记录各类被降级/丢弃项的计数 */
export interface ImportReport {
  skipped: Record<string, number>
}
```

- [ ] **Step 3: 运行 typecheck 验证**

Run: `npx tsc -b --noEmit 2>&1 | head -20`（若 `-b` 不支持 `--noEmit`，直接 `npx tsc -p tsconfig.json --noEmit`）
Expected: 无错误输出（path 为可选字段，无破坏）

- [ ] **Step 4: Commit**

```bash
git add src/types/slides.ts
git commit -m "feat: 类型增量 ShapeElement.path 与 ImportReport（pptx 导入第一期）"
```

---

### Task 2: xml.ts 共享 XML 助手

**Files:**
- Create: `src/core/import/xml.ts`
- Test: `src/core/import/xml.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/xml.test.ts
import { describe, expect, it } from 'vitest'
import { attr, directChild, descendants, firstDescendant } from './xml'

const doc = new DOMParser().parseFromString(
  `<root xmlns:a="urn:a"><item id="1" w="25400"><a:ln w="12700"><a:srgbClr val="FF0000"/></a:ln></item></root>`,
  'text/xml',
)

describe('xml helpers', () => {
  it('attr 读取属性，缺失返回 null', () => {
    const item = directChild(doc.documentElement, 'item')!
    expect(attr(item, 'id')).toBe('1')
    expect(attr(item, 'missing')).toBeNull()
    expect(attr(null, 'id')).toBeNull()
  })

  it('directChild 只找直接子级', () => {
    const item = directChild(doc.documentElement, 'item')!
    expect(directChild(item, 'a:ln')).not.toBeNull()
    expect(directChild(doc.documentElement, 'a:ln')).toBeNull()
  })

  it('descendants/firstDescendant 递归查找', () => {
    const clr = firstDescendant(doc.documentElement, 'a:srgbClr')
    expect(attr(clr, 'val')).toBe('FF0000')
    expect(descendants(doc.documentElement, 'a:srgbClr')).toHaveLength(1)
    expect(firstDescendant(doc.documentElement, 'nope')).toBeNull()
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/xml.test.ts`
Expected: FAIL（模块不存在）

- [ ] **Step 3: 实现**

```typescript
// src/core/import/xml.ts
/** XML 安全访问助手：所有 OOXML 节点访问必须经由此模块，WPS 非标准结构不抛异常 */

export function attr(el: Element | null, name: string): string | null {
  return el?.getAttribute(name) ?? null
}

export function directChild(parent: Element, name: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (child.nodeName === name) return child
  }
  return null
}

export function directChildren(parent: Element, name: string): Element[] {
  return Array.from(parent.children).filter((c) => c.nodeName === name)
}

export function descendants(parent: Element, name: string): Element[] {
  return Array.from(parent.getElementsByTagName(name))
}

export function firstDescendant(parent: Element, name: string): Element | null {
  return descendants(parent, name)[0] ?? null
}

export function parseXML(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml')
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/xml.test.ts`
Expected: PASS（3 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/xml.ts src/core/import/xml.test.ts
git commit -m "feat: 导入器共享 XML 助手 xml.ts"
```

---

### Task 3: package.ts（PptxPackage 部件与关系图）

**Files:**
- Create: `src/core/import/package.ts`
- Test: `src/core/import/package.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/package.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage, resolveTarget } from './package'

async function buildPkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation/>')
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`)
  zip.file('ppt/media/image1.png', 'PNGDATA')
  const blob = await zip.generateAsync({ type: 'blob' })
  return PptxPackage.load(blob)
}

describe('PptxPackage', () => {
  it('text 读取部件文本，缺失返回 null', async () => {
    const pkg = await buildPkg()
    expect(await pkg.text('ppt/presentation.xml')).toContain('presentation')
    expect(await pkg.text('ppt/missing.xml')).toBeNull()
  })

  it('rels 解析关系并区分 External', async () => {
    const pkg = await buildPkg()
    const rels = await pkg.rels('ppt/slides/slide1.xml')
    expect(rels.get('rId1')?.target).toBe('ppt/media/image1.png')
    expect(rels.get('rId2')?.mode).toBe('External')
  })

  it('relTarget 解析为包内绝对路径', async () => {
    const pkg = await buildPkg()
    expect(await pkg.relTarget('ppt/slides/slide1.xml', 'rId1')).toBe('ppt/media/image1.png')
    expect(await pkg.relTarget('ppt/slides/slide1.xml', 'rIdX')).toBeNull()
  })

  it('mediaDataUrl 产出 base64 data URL 并缓存', async () => {
    const pkg = await buildPkg()
    const url = await pkg.mediaDataUrl('ppt/media/image1.png')
    expect(url).toMatch(/^data:image\/png;base64,/)
    expect(await pkg.mediaDataUrl('ppt/media/none.png')).toBeUndefined()
  })

  it('resolveTarget 处理相对/绝对路径', () => {
    expect(resolveTarget('ppt/slides', '../media/a.png')).toBe('ppt/media/a.png')
    expect(resolveTarget('ppt', '/ppt/media/a.png')).toBe('ppt/media/a.png')
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/package.test.ts`
Expected: FAIL（模块不存在）

- [ ] **Step 3: 实现**

```typescript
// src/core/import/package.ts
/** 部件模型：zip 解压 + OOXML 关系（rels）图 + 媒体读取 */
import JSZip from 'jszip'

export interface PartRel {
  /** 包内绝对路径（External 关系为原始目标） */
  target: string
  mode?: string
}

/** 相对路径解析：'../media/a.png'（基于 'ppt/slides'）→ 'ppt/media/a.png' */
export function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = `${baseDir}/${target}`.split('/')
  const out: string[] = []
  for (const part of parts) {
    if (part === '..') out.pop()
    else if (part && part !== '.') out.push(part)
  }
  return out.join('/')
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp', tiff: 'image/tiff',
  mp4: 'video/mp4', webm: 'video/webm', mov: 'video/quicktime',
  mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', wma: 'audio/x-ms-wma',
}

export class PptxPackage {
  private zip: JSZip
  private relsCache = new Map<string, Map<string, PartRel>>()
  private mediaCache = new Map<string, string>()

  private constructor(zip: JSZip) {
    this.zip = zip
  }

  static async load(file: Blob): Promise<PptxPackage> {
    return new PptxPackage(await JSZip.loadAsync(file))
  }

  async text(path: string): Promise<string | null> {
    return (await this.zip.file(path)?.async('text')) ?? null
  }

  /** 部件的 rels：'ppt/slides/slide1.xml' → 解析 'ppt/slides/_rels/slide1.xml.rels'，target 已解析为包内绝对路径 */
  async rels(partPath: string): Promise<Map<string, PartRel>> {
    const cached = this.relsCache.get(partPath)
    if (cached) return cached
    const map = new Map<string, PartRel>()
    const idx = partPath.lastIndexOf('/')
    const relPath = `${partPath.slice(0, idx)}/_rels/${partPath.slice(idx + 1)}.rels`
    const xml = await this.text(relPath)
    if (xml) {
      const doc = new DOMParser().parseFromString(xml, 'text/xml')
      const baseDir = partPath.slice(0, idx)
      for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
        const id = rel.getAttribute('Id')
        const target = rel.getAttribute('Target')
        if (!id || !target) continue
        const mode = rel.getAttribute('TargetMode') ?? undefined
        map.set(id, {
          target: mode === 'External' ? target : resolveTarget(baseDir, target),
          mode,
        })
      }
    }
    this.relsCache.set(partPath, map)
    return map
  }

  /** rId → 包内绝对路径（External 或缺失返回 null） */
  async relTarget(partPath: string, rId: string): Promise<string | null> {
    const rel = (await this.rels(partPath)).get(rId)
    if (!rel || rel.mode === 'External') return null
    return rel.target
  }

  /** 包内媒体文件 → data URL（带缓存） */
  async mediaDataUrl(packagePath: string): Promise<string | undefined> {
    const cached = this.mediaCache.get(packagePath)
    if (cached) return cached
    const file = this.zip.file(packagePath)
    if (!file) return undefined
    try {
      const base64 = await file.async('base64')
      const ext = packagePath.split('.').pop()?.toLowerCase() ?? ''
      const mime = MIME_BY_EXT[ext] ?? 'application/octet-stream'
      const dataUrl = `data:${mime};base64,${base64}`
      this.mediaCache.set(packagePath, dataUrl)
      return dataUrl
    } catch {
      return undefined
    }
  }
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/package.test.ts`
Expected: PASS（5 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/package.ts src/core/import/package.test.ts
git commit -m "feat: PptxPackage 部件与关系图模型"
```

---

### Task 4: context.ts（解析上下文与公共类型）

**Files:**
- Create: `src/core/import/context.ts`

- [ ] **Step 1: 实现**

```typescript
// src/core/import/context.ts
/** 解析上下文：贯穿整个导入管线的共享状态 */
import type { ImportReport } from '../../types/slides'
import type { PptxPackage } from './package'
import type { PptxTheme } from './theme'

export interface ParseContext {
  pkg: PptxPackage
  /** 当前部件路径（rels 查询基准，如 'ppt/slides/slide1.xml'） */
  partPath: string
  theme: PptxTheme
  report: ImportReport
  /** EMU→px 缩放（画布适配） */
  scale: { x: number; y: number }
  /** 版式/母版占位符位置表（EMU 空间，key = idx 或 type） */
  placeholders: Map<string, { x: number; y: number; w: number; h: number }>
}

/** 记录一类降级/丢弃项 */
export function addSkipped(report: ImportReport, kind: string): void {
  report.skipped[kind] = (report.skipped[kind] ?? 0) + 1
}

/** 组合变换：EMU 空间的平移 + 轴缩放仿射（rot 为需要叠加到子元素的角度 1/60000 deg） */
export interface GroupXform {
  ox: number
  oy: number
  sx: number
  sy: number
  rot: number
}

export const IDENTITY_XFORM: GroupXform = { ox: 0, oy: 0, sx: 1, sy: 1, rot: 0 }

export function mapX(t: GroupXform, v: number): number {
  return t.ox + v * t.sx
}

export function mapY(t: GroupXform, v: number): number {
  return t.oy + v * t.sy
}
```

- [ ] **Step 2: typecheck**

Run: `npx tsc -p tsconfig.json --noEmit 2>&1 | grep import/context || true`
Expected: 无输出（theme.ts 尚不存在会报错，此任务允许暂时报 theme 未定义；若报错则先创建空壳 `src/core/import/theme.ts` 导出 `export interface PptxTheme { schemeColors: Record<string, string>; majorFont: string; minorFont: string; colorMap: Record<string, string> }`，Task 5 再补全）

- [ ] **Step 3: Commit**

```bash
git add src/core/import/context.ts src/core/import/theme.ts
git commit -m "feat: 导入解析上下文 ParseContext 与 GroupXform"
```

---

### Task 5: theme.ts（主题色表 / 字体 / clrMap）

**Files:**
- Modify: `src/core/import/theme.ts`（替换 Task 4 的空壳）
- Test: `src/core/import/theme.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/theme.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PptxPackage } from './package'
import { parseThemeForMaster, resolveSchemeColor } from './theme'

async function buildPkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>`)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`)
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`)
  const blob = await zip.generateAsync({ type: 'blob' })
  return PptxPackage.load(blob)
}

describe('parseThemeForMaster', () => {
  it('解析主题色表 / 字体 / clrMap，sysClr 取 lastClr', async () => {
    const theme = await parseThemeForMaster(await buildPkg(), 'ppt/slideMasters/slideMaster1.xml')
    expect(theme.schemeColors.dk1).toBe('#000000')
    expect(theme.schemeColors.lt1).toBe('#FFFFFF')
    expect(theme.schemeColors.accent1).toBe('#4472C4')
    expect(theme.majorFont).toBe('Calibri Light')
    expect(theme.minorFont).toBe('Calibri')
    expect(theme.colorMap.bg1).toBe('lt1')
  })

  it('resolveSchemeColor 过 clrMap 再查色表；缺失回退黑色', async () => {
    const theme = await parseThemeForMaster(await buildPkg(), 'ppt/slideMasters/slideMaster1.xml')
    expect(resolveSchemeColor('bg1', theme)).toBe('#FFFFFF')
    expect(resolveSchemeColor('accent2', theme)).toBe('#ED7D31')
    expect(resolveSchemeColor('nope', theme)).toBe('#000000')
  })

  it('缺 theme 部件时回退默认色表', async () => {
    const zip = new JSZip()
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:x"/>')
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const theme = await parseThemeForMaster(pkg, 'ppt/slideMasters/slideMaster1.xml')
    expect(theme.schemeColors.accent1).toBe('#4472C4')
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/theme.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现（整体替换 theme.ts）**

```typescript
// src/core/import/theme.ts
/** 主题解析：clrScheme / fontScheme / 母版 clrMap */
import { attr, directChild, firstDescendant, parseXML } from './xml'
import type { PptxPackage } from './package'

export interface PptxTheme {
  /** dk1 lt1 dk2 lt2 accent1-6 hlink folHlink → '#RRGGBB' */
  schemeColors: Record<string, string>
  majorFont: string
  minorFont: string
  /** 母版 p:clrMap：bg1→lt1 等；未指定的键与 key 相同 */
  colorMap: Record<string, string>
}

export const DEFAULT_SCHEME: Record<string, string> = {
  dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6',
  accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000',
  accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72',
}

const CLR_MAP_KEYS = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function readClrNode(clr: Element | null): string | undefined {
  if (!clr) return undefined
  const srgb = firstDescendant(clr, 'a:srgbClr')
  if (srgb) return `#${(attr(srgb, 'val') ?? '000000').toUpperCase()}`
  const sys = firstDescendant(clr, 'a:sysClr')
  if (sys) return `#${(attr(sys, 'lastClr') ?? '000000').toUpperCase()}`
  return undefined
}

function fontOf(fontScheme: Element | null, tag: string): string | undefined {
  const node = fontScheme ? directChild(fontScheme, `a:${tag}`) : null
  const latin = node ? directChild(node, 'a:latin') : null
  const typeface = attr(latin, 'typeface')
  return typeface && !typeface.startsWith('+') ? typeface : undefined
}

/** 解析 master 关联的主题部件 + master 自身 clrMap（master 缺失时返回默认主题） */
export async function parseThemeForMaster(pkg: PptxPackage, masterPath: string | null): Promise<PptxTheme> {
  const colorMap: Record<string, string> = {}
  if (masterPath) {
    const masterXml = await pkg.text(masterPath)
    if (masterXml) {
      const clrMap = firstDescendant(parseXML(masterXml).documentElement, 'p:clrMap')
      for (const k of CLR_MAP_KEYS) {
        const v = attr(clrMap, k)
        if (v) colorMap[k] = v
      }
    }
  }

  const schemeColors: Record<string, string> = { ...DEFAULT_SCHEME }
  let majorFont = 'Calibri'
  let minorFont = 'Calibri'

  if (masterPath) {
    const rels = await pkg.rels(masterPath)
    let themePath: string | null = null
    for (const rel of rels.values()) {
      if (rel.mode !== 'External' && rel.target.includes('/theme/')) {
        themePath = rel.target
        break
      }
    }
    if (themePath) {
      const xml = await pkg.text(themePath)
      if (xml) {
        const root = parseXML(xml).documentElement
        const scheme = firstDescendant(root, 'a:clrScheme')
        if (scheme) {
          for (const child of Array.from(scheme.children)) {
            const name = child.nodeName.replace(/^a:/, '')
            const hex = readClrNode(child)
            if (hex) schemeColors[name] = hex
          }
        }
        const fontScheme = firstDescendant(root, 'a:fontScheme')
        majorFont = fontOf(fontScheme, 'majorFont') ?? majorFont
        minorFont = fontOf(fontScheme, 'minorFont') ?? minorFont
      }
    }
  }

  return { schemeColors, majorFont, minorFont, colorMap }
}

/** schemeClr 名称求值：先过母版 clrMap（bg1→lt1 等），再查主题色表 */
export function resolveSchemeColor(name: string, theme: PptxTheme): string {
  const mapped = theme.colorMap[name] ?? name
  return theme.schemeColors[mapped] ?? theme.schemeColors[name] ?? '#000000'
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/theme.test.ts`
Expected: PASS（3 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/theme.ts src/core/import/theme.test.ts
git commit -m "feat: 主题解析 theme.ts（clrScheme/fontScheme/clrMap）"
```

---

### Task 6: styles.ts（颜色求值：schemeClr / alpha / lumMod / lumOff / shade / tint）

**Files:**
- Create: `src/core/import/styles.ts`
- Test: `src/core/import/styles.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/styles.test.ts
import { describe, expect, it } from 'vitest'
import { resolveColor, resolveColorOf } from './styles'
import type { PptxTheme } from './theme'

const theme: PptxTheme = {
  schemeColors: { dk1: '#000000', lt1: '#FFFFFF', accent1: '#4472C4' } as Record<string, string>,
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('styles.resolveColor', () => {
  it('srgbClr 直接取值', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"/></a:solidFill>`), theme)).toBe('#FF0000')
  })

  it('schemeClr 走主题色表', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="accent1"/></a:solidFill>`), theme)).toBe('#4472C4')
  })

  it('alpha 修饰 → 8 位 hex', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#FF000080')
  })

  it('lumMod/lumOff：白底 Darker 25%（lumMod 75000）→ #BFBFBF', () => {
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:schemeClr val="lt1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>`), theme)).toBe('#BFBFBF')
  })

  it('tint/shade 近似', () => {
    // shade 50% 黑色不变、白色减半
    expect(resolveColor(el(`<a:solidFill xmlns:a="urn:a"><a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr></a:solidFill>`), theme)).toBe('#808080')
  })

  it('空容器/无色子节点返回 undefined', () => {
    expect(resolveColor(el(`<a:ln xmlns:a="urn:a"/>`), theme)).toBeUndefined()
    expect(resolveColor(null, theme)).toBeUndefined()
    expect(resolveColorOf(el(`<a:noFill xmlns:a="urn:a"/>`), theme)).toBeUndefined()
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/styles.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/styles.ts
/** 颜色求值：OOXML 颜色节点 → 内部 hex（含 alpha 8 位形式） */
import { resolveSchemeColor } from './theme'
import type { PptxTheme } from './theme'

const BYTE = 255

function clampByte(n: number): number {
  return Math.min(BYTE, Math.max(0, Math.round(n)))
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((v) => clampByte(v).toString(16).padStart(2, '0')).join('')}`
}

function hexToRgb(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ]
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / BYTE, gn = g / BYTE, bn = b / BYTE
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6
  else if (max === gn) h = ((bn - rn) / d + 2) / 6
  else h = ((rn - gn) / d + 4) / 6
  return [h, s, l]
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * BYTE, l * BYTE, l * BYTE]
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const hue = (t: number): number => {
    let tt = t
    if (tt < 0) tt += 1
    if (tt > 1) tt -= 1
    if (tt < 1 / 6) return p + (q - p) * 6 * tt
    if (tt < 1 / 2) return q
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6
    return p
  }
  return [hue(h + 1 / 3) * BYTE, hue(h) * BYTE, hue(h - 1 / 3) * BYTE]
}

/** lumMod/lumOff：HSL 亮度缩放/偏移（mod 与 off 以 1 为满值） */
function applyLum(r: number, g: number, b: number, mod: number, off: number): [number, number, number] {
  const [h, s, l] = rgbToHsl(r, g, b)
  return hslToRgb(h, s, Math.min(1, Math.max(0, l * mod + off)))
}

/** 单个颜色节点（a:srgbClr / a:sysClr / a:schemeClr / a:prstClr）求值 */
export function resolveColorOf(clrEl: Element, theme: PptxTheme): string | undefined {
  const name = clrEl.nodeName
  let hex: string | undefined
  if (name === 'a:srgbClr') hex = `#${(clrEl.getAttribute('val') ?? '000000').toUpperCase()}`
  else if (name === 'a:sysClr') hex = `#${(clrEl.getAttribute('lastClr') ?? '000000').toUpperCase()}`
  else if (name === 'a:schemeClr') hex = resolveSchemeColor(clrEl.getAttribute('val') ?? 'tx1', theme)
  else if (name === 'a:prstClr') hex = `#${(clrEl.getAttribute('lastClr') ?? '000000').toUpperCase()}`
  else return undefined
  if (!hex || hex.length !== 7) return undefined

  let [r, g, b] = hexToRgb(hex)
  let mod = 1
  let off = 0
  for (const modEl of Array.from(clrEl.children)) {
    const val = parseInt(modEl.getAttribute('val') ?? '0', 10) / 100000
    switch (modEl.nodeName) {
      case 'a:alpha': {
        const a = Math.round(val * BYTE).toString(16).padStart(2, '0')
        return rgbToHex(r, g, b) + a
      }
      case 'a:lumMod': mod = val; break
      case 'a:lumOff': off = val; break
      case 'a:shade': r *= val; g *= val; b *= val; break
      case 'a:tint': r += (BYTE - r) * (1 - val); g += (BYTE - g) * (1 - val); b += (BYTE - b) * (1 - val); break
      default: break
    }
  }
  if (mod !== 1 || off !== 0) [r, g, b] = applyLum(r, g, b, mod, off)
  return rgbToHex(r, g, b)
}

/** 容器节点（a:solidFill / a:ln / a:bgPr 等）取第一个颜色子节点求值 */
export function resolveColor(container: Element | null, theme: PptxTheme): string | undefined {
  if (!container) return undefined
  for (const child of Array.from(container.children)) {
    const c = resolveColorOf(child, theme)
    if (c) return c
  }
  return undefined
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/styles.test.ts`
Expected: PASS（6 个用例；若 lumMod 用例偏差超过 ±2/通道，检查 HSL 转换实现）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/styles.ts src/core/import/styles.test.ts
git commit -m "feat: 颜色求值 styles.ts（schemeClr/alpha/lumMod/lumOff/shade/tint）"
```

---

### Task 7: geometry.ts（全量 prstGeom 预设表 + custGeom 路径）

**Files:**
- Create: `src/core/import/geometry.ts`
- Test: `src/core/import/geometry.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/geometry.test.ts
import { describe, expect, it } from 'vitest'
import { custGeomToPath, getShapeKey } from './geometry'

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('geometry', () => {
  it('getShapeKey：已知预设映射内部 key，未知回退 rect', () => {
    expect(getShapeKey('roundRect')).toBe('roundRect')
    expect(getShapeKey('rightArrow')).toBe('arrowRight')
    expect(getShapeKey('nonExistentPrst')).toBe('rect')
  })

  it('custGeomToPath：moveTo/lnTo/close 归一化到 0-100', () => {
    const cust = el(`<a:custGeom xmlns:a="urn:a">
      <a:pathLst>
        <a:path w="1000" h="500">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="1000" y="0"/></a:lnTo>
          <a:lnTo><a:pt x="500" y="500"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>`)
    expect(custGeomToPath(cust, 200, 100)).toBe('M0,0 L100,0 L50,100 Z')
  })

  it('custGeomToPath：cubicBezTo 与缺省 path 尺寸（用元素尺寸）', () => {
    const cust = el(`<a:custGeom xmlns:a="urn:a">
      <a:pathLst>
        <a:path>
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:cubicBezTo><a:pt x="0" y="952500"/><a:pt x="952500" y="952500"/><a:pt x="952500" y="0"/></a:cubicBezTo>
        </a:path>
      </a:pathLst>
    </a:custGeom>`)
    // 元素 96x96px → 952500 EMU；路径缺省 w/h 用元素尺寸
    expect(custGeomToPath(cust, 96, 96)).toBe('M0,0 C0,100 100,100 100,0')
  })

  it('custGeomToPath：无 pathLst 返回 null', () => {
    expect(custGeomToPath(el(`<a:custGeom xmlns:a="urn:a"/>`), 10, 10)).toBeNull()
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/geometry.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/geometry.ts
/** 几何解析：全量 OOXML prstGeom 预设映射 + custGeom → SVG path（0-100 视口空间） */
import { attr, directChildren, firstDescendant } from './xml'

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
```

**注意：** `PRST_MAP` 只列真实存在的 OOXML 预设名；未列出的预设由 `getShapeKey` 兜底为 `'rect'`（ connector 类预设映射为 `'rect'` 仅为占位，`p:cxnSp` 节点在 shape.ts 中会先按线条处理，不会走到该映射）。

```typescript
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
          if (pts[0]) cmds.push(`L${fmt(fx(pts[0][0]))},${fmt(fy(pts[0][1]))}`)
          break
        default:
          break
      }
    }
  }
  return cmds.length ? cmds.join(' ') : null
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/geometry.test.ts`
Expected: PASS（4 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/geometry.ts src/core/import/geometry.test.ts
git commit -m "feat: 几何解析 geometry.ts（全量 prst 预设表 + custGeom 路径）"
```

---

### Task 8: elements/text.ts（txBody → 富文本 HTML）

**Files:**
- Create: `src/core/elements/text.ts` → 实际路径 `src/core/import/elements/text.ts`
- Test: `src/core/import/elements/text.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/elements/text.test.ts
import { describe, expect, it } from 'vitest'
import { txBodyToHTML } from './text'
import type { PptxTheme } from '../theme'

const theme: PptxTheme = {
  schemeColors: { tx1: '#000000', accent1: '#4472C4' } as Record<string, string>,
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('txBodyToHTML', () => {
  it('基础 run：字号/加粗/颜色/对齐', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>标题</a:t></a:r>
      </a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('标题')
    expect(r.html).toContain('font-weight:bold')
    expect(r.html).toContain('color:#D14424')
    expect(r.html).toContain('text-align:center')
    expect(r.autoSize).toBe(false)
  })

  it('normAutofit → autoSize；vert=eaVert → vertical', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr vert="eaVert"><a:normAutofit fontScale="92500"/></a:bodyPr>
      <a:p><a:r><a:t>竖排</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.autoSize).toBe(true)
    expect(r.vertical).toBe(true)
  })

  it('项目符号段落聚合为 ul；编号聚合为 ol', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a">
      <a:bodyPr/>
      <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>项一</a:t></a:r></a:p>
      <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>项二</a:t></a:r></a:p>
      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>步一</a:t></a:r></a:p>
      <a:p><a:r><a:t>普通</a:t></a:r></a:p>
    </p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('<ul>')
    expect(r.html).toContain('<li')
    expect(r.html).toContain('<ol>')
    expect(r.html).toContain('普通')
    // 普通段落在列表之后
    expect(r.html.indexOf('</ol>')).toBeLessThan(r.html.indexOf('普通'))
  })

  it('空段落输出占位段落', async () => {
    const txBody = el(`<p:txBody xmlns:p="urn:p" xmlns:a="urn:a"><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>`)
    const r = await txBodyToHTML(txBody, theme)
    expect(r.html).toContain('<p')
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/elements/text.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/elements/text.ts
/** txBody → 富文本 HTML（项目符号/编号、超链接、autofit、竖排） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import type { PptxTheme } from '../theme'
import type { PptxPackage } from '../package'

export interface TextBodyResult {
  html: string
  autoSize: boolean
  vertical: boolean
}

function escapeHTML(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

interface ParaInfo {
  align: string
  kind: 'none' | 'bullet' | 'number'
  inner: string
}

function bulletKind(p: Element): 'none' | 'bullet' | 'number' {
  const pPr = directChild(p, 'a:pPr')
  if (!pPr) return 'none'
  if (directChild(pPr, 'a:buNone')) return 'none'
  if (directChild(pPr, 'a:buChar')) return 'bullet'
  if (directChild(pPr, 'a:buAutoNum')) return 'number'
  return 'none'
}

async function paragraphToInner(p: Element, theme: PptxTheme, pkg: PptxPackage, partPath: string): Promise<string> {
  let inner = ''
  for (const node of Array.from(p.children)) {
    if (node.nodeName === 'a:br') {
      inner += '<br>'
      continue
    }
    if (node.nodeName !== 'a:r') continue
    const t = firstDescendant(node, 'a:t')
    const text = t?.textContent ?? ''
    if (!text) continue
    const rPr = firstDescendant(node, 'a:rPr')
    const styles: string[] = []
    if (rPr) {
      const sz = attr(rPr, 'sz')
      if (sz) styles.push(`font-size:${Math.round(parseInt(sz, 10) / 100 / 0.75)}px`)
      if (attr(rPr, 'b') === '1') styles.push('font-weight:bold')
      if (attr(rPr, 'i') === '1') styles.push('font-style:italic')
      if (attr(rPr, 'u') === 'sng') styles.push('text-decoration:underline')
      if (attr(rPr, 'strike') === 'sng') styles.push('text-decoration:line-through')
      const color = resolveColor(directChild(rPr, 'a:solidFill'), theme)
      if (color) styles.push(`color:${color}`)
      const latin = directChild(rPr, 'a:latin')
      const typeface = attr(latin, 'typeface')
      if (typeface && !typeface.startsWith('+')) styles.push(`font-family:${typeface}`)
    }
    const styleAttr = styles.length ? ` style="${styles.join(';')}"` : ''
    let run = `<span${styleAttr}>${escapeHTML(text)}</span>`
    // 超链接：a:hlinkClick@r:id → rels（External）
    const hlink = rPr ? directChild(rPr, 'a:hlinkClick') : null
    const hlinkId = attr(hlink, 'r:id')
    if (hlinkId) {
      const rel = (await pkg.rels(partPath)).get(hlinkId)
      if (rel && rel.mode === 'External') {
        run = `<a href="${escapeHTML(rel.target)}"${styleAttr}>${escapeHTML(text)}</a>`
      }
    }
    inner += run
  }
  return inner
}

/** txBody → 富文本 HTML；autoSize 来自 normAutofit；vertical 来自 bodyPr@vert */
export async function txBodyToHTML(
  txBody: Element,
  theme: PptxTheme,
  pkg?: PptxPackage,
  partPath = '',
): Promise<TextBodyResult> {
  const bodyPr = directChild(txBody, 'a:bodyPr')
  const autoSize = Boolean(bodyPr && directChild(bodyPr, 'a:normAutofit'))
  const vert = attr(bodyPr, 'vert')
  const vertical = vert === 'eaVert' || vert === 'vert' || vert === 'mongolianVert'

  const paragraphs = directChildren(txBody, 'a:p')
  const paras: ParaInfo[] = []
  for (const p of paragraphs) {
    const pPr = directChild(p, 'a:pPr')
    const algn = attr(pPr, 'algn')
    const align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left'
    const inner = await paragraphToInner(p, theme, pkg ?? ({} as PptxPackage), partPath)
    paras.push({ align, kind: bulletKind(p), inner })
  }

  const html: string[] = []
  let i = 0
  while (i < paras.length) {
    const para = paras[i]
    if (para.kind === 'none') {
      html.push(`<p style="text-align:${para.align}">${para.inner || '&nbsp;'}</p>`)
      i += 1
      continue
    }
    const tag = para.kind === 'bullet' ? 'ul' : 'ol'
    const items: string[] = []
    while (i < paras.length && paras[i].kind === para.kind) {
      const cur = paras[i]
      items.push(`<li style="text-align:${cur.align}">${cur.inner || '&nbsp;'}</li>`)
      i += 1
    }
    html.push(`<${tag}>${items.join('')}</${tag}>`)
  }
  return { html: html.join('') || '<p></p>', autoSize, vertical }
}
```

注意：`pkg ?? ({} as PptxPackage)` 仅在无包上下文（单测便捷）时使用；超链接解析处对 `pkg.rels` 的调用需以 `typeof pkg.rels === 'function'` 守卫，无包时跳过：

```typescript
      if (rel && rel.mode === 'External') {
```

上方代码中rels取值行为改为：

```typescript
    if (hlinkId && pkg && typeof pkg.rels === 'function') {
      const rel = (await pkg.rels(partPath)).get(hlinkId)
      if (rel && rel.mode === 'External') {
        run = `<a href="${escapeHTML(rel.target)}"${styleAttr}>${escapeHTML(text)}</a>`
      }
    }
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/elements/text.test.ts`
Expected: PASS（4 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/elements/text.ts src/core/import/elements/text.test.ts
git commit -m "feat: txBody 富文本解析（项目符号/编号/超链接/autofit/竖排）"
```

---

### Task 9: elements/shape.ts（形状 / 线条 / 文本框，含渐变与阴影）

**Files:**
- Create: `src/core/import/elements/shape.ts`
- Test: `src/core/import/elements/shape.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/elements/shape.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseShapeEl } from './shape'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makeCtx(slideXml: string) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slideXml)
  const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
  return {
    pkg,
    ctx: {
      pkg,
      partPath: 'ppt/slides/slide1.xml',
      theme,
      report: { skipped: {} } as ImportReport,
      scale: { x: 1, y: 1 },
      placeholders: new Map(),
    },
  }
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('parseShapeEl', () => {
  it('渐变填充 + 阴影 + custGeom path', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Custom"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
      <a:custGeom><a:pathLst><a:path w="1000" h="1000">
        <a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1000" y="1000"/></a:lnTo>
      </a:path></a:pathLst></a:custGeom>
      <a:gradFill>
        <a:gsLst>
          <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
        </a:gsLst>
        <a:lin ang="2700000"/>
      </a:gradFill>
      <a:ln><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
      <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
  </p:sp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const sp = el(xml).getElementsByTagName('p:sp')[0]
    const result = await parseShapeEl(sp, ctx, IDENTITY_XFORM, pkg)
    expect(result).not.toBeNull()
    if (result?.type !== 'shape') throw new Error('expected shape')
    expect(result.path).toBe('M0,0 L100,100')
    expect(typeof result.fill).toBe('object')
    const grad = result.fill as { type: string; colors: Array<{ pos: number; color: string }>; rotate: number }
    expect(grad.colors[0].color).toBe('#FF0000')
    expect(grad.colors[1].color).toBe('#0000FF')
    expect(grad.rotate).toBe(45)
    expect(result.shadow?.blur).toBe(5)
    expect(result.shadow?.color).toBe('#00000066')
  })

  it('线条：prstDash 与箭头映射', async () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
  <p:cxnSp>
    <p:nvSpPr><p:cNvPr id="2" name="Line"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:schemeClr></a:ln>
    </p:spPr>
  </p:cxnSp>
</p:spTree></p:cSld></p:sld>`
    const { pkg, ctx } = await makeCtx(xml)
    const cxn = el(xml).getElementsByTagName('p:cxnSp')[0]
    const result = await parseShapeEl(cxn, ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'line') throw new Error('expected line')
    expect(result.lineWidth).toBe(2)
    expect(result.color).toBe('#4472C4')
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/elements/shape.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/elements/shape.ts
/** p:sp / p:cxnSp → shape / line / text 元素 */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { custGeomToPath, getShapeKey } from '../geometry'
import { txBodyToHTML } from './text'
import { genId } from '../../utils/id'
import { mapX, mapY, addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type {
  Gradient, LineElement, PPTElement, ShapeElement, ShadowEffect, TextElement,
} from '../../../types/slides'

/** EMU→px（96dpi） */
export function emu2px(emu: number): number {
  return emu / 9525
}

/** a:gradFill → 内部 Gradient（linear 取停站色，角度 1/60000 deg → rotate） */
export function parseGradient(gradFill: Element, ctx: ParseContext): Gradient | undefined {
  const gsLst = directChild(gradFill, 'a:gsLst')
  const stops = gsLst ? directChildren(gsLst, 'a:gs') : []
  if (!stops.length) return undefined
  const colors = stops.map((gs) => ({
    pos: (parseInt(attr(gs, 'pos') ?? '0', 10)) / 100000,
    color: resolveColor(gs, ctx.theme) ?? '#000000',
  }))
  const lin = directChild(gradFill, 'a:lin')
  const rotate = lin ? Math.round(parseInt(attr(lin, 'ang') ?? '0', 10) / 60000) : 0
  return { type: 'linear', colors, rotate }
}

/** a:outerShdw → 内部 ShadowEffect */
export function parseShadow(effectLst: Element | null, ctx: ParseContext): ShadowEffect | undefined {
  const shdw = effectLst ? directChild(effectLst, 'a:outerShdw') : null
  if (!shdw) return undefined
  const blur = Math.round(emu2px(parseInt(attr(shdw, 'blurRad') ?? '0', 10)))
  const dist = emu2px(parseInt(attr(shdw, 'dist') ?? '0', 10))
  const dirRad = (parseInt(attr(shdw, 'dir') ?? '0', 10) / 60000) * (Math.PI / 180)
  return {
    h: Math.round(dist * Math.cos(dirRad)),
    v: Math.round(dist * Math.sin(dirRad)),
    blur,
    color: resolveColor(shdw, ctx.theme) ?? '#00000000',
  }
}

function dashOf(ln: Element | null): 'solid' | 'dashed' | 'dotted' {
  const dash = ln ? attr(directChild(ln, 'a:prstDash'), 'val') : null
  if (dash === 'dash' || dash === 'sysDash' || dash === 'lgDash') return 'dashed'
  if (dash === 'dot' || dash === 'sysDot' || dash === 'lgDot') return 'dotted'
  return 'solid'
}

function arrowOf(end: Element | null): string {
  const type = attr(end, 'type')
  if (type === 'triangle' || type === 'stealth') return 'triangle'
  if (type === 'arrow') return 'arrow'
  if (type === 'oval') return 'dot'
  return ''
}

function alignFromAlgn(algn: string | null): 'left' | 'center' | 'right' {
  return algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left'
}

/** p:sp / p:cxnSp 解析；xf 为组合变换（EMU 空间） */
export async function parseShapeEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement | null> {
  const xfrm = firstDescendant(node, 'a:xfrm')
  const ph = firstDescendant(firstDescendant(node, 'p:nvSpPr') ?? node, 'p:ph')
  let emuX: number, emuY: number, emuW: number, emuH: number
  if (xfrm) {
    const off = directChild(xfrm, 'a:off')
    const ext = directChild(xfrm, 'a:ext')
    if (!off || !ext) return null
    emuX = parseInt(attr(off, 'x') ?? '0', 10)
    emuY = parseInt(attr(off, 'y') ?? '0', 10)
    emuW = parseInt(attr(ext, 'cx') ?? '0', 10)
    emuH = parseInt(attr(ext, 'cy') ?? '0', 10)
  } else if (ph) {
    // 占位符位置继承（版式/母版）
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    const pos = ctx.placeholders.get(key)
    if (!pos) return null
    emuX = pos.x; emuY = pos.y; emuW = pos.w; emuH = pos.h
  } else {
    return null
  }
  const x = Math.round(mapX(xf, emuX) / 9525 * ctx.scale.x)
  const y = Math.round(mapY(xf, emuY) / 9525 * ctx.scale.y)
  const w = Math.max(1, Math.round((mapX(xf, emuX + emuW) - mapX(xf, emuX)) / 9525 * ctx.scale.x))
  const h = Math.max(1, Math.round((mapY(xf, emuY + emuH) - mapY(xf, emuY)) / 9525 * ctx.scale.y))
  const rot = parseInt(attr(xfrm, 'rot') ?? '0', 10) / 60000 + (xf.rot ? xf.rot / 60000 : 0)
  const flipH = attr(xfrm, 'flipH') === '1' || undefined
  const flipV = attr(xfrm, 'flipV') === '1' || undefined

  const spPr = firstDescendant(node, 'a:spPr') ?? firstDescendant(node, 'p:spPr')
  const txBody = firstDescendant(node, 'p:txBody') ?? firstDescendant(node, 'a:txBody')
  const prstGeom = firstDescendant(node, 'a:prstGeom')
  const prst = attr(prstGeom, 'prst') ?? ''
  const custGeom = spPr ? directChild(spPr, 'a:custGeom') : null
  const solidFill = spPr ? directChild(spPr, 'a:solidFill') : null
  const noFill = spPr ? directChild(spPr, 'a:noFill') : null
  const gradFill = spPr ? directChild(spPr, 'a:gradFill') : null
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const effectLst = spPr ? directChild(spPr, 'a:effectLst') : null
  const shadow = parseShadow(effectLst, ctx)

  // 连接线
  if (node.nodeName === 'p:cxnSp' || prst === 'line' || prst === 'straightConnector1') {
    const line: LineElement = {
      id: genId('l-'), type: 'line', x, y, w, h,
      start: [flipH ? w : 0, flipV ? h : 0],
      end: [flipH ? 0 : w, flipV ? 0 : h],
      lineType: 'straight',
      color: resolveColor(ln ? directChild(ln, 'a:solidFill') : solidFill, ctx.theme) ?? '#333333',
      lineWidth: Math.max(1, Math.round(emu2px(parseInt(attr(ln, 'w') ?? '12700', 10)))),
      lineStyle: dashOf(ln),
      startArrow: arrowOf(ln ? directChild(ln, 'a:headEnd') : null) as LineElement['startArrow'],
      endArrow: arrowOf(ln ? directChild(ln, 'a:tailEnd') : null) as LineElement['endArrow'],
      name: '线条',
    }
    return line
  }

  // 纯文本框：rect/textbox 且无填充
  if (node.nodeName === 'p:sp' && (prst === 'rect' || prst === 'textbox' || (!prst && !custGeom)) && txBody && !solidFill && !gradFill && !noFill) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath)
    const text: TextElement = {
      id: genId('t-'), type: 'text', x, y, w, h,
      content: body.html,
      rotate: rot ? Math.round(rot) : undefined,
      defaultColor: '#333333', lineHeight: 1.5, padding: 8, name: '文本框',
    }
    if (body.autoSize) text.autoSize = true
    if (body.vertical) text.vertical = true
    if (shadow) text.shadow = shadow
    return text
  }

  // 形状
  const customPath = custGeom ? custGeomToPath(custGeom, w, h) : null
  let fill: string | Gradient | undefined
  if (gradFill) fill = parseGradient(gradFill, ctx)
  else if (noFill) fill = '#00000000'
  else if (solidFill) fill = resolveColor(solidFill, ctx.theme) ?? '#00000000'

  const outline = ln
    ? {
        color: resolveColor(directChild(ln, 'a:solidFill'), ctx.theme) ?? '#00000000',
        width: Math.max(1, Math.round(emu2px(parseInt(attr(ln, 'w') ?? '0', 10)))),
        style: dashOf(ln),
      }
    : undefined

  const shape: ShapeElement = {
    id: genId('s-'), type: 'shape', x, y, w, h,
    shapeKey: prst ? getShapeKey(prst) : 'rect',
    fill: fill ?? '#00000000',
    outline,
    flipH, flipV,
    rotate: rot ? Math.round(rot) : undefined,
    align: 'center', valign: 'middle', name: '形状',
  }
  if (customPath) shape.path = customPath
  if (shadow) shape.shadow = shadow

  // 形状内文本（首段样式近似）
  if (txBody) {
    const body = await txBodyToHTML(txBody, ctx.theme, pkg, ctx.partPath)
    const plain = body.html
      .replace(/<li[^>]*>/g, '\n').replace(/<p[^>]*>/g, '\n')
      .replace(/<[^>]+>/g, '').replace(/\n{2,}/g, '\n').trim()
    if (plain) {
      shape.text = plain
      const firstRun = firstDescendant(txBody, 'a:r')
      const rPr = firstRun ? firstDescendant(firstRun, 'a:rPr') : null
      shape.defaultColor = resolveColor(rPr ? directChild(rPr, 'a:solidFill') : null, ctx.theme) ?? '#FFFFFF'
      const sz = rPr ? attr(rPr, 'sz') : null
      if (sz) shape.fontSize = Math.round(parseInt(sz, 10) / 100 / 0.75)
      const firstP = firstDescendant(txBody, 'a:p')
      shape.align = alignFromAlgn(firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null)
      const bodyPr = directChild(txBody, 'a:bodyPr')
      const anchor = attr(bodyPr, 'anchor')
      shape.valign = anchor === 't' ? 'top' : anchor === 'b' ? 'bottom' : 'middle'
    }
  }
  return shape
}

/** 未识别的 shape 特性统一在此报告 */
export function skipShape(ctx: ParseContext, reason: string): void {
  addSkipped(ctx.report, reason)
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/elements/shape.test.ts`
Expected: PASS（2 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/elements/shape.ts src/core/import/elements/shape.test.ts
git commit -m "feat: 形状解析（渐变/阴影/custGeom/线条箭头/占位符继承）"
```

---

### Task 10: elements/picture.ts（图片 / 音视频）

**Files:**
- Create: `src/core/import/elements/picture.ts`
- Test: `src/core/import/elements/picture.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/elements/picture.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parsePictureEl } from './picture'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/a.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/v.mp4"/>
</Relationships>`)
  zip.file('ppt/media/a.png', 'PNGDATA')
  zip.file('ppt/media/v.mp4', 'MP4DATA')
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

const ctxBase = (pkg: PptxPackage) => ({
  pkg,
  partPath: 'ppt/slides/slide1.xml',
  theme,
  report: { skipped: {} } as ImportReport,
  scale: { x: 1, y: 1 },
  placeholders: new Map(),
})

describe('parsePictureEl', () => {
  it('图片：src + srcRect 裁剪', async () => {
    const pkg = await makePkg()
    const xml = `<p:pic xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:nvPicPr><p:cNvPr id="2" name="Img"/></p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="rId1"/>
        <a:srcRect l="10000" t="20000" r="10000" b="20000"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>`
    const result = await parsePictureEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'image') throw new Error('expected image')
    expect(result.src).toMatch(/^data:image\/png;base64,/)
    expect(result.clip).toEqual({ x: 0.1, y: 0.2, w: 0.8, h: 0.6 })
  })

  it('视频：videoFile → video 元素，blip 作海报帧', async () => {
    const pkg = await makePkg()
    const xml = `<p:pic xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">
      <p:nvPicPr><p:cNvPr id="2" name="Vid"/>
        <p:nvPr><a:videoFile r:link="rId2"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
    </p:pic>`
    const result = await parsePictureEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'video') throw new Error('expected video')
    expect(result.src).toMatch(/^data:video\/mp4;base64,/)
    expect(result.poster).toMatch(/^data:image\/png;base64,/)
    expect(result.loop).toBe(false)
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/elements/picture.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/elements/picture.ts
/** p:pic → image / video / audio */
import { attr, directChild, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { parseShadow } from './shape'
import { genId } from '../../utils/id'
import { mapX, mapY, addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { AudioElement, ImageElement, PPTElement, VideoElement } from '../../../types/slides'

function geomOf(node: Element, ctx: ParseContext, xf: GroupXform): { x: number; y: number; w: number; h: number } | null {
  const xfrm = firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!xfrm || !off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)
  return {
    x: Math.round(mapX(xf, ex) / 9525 * ctx.scale.x),
    y: Math.round(mapY(xf, ey) / 9525 * ctx.scale.y),
    w: Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x)),
    h: Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y)),
  }
}

/** a:srcRect（十万分之一百分比）→ 内部 clip 比例矩形 */
export function parseSrcRect(srcRect: Element | null): ImageElement['clip'] | undefined {
  if (!srcRect) return undefined
  const l = parseInt(attr(srcRect, 'l') ?? '0', 10) / 100000
  const t = parseInt(attr(srcRect, 't') ?? '0', 10) / 100000
  const r = parseInt(attr(srcRect, 'r') ?? '0', 10) / 100000
  const b = parseInt(attr(srcRect, 'b') ?? '0', 10) / 100000
  if (!l && !t && !r && !b) return undefined
  return { x: l, y: t, w: Math.max(0, 1 - l - r), h: Math.max(0, 1 - t - b) }
}

export async function parsePictureEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement | null> {
  const geom = geomOf(node, ctx, xf)
  if (!geom) return null

  const blipFill = firstDescendant(node, 'p:blipFill') ?? firstDescendant(node, 'a:blipFill')
  const blip = blipFill ? directChild(blipFill, 'a:blip') : null
  const embedId = attr(blip, 'r:embed')
  const poster = embedId ? await pkg.mediaDataUrl(await pkg.relTarget(ctx.partPath, embedId) ?? '') : undefined

  const nvPr = firstDescendant(firstDescendant(node, 'p:nvPicPr') ?? node, 'p:nvPr')
  const videoFile = nvPr ? directChild(nvPr, 'a:videoFile') : null
  const audioFile = nvPr ? directChild(nvPr, 'a:audioFile') : null
  const mediaLinkId = attr(videoFile, 'r:link') ?? attr(audioFile, 'r:link')
  const mediaPath = mediaLinkId ? await pkg.relTarget(ctx.partPath, mediaLinkId) : null
  const mediaSrc = mediaPath ? await pkg.mediaDataUrl(mediaPath) : undefined

  if ((videoFile || audioFile) && mediaSrc) {
    const base = { ...geom, loop: false, autoPlay: false }
    if (videoFile) {
      const video: VideoElement = { id: genId('v-'), type: 'video', src: mediaSrc, ...base, name: '视频' }
      if (poster) video.poster = poster
      return video
    }
    const audio: AudioElement = { id: genId('a-'), type: 'audio', src: mediaSrc, ...base, name: '音频' }
    return audio
  }

  if (!poster) {
    addSkipped(ctx.report, 'missingImage')
    return null
  }

  const spPr = firstDescendant(node, 'p:spPr') ?? firstDescendant(node, 'a:spPr')
  const xfrm = firstDescendant(node, 'a:xfrm')
  const ln = spPr ? directChild(spPr, 'a:ln') : null
  const image: ImageElement = {
    id: genId('i-'), type: 'image', src: poster, ...geom,
    name: '图片',
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: attr(xfrm, 'flipV') === '1' || undefined,
  }
  const clip = parseSrcRect(blipFill ? directChild(blipFill, 'a:srcRect') : null)
  if (clip) image.clip = clip
  const shadow = parseShadow(spPr ? directChild(spPr, 'a:effectLst') : null, ctx)
  if (shadow) image.shadow = shadow
  if (ln) {
    const color = resolveColor(directChild(ln, 'a:solidFill'), ctx.theme)
    const width = Math.max(1, Math.round(parseInt(attr(ln, 'w') ?? '0', 10) / 9525))
    if (color) image.outline = { color, width, style: 'solid' }
  }
  return image
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/elements/picture.test.ts`
Expected: PASS（2 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/elements/picture.ts src/core/import/elements/picture.test.ts
git commit -m "feat: 图片/音视频解析（srcRect 裁剪/海报帧/边框阴影）"
```

---

### Task 11: elements/table.ts（表格 + 首 run 样式 + 合并矩阵）

**Files:**
- Create: `src/core/import/elements/table.ts`
- Test: `src/core/import/elements/table.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/elements/table.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseTableEl, buildCellMatrix } from './table'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport, TableCell } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

const ctxBase = (pkg: PptxPackage) => ({
  pkg,
  partPath: 'ppt/slides/slide1.xml',
  theme,
  report: { skipped: {} } as ImportReport,
  scale: { x: 1, y: 1 },
  placeholders: new Map(),
})

describe('buildCellMatrix', () => {
  it('gridSpan/rowSpan/vMerge/hMerge → colspan/rowspan + null 填充', () => {
    const tbl = el(`<a:tbl xmlns:a="urn:a">
      <a:tr h="500000">
        <a:tc gridSpan="2" rowSpan="2"><a:txBody><a:p><a:r><a:t>合并</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
      <a:tr h="500000">
        <a:tc hMerge="1"/><a:tc vMerge="1"/>
        <a:tc><a:txBody><a:p><a:r><a:t>D</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
    </a:tbl>`)
    const { cells } = buildCellMatrix(tbl, theme)
    const expectCells = cells as TableCell[][]
    expect(expectCells[0][0]?.text).toBe('合并')
    expect(expectCells[0][0]?.colspan).toBe(2)
    expect(expectCells[0][0]?.rowspan).toBe(2)
    expect(expectCells[0][1]?.text).toBe('B')
    expect(expectCells[1][0]).toBeNull()
    expect(expectCells[1][1]).toBeNull()
    expect(expectCells[1][2]?.text).toBe('D')
  })
})

describe('parseTableEl', () => {
  it('首 run 样式进 style 字段；行高列宽归一化', async () => {
    const pkg = await makePkg()
    const xml = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblGrid><a:gridCol w="952500"/><a:gridCol w="952500"/></a:tblGrid>
          <a:tr h="952500">
            <a:tc><a:txBody><a:p><a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="D14424"/></a:solidFill></a:rPr><a:t>单元格</a:t></a:r></a:p></a:txBody></a:tc>
            <a:tc><a:txBody><a:p><a:r><a:t>普通</a:t></a:r></a:p></a:txBody></a:tc>
          </a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`
    const result = await parseTableEl(el(xml), ctxBase(pkg), IDENTITY_XFORM, pkg)
    if (result?.type !== 'table') throw new Error('expected table')
    expect(result.colSizes).toEqual([0.5, 0.5])
    expect(result.cells[0][0]?.style?.bold).toBe(true)
    expect(result.cells[0][0]?.style?.color).toBe('#D14424')
    expect(result.cells[0][0]?.style?.fontsize).toBe(27)
    expect(result.cells[0][1]?.text).toBe('普通')
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/elements/table.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/elements/table.ts
/** p:graphicFrame(a:tbl) → table 元素（首 run 样式提取；gridSpan/rowSpan/vMerge/hMerge 矩阵） */
import { attr, directChild, directChildren, firstDescendant } from '../xml'
import { resolveColor } from '../styles'
import { genId } from '../../utils/id'
import { mapX, mapY } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { PptxTheme } from '../theme'
import type { TableElement, TableCell } from '../../../types/slides'

function cellTextOf(tc: Element): string {
  const txBody = firstDescendant(tc, 'a:txBody')
  if (!txBody) return ''
  return directChildren(txBody, 'a:p')
    .map((p) => directChildren(p, 'a:r').map((r) => firstDescendant(r, 'a:t')?.textContent ?? '').join(''))
    .join('\n')
    .trim()
}

function cellStyleOf(tc: Element, theme: PptxTheme): TableCell['style'] {
  const style: NonNullable<TableCell['style']> = { color: '#333333', align: 'left', valign: 'middle' }
  const txBody = firstDescendant(tc, 'a:txBody')
  const firstRun = txBody ? firstDescendant(txBody, 'a:r') : null
  const rPr = firstRun ? firstDescendant(firstRun, 'a:rPr') : null
  if (rPr) {
    if (attr(rPr, 'b') === '1') style.bold = true
    if (attr(rPr, 'i') === '1') style.italic = true
    if (attr(rPr, 'u') === 'sng') style.underline = true
    const color = resolveColor(directChild(rPr, 'a:solidFill'), theme)
    if (color) style.color = color
    const sz = attr(rPr, 'sz')
    if (sz) style.fontsize = Math.round(parseInt(sz, 10) / 100 / 0.75)
  }
  const firstP = txBody ? firstDescendant(txBody, 'a:p') : null
  const algn = firstP ? attr(directChild(firstP, 'a:pPr'), 'algn') : null
  if (algn === 'ctr') style.align = 'center'
  else if (algn === 'r') style.align = 'right'
  const tcPr = directChild(tc, 'a:tcPr')
  const fill = tcPr ? resolveColor(directChild(tcPr, 'a:solidFill'), theme) : undefined
  if (fill) style.backcolor = fill
  return style
}

/** a:tbl → cells 矩阵（合并原点带 colspan/rowspan，被并格为 null） */
export function buildCellMatrix(tbl: Element, theme: PptxTheme): { cells: Array<Array<TableCell | null>> } {
  const trs = directChildren(tbl, 'a:tr')
  const cells: Array<Array<TableCell | null>> = []
  const occupied: Array<Array<boolean>> = []
  for (let r = 0; r < trs.length; r += 1) {
    cells.push([])
    occupied.push([])
  }
  trs.forEach((tr, r) => {
    let c = 0
    for (const tc of directChildren(tr, 'a:tc')) {
      while (occupied[r][c]) {
        cells[r][c] = null
        c += 1
      }
      const colspan = parseInt(attr(tc, 'gridSpan') ?? '1', 10) || 1
      const rowspan = parseInt(attr(tc, 'rowSpan') ?? '1', 10) || 1
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') {
        cells[r][c] = null
        occupied[r][c] = true
        c += 1
        continue
      }
      cells[r][c] = {
        text: cellTextOf(tc),
        colspan: colspan > 1 ? colspan : undefined,
        rowspan: rowspan > 1 ? rowspan : undefined,
        style: cellStyleOf(tc, theme),
      }
      for (let dr = 0; dr < rowspan; dr += 1) {
        for (let dc = 0; dc < colspan; dc += 1) {
          const rr = r + dr
          const cc = c + dc
          if (rr < cells.length) occupied[rr][cc] = true
        }
      }
      // 填充被并格为 null
      for (let dr = 0; dr < rowspan; dr += 1) {
        for (let dc = 0; dc < colspan; dc += 1) {
          if (dr === 0 && dc === 0) continue
          const rr = r + dr
          const cc = c + dc
          if (rr < cells.length) cells[rr][cc] = null
        }
      }
      c += colspan
    }
  })
  return { cells }
}

export async function parseTableEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  _pkg: PptxPackage,
): Promise<TableElement | null> {
  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)
  const x = Math.round(mapX(xf, ex) / 9525 * ctx.scale.x)
  const y = Math.round(mapY(xf, ey) / 9525 * ctx.scale.y)
  const w = Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x))
  const h = Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y))

  const tbl = firstDescendant(node, 'a:tbl')
  if (!tbl) return null
  const gridCols = directChildren(tbl, 'a:tblGrid').flatMap((g) => directChildren(g, 'a:gridCol'))
  const colSizesRaw = gridCols.map((c) => parseInt(attr(c, 'w') ?? '0', 10))
  const colTotal = colSizesRaw.reduce((a, b) => a + b, 0) || 1
  const trs = directChildren(tbl, 'a:tr')
  const rowSizesRaw = trs.map((tr) => parseInt(attr(tr, 'h') ?? '0', 10))
  const rowTotal = rowSizesRaw.reduce((a, b) => a + b, 0) || 1

  const { cells } = buildCellMatrix(tbl, ctx.theme)
  const table: TableElement = {
    id: genId('tb-'), type: 'table', x, y, w, h,
    colSizes: colSizesRaw.map((c) => c / colTotal),
    rowSizes: rowSizesRaw.map((r) => r / rowTotal),
    cells, name: '表格',
  }
  return table
}
```

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/elements/table.test.ts`
Expected: PASS（2 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/elements/table.ts src/core/import/elements/table.test.ts
git commit -m "feat: 表格解析（合并矩阵/首 run 样式提取）"
```

---

### Task 12: elements/chart.ts（图表导入）与 elements/group.ts（组合展开）与 elements/index.ts（分发）

**Files:**
- Create: `src/core/import/elements/chart.ts`
- Create: `src/core/import/elements/group.ts`
- Create: `src/core/import/elements/index.ts`
- Test: `src/core/import/elements/chart.test.ts`、`src/core/import/elements/group.test.ts`

- [ ] **Step 1: 写失败测试（chart）**

```typescript
// src/core/import/elements/chart.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseChartEl, mapChartType } from './chart'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

const CHART_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="urn:a">
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列一</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt idx="0"><c:v>一月</c:v></c:pt><c:pt idx="1"><c:v>二月</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`

async function makeCtx() {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="urn:p" xmlns:c="urn:c"/>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`)
  zip.file('ppt/charts/chart1.xml', CHART_XML)
  const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
  const ctx = {
    pkg,
    partPath: 'ppt/slides/slide1.xml',
    theme,
    report: { skipped: {} } as ImportReport,
    scale: { x: 1, y: 1 },
    placeholders: new Map(),
  }
  return { pkg, ctx }
}

const FRAME_XML = `<p:graphicFrame xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c">
  <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3" xmlns:r="urn:r"/>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('mapChartType', () => {
  it('barChart 组合映射', () => {
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'clustered' })).toBe('bar-cluster')
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'stacked' })).toBe('bar-stack')
    expect(mapChartType('barChart', { barDir: 'col', grouping: 'percentStacked' })).toBe('bar-percent')
    expect(mapChartType('barChart', { barDir: 'bar', grouping: 'clustered' })).toBe('bar-horizontal')
    expect(mapChartType('barChart', { barDir: 'bar', grouping: 'stacked' })).toBe('bar-horizontal-stack')
  })
  it('line/pie/area/scatter/radar 映射', () => {
    expect(mapChartType('lineChart', { marker: true })).toBe('line-marker')
    expect(mapChartType('lineChart', {})).toBe('line')
    expect(mapChartType('lineChart', { grouping: 'stacked' })).toBe('line-stack')
    expect(mapChartType('pieChart', {})).toBe('pie')
    expect(mapChartType('doughnutChart', {})).toBe('pie-doughnut')
    expect(mapChartType('areaChart', {})).toBe('area')
    expect(mapChartType('areaChart', { grouping: 'stacked' })).toBe('area-stack')
    expect(mapChartType('scatterChart', {})).toBe('scatter')
    expect(mapChartType('radarChart', {})).toBe('radar')
  })
})

describe('parseChartEl', () => {
  it('解析 rels 关联的 chart part → ChartElement', async () => {
    const { pkg, ctx } = await makeCtx()
    const result = await parseChartEl(el(FRAME_XML), ctx, IDENTITY_XFORM, pkg)
    if (result?.type !== 'chart') throw new Error('expected chart')
    expect(result.chartType).toBe('bar-cluster')
    expect(result.data.labels).toEqual(['一月', '二月'])
    expect(result.data.series).toEqual([{ name: '系列一', values: [10, 25] }])
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/elements/chart.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现 chart.ts**

```typescript
// src/core/import/elements/chart.ts
/** p:graphicFrame(c:chart) → chart 元素（OOXML chart part 缓存数据 → ChartData） */
import { attr, directChild, directChildren, firstDescendant, parseXML } from '../xml'
import { genId } from '../../utils/id'
import { mapX, mapY, addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { ChartElement, ChartType } from '../../../types/slides'

interface ChartKindHints {
  barDir?: string
  grouping?: string
  marker?: boolean
}

/** chart 部件节点名 + 提示 → 内部 ChartType */
export function mapChartType(kindNode: string, hints: ChartKindHints): ChartType | null {
  const grouping = hints.grouping ?? 'clustered'
  switch (kindNode) {
    case 'c:barChart':
      if (hints.barDir === 'bar') return grouping === 'stacked' ? 'bar-horizontal-stack' : grouping === 'percentStacked' ? 'bar-horizontal-stack' : 'bar-horizontal'
      return grouping === 'stacked' ? 'bar-stack' : grouping === 'percentStacked' ? 'bar-percent' : 'bar-cluster'
    case 'c:lineChart':
      if (grouping === 'stacked') return 'line-stack'
      return hints.marker ? 'line-marker' : 'line'
    case 'c:pieChart': return 'pie'
    case 'c:doughnutChart': return 'pie-doughnut'
    case 'c:areaChart': return grouping === 'stacked' ? 'area-stack' : 'area'
    case 'c:scatterChart': return 'scatter'
    case 'c:radarChart': return 'radar'
    default: return null
  }
}

function cacheValues(ref: Element | null, tag: string): string[] {
  const cache = ref ? firstDescendant(ref, tag) : null
  if (!cache) return []
  return directChildren(cache, 'c:pt')
    .map((pt) => firstDescendant(pt, 'c:v')?.textContent ?? '')
    .filter((v) => v !== '')
}

export async function parseChartEl(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<ChartElement | null> {
  const chartRef = firstDescendant(node, 'c:chart')
  const rId = attr(chartRef, 'r:id')
  if (!rId) return null
  const chartPath = await pkg.relTarget(ctx.partPath, rId)
  const xml = chartPath ? await pkg.text(chartPath) : null
  if (!xml) {
    addSkipped(ctx.report, 'missingChart')
    return null
  }
  const root = parseXML(xml).documentElement
  const plotArea = firstDescendant(root, 'c:plotArea')
  if (!plotArea) {
    addSkipped(ctx.report, 'missingChart')
    return null
  }

  const KIND_NODES = ['c:barChart', 'c:lineChart', 'c:pieChart', 'c:doughnutChart', 'c:areaChart', 'c:scatterChart', 'c:radarChart']
  const kindNode = directChildren(plotArea, '*').find((c) => KIND_NODES.includes(c.nodeName))
    ?? KIND_NODES.map((k) => directChild(plotArea, k)).find(Boolean) ?? null
  if (!kindNode) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  const hints: ChartKindHints = {
    barDir: attr(directChild(kindNode, 'c:barDir'), 'val') ?? undefined,
    grouping: attr(directChild(kindNode, 'c:grouping'), 'val') ?? undefined,
    marker: attr(directChild(kindNode, 'c:marker'), 'val') === '1',
  }
  const chartType = mapChartType(kindNode.nodeName, hints)
  if (!chartType) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }

  // scatter 使用 c:xVal/c:yVal，其余用 c:cat/c:val
  const sers = directChildren(kindNode, 'c:ser')
  if (!sers.length) {
    addSkipped(ctx.report, 'unknownChart')
    return null
  }
  const labels = cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:strCache')
  const labelsNum = labels.length ? labels : cacheValues(firstDescendant(sers[0], 'c:cat'), 'c:numCache')
  const series = sers.map((ser) => ({
    name: cacheValues(firstDescendant(ser, 'c:tx'), 'c:strCache')[0] ?? '系列',
    values: cacheValues(firstDescendant(ser, 'c:val'), 'c:numCache').map(Number),
  }))

  const xfrm = firstDescendant(node, 'p:xfrm') ?? firstDescendant(node, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  if (!off || !ext) return null
  const ex = parseInt(attr(off, 'x') ?? '0', 10)
  const ey = parseInt(attr(off, 'y') ?? '0', 10)
  const ew = parseInt(attr(ext, 'cx') ?? '0', 10)
  const eh = parseInt(attr(ext, 'cy') ?? '0', 10)

  const chart: ChartElement = {
    id: genId('ch-'), type: 'chart', chartType,
    data: { labels: labelsNum, series },
    x: Math.round(mapX(xf, ex) / 9525 * ctx.scale.x),
    y: Math.round(mapY(xf, ey) / 9525 * ctx.scale.y),
    w: Math.max(1, Math.round((mapX(xf, ex + ew) - mapX(xf, ex)) / 9525 * ctx.scale.x)),
    h: Math.max(1, Math.round((mapY(xf, ey + eh) - mapY(xf, ey)) / 9525 * ctx.scale.y)),
    name: '图表',
  }
  const titleEls = firstDescendant(root, 'c:title') ? directChildren(firstDescendant(root, 'c:title')!, 'c:tx') : []
  const title = titleEls.map((tx) => firstDescendant(tx, 'a:t')?.textContent ?? '').join('').trim()
  if (title) chart.title = title
  return chart
}
```

注意 scatterChart 的 `c:xVal/c:yVal`：`series` 映射时对 scatter 补充分支（实现时在 `sers.map` 内处理：`const xVal = cacheValues(firstDescendant(ser, 'c:xVal'), 'c:numCache').map(Number)`；labels 取 xVal，values 取 yVal）。

- [ ] **Step 4: 实现 group.ts 与 elements/index.ts**

```typescript
// src/core/import/elements/group.ts
/** p:grpSp 递归展开：组合变换（off/ext vs chOff/chExt）折算到子元素 */
import { attr, directChild, firstDescendant } from '../xml'
import { mapX, mapY } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PPTElement } from '../../../types/slides'
import { parseSpTreeNode } from './index'

/** 由 grpSp 的 xfrm 计算子空间 → 当前空间的复合变换 */
export function childXform(grpSp: Element, xf: GroupXform): GroupXform {
  const xfrm = firstDescendant(grpSp, 'a:xfrm')
  const off = xfrm ? directChild(xfrm, 'a:off') : null
  const ext = xfrm ? directChild(xfrm, 'a:ext') : null
  const chOff = xfrm ? directChild(xfrm, 'a:chOff') : null
  const chExt = xfrm ? directChild(xfrm, 'a:chExt') : null
  if (!off || !ext || !chOff || !chExt) return xf
  const offX = parseInt(attr(off, 'x') ?? '0', 10)
  const offY = parseInt(attr(off, 'y') ?? '0', 10)
  const chX = parseInt(attr(chOff, 'x') ?? '0', 10)
  const chY = parseInt(attr(chOff, 'y') ?? '0', 10)
  const cx = parseInt(attr(ext, 'cx') ?? '0', 10)
  const cy = parseInt(attr(ext, 'cy') ?? '0', 10)
  const chCx = parseInt(attr(chExt, 'cx') ?? '0', 10) || 1
  const chCy = parseInt(attr(chExt, 'cy') ?? '0', 10) || 1
  const sx = chCx ? cx / chCx : 1
  const sy = chCy ? cy / chCy : 1
  return {
    // 复合：outer( inner(v) ) = xf.ox + (offX + (v - chX) * s) * xf.sx
    ox: xf.ox + (offX - chX * sx) * xf.sx,
    oy: xf.oy + (offY - chY * sy) * xf.sy,
    sx: xf.sx * sx,
    sy: xf.sy * sy,
    rot: xf.rot + (parseInt(attr(xfrm, 'rot') ?? '0', 10)),
  }
}

/** 展开组合，返回打平后的元素列表（子元素失败逐个跳过） */
export async function parseGroupEl(
  grpSp: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  const cxf = childXform(grpSp, ctx, xf)
  const out: PPTElement[] = []
  const spTree = firstDescendant(grpSp, 'p:grpSpPr')?.parentElement ?? grpSp
  for (const child of Array.from(spTree.children)) {
    const parsed = await parseSpTreeNode(child, ctx, cxf, pkg)
    out.push(...parsed)
  }
  return out
}
```

**修正说明（实现时注意）：** `childXform(grpSp, ctx, xf)` 签名应为 `(grpSp, xf)` 二参（ctx 未使用）；`parseGroupEl` 中遍历应直接用 `Array.from(grpSp.children)`（组合子节点就是 grpSp 的直接子级，不需要经由 grpSpPr 的 parentElement 兜底）。按此二点修正后实现。

```typescript
// src/core/import/elements/index.ts
/** spTree 子节点分发：p:sp / p:cxnSp / p:pic / p:graphicFrame / p:grpSp */
import { addSkipped } from '../context'
import type { GroupXform, ParseContext } from '../context'
import type { PptxPackage } from '../package'
import type { PPTElement } from '../../../types/slides'
import { parseShapeEl } from './shape'
import { parsePictureEl } from './picture'
import { parseTableEl } from './table'
import { parseChartEl } from './chart'
import { parseGroupEl } from './group'

export async function parseSpTreeNode(
  node: Element,
  ctx: ParseContext,
  xf: GroupXform,
  pkg: PptxPackage,
): Promise<PPTElement[]> {
  try {
    switch (node.nodeName) {
      case 'p:sp':
      case 'p:cxnSp': {
        const el = await parseShapeEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'p:pic': {
        const el = await parsePictureEl(node, ctx, xf, pkg)
        return el ? [el] : []
      }
      case 'p:graphicFrame': {
        const el = (await parseTableEl(node, ctx, xf, pkg)) ?? (await parseChartEl(node, ctx, xf, pkg))
        return el ? [el] : []
      }
      case 'p:grpSp':
        return await parseGroupEl(node, ctx, xf, pkg)
      default:
        return []
    }
  } catch {
    addSkipped(ctx.report, 'elementParseFailed')
    return []
  }
}
```

- [ ] **Step 5: 写失败测试（group）**

```typescript
// src/core/import/elements/group.test.ts
import { describe, expect, it } from 'vitest'
import { parseSpTreeNode } from './index'
import { PptxPackage } from '../package'
import { DEFAULT_SCHEME, type PptxTheme } from '../theme'
import { IDENTITY_XFORM } from '../context'
import type { ImportReport } from '../../../types/slides'

const theme: PptxTheme = {
  schemeColors: { ...DEFAULT_SCHEME },
  majorFont: 'Calibri', minorFont: 'Calibri', colorMap: {},
}

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

describe('parseGroupEl（经 parseSpTreeNode 分发）', () => {
  it('组合子元素坐标按 off/chOff/ext/chExt 折算并打平', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="urn:p"/>`)
    const pkg = await PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
    const ctx = {
      pkg,
      partPath: 'ppt/slides/slide1.xml',
      theme,
      report: { skipped: {} } as ImportReport,
      scale: { x: 1, y: 1 },
      placeholders: new Map(),
    }
    // 组合：off=(0,0) ext=(2000000,1000000) chOff=(0,0) chExt=(1000000,1000000) → x 放大 2 倍
    // 子元素 chOff 空间 x=0,y=0,w=500000,h=500000 → 绝对 x=0,y=0,w=1000000,h=1000000 EMU → px(96dpi) w=105,h=105
    const grp = el(`<p:grpSp xmlns:p="urn:p" xmlns:a="urn:a">
      <p:nvGrpSpPr><p:cNvPr id="1" name="组"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>
        <a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/>
        <a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="子矩形"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="42A5F5"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>
    </p:grpSp>`)
    const els = await parseSpTreeNode(grp, ctx, IDENTITY_XFORM, pkg)
    expect(els).toHaveLength(1)
    const shape = els[0]
    if (shape.type !== 'shape') throw new Error('expected shape')
    expect(shape.w).toBe(105) // 1000000 EMU / 9525 = 104.98 → 105
    expect(shape.shapeKey).toBe('roundRect')
    expect(shape.fill).toBe('#42A5F5')
  })
})
```

- [ ] **Step 6: 运行测试通过**

Run: `npx vitest run src/core/import/elements/chart.test.ts src/core/import/elements/group.test.ts`
Expected: PASS（3 个用例）

- [ ] **Step 7: Commit**

```bash
git add src/core/import/elements/chart.ts src/core/import/elements/group.ts src/core/import/elements/index.ts src/core/import/elements/chart.test.ts src/core/import/elements/group.test.ts
git commit -m "feat: 图表导入/组合展开/元素分发器"
```

---

### Task 13: master.ts（版式/母版继承链、背景、装饰、占位符表）

**Files:**
- Create: `src/core/import/master.ts`
- Test: `src/core/import/master.test.ts`

- [ ] **Step 1: 写失败测试**

```typescript
// src/core/import/master.test.ts
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseBackgroundFill, collectPlaceholders, findAncestry } from './master'
import { PptxPackage } from './package'

function el(xml: string): Element {
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

async function makePkg(): Promise<PptxPackage> {
  const zip = new JSZip()
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout xmlns:p="urn:p"/>')
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
  return PptxPackage.load(await zip.generateAsync({ type: 'blob' }))
}

describe('master.ts', () => {
  it('findAncestry：slide → layout → master 链', async () => {
    const pkg = await makePkg()
    const chain = await findAncestry(pkg, 'ppt/slides/slide1.xml')
    expect(chain.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml')
    expect(chain.masterPath).toBe('ppt/slideMasters/slideMaster1.xml')
  })

  it('collectPlaceholders：ph 的 idx/type 作 key', () => {
    const tree = el(`<p:spTree xmlns:p="urn:p" xmlns:a="urn:a">
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3"/><p:nvPr><p:ph idx="12"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4"/></p:nvSpPr>
        <p:spPr/>
      </p:sp>
    </p:spTree>`)
    const map = collectPlaceholders(tree)
    expect(map.get('title')).toEqual({ x: 100, y: 200, w: 300, h: 400 })
    expect(map.get('12')).toEqual({ x: 1, y: 2, w: 3, h: 4 })
    expect(map.size).toBe(2)
  })

  it('parseBackgroundFill：solid → Background；无 bg 返回 undefined', () => {
    const theme = { schemeColors: { lt1: '#FFFFFF' } as Record<string, string>, majorFont: '', minorFont: '', colorMap: {} }
    const bg = el(`<p:bg xmlns:p="urn:p" xmlns:a="urn:a">
      <p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:effectLst/></p:bgPr>
    </p:bg>`)
    const result = parseBackgroundFill(bg, theme)
    expect(result).toEqual({ type: 'solid', color: '#112233' })
    expect(parseBackgroundFill(null, theme)).toBeUndefined()
  })
})
```

- [ ] **Step 2: 运行确认失败**

Run: `npx vitest run src/core/import/master.test.ts`
Expected: FAIL

- [ ] **Step 3: 实现**

```typescript
// src/core/import/master.ts
/** slide ancestry：layout/master 链、继承背景、装饰元素、占位符位置表 */
import { attr, directChild, firstDescendant, parseXML } from './xml'
import { resolveColor } from './styles'
import { parseThemeForMaster } from './theme'
import type { PptxTheme } from './theme'
import type { PptxPackage } from './package'
import type { Background, PPTElement } from '../../types/slides'
import { parseSpTreeNode } from './elements'
import { IDENTITY_XFORM, addSkipped } from './context'
import type { ImportReport } from '../../types/slides'

export interface SlideAncestry {
  layoutPath: string | null
  masterPath: string | null
  theme: PptxTheme
  /** 版式/母版继承背景（slide 自身 bg 优先，由 index.ts 处理） */
  background?: Background
  /** 版式 + 母版的非占位符装饰元素模板 */
  decorations: PPTElement[]
  /** 占位符位置表（EMU；key = idx ?? type） */
  placeholders: Map<string, { x: number; y: number; w: number; h: number }>
}

async function relTargetByType(pkg: PptxPackage, partPath: string, keyword: string): Promise<string | null> {
  const rels = await pkg.rels(partPath)
  for (const rel of rels.values()) {
    if (rel.mode !== 'External' && rel.target.includes(`/${keyword}/`)) return rel.target
  }
  return null
}

/** slide → layout → master 部件链 */
export async function findAncestry(pkg: PptxPackage, slidePath: string): Promise<{ layoutPath: string | null; masterPath: string | null }> {
  const layoutPath = await relTargetByType(pkg, slidePath, '/slideLayouts/')
  const masterPath = layoutPath ? await relTargetByType(pkg, layoutPath, '/slideMasters/') : null
  return { layoutPath, masterPath }
}

/** 收集 spTree 中带 p:ph 且有 xfrm 的占位符位置（EMU） */
export function collectPlaceholders(spTree: Element): Map<string, { x: number; y: number; w: number; h: number }> {
  const map = new Map<string, { x: number; y: number; w: number; h: number }>()
  for (const sp of Array.from(spTree.children)) {
    if (sp.nodeName !== 'p:sp') continue
    const ph = firstDescendant(firstDescendant(sp, 'p:nvSpPr') ?? sp, 'p:ph')
    if (!ph) continue
    const xfrm = firstDescendant(sp, 'a:xfrm')
    const off = xfrm ? directChild(xfrm, 'a:off') : null
    const ext = xfrm ? directChild(xfrm, 'a:ext') : null
    if (!off || !ext) continue
    const key = attr(ph, 'idx') ?? attr(ph, 'type') ?? ''
    if (!key || map.has(key)) continue
    map.set(key, {
      x: parseInt(attr(off, 'x') ?? '0', 10),
      y: parseInt(attr(off, 'y') ?? '0', 10),
      w: parseInt(attr(ext, 'cx') ?? '0', 10),
      h: parseInt(attr(ext, 'cy') ?? '0', 10),
    })
  }
  return map
}

/** p:bg → Background（solid/grad/blip；bgRef 以 srgbClr 走 solid） */
export function parseBackgroundFill(bg: Element | null, theme: PptxTheme): Background | undefined {
  if (!bg) return undefined
  const bgPr = directChild(bg, 'p:bgPr')
  const bgRef = directChild(bg, 'p:bgRef')
  if (bgPr) {
    const solid = directChild(bgPr, 'a:solidFill')
    if (solid) {
      const color = resolveColor(solid, theme)
      if (color) return { type: 'solid', color }
    }
    const grad = directChild(bgPr, 'a:gradFill')
    if (grad) return { type: 'gradient' }
    const blip = firstDescendant(bgPr, 'a:blip')
    return { type: 'image' } // src 由调用方（需 pkg rels）补充
  }
  if (bgRef) {
    const color = resolveColor(bgRef, theme)
    if (color) return { type: 'solid', color }
  }
  return undefined
}

/** 解析 slide 的 ancestry（layout/master 背景与装饰；装饰按 partPath 解析 rels） */
export async function parseSlideAncestry(pkg: PptxPackage, slidePath: string, report: ImportReport): Promise<SlideAncestry> {
  const { layoutPath, masterPath } = await findAncestry(pkg, slidePath)
  const theme = await parseThemeForMaster(pkg, masterPath)

  const placeholders = new Map<string, { x: number; y: number; w: number; h: number }>()
  const decorations: PPTElement[] = []
  let background: Background | undefined

  for (const partPath of [layoutPath, masterPath]) {
    if (!partPath) continue
    const xml = await pkg.text(partPath)
    if (!xml) continue
    const doc = parseXML(xml)
    const spTree = firstDescendant(firstDescendant(doc.documentElement, 'p:cSld') ?? doc.documentElement, 'p:spTree')
    if (!spTree) continue
    for (const [k, v] of collectPlaceholders(spTree)) {
      if (!placeholders.has(k)) placeholders.set(k, v)
    }
    const bg = firstDescendant(doc.documentElement, 'p:bg')
    const partBg = parseBackgroundFill(bg, theme)
    if (partBg && !background) background = partBg
    // 背景图片 src：bgPr/a:blip@r:embed → rels
    if (partBg?.type === 'image' && bg) {
      const blip = firstDescendant(bg, 'a:blip')
      const embedId = attr(blip, 'r:embed')
      if (embedId) {
        const target = await pkg.relTarget(partPath, embedId)
        const src = target ? await pkg.mediaDataUrl(target) : undefined
        if (src) background = { type: 'image', image: { src, size: 'cover' } }
      }
    }
    // 非占位符装饰元素
    const ctx = {
      pkg, partPath, theme, report,
      scale: { x: 1, y: 1 },
      placeholders,
    }
    for (const child of Array.from(spTree.children)) {
      if (!['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp'].includes(child.nodeName)) continue
      const ph = firstDescendant(child, 'p:ph')
      if (ph) continue
      const els = await parseSpTreeNode(child, ctx, IDENTITY_XFORM, pkg)
      decorations.push(...els)
    }
  }

  // 装饰元素需要按画布缩放：装饰在母版 EMU 空间，与 slide 相同 sldSz，无需额外缩放（index.ts 统一缩放）
  if (!background) {
    const lt1 = theme.schemeColors.lt1 ?? '#FFFFFF'
    background = { type: 'solid', color: lt1 }
  }
  return { layoutPath, masterPath, theme, background, decorations, placeholders }
}

export function skipBy(report: ImportReport, kind: string): void {
  addSkipped(report, kind)
}
```

**注意（实现时）：** 装饰元素经 `parseSpTreeNode` 解析时 ctx.scale 传 `{x:1,y:1}`，产出的坐标是 **px（源 96dpi 空间）**；index.ts 中须将装饰元素与 slide 元素用同一 scale 缩放 —— 由于 index.ts 会计算 `scale = {x: 1280/srcWpx, y: ...}`，装饰解析应改用该 scale（实现时给 `parseSlideAncestry` 增加 `scale` 参数传入）。

- [ ] **Step 4: 运行测试通过**

Run: `npx vitest run src/core/import/master.test.ts`
Expected: PASS（3 个用例）

- [ ] **Step 5: Commit**

```bash
git add src/core/import/master.ts src/core/import/master.test.ts
git commit -m "feat: 版式/母版继承链与背景、装饰、占位符表"
```

---

### Task 14: index.ts 管线组装 + 替换旧实现 + 测试迁移

**Files:**
- Create: `src/core/import/index.ts`
- Delete: `src/core/import/pptx.ts`
- Modify: `src/core/import/pptx.test.ts`（仅 import 路径）
- Modify: `src/components/dialogs/ImportDialog.tsx:6`（import 路径）

- [ ] **Step 1: 实现 index.ts**

```typescript
// src/core/import/index.ts
/** 导入入口：解压 → ancestry → 逐页解析 → Presentation（分层容错，失败项计入 report） */
import type { ImportReport, PPTElement, Presentation, Slide, Theme } from '../../types/slides'
import { createDefaultTheme } from '../../types/slides'
import { PptxPackage } from './package'
import { parseThemeForMaster } from './theme'
import type { PptxTheme } from './theme'
import { findAncestry, parseSlideAncestry, parseBackgroundFill } from './master'
import { parseSpTreeNode } from './elements'
import { IDENTITY_XFORM } from './context'
import { attr, directChild, firstDescendant, parseXML } from './xml'
import { genId } from '../utils/id'

export { resolveTarget } from './package'
export type { ImportReport } from '../../types/slides'

export interface ImportResult {
  presentation: Presentation
  report: ImportReport
}

const genIdPrefixOf = (id: string): string => `${id.slice(0, Math.max(1, id.indexOf('-') + 1))}`

/** 装饰元素模板克隆：深拷贝并重编 id，避免多页共享同一对象 */
function cloneDecoration(el: PPTElement): PPTElement {
  const clone = JSON.parse(JSON.stringify(el)) as PPTElement
  clone.id = genId(genIdPrefixOf(el.id))
  return clone
}

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
  }
}

/** 解析 pptx → Presentation + 兼容性报告 */
export async function importPPTXDetailed(file: File): Promise<ImportResult> {
  const pkg = await PptxPackage.load(file)
  const presXml = await pkg.text('ppt/presentation.xml')
  if (!presXml) throw new Error('不是有效的 PPTX 文件（缺少 presentation.xml）')
  const presDoc = parseXML(presXml)
  const root = presDoc.documentElement

  const sldSz = firstDescendant(root, 'p:sldSz')
  const srcW = parseInt(attr(sldSz, 'cx') ?? '12192000', 10)
  const srcH = parseInt(attr(sldSz, 'cy') ?? '6858000', 10)
  const ratio = srcW / srcH || 16 / 9
  const scale = { x: 1280 / (srcW / 9525), y: (1280 / ratio) / (srcH / 9525) }

  // 页面顺序：sldIdLst → rels；缺失时按 slideN.xml 编号兜底
  const sldIdLst = firstDescendant(root, 'p:sldIdLst')
  const slideIds = sldIdLst ? Array.from(sldIdLst.getElementsByTagName('p:sldId')) : []
  const rels = await pkg.rels('ppt/presentation.xml')
  const slidePaths = slideIds
    .map((id) => attr(id, 'r:id'))
    .map((rid) => (rid ? rels.get(rid)?.target : undefined))
    .filter((p): p is string => Boolean(p && !p.startsWith('http')))
  const paths = slidePaths.length
    ? slidePaths
    : Object.keys(zipFileIndex(pkg))
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => (parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10) - parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10)))
  if (!paths.length) throw new Error('PPTX 中没有幻灯片')

  const report: ImportReport = { skipped: {} }
  const themeCache = new Map<string, PptxTheme>()
  const slides: Slide[] = []

  for (const slidePath of paths) {
    const xml = await pkg.text(slidePath)
    if (!xml) continue
    try {
      const doc = parseXML(xml)
      const cSld = firstDescendant(doc.documentElement, 'p:cSld')
      const spTree = cSld ? firstDescendant(cSld, 'p:spTree') : null

      const { masterPath } = await findAncestry(pkg, slidePath)
      const cacheKey = masterPath ?? ''
      if (!themeCache.has(cacheKey)) {
        themeCache.set(cacheKey, await parseThemeForMaster(pkg, masterPath))
      }
      const theme = themeCache.get(cacheKey)!

      const ancestry = await parseSlideAncestry(pkg, slidePath, report, scale)
      const ctx = {
        pkg, partPath: slidePath, theme, report,
        scale, placeholders: ancestry.placeholders,
      }

      const elements: PPTElement[] = ancestry.decorations.map(cloneDecoration)
      if (spTree) {
        for (const child of Array.from(spTree.children)) {
          elements.push(...(await parseSpTreeNode(child, ctx, IDENTITY_XFORM, pkg)))
        }
      }

      // 背景：slide 自身 bg 优先 → 版式/母版 → 白色
      const bgEl = firstDescendant(doc.documentElement, 'p:bg')
      let background = parseBackgroundFill(bgEl, theme)
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

      slides.push({
        id: genId('slide-'),
        elements,
        background: background ?? ancestry.background ?? { type: 'solid', color: '#ffffff' },
      })
    } catch {
      report.skipped['slideParseFailed'] = (report.skipped['slideParseFailed'] ?? 0) + 1
    }
  }
  if (!slides.length) throw new Error('PPTX 中没有可解析的幻灯片')

  const firstTheme = themeCache.values().next().value
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

/** zip 内文件路径列表（JSZip files 视图的薄封装，便于测试注入） */
function zipFileIndex(pkg: PptxPackage): Record<string, unknown> {
  return (pkg as unknown as { zip: { files: Record<string, unknown> } }).zip.files
}

/** 兼容既有调用方：仅返回 Presentation */
export async function importPPTX(file: File): Promise<Presentation> {
  return (await importPPTXDetailed(file)).presentation
}
```

- [ ] **Step 2: 删除旧实现并迁移引用**

```bash
rm src/core/import/pptx.ts
```

- 修改 `src/core/import/pptx.test.ts` 首行 import：`from './pptx'` → `from './index'`（断言全部保留不动）
- 修改 `src/components/dialogs/ImportDialog.tsx:6`：`from '../../core/import/pptx'` → `from '../../core/import'`

- [ ] **Step 3: 运行导入相关全部测试**

Run: `npx vitest run src/core/import src/core/export`
Expected: 全部 PASS。**既有 pptx.test.ts 的断言是兼容性底线**：最小 pptx 的文本/形状解析结果（96px 坐标、`font-weight:bold`、`D14424`、roundRect 旋转 30°）必须不变。若失败，逐条修正新实现而非改断言。

- [ ] **Step 4: typecheck**

Run: `npx tsc -p tsconfig.json --noEmit 2>&1 | head -20`
Expected: 无错误

- [ ] **Step 5: Commit**

```bash
git add -A src/core/import src/components/dialogs/ImportDialog.tsx
git commit -m "feat: 分层导入管线 index.ts 替换旧实现，对外 API 不变"
```

---

### Task 15: 渲染器支持 ShapeElement.path（custGeom 可视化）

**Files:**
- Modify: `src/components/canvas/ElementRenderer.tsx:171`（ShapeItem 内）

- [ ] **Step 1: 修改路径取值**

`ShapeItem` 中：

```typescript
  const meta = getShapePath(el.shapeKey)
```

改为：

```typescript
  const meta = getShapePath(el.shapeKey)
  // custGeom 导入产出：优先使用元素自带的自定义路径（0-100 视口空间）
  const shapeD = el.path ?? meta.path
```

并将下方 `<path d={meta.path}` 改为 `<path d={shapeD}`。

- [ ] **Step 2: typecheck + 既有测试**

Run: `npx tsc -p tsconfig.json --noEmit && npx vitest run`
Expected: 全部 PASS（ShapeElement.path 为可选字段，其他 shape 消费方不受影响）

- [ ] **Step 3: Commit**

```bash
git add src/components/canvas/ElementRenderer.tsx
git commit -m "feat: 形状渲染支持自定义 path（custGeom 导入可视化）"
```

---

### Task 16: ImportDialog 兼容性报告

**Files:**
- Modify: `src/components/dialogs/ImportDialog.tsx`

- [ ] **Step 1: 修改 handleFile 的 pptx 分支**

```typescript
// import 处新增：
import { importPPTXDetailed } from '../../core/import'

// 组件外新增（模块级常量）：
const SKIP_LABELS: Record<string, string> = {
  missingImage: '缺失图片',
  missingChart: '缺失图表数据',
  unknownChart: '未识别图表',
  elementParseFailed: '无法解析的元素',
  slideParseFailed: '无法解析的页面',
  smartartFallback: 'SmartArt 已转图片',
  groupParseFailed: '组合解析失败',
}

// handleFile 的 else if (.pptx) 分支改为：
      } else if (file.name.toLowerCase().endsWith('.pptx')) {
        toast('正在解析 PPTX…')
        const { presentation: pres, report } = await importPPTXDetailed(file)
        useEditorStore.getState().pushHistory()
        useEditorStore.getState().replacePresentation(pres)
        const items = Object.entries(report.skipped).map(([k, v]) => `${SKIP_LABELS[k] ?? k} ×${v}`)
        if (items.length) {
          toast(`已导入 PPTX（${pres.slides.length} 页）；部分内容未完整还原：${items.join('、')}`, 'success')
        } else {
          toast(`已导入 PPTX（${pres.slides.length} 页）`, 'success')
        }
      }
```

- [ ] **Step 2: typecheck + 手动冒烟**

Run: `npx tsc -p tsconfig.json --noEmit`
Expected: 无错误
手动验证：`pnpm dev:demo`（根目录），导入一个真实 pptx，确认导入成功且 toast 文案正确。

- [ ] **Step 3: Commit**

```bash
git add src/components/dialogs/ImportDialog.tsx
git commit -m "feat: 导入完成展示兼容性报告"
```

---

### Task 17: 全量验证收尾

**Files:** 无新文件

- [ ] **Step 1: 全量单测**

Run: `npx vitest run`
Expected: 全部 PASS（含既有 export 测试与 store/组件测试）

- [ ] **Step 2: typecheck + lint**

Run: `npx tsc -p tsconfig.json --noEmit && npx oxlint src/core/import src/components/dialogs src/components/canvas`
Expected: 无 error（warning 逐条判断是否处理）

- [ ] **Step 3: demo 构建冒烟**

Run（根目录）: `cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx && pnpm build`
Expected: 构建成功

- [ ] **Step 4: 真实文件手工回归**

用 Office 模板 / WPS / 教育课件各 1 个真实 pptx（用户提供），在 demo 中导入并检查：
- 页数一致、主题色正确（不再是默认橙红）
- 母版装饰/背景出现
- 组合形状内的元素可见
- 图表出现（或 toast 中报告 unknownChart）
- 表格合并与样式正确

发现的问题按「单元素失败仅跳过」原则逐条修复后再次提交。

- [ ] **Step 5: 收尾提交与记忆更新**

```bash
git add -A
git commit -m "chore: pptx 导入第一期收尾（真实文件回归修复）"
```

并更新项目记忆 `eflink-pptx-architecture.md`：第一期完成状态、遗留项（富文本表格单元格、SmartArt 完整解析、lumMod 高级修饰、动画不支持）。

---

## 自查记录（writing-plans Self-Review）

1. **Spec 覆盖**：主题色（Task 5/6）、母版/版式（Task 13/14）、组合（Task 12）、渐变/阴影/custGeom（Task 7/9）、项目符号/超链接/autofit/竖排（Task 8）、图片裁剪（Task 10）、表格首 run 样式（Task 11）、图表导入（Task 12）、兼容性报告 UI（Task 16）、错误分层容错（Task 12 index 兜底/Task 14 slide 兜底）、测试策略（各任务 TDD + Task 17 回归）。第二期（SVG 预览）不在本计划范围。
2. **占位符扫描**：无 TBD/TODO；Task 12 group.ts 的两处修正说明与 Task 13 的 scale 参数说明为实现时明确指令，非占位符。
3. **类型一致性**：`ParseContext`（context.ts）在 Task 8-14 使用一致（pkg/partPath/theme/report/scale/placeholders）；`parseSpTreeNode(node, ctx, xf, pkg)` 四参签名在 group/index/master 调用处一致；`ImportReport` 自 types/slides.ts 导入；`resolveColor(container, theme)` 二参一致；`txBodyToHTML(txBody, theme, pkg?, partPath?)` 调用处一致。
