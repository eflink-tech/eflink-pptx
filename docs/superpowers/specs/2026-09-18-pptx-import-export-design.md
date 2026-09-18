# PPTX 导入导出兼容度提升 — 设计文档

日期：2026-09-18
状态：已与用户逐节确认
范围：`eflink-pptx/packages/pptx`（`@eflink-tech/pptx`）

## 1. 背景与目标

现有实现（`src/core/import/pptx.ts` 411 行手写 OOXML 解析、`src/core/export/pptx.ts` 基于 pptxgenjs）为「尽力还原」级别，面对真实 PPT（Office 模板、WPS、教育课件）兼容度低。

**已确认的需求决策**：

| 决策点 | 结论 |
|---|---|
| 总目标 | 重写导入技术路线（导出维持 pptxgenjs） |
| 架构约束 | 必须纯前端，不依赖后端 |
| 可编辑性 | 双模式：可编辑导入 + 高保真只读预览 |
| 文件来源 | Office 模板类、WPS 文件、教育课件类 |
| 优先级 | 先可编辑导入（第一期），后高保真预览（第二期） |
| 技术路线 | 方案 B：自研分层重写，保持 Apache-2.0 |

**许可证约束（已核实）**：PPTist 为 AGPL-3.0，本项目为 Apache-2.0 且发布 npm，**不得移植/复制 PPTist 代码**；仅可参考其公开的特性范围与优先级信息。

**兼容度预期**：Office 模板类 80-90% 视觉还原；WPS 特有文本效果（艺术字、发光）降级或丢弃；动画/切换效果不支持。

## 2. 总体架构（第一期）

替换单文件解析器为分层模块（对外 API `importPPTX(file): Promise<Presentation>` 不变，`ImportDialog.tsx` 调用方式不变）：

```
src/core/import/
├── index.ts          # 入口 importPPTX(file) → Presentation
├── package.ts        # 解压 + 关系图：JSZip → Part[]（完整 parts → rels 图）
├── theme.ts          # ppt/theme/themeN.xml → 主题色表/字体表/填充样式表
├── master.ts         # slideMaster + slideLayout 链 → 背景、占位符、装饰元素
├── styles.ts         # 样式继承解析器：schemeClr/lstStyle/placeholderStyle 求值
├── geometry.ts       # prstGeom 完整预设表 + custGeom 路径转换
├── elements/         # 各元素类型解析器
│   ├── text.ts  shape.ts  picture.ts  table.ts  chart.ts  group.ts
└── pptx.test.ts
```

解析顺序：

1. 解压 zip，构建完整部件关系图（支持任意层级 rels 跳转）
2. 解析主题 → 主题字典（`accent1-6`、`dk1/lt1`、`tx1/tx2`、字体、预设填充）
3. 解析 slideMaster → slideLayout → 每页样式继承链（页面 → 版式 → 母版 → 主题）
4. 逐页递归解析 `spTree`：组合 `grpSp` 递归展开，坐标按组合变换矩阵折算；元素解析时按继承链求值颜色/字体/填充

画布约定不变：目标 1280×(1280/viewportRatio)，EMU→px 1px=9525EMU。

导出侧维持 pptxgenjs 不动。注意：pptxgenjs 4 不支持形状渐变填充，渐变导出维持现有「离屏截图兜底」策略，不做原生 gradFill 改造。

## 3. 元素级兼容度与降级策略（第一期）

| 类型 | 第一期目标 | 降级策略 |
|---|---|---|
| 文本 | 字体/字号/颜色/加粗斜体下划线/对齐/行距/项目符号与编号/文本自动缩放（normAutofit）近似/竖排 | 主题字体按映射表落到系统字体 |
| 形状 | 全部预设 prstGeom（180+，现为 50）；渐变填充（linear 取首尾色）；阴影（outerShdw → shadow 字段）；custGeom → SVG path 存 `path` 字段 | 无对应形状 → 离屏渲染降级为图片 |
| 组合 grpSp | 递归展开打平，子元素坐标按 grpSp 变换矩阵折算 | 解析失败丢弃该组合并计入报告 |
| 图片 | 保留现状 + 裁剪（srcRect）/圆角/边框 | — |
| 表格 | 单元格首 run 样式提取（加粗/斜体/颜色/字号 → style 字段；渲染器 cell.text 为纯文本，富文本渲染留二期） /合并/边框色宽 | — |
| 图表 | 新增图表导入：bar/line/pie/doughnut/area/scatter → 内部 ChartElement | 不识别类型 → chart part 缓存图片或占位框 |
| 母版/版式 | 背景（纯色/渐变/图片）、母版装饰元素合入每页、占位符位置作为文本框默认样式 | — |
| 超链接 | 文本 run 级 r:link → TipTap link mark | — |
| 音视频 | 保留现状（海报帧 + src） | — |
| 明确丢弃 | 动画/切换/SmartArt→缓存图片兜底/艺术字特效 | 导入完成弹「兼容性报告」列出丢弃项计数 |

SmartArt：diagram part 不做完整解析，优先用其缓存预览图（`preview`），无图时画灰色占位框。

### 数据模型增量（`types/slides.ts`，全部可选字段，向后兼容）

- `ShapeElement.path?: string`（custGeom SVG 路径）
- `ShapeElement.shadow?: { type: 'outer'; blur: number; offsetX: number; offsetY: number; color: string }`
- `ShapeElement.fill` 渐变对象（模型已支持，导入直接产出）
- 新增 `ImportReport { skipped: Record<string, number> }` 供 UI 展示兼容性报告

## 4. 高保真预览（第二期）

新增 `src/core/preview/` OOXML→SVG 渲染器：

- 每页 slide → 一个 `<svg>`（viewBox 按源画布尺寸），文本/形状/图片/表格渲染为 SVG 原生节点
- 复用一期解析层（theme/master/styles/geometry）；与可编辑导入的差别是不做向编辑器模型的映射取舍，阴影、渐变中间站、精确字距、艺术字轮廓按 OOXML 语义直接渲染
- 入口 `previewPPTX(file): Promise<SVGSVGElement[]>`，用于 ImportDialog「先预览再导入」及独立只读预览
- 字体缺失时按主题字体映射表 fallback，预览与编辑保持同一映射

## 5. 错误处理

- 分层容错：单个 slide/元素解析失败仅跳过并计入 `ImportReport`，不中断整个文件
- 文件级错误：非 zip / 缺 presentation.xml → 明确中文错误信息
- 防御性解析：所有 XML 访问走 `attr/firstDescendant` 安全助手；WPS 非标准属性不抛异常

## 6. 测试策略

- 单元测试（vitest）：theme 色表求值、rels 路径解析、组合矩阵折算、prstGeom 映射表、HTML↔runs 转换；fixture 用 JSZip 在测试内动态构造最小 pptx，不提交二进制文件
- 真实文件回归：`e2e/fixtures/` 放 Office 模板/WPS/教育课件各 1-2 个真实 pptx，e2e 断言页数、元素类型计数、关键颜色值、无崩溃
- 往返测试：内部模型 → exportPPTX → importPPTX，断言核心属性不丢失
- 现有 `pptx.test.ts` 全部保留并通过

## 7. 交付划分

1. 第一期：分层解析器重写 + 兼容度提升 + 兼容性报告 UI
2. 第二期：SVG 高保真预览渲染器 + 预览入口
