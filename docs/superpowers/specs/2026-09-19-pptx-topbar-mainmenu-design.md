# PPTX 顶栏菜单整理设计（对齐 Word/Excel 主菜单风格）

日期：2026-09-19
状态：已确认（用户已批准设计）

## 背景与目标

eflink-pptx 当前顶栏为约 18 个纯图标按钮平铺（`TopBar.tsx`），与 eflink-word（☰ 主菜单级联：文件/插入/页面/视图/快捷键）和 eflink-excel（☰ 主菜单级联：文件/编辑/查看/快捷键）的菜单风格差距较大，功能入口依赖 tooltip 识别。

目标：

1. 引入 ☰ 主菜单（级联二级菜单），把低频/文件类功能收进菜单，顶栏只留高频图标
2. 彻底删除「文档管理」功能（用户确认不需要）

## 顶栏新布局（单行）

```
[←] [logo 易飞演示文稿] [☰ 主菜单▾] | [撤销][重做] | [文本][形状][线条][表格][图表][图片][视频][音频][公式] | … | [AI✨] | [▶放映]
```

- 左侧品牌区不动：返回按钮（宿主注入时）+ logo + 产品名
- 新增 ☰ 主菜单按钮，级联交互复用 word `ToolbarMenu.tsx` 的模式：悬停一级展开二级、分组分隔线、✓ 开关态、快捷键右侧提示、Escape 关闭、点击外部关闭
- 顶栏保留的高频图标：撤销/重做、插入 9 件套（现有 `InsertMenu` 原样保留）、AI 助手、红色放映主按钮

### 主菜单五分类

| 分类 | 菜单项 | 说明 |
|------|--------|------|
| 文件 | 新建文档 / 保存（Ctrl S）＼ 导入（PPTX/JSON）… / 导出… ＼ 分享 / 反馈 | 分享、反馈仅在宿主注入分享实现（`getPptxShareHandler() !== null`）时显示；导入/导出打开现有 `import`/`export` 弹窗 |
| 设计 | 模板库… / 主题配色… | 打开现有 `template`/`theme` 弹窗 |
| 视图 | 网格 ✓ / 缩略图面板 ✓ / 查找替换…（Ctrl F） | 网格、缩略图为开关态（✓）；查找替换打开现有 `findReplace` 弹窗 |
| 放映 | 放映 / 演讲者视图 | 行为同现有两个按钮 |
| 快捷键 | 查看全部快捷键 | 叶子分类（无二级），点击打开现有 `hotkey` 弹窗；对齐 word/excel 的叶子分类做法 |

行为约定：

- 「保存」从工具栏图标移入文件菜单（与 Word/Excel 一致），Ctrl+S 快捷键行为不变
- 「新建文档」为直接新建（沿用现有 `createDoc` + `loadDocument`，不加 Word 式确认框）
- 现有顶栏中移除的入口：新建图标、文档管理图标、保存图标、查找替换/模板库/主题配色/导入/导出图标、网格/缩略图图标、演讲者视图按钮
- 关键 `data-testid` 保留：`topbar`、`ai-toggle`、`play`、`insert-*`

## 删除「文档管理」（共 4 处）

1. `src/components/layout/TopBar.tsx` — 入口按钮
2. `src/components/dialogs/ModalHost.tsx` — `'fileManager'` 分支与 `FileManagerDialog` import
3. `src/store/uiStore.ts` — modal 类型中的 `'fileManager'`
4. `src/components/dialogs/FileManagerDialog.tsx` — 组件文件本体（删除前确认无其他引用）

## 组件设计

- 新增 `src/components/menus/MainMenu.tsx`：主菜单组件（参照 word `ToolbarMenu.tsx` 裁剪，使用 lucide 图标，不依赖 canvas-editor）
- `src/components/layout/TopBar.tsx`：移除已收进菜单的图标，挂载 `MainMenu`；分享弹窗相关逻辑（`openShare`/`ShareDialog`）随「分享」菜单项迁入 `MainMenu`
- 菜单项均为既有功能的重新组织，不新增编辑能力，不改 store 状态逻辑（仅 uiStore modal 类型收窄）

## 错误处理

- 分享前的强制保存沿用现有逻辑：保存失败 toast「保存失败，无法分享」并中止
- 菜单项动作全部复用现有弹窗/store 动作，无新增失败路径

## 测试与验收

- `tsc` 通过；既有 vitest 全量通过（当前无测试引用 TopBar）
- Playwright 手动验收：主菜单五个分类逐项打开对应弹窗 / 触发动作；开关项 ✓ 状态正确；分享入口在注入/未注入两种宿主下显示正确
- 「文档管理」全局无残留引用（grep 验证）
