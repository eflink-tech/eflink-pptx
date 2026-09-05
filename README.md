# EFLink PPTX 易飞演示文稿

开箱即用的在线演示文稿（PPT）编辑器。**既可以独立运行（本仓库 demo 应用），也可以作为 React 组件嵌入任意应用**（npm 包 `@eflink-tech/pptx`）。

A full-featured presentation editor for the web. Run it standalone, or embed `<PptxEditor />` into your React app.

## 功能特性

- 幻灯片管理：新建 / 复制 / 删除 / 拖拽排序 / 右键菜单（左侧缩略图面板）
- 元素：文本（TipTap 富文本：字体色 / 粗斜下划线删除线 / 上下标 / 行高 / 字间距 / 竖排 / 对齐 / 列表 / 引用 / 代码 / 链接 / 文字阴影 / 自动收缩）、图片（裁剪 / 圆角 / 翻转 / 阴影 / 边框）、形状（80+ 预设几何图形，渐变填充 / 内嵌文本 / 阴影）、线条（直线 / 折线 / 曲线，控制点拖拽 / 箭头 / 虚线）、表格（合并单元格 / 表头样式 / 单元格级编辑）、图表（柱 / 条 / 折线 / 面积 / 饼 / 环形 / 散点，CSV 数据编辑 / 系列配色）、音视频、公式（LaTeX，MathLive 渲染）
- 操作：拖拽 / 8 向缩放 / 旋转 / 多选框选 / 组合 / 层级 / 对齐与等距分布 / 智能吸附参考线 / 锁定 / 不透明度 / 超链接 / 批注 / 撤销重做 / 复制粘贴跨页偏移 / 全套快捷键
- 页面效果：纯色 / 渐变 / 图片背景，11 种幻灯片切换动画，元素进入 / 强调 / 退出动画
- 放映：全屏放映（B/W 黑白屏、超链接跳转）、演讲者视图（备注 / 计时器 / 下一页预览）
- 导入导出：导出 `.pptx`（pptxgenjs，可被 PowerPoint / WPS 打开）、PNG / JPG / PDF / JSON 工程；导入 `.pptx`（自研 OOXML 解析）与 JSON
- 模板与主题：5 套内置成套模板，主题色板一键换色
- AI 助手：主题 → 大纲 → 整套页面版式（OpenAI 兼容接口，用户自备 API Key 与 BaseURL），选中文本润色 / 扩写 / 翻译
- 持久化：多文档 IndexedDB（Dexie）+ localStorage 崩溃恢复镜像，自动保存，刷新不丢内容
- 工程化：Vite + TypeScript 严格模式 + Vitest 单测 + Playwright e2e + oxlint

## 使用组件

```bash
npm install @eflink-tech/pptx
```

```tsx
import { PptxEditor } from '@eflink-tech/pptx';
import '@eflink-tech/pptx/styles.css';

function Page() {
  return (
    <div style={{ width: '100vw', height: '100vh' }}>
      <PptxEditor />
    </div>
  );
}
```

组件自带完整编辑器 UI（顶部工具栏、缩略图面板、画布、样式面板、放映、AI 助手），挂载后自动恢复上次编辑的文档。

### AI 助手配置

点击「AI 助手」进入设置，填入任意 OpenAI 兼容服务的信息即可使用（Key 仅保存在浏览器 IndexedDB，不会上传到任何第三方）：

- BaseURL：OpenAI 兼容接口地址（如 `https://api.openai.com/v1`）
- API Key：对应服务的密钥
- 模型：对话与视觉模型名称

不配置 AI 时，编辑器的全部手工编辑功能不受影响。

### 导出面

- `PptxEditor`：编辑器组件
- `useEditorStore`：zustand 状态（页面、选中集、历史等），可用于程序化操作幻灯片；`useUIStore`：面板 / 模态 / 放映 UI 状态
- `loadStartupDoc / saveDoc / listDocs / createDoc / deleteDoc / duplicateDoc / renameDoc / scheduleAutosave`：多文档持久化（localStorage 镜像 + IndexedDB）
- `createTextElement / createShapeElement / createChartElement / createElementByType ...`：元素工厂
- `SLIDE_WIDTH / SLIDE_HEIGHT / PX_PER_INCH / createPresentation / createSlide`：画布常量与数据模型（`Presentation` / `Slide` / `PPTElement` / `Theme`）

> **Tailwind 说明**：组件库内部布局用到 Tailwind 工具类（已随 `styles.css` 提供回退样式）。若宿主使用 Tailwind v4 且希望得到与 demo 一致的布局，请在入口 CSS 中显式扫描组件包：
>
> ```css
> @import "tailwindcss";
> @source "../node_modules/@eflink-tech/pptx";
> ```

## 本地运行 Demo

```bash
pnpm install
pnpm dev          # 并行：组件库 watch 构建 + demo dev server
# 或
pnpm dev:demo     # 仅 demo（直连组件库源码，无需先构建）
```

打开终端提示的地址即为完整独立应用：富文本、图表、公式、模板、放映、导入导出、AI 助手、自动保存，开箱即用。

## 命令速查

```bash
pnpm lint         # oxlint（monorepo 全量）
pnpm typecheck    # TypeScript 严格类型检查
pnpm test         # Vitest 单元测试
pnpm build        # 构建组件库 + demo
pnpm test:e2e     # Playwright 端到端（生产构建 + preview）
```

## 目录结构

```
eflink-pptx/
├── packages/pptx/      # @eflink-tech/pptx 组件库（开源主体）
│   └── src/
│       ├── components/ # UI 组件（画布 / 面板 / 菜单 / 对话框 / 放映 / AI 助手）
│       ├── core/       # 编辑器内核（元素工厂 / 持久化 / 导入导出 / 形状库 / 几何工具）
│       ├── ai/         # AI 助手（服务调用 / 动作执行 / system prompt）
│       ├── store/      # zustand 编辑器状态
│       ├── files/      # 本地多文档 IndexedDB
│       └── types/      # 数据模型（PPTElement 判别联合 / Slide / Presentation / 动画）
├── apps/demo/          # 独立 demo 应用（易飞演示文稿）
├── e2e/                # Playwright 端到端测试
└── scripts/            # CI 发布辅助脚本
```

## 发版流程

使用 [Changesets](https://github.com/changesets/changesets) 管理版本：

```bash
pnpm changeset    # 记录变更（选择包与版本级别）
git push          # 推送后机器人自动开 Version PR
# 合并 Version PR → 自动发布 npm 并打 tag
```

## License

[Apache-2.0](./LICENSE) © 2026 eflink-tech
