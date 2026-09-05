# @eflink-tech/pptx

开箱即用的在线演示文稿编辑器 React 组件。既可独立运行，也可作为组件嵌入任意 React 应用。

An out-of-the-box presentation (PPT) editor for the web, usable standalone or embedded into any React app.

## 安装

```bash
npm install @eflink-tech/pptx
```

## 使用

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

组件自带完整编辑器 UI（顶部工具栏、缩略图面板、画布、样式面板、放映、AI 助手），挂载后自动恢复上次编辑的文档。文档保存在浏览器 IndexedDB，密钥等敏感信息不出浏览器。

## 主要能力

- 元素：TipTap 富文本、图片、形状（80+ 预设）、线条、表格、图表（ECharts）、音视频、LaTeX 公式
- 操作：多选框选 / 组合 / 层级 / 对齐分布 / 吸附参考线 / 撤销重做 / 快捷键
- 页面效果：背景填充、切换动画、元素进入 / 强调 / 退出动画
- 放映：全屏放映 + 演讲者视图
- 导入导出：`.pptx`（pptxgenjs 导出 / 自研 OOXML 导入）、PNG / JPG / PDF / JSON
- 模板与主题：内置成套模板、主题色板一键换色
- AI 助手：主题 → 大纲 → 整套版式（OpenAI 兼容接口，用户自备 Key）

详见 [GitHub 仓库](https://github.com/eflink-tech/eflink-pptx)。

## Tailwind 说明

组件库内部布局用到 Tailwind 工具类（已随 `styles.css` 提供回退样式）。若宿主使用 Tailwind v4 且希望得到与 demo 一致的布局，请在入口 CSS 中显式扫描组件包：

```css
@import "tailwindcss";
@source "../node_modules/@eflink-tech/pptx";
```

## License

[Apache-2.0](./LICENSE)
