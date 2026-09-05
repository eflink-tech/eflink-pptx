import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { defineConfig } from 'vite'

// 库构建：ESM + 类型（tsc 单独产出）+ 单文件 styles.css（含 tailwind 工具类与 MathLive 公式样式）
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    lib: {
      entry: 'src/index.ts',
      formats: ['es'],
      fileName: () => 'index.js',
    },
    // watch 模式（pnpm dev）不清空 dist，保留 tsc 产出的 dist/types；生产构建经 --emptyOutDir 显式清空
    emptyOutDir: false,
    cssCodeSplit: false,
    rollupOptions: {
      // 仅外部化直接依赖：传递依赖打包进产物，避免 pnpm 严格隔离下消费方无法解析。
      // mathlive/static.css 例外：随包打进 styles.css（含公式字体），避免产物出现裸 CSS 外部导入
      external: (id) =>
        /^(react|react-dom|react\/jsx-runtime|@tiptap\/|dexie|echarts|html-to-image|immer|jszip|lucide-react|mathlive|openai|pptxgenjs|zustand)/.test(id) &&
        !id.startsWith('mathlive/static.css'),
      output: {
        assetFileNames: (asset) =>
          asset.names?.[0]?.endsWith('.css') ? 'styles.css' : 'assets/[name][extname]',
      },
    },
  },
})
