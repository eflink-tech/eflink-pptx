import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// demo 直连组件库源码（无需先构建库）
const pptxSrc = fileURLToPath(new URL('../../packages/pptx/src/', import.meta.url))

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: /^@eflink-tech\/pptx\/styles\.css$/, replacement: `${pptxSrc}styles.css` },
      { find: /^@eflink-tech\/pptx$/, replacement: `${pptxSrc}index.ts` },
    ],
  },
})
