import { defineConfig } from '@playwright/test'

// 注意用 localhost 而非 127.0.0.1：vite 8 默认可能只绑 IPv6 ::1
// 端口默认 5182：避开 word(3000)/zitie·frontend(5173)/process-on(5174)/mindmap(5181)/draw·excel(5176/5179)/draw demo(5180)
const port = Number(process.env.E2E_PORT ?? 5182)

export default defineConfig({
  testDir: './e2e',
  workers: 1,
  reporter: [['list'], ['html', { open: 'never' }]],
  use: {
    baseURL: `http://localhost:${port}`,
    launchOptions: { channel: 'chrome' },
  },
  webServer: {
    command: `pnpm --filter eflink-pptx-demo build && pnpm --filter eflink-pptx-demo exec vite preview --port ${port} --strictPort`,
    url: `http://localhost:${port}`,
    reuseExistingServer: !process.env.CI,
    timeout: 120000,
  },
})
