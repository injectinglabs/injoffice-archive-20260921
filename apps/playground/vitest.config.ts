import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: {
      '@injoffice/agent-office/xlsx': fileURLToPath(new URL('../../packages/agent-office/src/xlsx.ts', import.meta.url)),
      '@injoffice/sheets/browser': fileURLToPath(new URL('../../packages/sheets/src/browser.ts', import.meta.url)),
      '@injoffice/agent-tools': fileURLToPath(new URL('../../packages/agent-tools/src/index.ts', import.meta.url)),
    },
  },
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts'],
  },
})
