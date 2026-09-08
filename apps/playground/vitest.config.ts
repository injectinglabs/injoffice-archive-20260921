import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: {
      '@injoffice/agent-office/xlsx': fileURLToPath(new URL('../../packages/agent-office/src/xlsx.ts', import.meta.url)),
      '@injoffice/agent-office/docx': fileURLToPath(new URL('../../packages/agent-office/src/docx.ts', import.meta.url)),
      '@injoffice/agent-office/pptx': fileURLToPath(new URL('../../packages/agent-office/src/pptx.ts', import.meta.url)),
      '@injoffice/agent-office/pdf': fileURLToPath(new URL('../../packages/agent-office/src/pdf.ts', import.meta.url)),
      '@injoffice/docs/native-docx': fileURLToPath(new URL('../../packages/docs/src/nativeDocx.ts', import.meta.url)),
      '@injoffice/slides/authoring': fileURLToPath(new URL('../../packages/slides/src/authoring.ts', import.meta.url)),
      '@injoffice/pdf/browser': fileURLToPath(new URL('../../packages/pdf/src/browser.ts', import.meta.url)),
      '@injoffice/pptx-native': fileURLToPath(new URL('../../packages/pptx-native/src/index.ts', import.meta.url)),
      '@injoffice/sheets/browser': fileURLToPath(new URL('../../packages/sheets/src/browser.ts', import.meta.url)),
      '@injoffice/agent-tools': fileURLToPath(new URL('../../packages/agent-tools/src/index.ts', import.meta.url)),
    },
  },
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts'],
  },
})
