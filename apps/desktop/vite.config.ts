import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { nativeRendererModules } from './scripts/renderer-modules.cjs'

const source = (name: string, file = 'index.ts') => fileURLToPath(new URL(`../../packages/${name}/src/${file}`, import.meta.url))

export default defineConfig({
  plugins: [react(), nativeRendererModules()],
  base: './',
  // electron/main.cjs on this branch serves injoffice://app from ../renderer.
  build: { outDir: 'renderer', assetsInlineLimit: 0 },
  // @injoffice/{docx,xlsx,pptx}-wasm are deliberately not aliased to src/: their .wasm,
  // wasm_exec.js and worker assets exist only next to dist/index.js after the engine build
  // (see the prebuild script), and the renderer must bundle those or fail loudly.
  resolve: { alias: [
    { find: '@injoffice/font-metrics/layout', replacement: source('font-metrics', 'layout.ts') },
    { find: '@injoffice/sheets/browser', replacement: source('sheets', 'browser.ts') },
    { find: '@injoffice/docs/native-docx', replacement: source('docs', 'nativeDocx.ts') },
    { find: '@injoffice/native-runtime', replacement: source('native-runtime') },
    { find: '@injoffice/pptx-native', replacement: source('pptx-native') },
  ] },
})
