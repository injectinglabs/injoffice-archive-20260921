import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'

const source = (name: string, file = 'index.ts') => fileURLToPath(new URL(`../../packages/${name}/src/${file}`, import.meta.url))
const pptxRenderEntry = source('pptx-render')

function optionalRendererModules(): Plugin {
  return {
    name: 'injoffice-desktop-optional-modules',
    enforce: 'pre',
    resolveId(id, importer) {
      if (id === './SpreadsheetEditor' || id === './SpreadsheetEditor.tsx' || id.endsWith('/SpreadsheetEditor')) {
        if (!existsSync(fileURLToPath(new URL('./src/SpreadsheetEditor.tsx', import.meta.url)))) return '\0injoffice-spreadsheet-editor'
      }
      if (id === '@injoffice/pptx-render') {
        if (!existsSync(pptxRenderEntry) || !readFileSync(pptxRenderEntry, 'utf8').includes('exportNativePptxSlideSvg')) return '\0injoffice-pptx-render'
      }
      if (!importer) return
      const spec = id.split('?')[0]
      if (!spec.endsWith('.wasm') && !spec.endsWith('wasm_exec.js') && !spec.endsWith('.worker.js')) return
      const from = importer.startsWith('\0') ? '' : importer.split('?')[0]
      const candidate = spec.startsWith('/') || spec.startsWith('\0') ? spec : resolve(dirname(from), spec)
      if (!existsSync(candidate)) return `\0injoffice-missing-asset:${spec}`
    },
    load(id) {
      if (id === '\0injoffice-spreadsheet-editor') return 'export function SpreadsheetEditor() { return null }\n'
      if (id === '\0injoffice-pptx-render') {
        return 'export async function exportNativePptxSlideSvg() { throw new Error("SVG export is unavailable in this app build."); }\n'
      }
      if (id.startsWith('\0injoffice-missing-asset:')) return 'export default ""\n'
    },
  }
}

export default defineConfig({
  plugins: [react(), optionalRendererModules()],
  base: './',
  // electron/main.cjs on this branch serves injoffice://app from ../renderer.
  build: { outDir: 'renderer', assetsInlineLimit: 0 },
  resolve: { alias: [
    { find: '@injoffice/font-metrics/layout', replacement: source('font-metrics', 'layout.ts') },
    { find: '@injoffice/sheets/browser', replacement: source('sheets', 'browser.ts') },
    { find: '@injoffice/docs/native-docx', replacement: source('docs', 'nativeDocx.ts') },
    { find: '@injoffice/native-runtime', replacement: source('native-runtime') },
    { find: '@injoffice/pptx-native', replacement: source('pptx-native') },
    { find: '@injoffice/docx-wasm', replacement: source('docx-wasm') },
    { find: '@injoffice/pptx-wasm', replacement: source('pptx-wasm') },
    { find: '@injoffice/xlsx-wasm', replacement: source('xlsx-wasm') },
  ] },
})
