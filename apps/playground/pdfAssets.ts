import { readFileSync, readdirSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import type { Plugin } from 'vite'

const require = createRequire(import.meta.url)
const root = dirname(require.resolve('pdfjs-dist/package.json'))
const version: string = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version
const directory = `assets/pdfjs-${version}/`
const virtualId = 'virtual:injoffice-pdf-resources'
const resolvedId = '\0' + virtualId

/** Copy only installed PDF.js resources, including their license files. */
export function pdfResourceFiles(): Map<string, Buffer> {
  const files = new Map<string, Buffer>()
  for (const folder of ['cmaps', 'standard_fonts', 'wasm']) {
    for (const entry of readdirSync(join(root, folder), { withFileTypes: true })) {
      if (entry.isFile()) files.set(`${directory}${folder}/${entry.name}`, readFileSync(join(root, folder, entry.name)))
    }
  }
  return files
}

/** Same resource URLs in development and production, respecting a deployment base. */
export function pdfAssets(): Plugin {
  let base = '/'
  return {
    name: 'injoffice-pdf-resources',
    configResolved(config) { base = config.base },
    resolveId(id) { if (id === virtualId) return resolvedId },
    load(id) {
      if (id !== resolvedId) return
      const prefix = `${base}${directory}`
      return `export function getPdfLoadOptions() { return {
        cMapUrl: new URL(${JSON.stringify(prefix + 'cmaps/')}, document.baseURI).href,
        standardFontDataUrl: new URL(${JSON.stringify(prefix + 'standard_fonts/')}, document.baseURI).href,
        wasmUrl: new URL(${JSON.stringify(prefix + 'wasm/')}, document.baseURI).href,
        useSystemFonts: false
      }; }`
    },
    configureServer(server) {
      const files = pdfResourceFiles()
      const prefix = new URL(base, 'http://localhost').pathname
      server.middlewares.use((req, res, next) => {
        if (req.method !== 'GET' && req.method !== 'HEAD') return next()
        const path = (req.url ?? '').split('?')[0]
        if (!path.startsWith(prefix)) return next()
        const key = path.slice(prefix.length)
        const source = files.get(key)
        if (!source) return next()
        res.setHeader('Content-Type', key.endsWith('.wasm') ? 'application/wasm' : key.endsWith('.js') ? 'text/javascript' : 'application/octet-stream')
        res.setHeader('Content-Length', source.byteLength)
        res.end(req.method === 'HEAD' ? undefined : source)
      })
    },
    generateBundle() {
      for (const [fileName, source] of pdfResourceFiles()) this.emitFile({ type: 'asset', fileName, source })
    },
  }
}
