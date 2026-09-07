import { fileURLToPath } from 'node:url'
import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'

const pkgSrc = (name: string) => fileURLToPath(new URL(`../../packages/${name}/src/index.ts`, import.meta.url))
const pkgFile = (name: string, file: string) => fileURLToPath(new URL(`../../packages/${name}/src/${file}`, import.meta.url))

function pdfNodeHostPlugin(
  handlePdfNodeRequest: (req: unknown, res: unknown) => Promise<boolean>,
): Plugin {
  return {
    name: 'injoffice-pdf-node-host',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        try {
          if (await handlePdfNodeRequest(req, res)) return
        } catch (error) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: error instanceof Error ? error.message : String(error) }))
          return
        }
        next()
      })
    },
  }
}

export default defineConfig(async ({ command }) => {
  const plugins: Plugin[] = [react()]
  if (command === 'serve') {
    const { handlePdfNodeRequest } = await import('./pdfNodeHost.ts')
    plugins.push(pdfNodeHostPlugin(handlePdfNodeRequest))
  }
  return {
    plugins,
    // Lazy route modules are outside Vite's default index.html dependency
    // crawl. Discover them at startup so visiting a demo never triggers a
    // one-time dependency optimizer reload in the middle of navigation.
    optimizeDeps: command === 'serve' ? {
      entries: ['index.html', 'src/**/*.{ts,tsx}', '../docs/src/**/*.{ts,tsx}'],
    } : undefined,
    resolve: {
      alias: {
        '@injoffice/agent-tools': pkgSrc('agent-tools'),
        '@injoffice/font-metrics/layout': pkgFile('font-metrics', 'layout.ts'),
        '@injoffice/charts': pkgSrc('charts'),
        '@injoffice/collab': pkgSrc('collab'),
        '@injoffice/connectors': pkgSrc('connectors'),
        '@injoffice/formulas': pkgSrc('formulas'),
        '@injoffice/univer-sheets/browser': pkgFile('univer-sheets', 'browser.ts'),
        '@injoffice/univer-sheets/styles.css': fileURLToPath(new URL('../../packages/univer-sheets/styles.css', import.meta.url)),
        '@injoffice/history': pkgSrc('history'),
        '@injoffice/pdf/browser': pkgFile('pdf', 'browser.ts'),
        '@injoffice/pivots': pkgSrc('pivots'),
        '@injoffice/pptx-authored': pkgSrc('pptx-authored'),
        '@injoffice/pptx-native': pkgSrc('pptx-native'),
        '@injoffice/pptx-render': pkgSrc('pptx-render'),
        '@injoffice/shapes': pkgSrc('shapes'),
        '@injoffice/slides/authoring': pkgFile('slides', 'authoring.ts'),
        '@injoffice/slides': pkgSrc('slides'),
      },
    },
    server: {
      host: '127.0.0.1',
      port: 3100,
      proxy: {
        '/healthz': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/capabilities': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/xlsx': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/docx': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/pptx': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/artifacts': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
        '/v1/collab': process.env.INJOFFICE_SERVER || 'http://127.0.0.1:18765',
      },
    },
  }
})
