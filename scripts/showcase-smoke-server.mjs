import { readFile, stat } from 'node:fs/promises'
import { createServer } from 'node:http'
import { extname, resolve, sep } from 'node:path'

const mime = new Map([
  ['.css', 'text/css'], ['.html', 'text/html'], ['.js', 'text/javascript'], ['.mjs', 'text/javascript'],
  ['.json', 'application/json'], ['.svg', 'image/svg+xml'], ['.wasm', 'application/wasm'],
])

/** Serve the already-built playground, exercising the same subpath as the Office smoke. */
export async function startShowcaseServer(dist, base = '/injoffice-smoke/') {
  const root = resolve(dist)
  await stat(resolve(root, 'index.html'))
  const server = createServer(async (request, response) => {
    try {
      const pathname = new URL(request.url ?? '/', 'http://127.0.0.1').pathname
      if (!pathname.startsWith(base)) return response.writeHead(404).end('not found')
      const relative = pathname === base ? 'index.html' : decodeURIComponent(pathname.slice(base.length))
      const file = resolve(root, relative)
      if (!file.startsWith(`${root}${sep}`)) return response.writeHead(400).end('invalid path')
      if (!(await stat(file)).isFile()) return response.writeHead(404).end('not found')
      response.writeHead(200, { 'Content-Type': mime.get(extname(file)) ?? 'application/octet-stream' })
      response.end(await readFile(file))
    } catch {
      response.writeHead(404).end('not found')
    }
  })
  await new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', resolve)
  })
  return {
    url: `http://127.0.0.1:${server.address().port}${base}`,
    close: () => new Promise((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve())
      server.closeAllConnections()
    }),
  }
}
