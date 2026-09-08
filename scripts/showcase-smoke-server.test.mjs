import assert from 'node:assert/strict'
import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { test } from 'node:test'
import { startShowcaseServer } from './showcase-smoke-server.mjs'

test('built showcase server respects its base and serves WASM MIME without SPA fallback', async () => {
  const dist = await mkdtemp(resolve(tmpdir(), 'showcase-server-test-'))
  let server
  try {
    await writeFile(resolve(dist, 'index.html'), '<main>showcase</main>')
    await writeFile(resolve(dist, 'test.wasm'), new Uint8Array([0, 97, 115, 109]))
    server = await startShowcaseServer(dist)
    assert.equal(await (await fetch(server.url)).text(), '<main>showcase</main>')
    assert.equal((await fetch(new URL('/', server.url))).status, 404)
    assert.equal((await fetch(`${server.url}missing.js`)).status, 404)
    assert.equal((await fetch(`${server.url}%ZZ`)).status, 404)
    assert.equal((await fetch(`${server.url}%2e%2e%2foutside`)).status, 400)
    assert.equal((await fetch(`${server.url}test.wasm`)).headers.get('content-type'), 'application/wasm')
  } finally {
    await server?.close()
    await rm(dist, { recursive: true, force: true })
  }
})
