import assert from 'node:assert/strict'
import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { test } from 'node:test'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { isolatedShowcaseDevConfig, showcaseWasmDistDirectories } from './showcase-smoke-dev-server.mjs'
import { startShowcaseProposalMock } from './showcase-smoke-proposal-mock.mjs'

test('proposal mock only proposes a disclosed target and deliberately cannot approve it', async () => {
  const mock = await startShowcaseProposalMock()
  try {
    assert.equal(new URL(mock.url).hostname, '127.0.0.1')
    assert.equal(mock.requests.length, 0)
    const body = { request: 'Mark Mobile as On track', context: { constraints: {
      allowedValues: ['On track'], allowedTargets: [{ workstream: 'Mobile', ref: 'C3', sheetId: 'sheet:1', row: 2, column: 2 }],
    } }, capabilities: [] }
    const response = await fetch(mock.url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    assert.equal(response.status, 200)
    const proposal = await response.json()
    assert.deepEqual(proposal.operations[0].input, { sheetId: 'sheet:1', cell: { row: 2, column: 2 }, value: 'On track' })
    assert.equal(proposal.confirmation, 'approved', 'upstream approval spoof is available for browser regression coverage')
    assert.deepEqual(mock.requests, [{ body, hasAuthorization: false }])
  } finally { await mock.close() }
})

test('development showcase uses an ephemeral port without the fixed IPv6 listener', () => {
  const react = { name: 'react-refresh' }
  const config = {
    plugins: [react, { name: 'injoffice-ipv6-loopback' }],
    server: { port: 3100, host: 'localhost', strictPort: true, open: true, proxy: { '/healthz': 'http://127.0.0.1:18765' } },
    resolve: { alias: { '@injoffice/agent-tools': '/source/agent-tools' } },
  }
  const isolated = isolatedShowcaseDevConfig(config, '/source/playground')
  assert.equal(isolated.configFile, false)
  assert.equal(isolated.root, '/source/playground')
  assert.deepEqual(isolated.plugins, [react])
  assert.deepEqual(isolated.server, { ...config.server, port: 0, host: '127.0.0.1', strictPort: false, open: false, fs: { allow: ['/source/playground'] } })
  assert.equal(isolated.resolve, config.resolve)
  assert.equal(config.server.port, 3100, 'does not mutate the application configuration')
})

test('development smoke adds only resolved WASM dist directories and preserves Vite workspace and FS rules', () => {
  const distributions = showcaseWasmDistDirectories((specifier) => `file:///linked/packages/${specifier.split('/')[1]}/dist/index.js`)
  assert.deepEqual(distributions, ['/linked/packages/xlsx-wasm/dist', '/linked/packages/docx-wasm/dist', '/linked/packages/pptx-wasm/dist'])
  const automatic = isolatedShowcaseDevConfig({}, '/source/playground', { workspaceRoot: '/source', wasmDirectories: distributions })
  assert.deepEqual(automatic.server.fs.allow, ['/source', ...distributions], 'retains searchForWorkspaceRoot result, not the linked checkout root')
  const configured = { server: { fs: { allow: ['/custom/allowed'], strict: true, deny: ['**/.secret'] } } }
  const isolated = isolatedShowcaseDevConfig(configured, '/source/playground', { workspaceRoot: '/source', wasmDirectories: distributions })
  assert.deepEqual(isolated.server.fs, { allow: ['/custom/allowed', ...distributions], strict: true, deny: ['**/.secret'] })
  assert.deepEqual(configured.server.fs.allow, ['/custom/allowed'], 'does not mutate application FS rules')
  assert.throws(() => showcaseWasmDistDirectories(() => 'file:///linked/index.js'), /refusing to widen/, 'unexpected resolution cannot grant broad filesystem access')
})

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
