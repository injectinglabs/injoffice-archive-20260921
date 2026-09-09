import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import test from 'node:test'
import { publishedIntegrity } from './release-registry.mjs'

test('registry preflight requests version JSON and preserves integrity', async () => {
  const integrity = await publishedIntegrity('@injoffice/collab', '0.1.0', async (url, options) => {
    assert.equal(url, 'https://registry.npmjs.org/%40injoffice%2Fcollab/0.1.0')
    assert.equal(options.headers.accept, 'application/json')
    assert.ok(options.signal instanceof AbortSignal)
    return new Response(JSON.stringify({ dist: { integrity: 'sha512-test' } }))
  })
  assert.equal(integrity, 'sha512-test')
})

test('only a registry 404 means a version is unpublished', async () => {
  assert.equal(await publishedIntegrity('example', '0.1.0', async () => new Response('', { status: 404 })), null)
  for (const status of [401, 403, 406, 429, 500]) {
    await assert.rejects(publishedIntegrity('example', '0.1.0', async () => new Response('', { status })), new RegExp(`HTTP ${status}`))
  }
})

test('registry preflight fails closed on invalid metadata or network failures', async () => {
  for (const metadata of [{}, { dist: { integrity: '' } }, { dist: { integrity: 123 } }]) {
    await assert.rejects(publishedIntegrity('example', '0.1.0', async () => new Response(JSON.stringify(metadata))), /missing dist.integrity/)
  }
  await assert.rejects(publishedIntegrity('example', '0.1.0', async () => { throw new Error('network failure') }), /network failure/)
  await assert.rejects(publishedIntegrity('example', '0.1.0', async () => new Response('invalid JSON')), SyntaxError)
})

const root = resolve(import.meta.dirname, '..')
const script = resolve(root, 'scripts/release-packages.mjs')

const run = (args, env = process.env) => spawnSync(process.execPath, [script, ...args], {
  cwd: root,
  encoding: 'utf8',
  env,
})

test('requires one explicit non-accidental release mode', () => {
  const missing = run(['v0.1.0'])
  assert.equal(missing.status, 1)
  assert.match(missing.stderr, /--dry-run.*--pack-destination.*--stage-from.*--bootstrap-from/)

  const conflicting = run(['v0.1.0', '--dry-run', '--pack-destination', 'release-tarballs'])
  assert.equal(conflicting.status, 1)
  assert.match(conflicting.stderr, /Usage:/)
})

test('refuses first-package bootstrap from CI', () => {
  const result = run(['v0.1.0', '--bootstrap-from', 'release-tarballs'], { ...process.env, CI: '1' })
  assert.equal(result.status, 1)
  assert.match(result.stderr, /first publish must be interactive and protected by 2FA/)
})

test('rejects a release artifact whose identity does not match the requested tag', () => {
  const temporary = mkdtempSync(join(tmpdir(), 'injoffice-release-test-'))
  try {
    writeFileSync(resolve(temporary, 'release-manifest.json'), JSON.stringify({
      schemaVersion: 1,
      tag: 'v0.1.1',
      npmTag: 'latest',
      packages: [],
    }))
    const result = run(['v0.1.0', '--stage-from', temporary])
    assert.equal(result.status, 1)
    assert.match(result.stderr, /Release manifest does not match the requested release/)
  } finally {
    rmSync(temporary, { recursive: true, force: true })
  }
})
