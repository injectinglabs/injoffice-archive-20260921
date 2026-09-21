import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
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
const releaseTag = `v${JSON.parse(readFileSync(resolve(root, 'packages/collab/package.json'), 'utf8')).version}`

const run = (args, env = process.env) => spawnSync(process.execPath, [script, ...args], {
  cwd: root,
  encoding: 'utf8',
  env,
})

test('requires one explicit non-accidental release mode', () => {
  const missing = run([releaseTag])
  assert.equal(missing.status, 1)
  assert.match(missing.stderr, /--dry-run.*--pack-destination.*--stage-from.*--bootstrap-from/)

  const conflicting = run([releaseTag, '--dry-run', '--pack-destination', 'release-tarballs'])
  assert.equal(conflicting.status, 1)
  assert.match(conflicting.stderr, /Usage:/)
})

test('refuses first-package bootstrap from CI', () => {
  const result = run([releaseTag, '--bootstrap-from', 'release-tarballs'], { ...process.env, CI: '1' })
  assert.equal(result.status, 1)
  assert.match(result.stderr, /first publish must be interactive and protected by 2FA/)
})

test('direct publication requires GitHub Actions OIDC', () => {
  for (const overrides of [
    { GITHUB_ACTIONS: '' },
    { GITHUB_ACTIONS: 'true', ACTIONS_ID_TOKEN_REQUEST_URL: '' },
    { GITHUB_ACTIONS: 'true', ACTIONS_ID_TOKEN_REQUEST_TOKEN: '' },
  ]) {
    const result = run([releaseTag, '--publish-from', 'release-tarballs'], { ...process.env, ...overrides })
    assert.equal(result.status, 1)
    assert.match(result.stderr, /Trusted publishing requires GitHub Actions with OIDC/)
  }
})

// Exercise the actual CLI with isolated manifests, a fake registry and an npm
// recorder. No network or real publishing credentials are used.
const withRelease = (fn) => {
  const temporary = mkdtempSync(join(tmpdir(), 'injoffice-publish-test-'))
  try {
    mkdirSync(join(temporary, 'packages', 'example'), { recursive: true })
    mkdirSync(join(temporary, 'bin'))
    writeFileSync(join(temporary, 'LICENSE'), 'test license')
    writeFileSync(join(temporary, 'packages', 'example', 'package.json'), JSON.stringify({ name: '@example/package', version: '1.2.3', license: 'MIT' }))
    const tarball = join(temporary, 'example.tgz')
    writeFileSync(tarball, 'synthetic test payload')
    const hash = createHash('sha512').update(readFileSync(tarball)).digest('hex')
    writeFileSync(join(temporary, 'release-manifest.json'), JSON.stringify({
      schemaVersion: 1, tag: 'v1.2.3', npmTag: 'latest',
      packages: [{ name: '@example/package', version: '1.2.3', filename: 'example.tgz', sha512: hash }],
    }))
    const preload = join(temporary, 'registry.mjs')
    writeFileSync(preload, `globalThis.fetch = async (url) => process.env.TEST_INTEGRITY && (!process.env.TEST_PARTIAL || url.includes('%2Fpackage/'))
      ? new Response(JSON.stringify({dist: {integrity: process.env.TEST_INTEGRITY}}))
      : new Response('', {status: 404})`)
    const calls = join(temporary, 'npm-calls.json')
    writeFileSync(join(temporary, 'bin', 'npm'), `#!${process.execPath}\nrequire('node:fs').writeFileSync(${JSON.stringify(calls)}, JSON.stringify(process.argv.slice(2)))\n`, { mode: 0o755 })
    const invoke = (overrides = {}) => spawnSync(process.execPath, ['--import', preload, script, 'v1.2.3', '--source-root', temporary, '--publish-from', temporary], {
      encoding: 'utf8',
      env: { ...process.env, PATH: `${join(temporary, 'bin')}:${process.env.PATH}`, GITHUB_ACTIONS: 'true', ACTIONS_ID_TOKEN_REQUEST_URL: 'https://example.invalid', ACTIONS_ID_TOKEN_REQUEST_TOKEN: 'test-only', ...overrides },
    })
    fn({ invoke, calls, tarball, temporary, hash, integrity: `sha512-${Buffer.from(hash, 'hex').toString('base64')}` })
  } finally {
    rmSync(temporary, { recursive: true, force: true })
  }
}

test('trusted publication sends the verified tarball to npm without lifecycle scripts', () => withRelease(({ invoke, calls, tarball }) => {
  const result = invoke()
  assert.equal(result.status, 0, result.stderr)
  assert.deepEqual(JSON.parse(readFileSync(calls, 'utf8')), ['publish', tarball, '--access', 'public', '--tag', 'latest', '--ignore-scripts', '--fetch-retries=0'])
}))

test('trusted publication skips matching immutable versions and rejects mismatches', () => withRelease(({ invoke, calls, integrity }) => {
  assert.equal(invoke({ TEST_INTEGRITY: integrity }).status, 0)
  assert.equal(existsSync(calls), false)
  const mismatch = invoke({ TEST_INTEGRITY: 'sha512-wrong' })
  assert.equal(mismatch.status, 1)
  assert.match(mismatch.stderr, /published registry integrity does not match/)
  assert.equal(existsSync(calls), false)
}))

test('trusted publication rejects modified tarballs before publishing', () => withRelease(({ invoke, calls, tarball }) => {
  writeFileSync(tarball, 'modified payload')
  const result = invoke()
  assert.equal(result.status, 1)
  assert.match(result.stderr, /failed its SHA-512 check/)
  assert.equal(existsSync(calls), false)
}))

test('trusted publication resumes a partial release after verifying every package', () => withRelease(({ invoke, calls, temporary, hash, integrity }) => {
  mkdirSync(join(temporary, 'packages', 'dependent'))
  writeFileSync(join(temporary, 'packages', 'dependent', 'package.json'), JSON.stringify({ name: '@example/dependent', version: '1.2.3', license: 'MIT', dependencies: { '@example/package': '1.2.3' } }))
  writeFileSync(join(temporary, 'dependent.tgz'), 'synthetic test payload')
  const manifestPath = join(temporary, 'release-manifest.json')
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
  manifest.packages.push({ name: '@example/dependent', version: '1.2.3', filename: 'dependent.tgz', sha512: hash })
  writeFileSync(manifestPath, JSON.stringify(manifest))
  const result = invoke({ TEST_INTEGRITY: integrity, TEST_PARTIAL: '1' })
  assert.equal(result.status, 0, result.stderr)
  assert.match(result.stdout, /Already published @example\/package/)
  assert.equal(JSON.parse(readFileSync(calls, 'utf8'))[1], join(temporary, 'dependent.tgz'))
}))

test('rejects a release artifact whose identity does not match the requested tag', () => {
  const temporary = mkdtempSync(join(tmpdir(), 'injoffice-release-test-'))
  try {
    writeFileSync(resolve(temporary, 'release-manifest.json'), JSON.stringify({
      schemaVersion: 1,
      tag: 'v0.0.0-mismatched-artifact',
      npmTag: 'latest',
      packages: [],
    }))
    const result = run([releaseTag, '--stage-from', temporary])
    assert.equal(result.status, 1)
    assert.match(result.stderr, /Release manifest does not match the requested release/)
  } finally {
    rmSync(temporary, { recursive: true, force: true })
  }
})
