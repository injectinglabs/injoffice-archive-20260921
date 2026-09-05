import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import test from 'node:test'

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
