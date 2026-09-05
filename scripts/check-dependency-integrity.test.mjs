import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import test from 'node:test'
import { integrityHash, packageNameFromLockPath, parseGoMod, shortestTrace } from './dependency-integrity-lib.mjs'

test('resolves scoped and nested package names from npm lock paths', () => {
  assert.equal(packageNameFromLockPath('node_modules/@scope/name'), '@scope/name')
  assert.equal(packageNameFromLockPath('node_modules/parent/node_modules/child'), 'child')
  assert.equal(packageNameFromLockPath('node_modules/alias', { name: 'actual' }), 'actual')
})

test('parses local Go requirements and replacements', () => {
  const parsed = parseGoMod(`
    module example.test/root
    require (
      example.test/local v0.0.0
      example.test/external v1.2.3
    )
    replace example.test/local => ../local
  `)
  assert.deepEqual(parsed.requires, [
    { module: 'example.test/local', version: 'v0.0.0' },
    { module: 'example.test/external', version: 'v1.2.3' },
  ])
  assert.deepEqual(parsed.replacements.get('example.test/local'), { target: '../local', version: null })
})

test('finds a deterministic shortest dependency trace', () => {
  const graph = new Map([
    ['root', new Set(['slow', 'fast'])],
    ['slow', new Set(['middle'])],
    ['middle', new Set(['target'])],
    ['fast', new Set(['target'])],
  ])
  assert.deepEqual(shortestTrace(graph, ['root'], 'target'), ['root', 'fast', 'target'])
})

test('converts npm integrity values into CycloneDX hashes', () => {
  assert.deepEqual(integrityHash('sha512-YQ=='), { alg: 'SHA-512', content: '61' })
  assert.equal(integrityHash('not-an-integrity'), null)
})

test('generates a parseable CycloneDX inventory from the checked-in lockfile', () => {
  const output = execFileSync(process.execPath, ['scripts/check-dependency-integrity.mjs', '--sbom'], { encoding: 'utf8' })
  const sbom = JSON.parse(output)
  assert.equal(sbom.bomFormat, 'CycloneDX')
  assert.equal(sbom.specVersion, '1.5')
  assert.ok(sbom.components.length > 0)
  assert.ok(sbom.metadata.properties.some((property) => property.name === 'injoffice:release-blocker-count'))
})
