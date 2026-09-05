import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import { BASELINE, BASELINE_V2, BASELINE_V3, findForbiddenNativeAuthority, loadCompletionManifest, normalizeSourceComments, validateCompletionManifest, validateNativeAuthorityClosure } from './validate-native-office-completion.mjs'

const testRoot = resolve(fileURLToPath(new URL('..', import.meta.url)))

const clone = (value) => structuredClone(value)
const validate = (mutate) => {
  const manifest = clone(loadCompletionManifest())
  mutate(manifest)
  return validateCompletionManifest(manifest)
}
const includes = (errors, text) => assert(errors.some((error) => error.includes(text)), errors.join('\n'))

test('canonical matrix is valid', () => {
  const schema = JSON.parse(readFileSync(resolve(testRoot, 'schemas/native-office-completion-v1.schema.json'), 'utf8'))
  const manifest = loadCompletionManifest()
  assert.equal(schema.properties.protocol.const, 'injoffice.native-office-completion/v1')
  assert.equal(schema.properties.baseline.properties.commit.const, BASELINE)
  assert.deepEqual(Object.keys(schema.properties.qualifications.properties).sort(), ['differential', 'resource'])
  assert(new RegExp(schema.$defs.capability.properties.id.pattern).test('docx.contract'))
  assert.equal(manifest.capabilities.length, 15)
  assert.equal(manifest.capabilities.every((capability) => capability.status === 'partial'), true)
  assert.equal(manifest.capabilities.filter((capability) => capability.gates.productionE2E.status === 'present').length, 15)
  assert.equal(manifest.fixtures.filter((fixture) => fixture.kind === 'office-export').map((fixture) => fixture.id).join(','), 'docx-word-mac-list-numbering,xlsx-excel-authored-happy-tree')
  assert.equal(manifest.fixtures.filter((fixture) => fixture.kind === 'office-export' && fixture.format === 'pptx').length, 0)
  assert.deepEqual(validateCompletionManifest(manifest), [])
})

test('v2 matrix binds the PowerPoint-authored office-export and stays partial', () => {
  const schema = JSON.parse(readFileSync(resolve(testRoot, 'schemas/native-office-completion-v2.schema.json'), 'utf8'))
  const manifest = loadCompletionManifest(resolve(testRoot, 'testdata/native-office-completion/v2/manifest.json'))
  assert.equal(schema.properties.protocol.const, 'injoffice.native-office-completion/v2')
  assert.equal(schema.properties.baseline.properties.commit.const, BASELINE_V2)
  assert.equal(manifest.version, 2)
  assert.equal(manifest.baseline.commit, BASELINE_V2)
  assert.equal(manifest.capabilities.every((capability) => capability.status === 'partial'), true)
  assert.equal(
    manifest.fixtures.filter((fixture) => fixture.kind === 'office-export').map((fixture) => fixture.id).join(','),
    'docx-word-mac-list-numbering,pptx-powerpoint-authored-attendee-survey-qr,xlsx-excel-authored-happy-tree',
  )
  assert.deepEqual(validateCompletionManifest(manifest), [])
})

test('v3 matrix binds office-export into every capability and is complete', () => {
  const schema = JSON.parse(readFileSync(resolve(testRoot, 'schemas/native-office-completion-v3.schema.json'), 'utf8'))
  const manifest = loadCompletionManifest(resolve(testRoot, 'testdata/native-office-completion/v3/manifest.json'))
  assert.equal(schema.properties.protocol.const, 'injoffice.native-office-completion/v3')
  assert.equal(schema.properties.baseline.properties.commit.const, BASELINE_V3)
  assert.equal(manifest.version, 3)
  assert.equal(manifest.baseline.commit, BASELINE_V3)
  assert.equal(manifest.capabilities.length, 15)
  assert.equal(manifest.capabilities.every((capability) => capability.status === 'complete'), true)
  assert.deepEqual(validateCompletionManifest(manifest), [])
})

test('missing and stale evidence paths fail deterministically', () => {
  includes(validate((manifest) => { manifest.capabilities[0].authoritativeModules[0] = 'missing/native.go' }), 'is missing')
  includes(validate((manifest) => { manifest.capabilities[0].gates.structuralPreservation.evidence[0].selector = 'TestThatNoLongerExists' }), 'selector is stale')
  includes(validate((manifest) => { manifest.capabilities[0].gates.structuralPreservation.evidence[0].selector = 'func' }), 'selector is too generic')
})

test('fixture provenance and digests are mandatory', () => {
  includes(validate((manifest) => { manifest.fixtures[0].sha256 = '0'.repeat(64) }), 'sha256 is stale')
  includes(validate((manifest) => { manifest.fixtures[0].provenance.source = '' }), 'provenance.source')
})

test('capabilities cannot omit tests or claim complete before gates exist', () => {
  includes(validate((manifest) => { manifest.capabilities[0].gates.unsupportedObjectLoss.status = 'missing'; manifest.capabilities[0].gates.unsupportedObjectLoss.evidence = []; manifest.capabilities[0].gates.unsupportedObjectLoss.reason = 'removed' }), 'silent unsupported-object loss gate')
  includes(validate((manifest) => {
    manifest.capabilities[0].status = 'complete'
    manifest.capabilities[0].gates.productionE2E = { status: 'missing', owner: 'host-integration', evidence: [], reason: 'removed' }
  }), 'complete before productionE2E exists')
  includes(validate((manifest) => {
    manifest.capabilities[0].status = 'complete'
    const officeIDs = new Set(manifest.fixtures.filter((fixture) => fixture.kind === 'office-export').map((fixture) => fixture.id))
    for (const gate of Object.values(manifest.capabilities[0].gates)) {
      gate.evidence = (gate.evidence ?? []).filter((evidence) => !evidence.fixtures.some((id) => officeIDs.has(id)))
    }
  }), 'real Office-export fixture')
  includes(validate((manifest) => {
    manifest.capabilities[0].gates.productionE2E = structuredClone(manifest.capabilities[0].gates.structuralPreservation)
  }), 'host-owned production E2E kit')
  includes(validate((manifest) => { manifest.capabilities[0].semantics.supported.push('Unbound second promise') }), 'maxItems')
  includes(validate((manifest) => { manifest.qualifications.resource.status = 'missing' }), 'must equal "present"')
  includes(validate((manifest) => { manifest.qualifications.resource.evidence[0].selector = 'TestMissingResourceGate' }), 'selector is stale')
})

test('native production authority refuses DOM and legacy surfaces', () => {
  includes(validateNativeAuthorityClosure(['scripts/testdata/native-office-completion/forbidden-entrypoint.js'], { root: testRoot }), 'forbidden DOM creation authority')
  includes(validateNativeAuthorityClosure(['scripts/testdata/native-office-completion/closure-entrypoint.js'], { root: testRoot }), 'forbidden DOM creation authority')
  includes(validateNativeAuthorityClosure(['scripts/testdata/native-office-completion/commonjs-entrypoint.cjs'], { root: testRoot }), 'forbidden DOM creation authority')
  includes(validate((manifest) => { manifest.policies.nativeProductionEntrypoints = [] }), 'non-empty sorted unique list')
  assert.deepEqual(findForbiddenNativeAuthority('// document.createElement("canvas")'), ['DOM creation'])
  assert(!validateNativeAuthorityClosure(['scripts/testdata/native-office-completion/comment-entrypoint.js'], { root: testRoot }).some((error) => error.includes('forbidden DOM creation authority')))
})

test('dependency closure covers static module forms after lexical comment normalization', () => {
  const cases = [
    'scripts/testdata/native-office-completion/closure-entrypoint.js',
    'scripts/testdata/native-office-completion/commonjs-entrypoint.cjs',
    'scripts/testdata/native-office-completion/import-equals-entrypoint.ts',
    'scripts/testdata/native-office-completion/interstitial-entrypoint.js',
    'scripts/testdata/native-office-completion/string-comment-entrypoint.cjs',
  ]
  for (const entrypoint of cases) {
    includes(validateNativeAuthorityClosure([entrypoint], { root: testRoot }), 'forbidden DOM creation authority')
  }
  includes(validateNativeAuthorityClosure(['scripts/testdata/native-office-completion/interstitial-package-entrypoint.js'], { root: testRoot }), 'forbidden Mammoth authority')
  for (const poisoned of [
    "const line = '//'; const block = '/*'; document.createElement('div'); const end = '*/'",
    "const matcher = /\\/\\//; document.createElement('div')",
    "const matcher = /[/*]/; document.createElement('div')",
  ]) {
    assert.equal(normalizeSourceComments(poisoned), poisoned)
    assert.deepEqual(findForbiddenNativeAuthority(normalizeSourceComments(poisoned)), ['DOM creation'])
  }
  for (const source of ["import 'mammoth'", "import('mammoth')"]) assert.deepEqual(findForbiddenNativeAuthority(source), ['Mammoth'])
  assert.deepEqual(findForbiddenNativeAuthority("import 'react-dom'"), ['HTML renderer authority'])
})

test('dependency closure rejects deterministic nested-parent symlink escapes by realpath', () => {
  const temporary = mkdtempSync(join(tmpdir(), 'injoffice-authority-'))
  const root = join(temporary, 'root')
  const outside = join(temporary, 'outside')
  try {
    mkdirSync(join(root, 'src'), { recursive: true })
    mkdirSync(outside, { recursive: true })
    writeFileSync(join(root, 'src', 'entrypoint.js'), "import './nested/forbidden.js'\n")
    writeFileSync(join(outside, 'forbidden.js'), "document.createElement('div')\n")
    symlinkSync(outside, join(root, 'src', 'nested'), 'dir')
    const first = validateNativeAuthorityClosure(['src/entrypoint.js'], { root })
    const second = validateNativeAuthorityClosure(['src/entrypoint.js'], { root })
    assert.deepEqual(first, second)
    includes(first, 'resolves outside the repository')
  } finally {
    rmSync(temporary, { recursive: true, force: true })
  }
})

test('format parity and exact pending references are enforced offline', () => {
  includes(validate((manifest) => { for (const capability of manifest.capabilities) capability.dimensions = capability.dimensions.filter((dimension) => !(capability.format === 'xlsx' && dimension === 'security')) }), 'xlsx has no capability for required security dimension')
  includes(validate((manifest) => { manifest.capabilities[0].dimensions.push('typo-dimension'); manifest.capabilities[0].dimensions.sort() }), 'schema enum')
  const pendingBranch = 'test/synthetic-completion-pending'
  const pendingHeadResult = spawnSync('git', ['-C', testRoot, 'rev-parse', 'HEAD'], { encoding: 'utf8' })
  assert.equal(pendingHeadResult.status, 0, pendingHeadResult.stderr)
  const pendingHead = pendingHeadResult.stdout.trim()
  const updateRef = spawnSync('git', ['-C', testRoot, 'update-ref', `refs/heads/${pendingBranch}`, pendingHead], { encoding: 'utf8' })
  assert.equal(updateRef.status, 0, updateRef.stderr)
  try {
    const withPending = (mutate) => validate((manifest) => {
      if (!manifest.pendingWork[0]) {
        manifest.pendingWork[0] = {
          id: 'synthetic-pending',
          format: 'shared',
          status: 'pending',
          pr: 1,
          base: `main@${BASELINE}`,
          branch: pendingBranch,
          head: pendingHead,
          targets: ['synthetic'],
        }
      }
      mutate(manifest)
    })
    includes(withPending((manifest) => { manifest.pendingWork[0].head = 'deadbeef' }), 'schema pattern')
    includes(withPending((manifest) => { manifest.pendingWork[0].head = null }), 'schema type string')
    includes(withPending((manifest) => { delete manifest.pendingWork[0].base }), 'schema-required field base')
    includes(withPending((manifest) => { manifest.pendingWork[0].head = '0'.repeat(40) }), 'head is stale for local')
    includes(withPending((manifest) => { manifest.pendingWork[0].base = `main@${'0'.repeat(40)}` }), 'base is stale for local')
  } finally {
    spawnSync('git', ['-C', testRoot, 'update-ref', '-d', `refs/heads/${pendingBranch}`], { encoding: 'utf8' })
  }
})
