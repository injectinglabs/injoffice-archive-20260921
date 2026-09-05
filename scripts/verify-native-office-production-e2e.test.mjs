import assert from 'node:assert/strict'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import {
  ADAPTER_PROTOCOL,
  DEFAULT_MANIFEST,
  OBSERVATION_PROTOCOL,
  canonicalJson,
  createAdapterRequest,
  loadFixtureBytes,
  loadManifest,
  parseStrictJson,
  semanticSha256,
  validateManifest,
  verifyFixtureBytes,
  verifyObservation,
  verifyObservationSet,
} from './verify-native-office-production-e2e.mjs'
import {
  HOST_OBSERVATION_CHECKIN,
  HOST_OBSERVATION_DIR,
  listAdapterRequestTargets,
  requestFileName,
  writeAdapterRequests,
} from './record-native-office-production-e2e.mjs'

const root = new URL('..', import.meta.url).pathname.replace(/\/$/, '')

function validatedManifest() {
  return validateManifest(loadManifest(), { root })
}

function clone(value) {
  return structuredClone(value)
}

function observation(validated, testCase, runner, replayIndex) {
  const fixture = validated.fixtures.get(testCase.fixtureId)
  const formatContract = validated.manifest.contracts[testCase.format]
  const nativeOutput = testCase.workflow === 'fixture-integrity' || fixture.corpusId === null || runner.nativeProtocol !== formatContract.protocol
    ? { outputProfile: testCase.expect.semantic.profile, outputSha256: testCase.expect.semantic.sha256 }
    : { outputProfile: fixture.canonicalExpectedProfile, outputSha256: fixture.canonicalExpectedSha256 }
  const components = runner.componentIds.map((identity) => {
    const separator = identity.indexOf(':')
    return { role: identity.slice(0, separator), id: identity.slice(separator + 1), revision: 'observed-revision' }
  })
  return {
    protocol: OBSERVATION_PROTOCOL,
    version: 1,
    caseId: testCase.id,
    runnerId: runner.id,
    replayIndex,
    fixture: { sha256: fixture.sha256, byteLength: fixture.byteLength },
    entrypoint: { repository: runner.repository, path: runner.entrypoint, operation: runner.operation },
    engine: { id: runner.engineId, revision: runner.engineRevision, components },
    native: { protocol: runner.nativeProtocol, version: runner.nativeVersion, status: testCase.expect.outcome, ...nativeOutput },
    semantic: clone(testCase.expect.semantic),
    evidence: {
      authoritySource: 'native-output',
      preservation: clone(testCase.expect.preservation),
      sourceUnchanged: true,
      refusal: clone(testCase.expect.refusal),
      candidateBytesProduced: false,
      forbiddenAuthoritySignals: [],
      productionSignals: clone(runner.requiredSignals),
    },
    resources: { inputBytes: fixture.byteLength, outputBytes: 1, outputItems: 1, durationMs: 1 },
  }
}

test('published manifest binds exact local fixtures, corpus expectations, schemas, and semantic digests', () => {
  const validated = validatedManifest()
  assert.equal(validated.manifest.protocol, 'injoffice.native-office-production-e2e/v1')
  assert.equal(validated.cases.size, 9)
  assert.equal(validated.fixtures.size, 8)
  for (const testCase of validated.cases.values()) {
    assert.equal(semanticSha256(testCase.expect.semantic.value), testCase.expect.semantic.sha256)
  }
})

test('strict JSON rejects duplicate keys and trailing content', () => {
  assert.throws(() => parseStrictJson('{"protocol":1,"protocol":1}'), /duplicate key/)
  assert.throws(() => parseStrictJson('{"protocol":1} null'), /trailing content/)
  assert.deepEqual(parseStrictJson('{"a":[true,false,null,-1.5e2]}'), { a: [true, false, null, -150] })
})

test('fixture tamper is rejected before an adapter can be invoked', () => {
  const validated = validatedManifest()
  const fixture = validated.fixtures.get('docx-inline-png-page-paint')
  const bytes = Buffer.from(loadFixtureBytes(validated, fixture.id))
  let adapterCalls = 0
  bytes[bytes.byteLength - 1] ^= 1
  assert.throws(() => {
    verifyFixtureBytes(fixture, bytes)
    adapterCalls += 1
  }, /FIXTURE_DIGEST_MISMATCH/)
  assert.equal(adapterCalls, 0)
  assert.throws(
    () => createAdapterRequest(validated, 'docx.atomic-fixture-tamper', 'contract.verifier', 0),
    /FIXTURE_DIGEST_MISMATCH/,
  )
})

test('adapter request is offline, bounded, and carries verifier-owned exact bytes', () => {
  const validated = validatedManifest()
  const request = createAdapterRequest(validated, 'xlsx.sheet-native-browser', 'browser.production-client', 0)
  assert.equal(request.protocol, ADAPTER_PROTOCOL)
  assert.equal(request.fixture.sha256, 'd5b0e821886fd43d7d484310e1a6c86ead6af813ce5c47bb2a2f33a31eea6bb2')
  assert.equal(Buffer.from(request.fixture.bytesBase64, 'base64').byteLength, request.fixture.byteLength)
  assert.equal(request.limits.maxObservationBytes, 67_108_864)
  assert.equal(request.target.entrypoint.path, 'browser/native-office-adapter')
  assert.equal(request.target.native.protocol, 'injoffice.xlsx.render-model')
  assert.equal(request.target.native.outputProfile, 'xlsx-native-canonical-v1')
  assert.equal(request.target.native.outputSha256, request.expect.semantic.sha256)
  assert.deepEqual(request.target.requiredSignals, ['hashed-production-chunk', 'native-root-or-refusal', 'no-raw-browser-source-read'])
  assert.equal(request.expect.semantic.sha256, 'b1a247a0127e81cc694d5a9c2e8e84f31fa97d62f414ad3c677bdab2e25031c5')

  const parserRequest = createAdapterRequest(validated, 'pptx.picture-native-browser', 'application.production-service', 0)
  assert.equal(parserRequest.target.native.outputSha256, '8b794073933bfa14554e34d528dee342b00fb5e0bffac9e17682a9714b656b4e')
  assert.equal(canonicalJson(request).includes('http://'), false)
  assert.equal(canonicalJson(request).includes('https://'), false)
})

test('manifest tamper, authority weakening, path escape, and output-digest drift fail closed', () => {
  const cases = [
    [(candidate) => { candidate.authority.forbidden.pop() }, /exact mandatory no-legacy set/],
    [(candidate) => { candidate.cases[0].expect.semantic.sha256 = '0'.repeat(64) }, /semantic\.sha256 is stale/],
    [(candidate) => { candidate.fixtures[0].path = '../escape.docx' }, /not canonical-safe|escapes/],
  ]
  for (const [mutate, pattern] of cases) {
    const candidate = clone(loadManifest())
    mutate(candidate)
    assert.throws(() => validateManifest(candidate, { root }), pattern)
  }
})

test('observation verifier binds protocol, engine, semantic output, preservation, refusal, and budgets', () => {
  const validated = validatedManifest()
  const testCase = validated.cases.get('docx.image-page-paint')
  const runner = testCase.runners[0]
  const valid = observation(validated, testCase, runner, 0)
  assert.equal(verifyObservation(validated, valid), valid)
  const mutations = [
    [(candidate) => { candidate.engine.id = 'legacy-renderer' }, /engine identity mismatch/],
    [(candidate) => { candidate.native.protocol = 'html' }, /native identity/],
    [(candidate) => { candidate.native.outputSha256 = '0'.repeat(64) }, /authoritative output digest/],
    [(candidate) => { candidate.semantic.value.status = 'refused' }, /semantic value/],
    [(candidate) => { candidate.evidence.authoritySource = 'screenshot' }, /non-native semantic authority/],
    [(candidate) => { candidate.evidence.forbiddenAuthoritySignals = ['canvas-measurement'] }, /forbidden authority/],
    [(candidate) => { candidate.resources.outputBytes = 67_108_865 }, /exceeds/],
  ]
  for (const [mutate, pattern] of mutations) {
    const candidate = clone(valid)
    mutate(candidate)
    assert.throws(() => verifyObservation(validated, candidate), pattern)
  }
})

test('refusal observations cannot expose partial semantic or candidate output', () => {
  const validated = validatedManifest()
  const testCase = validated.cases.get('xlsx.unsupported-numeric-refusal')
  const valid = observation(validated, testCase, testCase.runners[0], 0)
  verifyObservation(validated, valid)
  const candidate = clone(valid)
  candidate.evidence.candidateBytesProduced = true
  assert.throws(() => verifyObservation(validated, candidate), /candidate bytes/)
  candidate.evidence.candidateBytesProduced = false
  candidate.evidence.refusal.atomic = false
  assert.throws(() => verifyObservation(validated, candidate), /refusal mismatch/)
})

test('complete observation sets require deterministic replay for every case and runner', () => {
  const validated = validatedManifest()
  const observations = []
  for (const testCase of validated.cases.values()) {
    for (const runner of testCase.runners) {
      for (let replayIndex = 0; replayIndex < testCase.replay.runs; replayIndex += 1) observations.push(observation(validated, testCase, runner, replayIndex))
    }
  }
  assert.equal(verifyObservationSet(validated, observations), observations)

  const missing = observations.slice(1)
  assert.throws(() => verifyObservationSet(validated, missing), /has 1 observations, want 2/)

  const nondeterministic = clone(observations)
  const pair = nondeterministic.filter((entry) => entry.caseId === 'pptx.authored-deck' && entry.runnerId === 'browser.production-client')
  pair[1].engine.components[0].revision = 'replay-drift'
  assert.throws(() => verifyObservationSet(validated, nondeterministic), /engine changed across replay/)
})

test('manifest file itself remains strict JSON with no placeholder digests', () => {
  const source = readFileSync(DEFAULT_MANIFEST, 'utf8')
  parseStrictJson(source, DEFAULT_MANIFEST)
  assert.equal(source.includes('REPLACE_'), false)
  assert.equal(source.includes('toHaveScreenshot'), false)
  assert.equal(source.includes('page.screenshot'), false)
})

test('host observation recorder writes adapter requests the downstream host must fill', () => {
  const validated = validatedManifest()
  const { targets, refusals } = listAdapterRequestTargets(validated)
  assert.equal(HOST_OBSERVATION_CHECKIN, `${HOST_OBSERVATION_DIR}/observations.json`)
  assert.equal(refusals.length, 2)
  assert.equal(refusals[0].caseId, 'docx.atomic-fixture-tamper')
  assert.equal(targets.some((target) => target.caseId === 'docx.atomic-fixture-tamper'), false)

  const directory = mkdtempSync(join(tmpdir(), 'injoffice-e2e-requests-'))
  try {
    const result = writeAdapterRequests(validated, directory)
    assert.equal(result.written.length, targets.length)
    const sample = result.written.find((entry) => entry.caseId === 'xlsx.sheet-native-browser' && entry.runnerId === 'browser.production-client' && entry.replayIndex === 0)
    assert.equal(sample.protocol, ADAPTER_PROTOCOL)
    const request = parseStrictJson(readFileSync(join(directory, sample.fileName), 'utf8'))
    assert.equal(canonicalJson(request), canonicalJson(createAdapterRequest(validated, sample.caseId, sample.runnerId, sample.replayIndex)))
    assert.equal(requestFileName(sample.caseId, sample.runnerId, sample.replayIndex), sample.fileName)
    const tamper = parseStrictJson(readFileSync(join(directory, 'tamper-refusals.json'), 'utf8'))
    assert.equal(tamper.observationProtocol, OBSERVATION_PROTOCOL)
    assert.equal(tamper.observationCheckIn, HOST_OBSERVATION_CHECKIN)
    assert.equal(tamper.code, 'FIXTURE_DIGEST_MISMATCH')
    assert.equal(tamper.refusals.length, 2)
    assert.throws(
      () => createAdapterRequest(validated, 'docx.atomic-fixture-tamper', 'contract.verifier', 0),
      /FIXTURE_DIGEST_MISMATCH/,
    )
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})
