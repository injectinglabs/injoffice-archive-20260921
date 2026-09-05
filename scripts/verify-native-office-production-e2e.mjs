#!/usr/bin/env node

import { createHash } from 'node:crypto'
import { lstatSync, readFileSync, realpathSync, statSync } from 'node:fs'
import { dirname, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

export const MANIFEST_PROTOCOL = 'injoffice.native-office-production-e2e/v1'
export const ADAPTER_PROTOCOL = 'injoffice.native-office-production-e2e-adapter/v1'
export const OBSERVATION_PROTOCOL = 'injoffice.native-office-production-e2e-observation/v1'
export const BASELINE_COMMIT = 'ddf2e523a9d6422e2dfccd93e263fcbb10322cca'

const scriptRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
export const DEFAULT_MANIFEST = resolve(scriptRoot, 'testdata/native-office-production-e2e/v1/manifest.json')
const SHA256 = /^[a-f0-9]{64}$/
const ID = /^[a-z0-9]+(?:[.-][a-z0-9]+)*$/
const FORMAT = new Set(['docx', 'pptx', 'xlsx'])
const REQUIRED_FORBIDDEN = [
  'canvas-measurement',
  'dom',
  'html',
  'konva',
  'legacy-deckview',
  'mammoth',
  'screenshot',
]

function fail(message) {
  throw new Error(message)
}

function object(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function exactKeys(value, required, optional, context) {
  if (!object(value)) fail(`${context} must be an object`)
  const allowed = new Set([...required, ...optional])
  for (const key of Object.keys(value)) if (!allowed.has(key)) fail(`${context} has unknown field ${key}`)
  for (const key of required) if (!(key in value)) fail(`${context} is missing ${key}`)
}

function nonemptyString(value, context) {
  if (typeof value !== 'string' || value.length === 0 || value.length > 512) fail(`${context} must be a bounded non-empty string`)
}

function canonicalStrings(value, context, { nonempty = true } = {}) {
  if (!Array.isArray(value) || (nonempty && value.length === 0)) fail(`${context} must be ${nonempty ? 'a non-empty' : 'an'} array`)
  for (const entry of value) nonemptyString(entry, `${context} entry`)
  const sorted = [...value].sort()
  if (new Set(value).size !== value.length || JSON.stringify(value) !== JSON.stringify(sorted)) fail(`${context} must be sorted and unique`)
}

export function parseStrictJson(source, label = 'JSON') {
  if (typeof source !== 'string') fail(`${label} must be text`)
  let index = 0
  const whitespace = () => { while (/\s/.test(source[index] ?? '')) index += 1 }
  const value = (path) => {
    whitespace()
    const start = index
    const token = source[index]
    if (token === '{') return record(path)
    if (token === '[') return list(path)
    if (token === '"') return string(path)
    for (const [literal, parsed] of [['true', true], ['false', false], ['null', null]]) {
      if (source.startsWith(literal, index)) { index += literal.length; return parsed }
    }
    const match = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(source.slice(index))
    if (!match) fail(`${label} has invalid token at ${path} byte ${start}`)
    index += match[0].length
    const parsed = Number(match[0])
    if (!Number.isFinite(parsed) || Object.is(parsed, -0)) fail(`${label} has unsupported number at ${path}`)
    return parsed
  }
  const string = (path) => {
    const start = index++
    let escaped = false
    while (index < source.length) {
      const code = source.charCodeAt(index)
      if (!escaped && code === 0x22) {
        index += 1
        try { return JSON.parse(source.slice(start, index)) } catch { fail(`${label} has invalid string at ${path}`) }
      }
      if (!escaped && code < 0x20) fail(`${label} has unescaped control character at ${path}`)
      escaped = !escaped && code === 0x5c
      if (code !== 0x5c) escaped = false
      index += 1
    }
    fail(`${label} has unterminated string at ${path}`)
  }
  const record = (path) => {
    index += 1
    const result = {}
    const keys = new Set()
    whitespace()
    if (source[index] === '}') { index += 1; return result }
    while (index < source.length) {
      whitespace()
      if (source[index] !== '"') fail(`${label} has non-string object key at ${path}`)
      const key = string(path)
      if (keys.has(key)) fail(`${label} has duplicate key ${JSON.stringify(key)} at ${path}`)
      keys.add(key)
      whitespace()
      if (source[index++] !== ':') fail(`${label} is missing ':' after ${path}.${key}`)
      result[key] = value(`${path}.${key}`)
      whitespace()
      const next = source[index++]
      if (next === '}') return result
      if (next !== ',') fail(`${label} is missing ',' at ${path}`)
    }
    fail(`${label} has unterminated object at ${path}`)
  }
  const list = (path) => {
    index += 1
    const result = []
    whitespace()
    if (source[index] === ']') { index += 1; return result }
    while (index < source.length) {
      result.push(value(`${path}[${result.length}]`))
      whitespace()
      const next = source[index++]
      if (next === ']') return result
      if (next !== ',') fail(`${label} is missing ',' at ${path}`)
    }
    fail(`${label} has unterminated array at ${path}`)
  }
  const parsed = value('$')
  whitespace()
  if (index !== source.length) fail(`${label} has trailing content at byte ${index}`)
  return parsed
}

export function canonicalJson(value) {
  if (value === null || typeof value !== 'object') {
    const encoded = JSON.stringify(value)
    if (encoded === undefined) fail('canonical JSON cannot encode undefined or non-JSON values')
    return encoded
  }
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`
  const keys = Object.keys(value).sort()
  return `{${keys.map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(',')}}`
}

export function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

export function semanticSha256(value) {
  return sha256(canonicalJson(value))
}

function safeFile(root, relativePath, context) {
  if (typeof relativePath !== 'string' || relativePath.length === 0 || relativePath.startsWith('/') || relativePath.includes('\\')) fail(`${context} must be repository-relative`)
  const parts = relativePath.split('/')
  if (parts.some((part) => part === '' || part === '.' || part === '..' || part.includes('%'))) fail(`${context} is not canonical-safe`)
  const absolute = resolve(root, relativePath)
  if (absolute !== root && !absolute.startsWith(root + sep)) fail(`${context} escapes the repository`)
  let cursor = root
  for (const part of parts) {
    cursor = resolve(cursor, part)
    if (lstatSync(cursor).isSymbolicLink()) fail(`${context} traverses a symbolic link`)
  }
  if (!statSync(absolute).isFile()) fail(`${context} is not a file`)
  if (realpathSync(absolute) !== absolute) fail(`${context} is not a canonical file path`)
  return absolute
}

function checkedFile(root, relativePath, expectedBytes, expectedSha256, context) {
  if (!SHA256.test(expectedSha256)) fail(`${context}.sha256 must be lowercase SHA-256`)
  const path = safeFile(root, relativePath, `${context}.path`)
  const bytes = readFileSync(path)
  if (expectedBytes !== null && bytes.byteLength !== expectedBytes) fail(`${context} byte length drift: got ${bytes.byteLength}, want ${expectedBytes}`)
  const actual = sha256(bytes)
  if (actual !== expectedSha256) fail(`${context} digest drift: got ${actual}, want ${expectedSha256}`)
  return bytes
}

function validateRunner(runner, context) {
  exactKeys(runner, ['id', 'repository', 'entrypoint', 'operation', 'engineId', 'engineRevision', 'componentIds', 'requiredSignals', 'nativeProtocol', 'nativeVersion'], [], context)
  if (!ID.test(runner.id)) fail(`${context}.id is not canonical`)
  for (const field of ['repository', 'entrypoint', 'operation', 'engineId', 'engineRevision', 'nativeProtocol']) nonemptyString(runner[field], `${context}.${field}`)
  if (runner.entrypoint.startsWith('/') || runner.entrypoint.split('/').some((part) => part === '..' || part === '')) fail(`${context}.entrypoint must be repository-relative`)
  canonicalStrings(runner.componentIds, `${context}.componentIds`)
  canonicalStrings(runner.requiredSignals, `${context}.requiredSignals`)
  for (const component of runner.componentIds) if (!/^[a-z0-9-]+:.+/.test(component)) fail(`${context}.componentIds entries must be role:id identities`)
  if (runner.nativeVersion !== 1) fail(`${context}.nativeVersion must equal 1`)
}

function validateExpectation(expectation, context) {
  exactKeys(expectation, ['outcome', 'semantic', 'preservation', 'refusal'], [], context)
  if (!['accepted', 'refused'].includes(expectation.outcome)) fail(`${context}.outcome is invalid`)
  exactKeys(expectation.semantic, ['profile', 'value', 'sha256'], [], `${context}.semantic`)
  nonemptyString(expectation.semantic.profile, `${context}.semantic.profile`)
  if (!object(expectation.semantic.value)) fail(`${context}.semantic.value must be an object`)
  if (semanticSha256(expectation.semantic.value) !== expectation.semantic.sha256) fail(`${context}.semantic.sha256 is stale`)
  canonicalStrings(expectation.preservation, `${context}.preservation`)
  if (expectation.outcome === 'accepted') {
    if (expectation.refusal !== null) fail(`${context}.refusal must be null for accepted cases`)
  } else {
    exactKeys(expectation.refusal, ['code', 'stage', 'atomic', 'candidateBytesProduced'], [], `${context}.refusal`)
    nonemptyString(expectation.refusal.code, `${context}.refusal.code`)
    nonemptyString(expectation.refusal.stage, `${context}.refusal.stage`)
    if (expectation.refusal.atomic !== true || expectation.refusal.candidateBytesProduced !== false) fail(`${context}.refusal must be atomic with no candidate bytes`)
  }
}

export function loadManifest(path = DEFAULT_MANIFEST) {
  return parseStrictJson(readFileSync(path, 'utf8'), path)
}

export function validateManifest(manifest, { root = scriptRoot } = {}) {
  exactKeys(manifest, ['protocol', 'version', 'baseline', 'authority', 'adapter', 'contracts', 'resourceProfiles', 'fixtures', 'cases'], [], 'manifest')
  if (manifest.protocol !== MANIFEST_PROTOCOL || manifest.version !== 1) fail('unsupported production E2E manifest protocol/version')
  exactKeys(manifest.baseline, ['repository', 'commit', 'corpusManifest', 'completionMatrix'], [], 'manifest.baseline')
  if (manifest.baseline.repository !== 'injectinglabs/injoffice' || manifest.baseline.commit !== BASELINE_COMMIT) fail('manifest baseline drift')
  exactKeys(manifest.baseline.completionMatrix, ['protocol', 'relationship'], [], 'manifest.baseline.completionMatrix')
  if (manifest.baseline.completionMatrix.protocol !== 'injoffice.native-office-completion/v1') fail('completion matrix protocol drift')
  if (/capabilit(?:y|ies)|pending work|test selector/i.test(manifest.baseline.completionMatrix.relationship) === false) fail('completion matrix relationship must state the non-duplication boundary')

  exactKeys(manifest.authority, ['semanticSource', 'digestProfile', 'allowedPaintSinks', 'forbidden', 'screenshotPolicy'], [], 'manifest.authority')
  if (manifest.authority.semanticSource !== 'native-output' || manifest.authority.digestProfile !== 'sha256-canonical-json/v1') fail('native canonical output must be semantic authority')
  canonicalStrings(manifest.authority.allowedPaintSinks, 'manifest.authority.allowedPaintSinks')
  canonicalStrings(manifest.authority.forbidden, 'manifest.authority.forbidden')
  if (JSON.stringify(manifest.authority.forbidden) !== JSON.stringify(REQUIRED_FORBIDDEN)) fail('manifest.authority.forbidden must contain the exact mandatory no-legacy set')
  if (manifest.authority.screenshotPolicy !== 'diagnostic-only-never-semantic-authority') fail('screenshots must never be semantic authority')

  exactKeys(manifest.adapter, ['protocol', 'version', 'requestMode', 'networkPolicy', 'observationProtocol'], [], 'manifest.adapter')
  if (manifest.adapter.protocol !== ADAPTER_PROTOCOL || manifest.adapter.version !== 1 || manifest.adapter.observationProtocol !== OBSERVATION_PROTOCOL) fail('adapter protocol drift')
  if (manifest.adapter.networkPolicy !== 'fixture-local-or-intercepted-only') fail('adapter must remain offline/intercepted')

  exactKeys(manifest.baseline.corpusManifest, ['path', 'protocol', 'selectedRecordsSha256'], [], 'manifest.baseline.corpusManifest')
  if (!SHA256.test(manifest.baseline.corpusManifest.selectedRecordsSha256)) fail('manifest.baseline.corpusManifest.selectedRecordsSha256 must be lowercase SHA-256')
  const corpusPath = safeFile(root, manifest.baseline.corpusManifest.path, 'manifest.baseline.corpusManifest.path')
  const corpusBytes = readFileSync(corpusPath)
  const corpus = parseStrictJson(corpusBytes.toString('utf8'), 'officecompat corpus manifest')
  if (corpus.protocol !== manifest.baseline.corpusManifest.protocol) fail('officecompat corpus protocol drift')
  const corpusByID = new Map(corpus.fixtures.map((fixture) => [fixture.id, fixture]))

  exactKeys(manifest.contracts, ['docx', 'pptx', 'xlsx'], [], 'manifest.contracts')
  for (const format of FORMAT) {
    const contract = manifest.contracts[format]
    exactKeys(contract, ['protocol', 'version', 'schemaPath', 'schemaSha256'], [], `manifest.contracts.${format}`)
    if (contract.version !== 1) fail(`${format} contract version drift`)
    checkedFile(root, contract.schemaPath, null, contract.schemaSha256, `manifest.contracts.${format}`)
  }

  const profiles = new Map()
  for (const [id, profile] of Object.entries(manifest.resourceProfiles)) {
    if (!ID.test(id)) fail(`resource profile ${id} is not canonical`)
    exactKeys(profile, ['maxFixtureBytes', 'maxObservationBytes', 'maxOutputItems', 'timeoutMs'], [], `resourceProfiles.${id}`)
    for (const field of ['maxFixtureBytes', 'maxObservationBytes', 'maxOutputItems', 'timeoutMs']) if (!Number.isSafeInteger(profile[field]) || profile[field] <= 0) fail(`resourceProfiles.${id}.${field} must be a positive safe integer`)
    profiles.set(id, profile)
  }

  const fixtures = new Map()
  const selectedCorpusRecords = []
  let priorFixture = ''
  for (const [index, fixture] of manifest.fixtures.entries()) {
    const context = `manifest.fixtures[${index}]`
    exactKeys(fixture, ['id', 'format', 'mediaType', 'path', 'byteLength', 'sha256', 'expectedPath', 'expectedSha256', 'canonicalExpectedProfile', 'canonicalExpectedSha256', 'corpusId'], [], context)
    if (!ID.test(fixture.id) || fixture.id <= priorFixture) fail(`${context}.id must be canonical and strictly ordered`)
    priorFixture = fixture.id
    if (!FORMAT.has(fixture.format)) fail(`${context}.format is invalid`)
    nonemptyString(fixture.mediaType, `${context}.mediaType`)
    if (!Number.isSafeInteger(fixture.byteLength) || fixture.byteLength <= 0) fail(`${context}.byteLength must be positive`)
    checkedFile(root, fixture.path, fixture.byteLength, fixture.sha256, context)
    nonemptyString(fixture.canonicalExpectedProfile, `${context}.canonicalExpectedProfile`)
    if (!SHA256.test(fixture.canonicalExpectedSha256)) fail(`${context}.canonicalExpectedSha256 must be lowercase SHA-256`)
    if (fixtures.has(fixture.id)) fail(`${context}.id is duplicated`)
    fixtures.set(fixture.id, fixture)
    if (fixture.corpusId === null) {
      if (fixture.expectedPath !== null || fixture.expectedSha256 !== null) fail(`${context} non-corpus fixture must not claim corpus expectations`)
    } else {
      const corpusFixture = corpusByID.get(fixture.corpusId)
      if (!corpusFixture) fail(`${context}.corpusId is missing from the corpus`)
      selectedCorpusRecords.push(corpusFixture)
      const prefix = dirname(manifest.baseline.corpusManifest.path)
      const expectedPackagePath = `${prefix}/${corpusFixture.package}`
      const expectedOutputPath = `${prefix}/${corpusFixture.expected}`
      if (fixture.id !== corpusFixture.id || fixture.path !== expectedPackagePath || fixture.sha256 !== corpusFixture.sha256 || fixture.byteLength !== corpusFixture.bytes) fail(`${context} drifts from corpus package identity`)
      if (fixture.expectedPath !== expectedOutputPath || fixture.expectedSha256 !== corpusFixture.expectedSha256) fail(`${context} drifts from corpus expected output identity`)
      const expectedBytes = checkedFile(root, fixture.expectedPath, null, fixture.expectedSha256, `${context}.expected`)
      const expected = parseStrictJson(expectedBytes.toString('utf8'), `${context}.expected`)
      const canonicalExpected = expected.native ?? expected.refusal
      if (!object(canonicalExpected) || semanticSha256(canonicalExpected) !== fixture.canonicalExpectedSha256) fail(`${context}.canonicalExpectedSha256 drifts from corpus native/refusal output`)
    }
  }
  if (semanticSha256(selectedCorpusRecords) !== manifest.baseline.corpusManifest.selectedRecordsSha256) fail('manifest.baseline.corpusManifest selected record digest drift')

  const cases = new Map()
  let priorCase = ''
  const positiveFormats = new Set()
  const refusedFormats = new Set()
  for (const [index, testCase] of manifest.cases.entries()) {
    const context = `manifest.cases[${index}]`
    exactKeys(testCase, ['id', 'format', 'workflow', 'fixtureId', 'input', 'resourceProfile', 'runners', 'expect', 'replay'], [], context)
    if (!ID.test(testCase.id) || testCase.id <= priorCase) fail(`${context}.id must be canonical and strictly ordered`)
    priorCase = testCase.id
    if (!FORMAT.has(testCase.format) || !testCase.id.startsWith(`${testCase.format}.`)) fail(`${context} format/id mismatch`)
    nonemptyString(testCase.workflow, `${context}.workflow`)
    const fixture = fixtures.get(testCase.fixtureId)
    if (!fixture || fixture.format !== testCase.format) fail(`${context}.fixtureId is unknown or format-mismatched`)
    if (testCase.input?.mode === 'exact') {
      exactKeys(testCase.input, ['mode'], [], `${context}.input`)
    } else if (testCase.input?.mode === 'tamper-xor') {
      exactKeys(testCase.input, ['mode', 'byteOffset', 'xorMask'], [], `${context}.input`)
      if (testCase.workflow !== 'fixture-integrity') fail(`${context}.input tamper is only valid for fixture-integrity`)
      if (!Number.isSafeInteger(testCase.input.byteOffset) || testCase.input.byteOffset < 0 || testCase.input.byteOffset >= fixture.byteLength) fail(`${context}.input.byteOffset is outside the fixture`)
      if (!Number.isSafeInteger(testCase.input.xorMask) || testCase.input.xorMask < 1 || testCase.input.xorMask > 255) fail(`${context}.input.xorMask must be between 1 and 255`)
    } else {
      fail(`${context}.input.mode is invalid`)
    }
    const profile = profiles.get(testCase.resourceProfile)
    if (!profile) fail(`${context}.resourceProfile is unknown`)
    if (fixture.byteLength > profile.maxFixtureBytes) fail(`${context} fixture exceeds its resource profile`)
    if (!Array.isArray(testCase.runners) || testCase.runners.length === 0) fail(`${context}.runners must be non-empty`)
    let priorRunner = ''
    for (const [runnerIndex, runner] of testCase.runners.entries()) {
      validateRunner(runner, `${context}.runners[${runnerIndex}]`)
      if (runner.id <= priorRunner) fail(`${context}.runners must be strictly ordered by id`)
      priorRunner = runner.id
    }
    validateExpectation(testCase.expect, `${context}.expect`)
    exactKeys(testCase.replay, ['runs', 'requireIdentical'], [], `${context}.replay`)
    if (!Number.isSafeInteger(testCase.replay.runs) || testCase.replay.runs < 2 || testCase.replay.runs > 8) fail(`${context}.replay.runs must be between 2 and 8`)
    canonicalStrings(testCase.replay.requireIdentical, `${context}.replay.requireIdentical`)
    if (JSON.stringify(testCase.replay.requireIdentical) !== JSON.stringify(['engine', 'evidence', 'semanticSha256'])) fail(`${context}.replay must bind engine, evidence, and semanticSha256`)
    if (testCase.expect.outcome === 'accepted') positiveFormats.add(testCase.format)
    else refusedFormats.add(testCase.format)
    cases.set(testCase.id, testCase)
  }
  for (const format of FORMAT) {
    if (!positiveFormats.has(format)) fail(`${format} has no positive production case`)
    if (!refusedFormats.has(format)) fail(`${format} has no atomic/refusal case`)
  }
  for (const required of ['docx.image-page-paint', 'docx.table-font-page-paint', 'pptx.authored-deck', 'pptx.picture-native-browser', 'xlsx.sheet-native-browser', 'docx.atomic-fixture-tamper']) if (!cases.has(required)) fail(`required production case ${required} is missing`)

  return { manifest, root, fixtures, cases, profiles }
}

export function verifyFixtureBytes(fixture, bytes) {
  if (!(bytes instanceof Uint8Array)) fail('fixture bytes must be a Uint8Array')
  if (bytes.byteLength !== fixture.byteLength) fail(`FIXTURE_DIGEST_MISMATCH: fixture byte length is ${bytes.byteLength}, want ${fixture.byteLength}`)
  const actual = sha256(bytes)
  if (actual !== fixture.sha256) fail(`FIXTURE_DIGEST_MISMATCH: fixture SHA-256 is ${actual}, want ${fixture.sha256}`)
  return bytes
}

export function loadFixtureBytes(validated, fixtureId) {
  const fixture = validated.fixtures.get(fixtureId)
  if (!fixture) fail(`unknown fixture ${fixtureId}`)
  const path = safeFile(validated.root, fixture.path, `fixture ${fixtureId}`)
  return verifyFixtureBytes(fixture, readFileSync(path))
}

function expectedNativeOutput(validated, testCase, runner, fixture) {
  const formatContract = validated.manifest.contracts[testCase.format]
  if (testCase.workflow !== 'fixture-integrity' && fixture.corpusId !== null && runner.nativeProtocol === formatContract.protocol) {
    return { profile: fixture.canonicalExpectedProfile, sha256: fixture.canonicalExpectedSha256 }
  }
  return { profile: testCase.expect.semantic.profile, sha256: testCase.expect.semantic.sha256 }
}

export function createAdapterRequest(validated, caseId, runnerId, replayIndex) {
  const testCase = validated.cases.get(caseId)
  if (!testCase) fail(`unknown case ${caseId}`)
  const runner = testCase.runners.find((candidate) => candidate.id === runnerId)
  if (!runner) fail(`case ${caseId} has no runner ${runnerId}`)
  if (!Number.isSafeInteger(replayIndex) || replayIndex < 0 || replayIndex >= testCase.replay.runs) fail(`invalid replay index ${replayIndex}`)
  const fixture = validated.fixtures.get(testCase.fixtureId)
  const nativeOutput = expectedNativeOutput(validated, testCase, runner, fixture)
  let bytes = loadFixtureBytes(validated, fixture.id)
  if (testCase.input.mode === 'tamper-xor') {
    bytes = Buffer.from(bytes)
    bytes[testCase.input.byteOffset] ^= testCase.input.xorMask
    verifyFixtureBytes(fixture, bytes)
    fail('tampered fixture unexpectedly passed integrity verification')
  }
  return {
    protocol: ADAPTER_PROTOCOL,
    version: 1,
    caseId,
    runnerId,
    replayIndex,
    format: testCase.format,
    workflow: testCase.workflow,
    fixture: { mediaType: fixture.mediaType, byteLength: fixture.byteLength, sha256: fixture.sha256, bytesBase64: Buffer.from(bytes).toString('base64') },
    target: {
      entrypoint: { repository: runner.repository, path: runner.entrypoint, operation: runner.operation },
      engine: { id: runner.engineId, revision: runner.engineRevision, componentIds: structuredClone(runner.componentIds) },
      native: { protocol: runner.nativeProtocol, version: runner.nativeVersion, outputProfile: nativeOutput.profile, outputSha256: nativeOutput.sha256 },
      requiredSignals: structuredClone(runner.requiredSignals),
    },
    expect: structuredClone(testCase.expect),
    limits: structuredClone(validated.profiles.get(testCase.resourceProfile)),
  }
}

function validateObservationShape(observation, context) {
  exactKeys(observation, ['protocol', 'version', 'caseId', 'runnerId', 'replayIndex', 'fixture', 'entrypoint', 'engine', 'native', 'semantic', 'evidence', 'resources'], [], context)
  exactKeys(observation.fixture, ['sha256', 'byteLength'], [], `${context}.fixture`)
  exactKeys(observation.entrypoint, ['repository', 'path', 'operation'], [], `${context}.entrypoint`)
  exactKeys(observation.engine, ['id', 'revision', 'components'], [], `${context}.engine`)
  if (!Array.isArray(observation.engine.components)) fail(`${context}.engine.components must be an array`)
  for (const [index, component] of observation.engine.components.entries()) {
    exactKeys(component, ['role', 'id', 'revision'], [], `${context}.engine.components[${index}]`)
    for (const field of ['role', 'id', 'revision']) nonemptyString(component[field], `${context}.engine.components[${index}].${field}`)
  }
  exactKeys(observation.native, ['protocol', 'version', 'status', 'outputProfile', 'outputSha256'], [], `${context}.native`)
  exactKeys(observation.semantic, ['profile', 'value', 'sha256'], [], `${context}.semantic`)
  exactKeys(observation.evidence, ['authoritySource', 'preservation', 'sourceUnchanged', 'refusal', 'candidateBytesProduced', 'forbiddenAuthoritySignals', 'productionSignals'], [], `${context}.evidence`)
  exactKeys(observation.resources, ['inputBytes', 'outputBytes', 'outputItems', 'durationMs'], [], `${context}.resources`)
}

export function verifyObservation(validated, observation, context = 'observation') {
  validateObservationShape(observation, context)
  if (observation.protocol !== OBSERVATION_PROTOCOL || observation.version !== 1) fail(`${context} protocol/version drift`)
  const testCase = validated.cases.get(observation.caseId)
  if (!testCase) fail(`${context}.caseId is unknown`)
  const runner = testCase.runners.find((candidate) => candidate.id === observation.runnerId)
  if (!runner) fail(`${context}.runnerId is not allowed for ${testCase.id}`)
  if (!Number.isSafeInteger(observation.replayIndex) || observation.replayIndex < 0 || observation.replayIndex >= testCase.replay.runs) fail(`${context}.replayIndex is invalid`)
  const fixture = validated.fixtures.get(testCase.fixtureId)
  if (observation.fixture.sha256 !== fixture.sha256 || observation.fixture.byteLength !== fixture.byteLength) fail(`${context}.fixture does not echo exact verified bytes`)
  if (canonicalJson(observation.entrypoint) !== canonicalJson({ repository: runner.repository, path: runner.entrypoint, operation: runner.operation })) fail(`${context}.entrypoint identity mismatch`)
  if (observation.engine.id !== runner.engineId || observation.engine.revision !== runner.engineRevision) fail(`${context}.engine identity mismatch`)
  const componentRoles = observation.engine.components.map((component) => component.role)
  if (new Set(componentRoles).size !== componentRoles.length || JSON.stringify(componentRoles) !== JSON.stringify([...componentRoles].sort())) fail(`${context}.engine.components must have sorted unique roles`)
  const componentIds = observation.engine.components.map((component) => `${component.role}:${component.id}`)
  if (canonicalJson(componentIds) !== canonicalJson(runner.componentIds)) fail(`${context}.engine component identity mismatch`)
  if (observation.native.protocol !== runner.nativeProtocol || observation.native.version !== runner.nativeVersion || observation.native.status !== testCase.expect.outcome) fail(`${context}.native identity/outcome mismatch`)
  const nativeOutput = expectedNativeOutput(validated, testCase, runner, fixture)
  if (observation.native.outputProfile !== nativeOutput.profile || observation.native.outputSha256 !== nativeOutput.sha256) fail(`${context}.native authoritative output digest mismatch`)
  if (observation.semantic.profile !== testCase.expect.semantic.profile || observation.semantic.sha256 !== testCase.expect.semantic.sha256) fail(`${context}.semantic identity mismatch`)
  if (semanticSha256(observation.semantic.value) !== observation.semantic.sha256 || canonicalJson(observation.semantic.value) !== canonicalJson(testCase.expect.semantic.value)) fail(`${context}.semantic value is not authoritative expected output`)
  if (observation.evidence.authoritySource !== 'native-output') fail(`${context} uses non-native semantic authority`)
  canonicalStrings(observation.evidence.preservation, `${context}.evidence.preservation`)
  if (canonicalJson(observation.evidence.preservation) !== canonicalJson(testCase.expect.preservation)) fail(`${context}.evidence.preservation mismatch`)
  if (observation.evidence.sourceUnchanged !== true || observation.evidence.candidateBytesProduced !== false) fail(`${context} changed source or produced uncommitted candidate bytes`)
  canonicalStrings(observation.evidence.forbiddenAuthoritySignals, `${context}.evidence.forbiddenAuthoritySignals`, { nonempty: false })
  if (observation.evidence.forbiddenAuthoritySignals.length !== 0) fail(`${context} reported forbidden authority: ${observation.evidence.forbiddenAuthoritySignals.join(', ')}`)
  canonicalStrings(observation.evidence.productionSignals, `${context}.evidence.productionSignals`)
  if (canonicalJson(observation.evidence.productionSignals) !== canonicalJson(runner.requiredSignals)) fail(`${context}.evidence.productionSignals mismatch`)
  if (testCase.expect.refusal === null) {
    if (observation.evidence.refusal !== null) fail(`${context} accepted output carries refusal evidence`)
  } else if (canonicalJson(observation.evidence.refusal) !== canonicalJson(testCase.expect.refusal)) {
    fail(`${context}.evidence.refusal mismatch`)
  }
  const profile = validated.profiles.get(testCase.resourceProfile)
  const limits = { inputBytes: profile.maxFixtureBytes, outputBytes: profile.maxObservationBytes, outputItems: profile.maxOutputItems, durationMs: profile.timeoutMs }
  for (const [field, limit] of Object.entries(limits)) {
    const amount = observation.resources[field]
    if (!Number.isSafeInteger(amount) || amount < 0 || amount > limit) fail(`${context}.resources.${field} exceeds ${limit}`)
  }
  if (observation.resources.inputBytes !== fixture.byteLength) fail(`${context}.resources.inputBytes must equal exact fixture bytes`)
  if (Buffer.byteLength(canonicalJson(observation)) > profile.maxObservationBytes) fail(`${context} encoded observation exceeds ${profile.maxObservationBytes} bytes`)
  return observation
}

export function verifyObservationSet(validated, observations) {
  if (!Array.isArray(observations)) fail('observations must be an array')
  const groups = new Map()
  for (const [index, observation] of observations.entries()) {
    verifyObservation(validated, observation, `observations[${index}]`)
    const key = `${observation.caseId}\0${observation.runnerId}`
    const group = groups.get(key) ?? []
    group.push(observation)
    groups.set(key, group)
  }
  for (const testCase of validated.cases.values()) for (const runner of testCase.runners) {
    const key = `${testCase.id}\0${runner.id}`
    const group = groups.get(key) ?? []
    if (group.length !== testCase.replay.runs) fail(`${testCase.id}/${runner.id} has ${group.length} observations, want ${testCase.replay.runs}`)
    group.sort((left, right) => left.replayIndex - right.replayIndex)
    for (let index = 0; index < group.length; index += 1) if (group[index].replayIndex !== index) fail(`${testCase.id}/${runner.id} replay indices are incomplete or duplicated`)
    const first = group[0]
    for (const candidate of group.slice(1)) {
      if (canonicalJson(candidate.engine) !== canonicalJson(first.engine)) fail(`${testCase.id}/${runner.id} engine changed across replay`)
      if (canonicalJson(candidate.evidence) !== canonicalJson(first.evidence)) fail(`${testCase.id}/${runner.id} evidence changed across replay`)
      if (candidate.semantic.sha256 !== first.semantic.sha256) fail(`${testCase.id}/${runner.id} semantic digest changed across replay`)
    }
  }
  return observations
}

function parseCLI(arguments_) {
  let manifestPath = DEFAULT_MANIFEST
  let observationsPath
  let request
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index]
    if (argument === '--manifest') manifestPath = resolve(arguments_[++index] ?? fail('--manifest requires a path'))
    else if (argument === '--observations') observationsPath = arguments_[++index] ?? fail('--observations requires a path or -')
    else if (argument === '--request') request = [arguments_[++index], arguments_[++index], Number(arguments_[++index])]
    else fail(`unknown argument ${argument}`)
  }
  return { manifestPath, observationsPath, request }
}

async function main() {
  const options = parseCLI(process.argv.slice(2))
  const manifest = loadManifest(options.manifestPath)
  const validated = validateManifest(manifest, { root: scriptRoot })
  if (options.request) {
    process.stdout.write(`${canonicalJson(createAdapterRequest(validated, ...options.request))}\n`)
    return
  }
  if (options.observationsPath) {
    const source = options.observationsPath === '-' ? await new Promise((resolveInput, reject) => {
      const chunks = []
      process.stdin.on('data', (chunk) => chunks.push(chunk))
      process.stdin.on('end', () => resolveInput(Buffer.concat(chunks).toString('utf8')))
      process.stdin.on('error', reject)
    }) : readFileSync(resolve(options.observationsPath), 'utf8')
    verifyObservationSet(validated, parseStrictJson(source, 'observations'))
    process.stdout.write(`verified ${manifest.cases.length} native Office production E2E cases and ${manifest.fixtures.length} exact fixtures\n`)
    return
  }
  process.stdout.write(`checked ${manifest.protocol}: ${manifest.cases.length} cases, ${manifest.fixtures.length} fixtures, offline\n`)
}

if (process.argv[1] && realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url))) {
  main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exitCode = 1
  })
}
