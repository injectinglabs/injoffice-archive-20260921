#!/usr/bin/env node

import { mkdirSync, realpathSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  ADAPTER_PROTOCOL,
  OBSERVATION_PROTOCOL,
  canonicalJson,
  createAdapterRequest,
  loadManifest,
  validateManifest,
} from './verify-native-office-production-e2e.mjs'

const scriptRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
export const HOST_OBSERVATION_DIR = 'testdata/native-office-production-e2e/v1/host-observations'
export const DEFAULT_REQUEST_DIR = `${HOST_OBSERVATION_DIR}/adapter-requests`
export const HOST_OBSERVATION_CHECKIN = `${HOST_OBSERVATION_DIR}/observations.json`

function fail(message) {
  throw new Error(message)
}

export function requestFileName(caseId, runnerId, replayIndex) {
  return `${caseId}.${runnerId}.${replayIndex}.json`
}

export function listAdapterRequestTargets(validated) {
  const targets = []
  const refusals = []
  for (const testCase of validated.cases.values()) {
    for (const runner of testCase.runners) {
      for (let replayIndex = 0; replayIndex < testCase.replay.runs; replayIndex += 1) {
        const record = { caseId: testCase.id, runnerId: runner.id, replayIndex }
        if (testCase.input.mode === 'tamper-xor') refusals.push(record)
        else targets.push(record)
      }
    }
  }
  return { targets, refusals }
}

export function writeAdapterRequests(validated, directory = resolve(scriptRoot, DEFAULT_REQUEST_DIR)) {
  mkdirSync(directory, { recursive: true })
  const { targets, refusals } = listAdapterRequestTargets(validated)
  const written = []
  for (const target of targets) {
    const request = createAdapterRequest(validated, target.caseId, target.runnerId, target.replayIndex)
    const fileName = requestFileName(target.caseId, target.runnerId, target.replayIndex)
    writeFileSync(resolve(directory, fileName), `${canonicalJson(request)}\n`)
    written.push({ ...target, fileName, protocol: ADAPTER_PROTOCOL })
  }
  const refusalNote = {
    observationProtocol: OBSERVATION_PROTOCOL,
    observationCheckIn: HOST_OBSERVATION_CHECKIN,
    kind: 'verifier-owned-refusal',
    code: 'FIXTURE_DIGEST_MISMATCH',
    note: 'Request creation must fail before an adapter is invoked. The host copies these refusal rows into the checked-in observation array; it does not invent a passing adapter result.',
    refusals,
  }
  writeFileSync(resolve(directory, 'tamper-refusals.json'), `${canonicalJson(refusalNote)}\n`)
  return { directory, written, refusals }
}

function parseCLI(arguments_) {
  let writeRequests = false
  let requestDir = resolve(scriptRoot, DEFAULT_REQUEST_DIR)
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index]
    if (argument === '--write-requests') {
      writeRequests = true
      const next = arguments_[index + 1]
      if (next && !next.startsWith('--')) requestDir = resolve(next), index += 1
    } else fail(`unknown argument ${argument}`)
  }
  return { writeRequests: writeRequests || arguments_.length === 0, requestDir }
}

function main() {
  const options = parseCLI(process.argv.slice(2))
  const validated = validateManifest(loadManifest(), { root: scriptRoot })
  if (!options.writeRequests) return
  const result = writeAdapterRequests(validated, options.requestDir)
  process.stdout.write(`wrote ${result.written.length} ${ADAPTER_PROTOCOL} requests to ${result.directory}\n`)
  process.stdout.write(`host observation check-in: ${HOST_OBSERVATION_CHECKIN} (${OBSERVATION_PROTOCOL})\n`)
  process.stdout.write(`tamper refusals: ${result.refusals.length} verifier-owned FIXTURE_DIGEST_MISMATCH rows; live host adapters must fill the remaining observations\n`)
}

if (process.argv[1] && realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url))) {
  try {
    main()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exitCode = 1
  }
}
