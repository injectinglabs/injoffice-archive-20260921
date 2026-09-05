#!/usr/bin/env node
// Node harness for the xlsxnative WASM spike. Instantiates the same
// extract/apply JS bindings used by the playground worker.

import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { runInThisContext } from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))

function usage() {
  return `xlsxnative WASM contract harness

Usage:
  node node_contract.mjs extract --wasm PATH --wasm-exec PATH --input XLSX [--out FILE]
  node node_contract.mjs apply --wasm PATH --wasm-exec PATH --original XLSX --payload JSON --expected-revision SHA [--out FILE]
  node node_contract.mjs survive --wasm PATH --wasm-exec PATH --input XLSX --payload JSON --expected-revision SHA --rev-token REV --stale-revision SHA [--out FILE]

extract/apply write native JSON or mutated XLSX bytes.
survive refuses empty/CAS-mismatch calls on one instance, then extracts.

Go WASM returns {ok:true,value}|{ok:false,error,fatal}; this harness throws an
Error(error) with a boolean fatal property. A missing fatal field from an older
module is treated as false.
`
}

function arg(flag, args, required = true) {
  const i = args.indexOf(flag)
  if (i === -1 || i + 1 >= args.length) {
    if (required) {
      console.error(`missing ${flag}`)
      process.exit(2)
    }
    return undefined
  }
  return args[i + 1]
}

function loadWasmExec(path) {
  const src = readFileSync(path, 'utf8')
  runInThisContext(src, { filename: path })
  if (typeof globalThis.Go !== 'function') {
    throw new Error(`wasm_exec.js at ${path} did not define globalThis.Go`)
  }
}

function waitReady(timeoutMs = 60_000) {
  return new Promise((resolve, reject) => {
    if (globalThis.xlsxnative) {
      resolve()
      return
    }
    const timer = setTimeout(() => reject(new Error('xlsxnative WASM ready timeout')), timeoutMs)
    globalThis.xlsxnativeOnReady = () => {
      clearTimeout(timer)
      resolve()
    }
  })
}

async function instantiate(wasmPath, wasmExecPath) {
  loadWasmExec(wasmExecPath)
  const go = new globalThis.Go()
  const ready = waitReady()
  const { instance } = await WebAssembly.instantiate(readFileSync(wasmPath), go.importObject)
  // go.run resolves only when the Go program exits; the spike parks in select{}.
  void go.run(instance)
  await ready
  if (!globalThis.xlsxnative?.extract || !globalThis.xlsxnative?.apply) {
    throw new Error('xlsxnative extract/apply bindings were not installed')
  }
  return { api: globalThis.xlsxnative, go }
}

function unwrap(result, expect) {
  if (result == null || typeof result !== 'object') {
    throw new Error('xlsxnative returned an empty result')
  }
  if (result.ok !== true) {
    const message = typeof result.error === 'string' && result.error !== '' ? result.error : 'xlsxnative failed'
    const error = new Error(message)
    error.fatal = result.fatal === true
    throw error
  }
  if (expect === 'json') {
    if (typeof result.value !== 'string' || result.value === 'undefined') {
      throw new Error('xlsxnative extract did not return JSON')
    }
    return result.value
  }
  if (expect === 'bytes') {
    if (!(result.value instanceof Uint8Array)) {
      throw new Error('xlsxnative apply did not return Uint8Array')
    }
    return result.value
  }
  return result.value
}

function extractJSON(api, bytes) {
  return unwrap(api.extract(bytes), 'json')
}

function applyBytes(api, original, payload, expectedRevision) {
  return unwrap(api.apply(original, payload, expectedRevision), 'bytes')
}

function caught(fn) {
  try {
    fn()
    return null
  } catch (err) {
    if (err instanceof Error) return err
    return new Error(String(err))
  }
}

function refuse(name, fn, want, expectedFatal = false) {
  const error = caught(fn)
  if (!error) throw new Error(`${name} did not fail`)
  const message = error.message
  if (message === 'undefined' || /Go program has already exited/i.test(message)) {
    throw new Error(`${name} killed the WASM instance: ${message}`)
  }
  if (want && !message.includes(want)) {
    throw new Error(`${name} error ${JSON.stringify(message)} does not contain ${JSON.stringify(want)}`)
  }
  if (error.fatal !== expectedFatal) {
    throw new Error(`${name} fatal=${String(error.fatal)}, want ${String(expectedFatal)}`)
  }
  return { name, error: message, fatal: error.fatal }
}

function writeOutput(data, outPath) {
  if (outPath) {
    writeFileSync(outPath, data)
    return
  }
  process.stdout.write(data)
}

function readPayload(payloadArg) {
  return payloadArg.trimStart().startsWith('{')
    ? Buffer.from(payloadArg, 'utf8')
    : readFileSync(resolve(payloadArg))
}

const args = process.argv.slice(2)
if (args.length === 0 || args[0] === '-h' || args[0] === '--help') {
  process.stderr.write(usage())
  process.exit(args.length === 0 ? 2 : 0)
}

const command = args[0]
const wasm = resolve(arg('--wasm', args) || resolve(here, 'dist/xlsxnative.wasm'))
const wasmExec = resolve(arg('--wasm-exec', args) || resolve(here, 'dist/wasm_exec.js'))
const out = arg('--out', args, false)

try {
  const { api, go } = await instantiate(wasm, wasmExec)
  if (command === 'extract') {
    const input = readFileSync(resolve(arg('--input', args)))
    const json = extractJSON(api, new Uint8Array(input))
    writeOutput(Buffer.from(json, 'utf8'), out)
    process.exit(0)
  }
  if (command === 'apply') {
    const original = readFileSync(resolve(arg('--original', args)))
    const payload = readPayload(arg('--payload', args))
    const expectedRevision = arg('--expected-revision', args)
    const produced = applyBytes(api, new Uint8Array(original), new Uint8Array(payload), expectedRevision)
    writeOutput(Buffer.from(produced), out)
    process.exit(0)
  }
  if (command === 'survive') {
    const original = new Uint8Array(readFileSync(resolve(arg('--input', args))))
    const payload = new Uint8Array(readPayload(arg('--payload', args)))
    const expectedRevision = arg('--expected-revision', args)
    const revToken = arg('--rev-token', args)
    const stale = arg('--stale-revision', args)
    const empty = new Uint8Array()
    const refusals = [
      refuse('empty_extract', () => extractJSON(api, empty)),
      refuse('empty_original', () => applyBytes(api, empty, payload, expectedRevision)),
      refuse('empty_payload', () => applyBytes(api, original, empty, expectedRevision)),
      refuse('missing_revision', () => applyBytes(api, original, payload, ''), 'expectedRevision is required'),
      refuse('rev_token', () => applyBytes(api, original, payload, revToken), 'sha256'),
      refuse('stale_cas', () => applyBytes(api, original, payload, stale), 'stale outer revision'),
    ]
    // Exercise the fatal-envelope consumer without adding a production-only
    // callback that deliberately panics inside the Go WASM instance.
    const fatalProbe = refuse(
      'synthetic_fatal',
      () => unwrap({ ok: false, error: 'synthetic recovered panic', fatal: true }),
      'recovered panic',
      true,
    )
    const legacyProbe = refuse(
      'legacy_without_fatal',
      () => unwrap({ ok: false, error: 'legacy refusal' }),
      'legacy refusal',
    )
    if (go.exited) throw new Error('Go program has already exited after refused extract/apply')
    const json = extractJSON(api, original)
    const report = JSON.stringify({ refusals, fatalProbe, legacyProbe, extract: json })
    writeOutput(Buffer.from(report, 'utf8'), out)
    process.exit(0)
  }
  console.error(`unknown command ${command}\n\n${usage()}`)
  process.exit(2)
} catch (err) {
  console.error(err instanceof Error ? err.message : String(err))
  process.exit(1)
}
