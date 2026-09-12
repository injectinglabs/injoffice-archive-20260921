#!/usr/bin/env node
// Node harness for the docxnative WASM binding. It instantiates the same
// extract/apply JavaScript contract intended for a browser Web Worker.

import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { runInThisContext } from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))

function usage() {
  return `docxnative WASM contract harness

Usage:
  node node_contract.mjs extract --wasm PATH --wasm-exec PATH --input DOCX [--out FILE]
  node node_contract.mjs apply --wasm PATH --wasm-exec PATH --original DOCX --payload JSON --expected-revision SHA [--out FILE]
  node node_contract.mjs survive --wasm PATH --wasm-exec PATH --input DOCX --payload JSON --expected-revision SHA --rev-token REV --stale-revision SHA [--out FILE]

extract/apply write native JSON or mutated DOCX bytes.
survive refuses malformed/CAS-mismatch calls on one instance, then extracts.

Go WASM returns {ok:true,value}|{ok:false,error,fatal}; this harness throws an
Error(error) with a boolean fatal property. Missing fatal is treated as false.
`
}

function arg(flag, args, required = true) {
  const index = args.indexOf(flag)
  if (index === -1 || index + 1 >= args.length) {
    if (required) {
      console.error(`missing ${flag}`)
      process.exit(2)
    }
    return undefined
  }
  return args[index + 1]
}

function loadWasmExec(path) {
  runInThisContext(readFileSync(path, 'utf8'), { filename: path })
  if (typeof globalThis.Go !== 'function') {
    throw new Error(`wasm_exec.js at ${path} did not define globalThis.Go`)
  }
}

function waitReady(timeoutMs = 60_000) {
  return new Promise((resolveReady, reject) => {
    if (globalThis.docxnative) {
      resolveReady()
      return
    }
    const timer = setTimeout(() => reject(new Error('docxnative WASM ready timeout')), timeoutMs)
    globalThis.docxnativeOnReady = () => {
      clearTimeout(timer)
      resolveReady()
    }
  })
}

async function instantiate(wasmPath, wasmExecPath) {
  loadWasmExec(wasmExecPath)
  const go = new globalThis.Go()
  const ready = waitReady()
  const { instance } = await WebAssembly.instantiate(readFileSync(wasmPath), go.importObject)
  // go.run resolves only when the Go program exits; the binding parks forever.
  void go.run(instance)
  await ready
  if (!globalThis.docxnative?.extract || !globalThis.docxnative?.apply) {
    throw new Error('docxnative extract/apply bindings were not installed')
  }
  return { api: globalThis.docxnative, go }
}

function unwrap(result, expectedType) {
  if (result == null || typeof result !== 'object') {
    throw new Error('docxnative returned an empty result')
  }
  if (result.ok !== true) {
    const message = typeof result.error === 'string' && result.error !== '' ? result.error : 'docxnative failed'
    const error = new Error(message)
    error.fatal = result.fatal === true
    throw error
  }
  if (expectedType === 'json') {
    if (typeof result.value !== 'string' || result.value === 'undefined') {
      throw new Error('docxnative extract did not return JSON')
    }
    return result.value
  }
  if (expectedType === 'bytes') {
    if (!(result.value instanceof Uint8Array)) {
      throw new Error('docxnative apply did not return Uint8Array')
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
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error))
  }
}

function refuse(name, fn, wanted, expectedFatal = false) {
  const error = caught(fn)
  if (!error) throw new Error(`${name} did not fail`)
  const message = error.message
  if (message === 'undefined' || /Go program has already exited/i.test(message)) {
    throw new Error(`${name} killed the WASM instance: ${message}`)
  }
  if (wanted && !message.includes(wanted)) {
    throw new Error(`${name} error ${JSON.stringify(message)} does not contain ${JSON.stringify(wanted)}`)
  }
  if (error.fatal !== expectedFatal) {
    throw new Error(`${name} fatal=${String(error.fatal)}, want ${String(expectedFatal)}`)
  }
  return { name, error: message, fatal: error.fatal }
}

function writeOutput(data, path) {
  if (path) {
    writeFileSync(path, data)
    return
  }
  process.stdout.write(data)
}

function readPayload(value) {
  return value.trimStart().startsWith('{')
    ? Buffer.from(value, 'utf8')
    : readFileSync(resolve(value))
}

const args = process.argv.slice(2)
if (args.length === 0 || args[0] === '-h' || args[0] === '--help') {
  process.stderr.write(usage())
  process.exit(args.length === 0 ? 2 : 0)
}

const command = args[0]
const wasm = resolve(arg('--wasm', args) || resolve(here, 'dist/docxnative.wasm'))
const wasmExec = resolve(arg('--wasm-exec', args) || resolve(here, 'dist/wasm_exec.js'))
const out = arg('--out', args, false)

try {
  const { api, go } = await instantiate(wasm, wasmExec)
  if (command === 'extract') {
    const input = readFileSync(resolve(arg('--input', args)))
    writeOutput(Buffer.from(extractJSON(api, new Uint8Array(input)), 'utf8'), out)
    process.exit(0)
  }
  if(command==='inspect'){
    const input=readFileSync(resolve(arg('--input',args)))
    writeOutput(Buffer.from(unwrap(api.inspect(new Uint8Array(input)),'json'),'utf8'),out)
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
    const revisionToken = arg('--rev-token', args)
    const staleRevision = arg('--stale-revision', args)
    const empty = new Uint8Array()
    const refusals = [
      refuse('wrong_extract_arity', () => unwrap(api.extract(original, original), 'json'), 'requires 1 argument'),
      refuse('wrong_apply_arity', () => unwrap(api.apply(original, payload), 'bytes'), 'requires 3 arguments'),
      refuse('empty_extract', () => extractJSON(api, empty)),
      refuse('empty_original', () => applyBytes(api, empty, payload, expectedRevision)),
      refuse('empty_payload', () => applyBytes(api, original, empty, expectedRevision)),
      refuse('missing_revision', () => applyBytes(api, original, payload, ''), 'expectedRevision is required'),
      refuse('rev_token', () => applyBytes(api, original, payload, revisionToken), 'sha256'),
      refuse('stale_cas', () => applyBytes(api, original, payload, staleRevision), 'outer expected revision'),
    ]
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
    // Exercise ArrayBuffer input after the refused Uint8Array calls. `slice`
    // avoids exposing unrelated bytes when a Node Buffer has an offset.
    const arrayBuffer = original.buffer.slice(original.byteOffset, original.byteOffset + original.byteLength)
    const extract = extractJSON(api, arrayBuffer)
    writeOutput(Buffer.from(JSON.stringify({ refusals, fatalProbe, legacyProbe, extract }), 'utf8'), out)
    process.exit(0)
  }
  console.error(`unknown command ${command}\n\n${usage()}`)
  process.exit(2)
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
}
