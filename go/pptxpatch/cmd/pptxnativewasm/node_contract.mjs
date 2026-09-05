#!/usr/bin/env node
// Node contract harness for the pptxnative Go WASM binding.

import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { runInThisContext } from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))

function usage() {
  return `pptxnative WASM contract harness

Usage:
  node node_contract.mjs extract --wasm PATH --wasm-exec PATH --input PPTX [--out FILE]
  node node_contract.mjs apply --wasm PATH --wasm-exec PATH --original PPTX --payload JSON --expected-revision SHA [--out FILE]
  node node_contract.mjs survive --wasm PATH --wasm-exec PATH --input PPTX --payload JSON --expected-revision SHA --rev-token REV --stale-revision SHA [--out FILE]

extract/apply write native v1 JSON or mutated PPTX bytes. survive exercises
refusals on one instance and then proves that extraction still works.
`
}

function arg(flag, args, required = true) {
  const index = args.indexOf(flag)
  if (index === -1 || index + 1 >= args.length) {
    if (required) throw new Error(`missing ${flag}`)
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
    if (globalThis.pptxnative) {
      resolveReady()
      return
    }
    const timer = setTimeout(() => reject(new Error('pptxnative WASM ready timeout')), timeoutMs)
    globalThis.pptxnativeOnReady = () => {
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
  // The Go command intentionally parks forever after installing its callbacks.
  void go.run(instance)
  await ready
  if (!globalThis.pptxnative?.extract || !globalThis.pptxnative?.apply) {
    throw new Error('pptxnative extract/apply bindings were not installed')
  }
  return { api: globalThis.pptxnative, go }
}

function unwrap(result, expected) {
  if (result == null || typeof result !== 'object') {
    throw new Error('pptxnative returned an empty result')
  }
  if (result.ok !== true) {
    const message = typeof result.error === 'string' && result.error !== '' ? result.error : 'pptxnative failed'
    throw new Error(message)
  }
  if (expected === 'json') {
    if (typeof result.value !== 'string' || result.value === 'undefined') {
      throw new Error('pptxnative extract did not return JSON')
    }
    return result.value
  }
  if (expected === 'bytes') {
    if (!(result.value instanceof Uint8Array)) {
      throw new Error('pptxnative apply did not return Uint8Array')
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

function refusal(name, call, expectedSubstring) {
  let message = ''
  try {
    call()
  } catch (error) {
    message = error instanceof Error ? error.message : String(error)
  }
  if (!message) throw new Error(`${name} did not fail`)
  if (message === 'undefined' || /Go program has already exited/i.test(message)) {
    throw new Error(`${name} killed the WASM instance: ${message}`)
  }
  if (expectedSubstring && !message.includes(expectedSubstring)) {
    throw new Error(`${name} error ${JSON.stringify(message)} does not contain ${JSON.stringify(expectedSubstring)}`)
  }
  return { name, error: message }
}

function writeOutput(data, outputPath) {
  if (outputPath) {
    writeFileSync(outputPath, data)
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

try {
  const command = args[0]
  const wasm = resolve(arg('--wasm', args, false) || resolve(here, 'dist/pptxnative.wasm'))
  const wasmExec = resolve(arg('--wasm-exec', args, false) || resolve(here, 'dist/wasm_exec.js'))
  const output = arg('--out', args, false)
  const { api, go } = await instantiate(wasm, wasmExec)

  if (command === 'extract') {
    const input = new Uint8Array(readFileSync(resolve(arg('--input', args))))
    writeOutput(Buffer.from(extractJSON(api, input), 'utf8'), output)
    process.exit(0)
  }
  if (command === 'apply') {
    const original = new Uint8Array(readFileSync(resolve(arg('--original', args))))
    const payload = new Uint8Array(readPayload(arg('--payload', args)))
    const revision = arg('--expected-revision', args)
    writeOutput(Buffer.from(applyBytes(api, original, payload, revision)), output)
    process.exit(0)
  }
  if (command === 'survive') {
    const original = new Uint8Array(readFileSync(resolve(arg('--input', args))))
    const payload = new Uint8Array(readPayload(arg('--payload', args)))
    const revision = arg('--expected-revision', args)
    const revToken = arg('--rev-token', args)
    const stale = arg('--stale-revision', args)
    const empty = new Uint8Array()
    const refusals = [
      refusal('empty_extract', () => extractJSON(api, empty)),
      refusal('empty_original', () => applyBytes(api, empty, payload, revision)),
      refusal('empty_payload', () => applyBytes(api, original, empty, revision)),
      refusal('missing_revision', () => applyBytes(api, original, payload, ''), 'expectedRevision is required'),
      refusal('rev_token', () => applyBytes(api, original, payload, revToken), 'stale outer source revision'),
      refusal('stale_cas', () => applyBytes(api, original, payload, stale), 'stale outer source revision'),
    ]
    if (go.exited) throw new Error('Go program has already exited after refused extract/apply')
    writeOutput(Buffer.from(JSON.stringify({ refusals, extract: extractJSON(api, original) }), 'utf8'), output)
    process.exit(0)
  }

  throw new Error(`unknown command ${command}\n\n${usage()}`)
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
}
