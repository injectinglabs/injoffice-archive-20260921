#!/usr/bin/env node
import assert from 'node:assert/strict'
import { readFileSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { runInThisContext } from 'node:vm'

const args = process.argv.slice(2)
function arg(name) { const i = args.indexOf(name); if (i < 0 || !args[i + 1]) throw Error(`Missing ${name}`); return args[i + 1] }
runInThisContext(readFileSync(arg('--runtime'), 'utf8'), { filename: 'wasm_exec.js' })
const go = new globalThis.Go()
const ready = new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('Rich source WASM init timeout')), 60000); globalThis.xlsxnativeOnReady = () => { clearTimeout(timer); resolve() } })
const { instance } = await WebAssembly.instantiate(readFileSync(arg('--wasm')), go.importObject)
void go.run(instance)
await ready
const api = globalThis.xlsxnative
assert.deepEqual(Object.keys(api), ['previewRichSource'])
const refused = [[], [null], [true], ['bytes'], [new ArrayBuffer(1)], [new Uint16Array(1)], [new Uint8Array()], [new Uint8Array([1])], [new Uint8Array([1]), 2]]
const oversized = new Uint8Array(1)
Object.defineProperty(oversized, 'byteLength', { value: 128 * 1024 * 1024 + 1 })
refused.push([oversized])
for (const input of refused) { const r = api.previewRichSource(...input); assert.equal(r.ok, false); assert.equal(r.fatal, false); assert.ok(r.error) }
const bytes = new Uint8Array(readFileSync(arg('--input'))), before = new Uint8Array(bytes)
const result = api.previewRichSource(bytes)
assert.equal(result.ok, true, result.error)
assert.equal(typeof result.value, 'string')
const value = JSON.parse(result.value)
assert.deepEqual(value, JSON.parse(readFileSync(arg('--expected'), 'utf8')))
assert.equal(value.package_sha256, 'sha256:' + createHash('sha256').update(bytes).digest('hex'))
assert.deepEqual(bytes, before)
const trap = new Uint8Array(1)
Object.defineProperty(trap, 'byteLength', { value: 'invalid numeric property' })
const fatal = api.previewRichSource(trap)
assert.equal(fatal.ok, false); assert.equal(fatal.fatal, true); assert.ok(fatal.error)
if (args.includes('--output')) writeFileSync(arg('--output'), result.value)
console.log(JSON.stringify({ status: 'passed', refusals: refused.length, onlyReadOnlyBinding: true, sourceHash: value.package_sha256, cells: value.sheet.cells.length, richCells: value.rich_cells.length }))
process.exit(0)
