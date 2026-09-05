/*
 * Copyright 2026 Injecting Inc.
 * SPDX-License-Identifier: Apache-2.0
 *
 * Classic worker for the matching Go wasm_exec.js runtime supplied with this
 * package. The worker and both runtime assets are one versioned unit.
 */
/* global Go, importScripts */

const PROTOCOL = 'injoffice.native-wasm-worker'
const VERSION = 1
const FORMAT = 'docx'
const MAX_ERROR_LENGTH = 4096

let boot
let runtimeFailure
let initialized = false

class NativeBindingError extends Error {
  constructor(message, fatal) {
    super(message)
    this.name = 'NativeBindingError'
    this.fatal = fatal
  }
}

function respond(request, response, transfer = []) {
  postMessage({
    protocol: PROTOCOL,
    version: VERSION,
    id: request.id,
    format: FORMAT,
    op: request.op,
    ...response,
  }, transfer)
}

function errorMessage(reason, fallback) {
  const message = reason instanceof Error ? reason.message : typeof reason === 'string' ? reason : fallback
  return (message || fallback).slice(0, MAX_ERROR_LENGTH)
}

function refuse(request, code, reason, fatal) {
  respond(request, { ok: false, error: { code, message: errorMessage(reason, 'DOCX WASM failed.'), fatal } })
}

function notifyFatal(reason) {
  postMessage({
    protocol: PROTOCOL,
    version: VERSION,
    format: FORMAT,
    op: 'fatal',
    error: { code: 'ENGINE_EXITED', message: errorMessage(reason, 'DOCX Go runtime exited.'), fatal: true },
  })
}

function unwrap(result, expected) {
  if (result == null || typeof result !== 'object' || typeof result.ok !== 'boolean') {
    throw new NativeBindingError('docxnative returned a malformed result envelope', true)
  }
  if (result.ok !== true) {
    if (typeof result.error !== 'string' || result.error === '') throw new NativeBindingError('docxnative returned a malformed error envelope', true)
    if (result.fatal !== undefined && typeof result.fatal !== 'boolean') throw new NativeBindingError('docxnative returned a malformed fatal marker', true)
    throw new NativeBindingError(result.error, result.fatal === true)
  }
  if (expected === 'json' && typeof result.value !== 'string') throw new NativeBindingError('docxnative extract did not return JSON', true)
  if (expected === 'bytes' && !(result.value instanceof Uint8Array)) throw new NativeBindingError('docxnative apply did not return Uint8Array', true)
  return result.value
}

async function initialize(assets) {
  if (assets == null || typeof assets !== 'object' || typeof assets.wasmUrl !== 'string' || typeof assets.goRuntimeUrl !== 'string') {
    throw new Error('init requires wasmUrl and goRuntimeUrl')
  }
  importScripts(assets.goRuntimeUrl)
  if (typeof Go !== 'function') throw new Error('Go runtime did not define Go')
  const go = new Go()
  const ready = new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('docxnative WASM init timeout')), 60_000)
    self.docxnativeOnReady = () => {
      clearTimeout(timer)
      resolve()
    }
  })
  const response = await fetch(assets.wasmUrl)
  if (!response.ok) throw new Error(`could not fetch docxnative.wasm (${response.status})`)
  const { instance } = await WebAssembly.instantiate(await response.arrayBuffer(), go.importObject)
  let rejectRuntimeExit
  const runtimeExit = new Promise((_resolve, reject) => { rejectRuntimeExit = reject })
  const stopRuntime = (reason) => {
    if (runtimeFailure) return
    runtimeFailure = reason instanceof Error ? reason : new Error(errorMessage(reason, 'DOCX Go runtime exited.'))
    rejectRuntimeExit(runtimeFailure)
    if (initialized) {
      notifyFatal(runtimeFailure)
      initialized = false
      self.close()
    }
  }
  Promise.resolve()
    .then(() => go.run(instance))
    .then(
      () => stopRuntime(new Error('DOCX Go runtime exited unexpectedly.')),
      (reason) => stopRuntime(reason),
    )
  await Promise.race([ready, runtimeExit])
  if (!self.docxnative || typeof self.docxnative.extract !== 'function' || typeof self.docxnative.apply !== 'function') {
    throw new Error('docxnative extract/apply bindings were not installed')
  }
  if (runtimeFailure) throw runtimeFailure
  initialized = true
}

onmessage = async (event) => {
  const request = event.data
  if (request == null || typeof request !== 'object' || typeof request.id !== 'string') return
  if (request.protocol !== PROTOCOL || request.version !== VERSION || request.format !== FORMAT) {
    refuse(request, 'PROTOCOL_MISMATCH', 'Native worker protocol, version, or format does not match.', true)
    return
  }
  if (request.op === 'init') {
    if (!boot) boot = initialize(request.assets)
    try {
      await boot
      respond(request, { ok: true })
    } catch (reason) {
      refuse(request, 'INIT_FAILED', reason, true)
    }
    return
  }
  if (!boot) {
    refuse(request, 'NOT_INITIALIZED', 'DOCX WASM worker is not initialized.', true)
    return
  }
  try {
    await boot
    if (runtimeFailure) throw runtimeFailure
    if (request.op === 'extract') {
      const contractJson = unwrap(self.docxnative.extract(new Uint8Array(request.bytes)), 'json')
      respond(request, { ok: true, result: { contractJson } })
      return
    }
    if (request.op === 'apply') {
      const payload = typeof request.payload === 'string' ? request.payload : new Uint8Array(request.payload)
      const produced = unwrap(self.docxnative.apply(new Uint8Array(request.original), payload, request.expectedRevision), 'bytes')
      const bytes = produced.buffer.slice(produced.byteOffset, produced.byteOffset + produced.byteLength)
      respond(request, { ok: true, result: { bytes } }, [bytes])
      return
    }
    refuse(request, 'UNKNOWN_OPERATION', `Unknown DOCX WASM operation ${String(request.op)}.`, false)
  } catch (reason) {
    const fatal = reason instanceof NativeBindingError ? reason.fatal : true
    refuse(request, fatal ? 'NATIVE_FAILED' : 'NATIVE_REFUSED', reason, fatal)
    if (fatal) {
      initialized = false
      self.close()
    }
  }
}
