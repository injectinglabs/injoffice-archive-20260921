import {
  NativeWasmError,
  createNativeWasmClient,
  type NativeWasmAbortSignal,
  type NativeWasmClient,
  type NativeWasmWorker,
} from '@injoffice/native-runtime'
import {
  NativeWorkbookV2ValidationError,
  WORKBOOK_MUTATION_PROTOCOL,
  WORKBOOK_MUTATION_VERSION,
  WorkbookMutationValidationError,
  assertNativeWorkbookV2,
  decodeNativeWorkbookV2,
  decodeWorkbookMutationBatch,
  type NativeWorkbookV2,
  type SupportedWorkbookMutation,
  type WorkbookMutationBatchV1,
} from '@injoffice/sheets/browser'

export const XLSX_WASM_NATIVE_MAX_PACKAGE_BYTES = 128 * 1024 * 1024
export const XLSX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES = 3 * 1024 * 1024
export const DEFAULT_XLSX_WASM_MAX_PACKAGE_BYTES = 32 * 1024 * 1024
export const DEFAULT_XLSX_WASM_MAX_MUTATION_PAYLOAD_BYTES = 1024 * 1024

export type XlsxNativeCellMutationV1 = Extract<SupportedWorkbookMutation, { kind: `cell.${string}` }>
export type XlsxNativeStyleMutationV1 = Extract<SupportedWorkbookMutation, { kind: 'style.patch' }>
export type XlsxNativeLayoutMutationV1 = Extract<SupportedWorkbookMutation, { kind: 'row.set_height' | 'column.set_width' }>

/** Strict JSON shape consumed by Go's NativeWorkbookMutationTransactionV1. */
export interface XlsxNativeMutationTransactionV1 {
  /** Inner contract CAS: rev:<the exact source package SHA-256 digest>. */
  readonly expected_revision: string
  readonly cells?: ReadonlyArray<XlsxNativeCellMutationV1>
  readonly styles?: ReadonlyArray<XlsxNativeStyleMutationV1>
  readonly layout?: ReadonlyArray<XlsxNativeLayoutMutationV1>
}

export interface XlsxWasmAssetUrls {
  workerUrl: string
  wasmUrl: string
  goRuntimeUrl: string
}

export interface XlsxWasmClientOptions {
  workerUrl?: string | URL
  wasmUrl?: string | URL
  goRuntimeUrl?: string | URL
  operationTimeoutMs?: number
  /** Browser-side ceiling, capped by the native 128 MiB package limit. */
  maxPackageBytes?: number
  /** UTF-8 JSON ceiling, capped by the native 3 MiB mutation limit. */
  maxMutationPayloadBytes?: number
  workerFactory?: (workerUrl: string) => NativeWasmWorker
}

export interface XlsxWasmOperationOptions {
  signal?: NativeWasmAbortSignal
}

export interface XlsxWasmClient {
  extract(bytes: Uint8Array, options?: XlsxWasmOperationOptions): Promise<NativeWorkbookV2>
  apply(
    original: Uint8Array,
    workbook: NativeWorkbookV2,
    transaction: XlsxNativeMutationTransactionV1,
    options?: XlsxWasmOperationOptions,
  ): Promise<Uint8Array>
  terminate(): void
}

export function resolveXlsxWasmAssetUrls(options: Pick<XlsxWasmClientOptions, 'workerUrl' | 'wasmUrl' | 'goRuntimeUrl'> = {}): XlsxWasmAssetUrls {
  return {
    workerUrl: toUrl(options.workerUrl, new URL('./xlsxnative.worker.js', import.meta.url)),
    wasmUrl: toUrl(options.wasmUrl, new URL('./xlsxnative.wasm', import.meta.url)),
    goRuntimeUrl: toUrl(options.goRuntimeUrl, new URL('./wasm_exec.js', import.meta.url)),
  }
}

/** Validate a public Sheets batch and adapt the native engine's supported subset. */
export function adaptWorkbookMutationBatchV1(
  workbook: NativeWorkbookV2,
  input: WorkbookMutationBatchV1,
): XlsxNativeMutationTransactionV1 {
  assertNativeWorkbookV2(workbook)
  const decoded = decodeWorkbookMutationBatch(input)
  if (!decoded.ok) throw new WorkbookMutationValidationError(decoded.issues)
  const batch = decoded.value
  rejectUnpairedSurrogates(batch)
  if (batch.expected_revision !== workbook.source.package_sha256) {
    throw new NativeWasmError(
      'STALE_REVISION',
      `Mutation batch outer expected_revision must equal extracted source.package_sha256 ${JSON.stringify(workbook.source.package_sha256)}.`,
    )
  }
  const cells: XlsxNativeCellMutationV1[] = []
  const styles: XlsxNativeStyleMutationV1[] = []
  const layout: XlsxNativeLayoutMutationV1[] = []
  let lastFamily = 0
  for (const operation of batch.operations) {
    validateNativeOperation(operation)
    const family = operation.kind.startsWith('cell.') ? 0
      : operation.kind === 'style.patch' ? 1
        : operation.kind === 'row.set_height' || operation.kind === 'column.set_width' ? 2
          : -1
    if (family < 0) {
      throw new NativeWasmError(
        'UNSUPPORTED_OPERATION',
        `${operation.kind} is valid in the Sheets protocol but is not supported by the native XLSX transaction.`,
      )
    }
    if (family < lastFamily) {
      throw new NativeWasmError(
        'UNSUPPORTED_ORDER',
        'Native XLSX batches must keep cell operations before style operations and style operations before row/column layout operations.',
      )
    }
    lastFamily = family
    if (family === 0) cells.push(operation as XlsxNativeCellMutationV1)
    else if (family === 1) styles.push(operation as XlsxNativeStyleMutationV1)
    else layout.push(operation as XlsxNativeLayoutMutationV1)
  }
  return {
    expected_revision: workbook.revision,
    ...(cells.length > 0 ? { cells } : {}),
    ...(styles.length > 0 ? { styles } : {}),
    ...(layout.length > 0 ? { layout } : {}),
  }
}

export function createXlsxWasmClient(options: XlsxWasmClientOptions = {}): XlsxWasmClient {
  const assets = resolveXlsxWasmAssetUrls(options)
  const workerFactory = options.workerFactory ?? createBrowserWorker
  const maxPackageBytes = boundedLimit(options.maxPackageBytes, DEFAULT_XLSX_WASM_MAX_PACKAGE_BYTES, XLSX_WASM_NATIVE_MAX_PACKAGE_BYTES, 'maxPackageBytes')
  const maxMutationPayloadBytes = boundedLimit(options.maxMutationPayloadBytes, DEFAULT_XLSX_WASM_MAX_MUTATION_PAYLOAD_BYTES, XLSX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES, 'maxMutationPayloadBytes')
  const native = createNativeWasmClient({
    format: 'xlsx',
    workerFactory: () => workerFactory(assets.workerUrl),
    assets: { wasmUrl: assets.wasmUrl, goRuntimeUrl: assets.goRuntimeUrl },
    operationTimeoutMs: options.operationTimeoutMs,
  })
  return new XlsxWasmClientImpl(native, maxPackageBytes, maxMutationPayloadBytes)
}

class XlsxWasmClientImpl implements XlsxWasmClient {
  constructor(
    private readonly native: NativeWasmClient,
    private readonly maxPackageBytes: number,
    private readonly maxMutationPayloadBytes: number,
  ) {}

  extract(bytes: Uint8Array, options: XlsxWasmOperationOptions = {}): Promise<NativeWorkbookV2> {
    assertPackageSize(bytes, this.maxPackageBytes)
    return this.extractValidated(bytes, options)
  }

  apply(
    original: Uint8Array,
    workbook: NativeWorkbookV2,
    transaction: XlsxNativeMutationTransactionV1,
    options: XlsxWasmOperationOptions = {},
  ): Promise<Uint8Array> {
    assertPackageSize(original, this.maxPackageBytes)
    assertNativeWorkbookV2(workbook)
    const normalized = validateNativeTransaction(workbook, transaction)
    const payload = JSON.stringify(normalized)
    const payloadBytes = new TextEncoder().encode(payload).byteLength
    if (payloadBytes > this.maxMutationPayloadBytes) {
      throw new RangeError(`XLSX mutation payload is ${payloadBytes} UTF-8 bytes; maxMutationPayloadBytes is ${this.maxMutationPayloadBytes}.`)
    }
    return this.native.apply(original, payload, workbook.source.package_sha256, options)
  }

  terminate(): void {
    this.native.terminate()
  }

  private async extractValidated(bytes: Uint8Array, options: XlsxWasmOperationOptions): Promise<NativeWorkbookV2> {
    const contractJson = await this.native.extract(bytes, options)
    const decoded = decodeNativeWorkbookV2(contractJson)
    if (!decoded.ok) {
      this.native.terminate()
      throw new NativeWorkbookV2ValidationError(decoded.issues)
    }
    return decoded.value
  }
}

function validateNativeTransaction(workbook: NativeWorkbookV2, input: XlsxNativeMutationTransactionV1): XlsxNativeMutationTransactionV1 {
  if (!isRecord(input)) throw new TypeError('XLSX native transaction must be an object.')
  rejectUnknownKeys(input, ['expected_revision', 'cells', 'styles', 'layout'])
  if (input.expected_revision !== workbook.revision) {
    throw new NativeWasmError(
      'STALE_REVISION',
      `Native transaction expected_revision must equal extracted workbook.revision ${JSON.stringify(workbook.revision)}.`,
    )
  }
  const cells = optionalArray(input.cells, 'cells')
  const styles = optionalArray(input.styles, 'styles')
  const layout = optionalArray(input.layout, 'layout')
  const operations = [...cells, ...styles, ...layout]
  const decoded = decodeWorkbookMutationBatch({
    protocol: WORKBOOK_MUTATION_PROTOCOL,
    version: WORKBOOK_MUTATION_VERSION,
    batch_id: 'xlsx-wasm-native',
    expected_revision: workbook.source.package_sha256,
    operations,
  })
  if (!decoded.ok) throw new WorkbookMutationValidationError(decoded.issues)
  rejectUnpairedSurrogates(decoded.value)
  const normalized = decoded.value.operations
  for (const operation of normalized) validateNativeOperation(operation)
  const normalizedCells = normalized.slice(0, cells.length)
  const normalizedStyles = normalized.slice(cells.length, cells.length + styles.length)
  const normalizedLayout = normalized.slice(cells.length + styles.length)
  if (normalizedCells.some((item) => !item.kind.startsWith('cell.'))) throw new TypeError('XLSX native cells may contain only cell operations.')
  if (normalizedStyles.some((item) => item.kind !== 'style.patch')) throw new TypeError('XLSX native styles may contain only style.patch operations.')
  if (normalizedLayout.some((item) => item.kind !== 'row.set_height' && item.kind !== 'column.set_width')) {
    throw new TypeError('XLSX native layout may contain only row.set_height or column.set_width operations.')
  }
  return {
    expected_revision: workbook.revision,
    ...(normalizedCells.length > 0 ? { cells: normalizedCells as XlsxNativeCellMutationV1[] } : {}),
    ...(normalizedStyles.length > 0 ? { styles: normalizedStyles as XlsxNativeStyleMutationV1[] } : {}),
    ...(normalizedLayout.length > 0 ? { layout: normalizedLayout as XlsxNativeLayoutMutationV1[] } : {}),
  }
}

function createBrowserWorker(workerUrl: string): NativeWasmWorker {
  if (typeof globalThis.Worker !== 'function') {
    throw new NativeWasmError('WORKER_UNAVAILABLE', 'XLSX WASM requires a browser Worker; provide workerFactory in non-browser runtimes.', true)
  }
  return new globalThis.Worker(workerUrl) as unknown as NativeWasmWorker
}

function assertPackageSize(bytes: Uint8Array, limit: number): void {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('XLSX package must be a Uint8Array.')
  if (bytes.byteLength === 0) throw new RangeError('XLSX package must not be empty.')
  if (bytes.byteLength > limit) throw new RangeError(`XLSX package is ${bytes.byteLength} bytes; maxPackageBytes is ${limit}.`)
}

function boundedLimit(value: number | undefined, fallback: number, hardMaximum: number, name: string): number {
  const resolved = value ?? fallback
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > hardMaximum) {
    throw new RangeError(`${name} must be a positive safe integer no greater than ${hardMaximum}.`)
  }
  return resolved
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function rejectUnknownKeys(value: Record<string, unknown>, allowed: readonly string[]): void {
  const allowedKeys = new Set(allowed)
  const unknown = Object.keys(value).find((key) => !allowedKeys.has(key))
  if (unknown !== undefined) throw new TypeError(`XLSX native transaction contains unknown field ${JSON.stringify(unknown)}.`)
}

function optionalArray(value: unknown, name: string): unknown[] {
  if (value === undefined) return []
  if (!Array.isArray(value)) throw new TypeError(`XLSX native transaction ${name} must be an array.`)
  return value
}

function validateNativeOperation(operation: SupportedWorkbookMutation): void {
  if (operation.kind === 'cell.set_formula' && operation.formula.length === 1) {
    throw new NativeWasmError('INVALID_MUTATION', 'Native XLSX formulas must contain content after the leading equals sign.')
  }
}

function rejectUnpairedSurrogates(value: unknown): void {
  if (typeof value === 'string') {
    for (let index = 0; index < value.length; index++) {
      const code = value.charCodeAt(index)
      if (code >= 0xd800 && code <= 0xdbff) {
        const next = value.charCodeAt(index + 1)
        if (!(next >= 0xdc00 && next <= 0xdfff)) throw new NativeWasmError('INVALID_UNICODE', 'XLSX mutations must not contain unpaired UTF-16 surrogates.')
        index++
      } else if (code >= 0xdc00 && code <= 0xdfff) {
        throw new NativeWasmError('INVALID_UNICODE', 'XLSX mutations must not contain unpaired UTF-16 surrogates.')
      }
    }
    return
  }
  if (Array.isArray(value)) {
    for (const item of value) rejectUnpairedSurrogates(item)
    return
  }
  if (isRecord(value)) {
    for (const item of Object.values(value)) rejectUnpairedSurrogates(item)
  }
}

function toUrl(value: string | URL | undefined, fallback: URL): string {
  if (value === undefined) return fallback.href
  const resolved = typeof value === 'string' ? value : value.href
  if (resolved.length === 0) throw new TypeError('XLSX WASM asset URLs must not be empty')
  return resolved
}
