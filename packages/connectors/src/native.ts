import { isCredentialFreeConnectorSpec } from './commands'
import { isConnectorManagerSnapshot, type ConnectorManagerSnapshotV1 } from './manager'
import type { ConnectorSpec } from './types'

export const CONNECTOR_NATIVE_EXTENSION_VERSION = 1 as const
export const CONNECTOR_NATIVE_EXTENSION_NAMESPACE = 'https://schemas.injoffice.dev/xlsx/connectors/2026' as const

/** JSON shape shared with xlsxpatch. Runtime status, cells, and secrets are not part of it. */
export type NativeConnectorDefinitionV1 = ConnectorSpec

export interface ConnectorNativePersistenceCodec {
  readConnectorDefinitions(workbook: Uint8Array, options?: { readonly signal?: AbortSignal }): Promise<unknown>
  setConnectorDefinitions(
    workbook: Uint8Array,
    definitions: readonly NativeConnectorDefinitionV1[],
    options?: { readonly signal?: AbortSignal },
  ): Promise<Uint8Array>
}

export interface ConnectorNativeManager {
  serialize(): ConnectorSpec[]
  restoreState(snapshot: ConnectorManagerSnapshotV1): boolean
}

export type ConnectorNativePersistenceErrorCode = 'invalid-model' | 'invalid-response' | 'restore-failed' | 'canceled'

export class ConnectorNativePersistenceError extends Error {
  override readonly name = 'ConnectorNativePersistenceError'

  constructor(readonly code: ConnectorNativePersistenceErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
  }
}

/**
 * Validates and copies the credential-free wire subset used by xlsxpatch's
 * custom OPC extension. The Go codec remains the package-byte authority.
 */
export function toNativeConnectorDefinitions(specs: readonly ConnectorSpec[]): NativeConnectorDefinitionV1[] {
  return validateDefinitions(specs, 'invalid-model')
}

export function fromNativeConnectorDefinitions(value: unknown): ConnectorSpec[] {
  return validateDefinitions(value, 'invalid-response')
}

export class ConnectorNativePersistence {
  constructor(private readonly codec: ConnectorNativePersistenceCodec) {}

  async read(workbook: Uint8Array, options: { readonly signal?: AbortSignal } = {}): Promise<ConnectorSpec[]> {
    requireBytes(workbook)
    throwIfAborted(options.signal)
    const value = await this.codec.readConnectorDefinitions(workbook.slice(), options)
    throwIfAborted(options.signal)
    return fromNativeConnectorDefinitions(value)
  }

  async write(
    workbook: Uint8Array,
    specs: readonly ConnectorSpec[],
    options: { readonly signal?: AbortSignal } = {},
  ): Promise<Uint8Array> {
    requireBytes(workbook)
    const definitions = toNativeConnectorDefinitions(specs)
    throwIfAborted(options.signal)
    const output = await this.codec.setConnectorDefinitions(workbook.slice(), definitions, options)
    throwIfAborted(options.signal)
    if (!(output instanceof Uint8Array) || output.byteLength === 0) {
      throw new ConnectorNativePersistenceError('invalid-response', 'Native connector persistence returned no XLSX bytes.')
    }
    return output.slice()
  }
}

/** Hydrates definitions without triggering on-open fetches; the host decides
 * when authorization and network refresh are appropriate after file load. */
export async function hydrateConnectorManagerFromNative(
  manager: ConnectorNativeManager,
  persistence: ConnectorNativePersistence,
  workbook: Uint8Array,
  options: { readonly signal?: AbortSignal } = {},
): Promise<ConnectorSpec[]> {
  const specs = await persistence.read(workbook, options)
  const snapshot = snapshotOf(specs)
  if (!manager.restoreState(snapshot)) {
    throw new ConnectorNativePersistenceError('restore-failed', 'Connector manager refused the validated native definitions.')
  }
  return specs
}

export function persistConnectorManagerToNative(
  manager: ConnectorNativeManager,
  persistence: ConnectorNativePersistence,
  workbook: Uint8Array,
  options: { readonly signal?: AbortSignal } = {},
): Promise<Uint8Array> {
  return persistence.write(workbook, manager.serialize(), options)
}

function validateDefinitions(value: unknown, code: 'invalid-model' | 'invalid-response'): ConnectorSpec[] {
  if (!Array.isArray(value) || value.length > 1024) {
    throw new ConnectorNativePersistenceError(code, 'Native connector definitions must be an array of at most 1024 items.')
  }
  let specs: ConnectorSpec[]
  try {
    specs = structuredClone(value) as ConnectorSpec[]
  } catch (cause) {
    throw new ConnectorNativePersistenceError(code, 'Native connector definitions must be structured-cloneable plain data.', { cause })
  }
  if (specs.some((spec) => !isCredentialFreeConnectorSpec(spec) || !nativeBounds(spec))) {
    throw new ConnectorNativePersistenceError(code, 'Native connector definitions must use the credential-free, bounded v1 subset.')
  }
  if (!isConnectorManagerSnapshot(snapshotOf(specs))) {
    throw new ConnectorNativePersistenceError(code, 'Native connector definitions do not form a valid unique connector snapshot.')
  }
  return specs
}

function snapshotOf(specs: readonly ConnectorSpec[]): ConnectorManagerSnapshotV1 {
  return {
    version: 1,
    connectors: specs.map((spec) => ({ spec: structuredClone(spec), status: {}, lastRows: 0, lastColumns: 0 })),
  }
}

function nativeBounds(spec: ConnectorSpec): boolean {
  return boundedText(spec.id, 256) && boundedText(spec.name, 1024) && boundedText(spec.target.sheetId, 256)
    && spec.target.startRow <= 1_048_575 && spec.target.startColumn <= 16_383
    && utf8Length(spec.source.url) <= 8192 && !spec.source.url.includes('\\')
    && utf8Length(spec.source.path ?? '') <= 1024 && !/[\0\r\n]/.test(spec.source.path ?? '')
    && (spec.schedule === undefined || spec.schedule.intervalMs <= 31_536_000_000)
    && (spec.cache === undefined || spec.cache.ttlMs <= 31_536_000_000)
    && (spec.schema === undefined || spec.schema.columns.length <= 16_384
      && spec.schema.columns.every((column) => column.index <= 16_383))
}

function boundedText(value: string, maximumBytes: number): boolean {
  return value.length > 0 && value.trim() === value && !/[\0\r\n]/.test(value) && utf8Length(value) <= maximumBytes
}

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).byteLength
}

function requireBytes(workbook: Uint8Array): void {
  if (!(workbook instanceof Uint8Array) || workbook.byteLength === 0) {
    throw new TypeError('Native connector persistence requires nonempty XLSX bytes.')
  }
}

function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return
  const cause = signal.reason
  throw new ConnectorNativePersistenceError(
    'canceled',
    cause instanceof Error ? cause.message : String(cause || 'Native connector persistence canceled.'),
    cause === undefined ? undefined : { cause },
  )
}
