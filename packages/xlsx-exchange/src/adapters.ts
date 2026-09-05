import {
  assertNativeWorkbookV2,
  type NativeWorkbookV2,
} from '@injoffice/sheets/browser'
import type {
  XlsxExchangeCodec,
  XlsxExchangeContext,
  XlsxReadResult,
  XlsxSnapshotImportResult,
} from './types.js'

const XLSX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** Structural subset implemented by `@injoffice/xlsx-wasm`'s XlsxWasmClient. */
export interface NativeWorkbookExtractor {
  extract(bytes: Uint8Array, options?: { signal?: AbortSignal }): Promise<NativeWorkbookV2>
}

/**
 * Snapshot-import adapter for the existing Go/WASM extractor. Export methods
 * stay absent because the native engine only writes explicit mutation batches.
 */
export function createXlsxWasmImportCodec(client: NativeWorkbookExtractor): XlsxExchangeCodec<NativeWorkbookV2> {
  if (!client || typeof client.extract !== 'function') throw new TypeError('XLSX WASM import requires an extractor.')
  return {
    async importSnapshot(file, context) {
      context.report({ phase: 'converting', fraction: 0.55, totalBytes: file.bytes.byteLength })
      const snapshot = await client.extract(file.bytes, { signal: context.signal })
      assertNativeWorkbookV2(snapshot)
      return {
        snapshot,
        revision: snapshot.source.package_sha256,
        metadata: { engine: 'xlsx-wasm' },
      }
    },
  }
}

export interface InjOfficeServerImportOptions {
  readonly apiBase: string
  readonly fetch?: typeof fetch
}

/** Import-only adapter for the existing `/v1/xlsx/extract` native server route. */
export function createInjOfficeServerImportCodec(options: InjOfficeServerImportOptions): XlsxExchangeCodec<NativeWorkbookV2> {
  const base = options.apiBase?.trim().replace(/\/$/, '')
  if (!base) throw new TypeError('InjOffice server import requires an explicit nonempty apiBase.')
  const fetcher = options.fetch ?? globalThis.fetch?.bind(globalThis)
  if (!fetcher) throw new TypeError('InjOffice server import requires fetch.')

  async function extract(file: XlsxReadResult, context: XlsxExchangeContext, artifactId?: string): Promise<XlsxSnapshotImportResult<NativeWorkbookV2>> {
    context.report({ phase: 'transferring', fraction: 0.5, totalBytes: file.bytes.byteLength })
    const response = await fetcher(`${base}/v1/xlsx/extract`, {
      method: 'POST',
      headers: artifactId
        ? { 'Content-Type': 'application/json' }
        : { 'Content-Type': file.mediaType ?? XLSX_MEDIA_TYPE },
      body: artifactId ? JSON.stringify({ artifact_id: artifactId }) : copyArrayBuffer(file.bytes),
      signal: context.signal,
    })
    const body = await response.text()
    if (!response.ok) throw serverError(response, body)
    let snapshot: unknown
    try {
      snapshot = JSON.parse(body)
    } catch (cause) {
      throw new Error('InjOffice server returned invalid XLSX snapshot JSON.', { cause })
    }
    assertNativeWorkbookV2(snapshot)
    return {
      snapshot,
      artifactId: response.headers.get('X-InjOffice-Artifact-Id')?.trim() || artifactId,
      revision: snapshot.source.package_sha256,
      metadata: { engine: 'injoffice-server' },
    }
  }

  return {
    importSnapshot: (file, context) => extract(file, context),
    async importServerUnit(file, context) {
      const imported = await extract(file, context)
      if (!imported.artifactId) throw new Error('InjOffice server did not return an artifact ID for the imported XLSX unit.')
      return {
        unitId: imported.artifactId,
        revision: imported.revision,
        metadata: imported.metadata,
      }
    },
  }
}

function serverError(response: Response, body: string): Error {
  const trimmed = body.trim()
  if (trimmed) {
    try {
      const parsed = JSON.parse(trimmed) as unknown
      if (isRecord(parsed) && typeof parsed.error === 'string') return new Error(parsed.error)
    } catch {
      // Native helper errors may be plain text.
    }
  }
  return new Error(trimmed || `InjOffice server request failed with HTTP ${response.status}.`)
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
