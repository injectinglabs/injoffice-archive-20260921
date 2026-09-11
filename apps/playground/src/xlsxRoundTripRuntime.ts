import {
  createXlsxWasmClient,
  type XlsxNativeMutationTransactionV1,
  type XlsxWasmClient,
} from '@injoffice/xlsx-wasm'
import { decodeNativeWorkbookObjectsV1, type NativeWorkbookV2, type NativeWorkbookObjectsV1 } from '@injoffice/sheets/browser'
import { decodeNativeWorkbook } from './nativeRoundTrip'
import { readNativePreviewResponse } from './components/NativeDocxPages'

export type XlsxRoundTripMode = 'browser' | 'server'

export interface XlsxExtractResult {
  workbook: NativeWorkbookV2
  artifactId: string
}

export interface XlsxApplyInput {
  original: Uint8Array
  workbook: NativeWorkbookV2
  transaction: XlsxNativeMutationTransactionV1
  artifactId?: string
  sourceName?: string
}

export interface XlsxApplyResult {
  bytes: Uint8Array
  revision?: string
  packageSHA256?: string
}

export interface XlsxRoundTripRuntime {
  inspectObjects?(bytes: Uint8Array, packageSHA256: string): Promise<NativeWorkbookObjectsV1>
  extract(bytes: Uint8Array): Promise<XlsxExtractResult>
  apply(input: XlsxApplyInput): Promise<XlsxApplyResult>
  terminate(): void
}

export function createBrowserXlsxRoundTripRuntime(
  client: XlsxWasmClient = createXlsxWasmClient(),
): XlsxRoundTripRuntime {
  return {
    inspectObjects: (bytes, hash) => client.inspectObjects(bytes, hash),
    async extract(bytes) {
      return { workbook: await client.extract(bytes), artifactId: '' }
    },
    async apply({ original, workbook, transaction }) {
      return { bytes: await client.apply(original, workbook, transaction) }
    },
    terminate() {
      client.terminate()
    },
  }
}

export function createServerXlsxRoundTripRuntime(
  apiBase: string,
  fetcher: typeof fetch = globalThis.fetch.bind(globalThis),
): XlsxRoundTripRuntime {
  const base = apiBase.trim().replace(/\/$/, '')
  if (base.length === 0) {
    throw new Error('Server fallback requires an explicit nonempty VITE_INJOFFICE_API_BASE.')
  }
  return {
    async inspectObjects(bytes, hash) {
      const response = await fetcher(`${base}/v1/xlsx/preview-objects`, {method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:copyArrayBuffer(bytes),credentials:'omit',redirect:'error'})
      const body = await readNativePreviewResponse(response, 8 * 1024 * 1024)
      if (!response.ok) throw responseError(response, JSON.stringify(body))
      return decodeNativeWorkbookObjectsV1(body, hash)
    },
    async extract(bytes) {
      const response = await fetcher(`${base}/v1/xlsx/extract`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
        body: copyArrayBuffer(bytes),
      })
      const body = await response.text()
      if (!response.ok) throw responseError(response, body)
      return {
        workbook: decodeNativeWorkbook(JSON.parse(body) as unknown),
        artifactId: response.headers.get('X-InjOffice-Artifact-Id') ?? '',
      }
    },
    async apply({ original, workbook, transaction, artifactId, sourceName }) {
      const form = new FormData()
      if (artifactId) form.append('artifact_id', artifactId)
      else form.append('original', new Blob([copyArrayBuffer(original)]), sourceName || 'workbook.xlsx')
      form.append('payload', JSON.stringify(transaction))
      form.append('expected_revision', workbook.source.package_sha256)
      const response = await fetcher(`${base}/v1/xlsx/mutations`, { method: 'POST', body: form })
      if (!response.ok) throw responseError(response, await response.text())
      return {
        bytes: new Uint8Array(await response.arrayBuffer()),
        revision: response.headers.get('X-InjOffice-Revision') ?? undefined,
        packageSHA256: response.headers.get('X-InjOffice-Package-SHA256') ?? undefined,
      }
    },
    terminate() {},
  }
}

function responseError(response: Response, detail: string): Error {
  const trimmed = detail.trim()
  if (trimmed) {
    try {
      const parsed = JSON.parse(trimmed) as unknown
      if (isRecord(parsed) && typeof parsed.error === 'string') return new Error(parsed.error)
    } catch {
      // The handler may return a plain-text refusal.
    }
  }
  return new Error(trimmed || `Request failed with HTTP ${response.status}.`)
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
