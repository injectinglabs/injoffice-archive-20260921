import { createDocxWasmClient, type DocxWasmClient } from '@injoffice/docx-wasm'
import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'
import type { NativeDocxOfficeMutationEnvelopeV1 } from '../../../packages/docs/src/nativeTransactionAdapterV1'
import { DOCX_MEDIA_TYPE, decodeNativeDocx } from './docsNativePreview'

export type DocxRoundTripMode = 'browser' | 'server'

export interface DocxExtractResult {
  document: NativeDocxDocumentV1
  artifactId: string
}

export interface DocxApplyInput {
  original: Uint8Array
  document: NativeDocxDocumentV1
  envelope: NativeDocxOfficeMutationEnvelopeV1
  artifactId?: string
  sourceName?: string
}

export interface DocxApplyResult {
  bytes: Uint8Array
  revision?: string
  packageSHA256?: string
  artifactId?: string
}

export interface DocxRoundTripRuntime {
  extract(bytes: Uint8Array, artifactId?: string): Promise<DocxExtractResult>
  apply(input: DocxApplyInput): Promise<DocxApplyResult>
  terminate(): void
}

export function createBrowserDocxRoundTripRuntime(
  client: DocxWasmClient = createDocxWasmClient(),
): DocxRoundTripRuntime {
  return {
    async extract(bytes) {
      return { document: await client.extract(bytes), artifactId: '' }
    },
    async apply({ original, document, envelope }) {
      return { bytes: await client.apply(original, document, envelope) }
    },
    terminate() {
      client.terminate()
    },
  }
}

export function createServerDocxRoundTripRuntime(
  apiBase: string,
  fetcher: typeof fetch = globalThis.fetch.bind(globalThis),
): DocxRoundTripRuntime {
  const base = apiBase.trim().replace(/\/$/, '')
  if (base.length === 0) throw new Error('Server fallback requires an explicit nonempty VITE_INJOFFICE_API_BASE.')
  return {
    async extract(bytes, artifactId) {
      const response = await fetcher(`${base}/v1/docx/extract`, {
        method: 'POST',
        headers: { 'Content-Type': artifactId ? 'application/json' : DOCX_MEDIA_TYPE },
        body: artifactId ? JSON.stringify({ artifact_id: artifactId }) : copyArrayBuffer(bytes),
      })
      const body = await response.text()
      if (!response.ok) throw responseError(response, body)
      return {
        document: decodeNativeDocx(JSON.parse(body) as unknown),
        artifactId: response.headers.get('X-InjOffice-Artifact-Id') ?? '',
      }
    },
    async apply({ original, envelope, artifactId, sourceName }) {
      const form = new FormData()
      if (artifactId) form.append('artifact_id', artifactId)
      else form.append('original', new Blob([copyArrayBuffer(original)], { type: DOCX_MEDIA_TYPE }), sourceName || 'document.docx')
      form.append('payload', JSON.stringify(envelope.payload))
      form.append('expected_revision', envelope.expected_revision)
      const response = await fetcher(`${base}/v1/docx/mutations`, { method: 'POST', body: form })
      if (!response.ok) throw responseError(response, await response.text())
      return {
        bytes: new Uint8Array(await response.arrayBuffer()),
        revision: response.headers.get('X-InjOffice-Revision') ?? undefined,
        packageSHA256: response.headers.get('X-InjOffice-Package-SHA256') ?? undefined,
        artifactId: response.headers.get('X-InjOffice-Artifact-Id') ?? undefined,
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
      if (typeof parsed === 'object' && parsed !== null && 'error' in parsed && typeof parsed.error === 'string') return new Error(parsed.error)
    } catch {
      // The handler may return a bounded plain-text refusal.
    }
  }
  return new Error(trimmed || `Request failed with HTTP ${response.status}.`)
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}
