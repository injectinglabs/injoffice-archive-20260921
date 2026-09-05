import { assertNativePptx, type NativePptxDeck } from '@injoffice/pptx-native'
import {
  createPptxWasmClient,
  type PptxNativeMutationRequestV1,
  type PptxWasmClient,
} from '@injoffice/pptx-wasm'

export type PptxRoundTripMode = 'browser' | 'server'

export interface PptxExtractResult {
  deck: NativePptxDeck
  artifactId: string
}

export interface PptxApplyInput {
  original: Uint8Array
  deck: NativePptxDeck
  mutation: PptxNativeMutationRequestV1
  artifactId?: string
  sourceName?: string
}

export interface PptxApplyResult {
  bytes: Uint8Array
  revision?: string
  packageSHA256?: string
  artifactId?: string
}

export interface PptxRoundTripRuntime {
  extract(bytes: Uint8Array, artifactId?: string): Promise<PptxExtractResult>
  apply(input: PptxApplyInput): Promise<PptxApplyResult>
  terminate(): void
}

export function createBrowserPptxRoundTripRuntime(
  client: PptxWasmClient = createPptxWasmClient(),
): PptxRoundTripRuntime {
  return {
    async extract(bytes) {
      return { deck: await client.extract(bytes), artifactId: '' }
    },
    async apply({ original, deck, mutation }) {
      return { bytes: await client.apply(original, deck, mutation) }
    },
    terminate() {
      client.terminate()
    },
  }
}

export function createServerPptxRoundTripRuntime(
  apiBase: string,
  fetcher: typeof fetch = globalThis.fetch.bind(globalThis),
): PptxRoundTripRuntime {
  const base = apiBase.trim().replace(/\/$/, '')
  if (base.length === 0) throw new Error('Server fallback requires an explicit nonempty VITE_INJOFFICE_API_BASE.')
  return {
    async extract(bytes, artifactId) {
      const response = await fetcher(`${base}/v1/pptx/extract`, {
        method: 'POST',
        headers: { 'Content-Type': artifactId ? 'application/json' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        body: artifactId ? JSON.stringify({ artifact_id: artifactId }) : copyArrayBuffer(bytes),
      })
      const body = await response.text()
      if (!response.ok) throw responseError(response, body)
      const deck = JSON.parse(body) as unknown
      assertNativePptx(deck)
      return { deck, artifactId: response.headers.get('X-InjOffice-Artifact-Id') ?? '' }
    },
    async apply({ original, deck, mutation, artifactId, sourceName }) {
      const form = new FormData()
      if (artifactId) form.append('artifact_id', artifactId)
      else form.append('original', new Blob([copyArrayBuffer(original)]), sourceName || 'presentation.pptx')
      form.append('payload', JSON.stringify(mutation))
      form.append('expected_revision', outerRevision(deck.sourceRevision))
      const response = await fetcher(`${base}/v1/pptx/mutations`, { method: 'POST', body: form })
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

function outerRevision(sourceRevision: string | undefined): string {
  if (!sourceRevision || !/^rev-[0-9a-f]{64}$/.test(sourceRevision)) throw new Error('PPTX server mutation requires an exact parsed source revision.')
  return `sha256:${sourceRevision.slice(4)}`
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
  return value !== null && typeof value === 'object'
}
