import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { DocxWasmClient } from '@injoffice/docx-wasm'
import { decodeNativeDocx } from './docsNativePreview'
import { buildDocxRunMutation, editableDocxRuns } from './docxRoundTrip'
import {
  createBrowserDocxRoundTripRuntime,
  createServerDocxRoundTripRuntime,
} from './docxRoundTripRuntime'

const fixtureSource = readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url),
  'utf8',
)
const document = decodeNativeDocx((JSON.parse(fixtureSource) as { native: unknown }).native)
const target = editableDocxRuns(document)[0]!
const envelope = buildDocxRunMutation(document, target, 'After', 'save-1')

describe('DOCX round-trip runtimes', () => {
  it('keeps the browser path entirely inside the injected WASM client', async () => {
    const client: DocxWasmClient = {
      inspectPartialContent:vi.fn(),
      extract: vi.fn(async () => document),
      apply: vi.fn(async () => new Uint8Array([4, 5, 6])),
      terminate: vi.fn(),
    }
    const runtime = createBrowserDocxRoundTripRuntime(client)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original)).resolves.toEqual({ document, artifactId: '' })
    await expect(runtime.apply({ original, document, envelope })).resolves.toEqual({ bytes: new Uint8Array([4, 5, 6]) })
    expect(client.extract).toHaveBeenCalledWith(original)
    expect(client.apply).toHaveBeenCalledWith(original, document, envelope)
    runtime.terminate()
    expect(client.terminate).toHaveBeenCalledOnce()
  })

  it('does not introduce a remote retry when browser extraction fails', async () => {
    const refusal = new Error('native refusal')
    const client: DocxWasmClient = {
      inspectPartialContent:vi.fn(),
      extract: vi.fn(async () => { throw refusal }),
      apply: vi.fn(),
      terminate: vi.fn(),
    }
    const runtime = createBrowserDocxRoundTripRuntime(client)
    await expect(runtime.extract(new Uint8Array([1]))).rejects.toBe(refusal)
    expect(client.extract).toHaveBeenCalledOnce()
  })

  it('uses the HTTP API only for the explicit server runtime', async () => {
    const fetcher = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify(document), {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'artifact-1' },
      }))
      .mockResolvedValueOnce(new Response(new Uint8Array([7, 8, 9]), {
        status: 200,
        headers: {
          'X-InjOffice-Revision': 'rev:saved',
          'X-InjOffice-Package-SHA256': 'sha256:saved',
          'X-InjOffice-Artifact-Id': 'artifact-1',
        },
      }))
      .mockResolvedValueOnce(new Response(JSON.stringify(document), {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'artifact-1' },
      }))
    const runtime = createServerDocxRoundTripRuntime('https://api.example.test/', fetcher as typeof fetch)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original)).resolves.toEqual({ document, artifactId: 'artifact-1' })
    await expect(runtime.apply({ original, document, envelope, artifactId: 'artifact-1' })).resolves.toEqual({
      bytes: new Uint8Array([7, 8, 9]),
      revision: 'rev:saved',
      packageSHA256: 'sha256:saved',
      artifactId: 'artifact-1',
    })
    expect(fetcher).toHaveBeenNthCalledWith(1, 'https://api.example.test/v1/docx/extract', expect.objectContaining({ method: 'POST' }))
    expect(new Uint8Array(fetcher.mock.calls[0]![1]!.body as ArrayBuffer)).toEqual(original)
    const form = fetcher.mock.calls[1]![1]!.body as FormData
    expect(form.get('artifact_id')).toBe('artifact-1')
    expect(form.get('original')).toBeNull()
    expect(form.get('payload')).toBe(JSON.stringify(envelope.payload))
    expect(form.get('expected_revision')).toBe(document.source.package_sha256)

    await expect(runtime.extract(new Uint8Array([7, 8, 9]), 'artifact-1')).resolves.toEqual({ document, artifactId: 'artifact-1' })
    expect(fetcher).toHaveBeenNthCalledWith(3, 'https://api.example.test/v1/docx/extract', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ artifact_id: 'artifact-1' }),
    })
  })

  it('uploads original bytes only after the server runtime is selected', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response(new Uint8Array([4]), { status: 200 }))
    const runtime = createServerDocxRoundTripRuntime('https://api.example.test', fetcher as typeof fetch)
    await runtime.apply({ original: new Uint8Array([1, 2]), document, envelope, sourceName: 'input.docx' })
    const form = fetcher.mock.calls[0]![1]!.body as FormData
    expect(form.get('artifact_id')).toBeNull()
    expect(form.get('original')).toBeInstanceOf(Blob)
    expect((form.get('original') as File).name).toBe('input.docx')
  })

  it('refuses to create a same-origin server fallback without explicit configuration', () => {
    expect(() => createServerDocxRoundTripRuntime('')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
    expect(() => createServerDocxRoundTripRuntime('   ')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
  })
})
