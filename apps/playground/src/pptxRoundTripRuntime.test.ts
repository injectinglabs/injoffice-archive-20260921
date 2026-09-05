import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { NativePptxDeck } from '@injoffice/pptx-native'
import type { PptxNativeMutationRequestV1, PptxWasmClient } from '@injoffice/pptx-wasm'
import {
  createBrowserPptxRoundTripRuntime,
  createServerPptxRoundTripRuntime,
} from './pptxRoundTripRuntime'

const deck = JSON.parse(readFileSync(
  new URL('../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json', import.meta.url),
  'utf8',
)) as NativePptxDeck
deck.sourceRevision = `rev-${'a'.repeat(64)}`

const mutation: PptxNativeMutationRequestV1 = {
  expectedSourceRevision: deck.sourceRevision,
  operations: [{
    operationId: 'edit-title',
    kind: 'text.replace',
    elementId: 'el-title',
    expectedFingerprintSha256: '5'.repeat(64),
    paragraphs: [{
      align: 'center', level: 0, bullet: false,
      runs: [{ text: 'After', bold: true, italic: false, fontSizeHundredthPt: 3200, color: '112233', fontFamily: 'Aptos' }],
    }],
  }],
}

describe('PPTX round-trip runtimes', () => {
  it('keeps the browser path entirely inside the injected WASM client', async () => {
    const client: PptxWasmClient = {
      extract: vi.fn(async () => deck),
      apply: vi.fn(async () => new Uint8Array([4, 5, 6])),
      terminate: vi.fn(),
    }
    const runtime = createBrowserPptxRoundTripRuntime(client)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original, 'ignored-server-artifact')).resolves.toEqual({ deck, artifactId: '' })
    await expect(runtime.apply({ original, deck, mutation })).resolves.toEqual({ bytes: new Uint8Array([4, 5, 6]) })
    expect(client.extract).toHaveBeenCalledWith(original)
    expect(client.apply).toHaveBeenCalledWith(original, deck, mutation)
    runtime.terminate()
    expect(client.terminate).toHaveBeenCalledOnce()
  })

  it('does not introduce a remote retry when browser extraction fails', async () => {
    const refusal = new Error('native refusal')
    const client: PptxWasmClient = {
      extract: vi.fn(async () => { throw refusal }),
      apply: vi.fn(),
      terminate: vi.fn(),
    }
    const runtime = createBrowserPptxRoundTripRuntime(client)

    await expect(runtime.extract(new Uint8Array([1]))).rejects.toBe(refusal)
    expect(client.extract).toHaveBeenCalledOnce()
  })

  it('uses HTTP only for the explicit server runtime and sends exact CAS fields', async () => {
    const fetcher = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify(deck), {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'artifact-1' },
      }))
      .mockResolvedValueOnce(new Response(new Uint8Array([7, 8, 9]), {
        status: 200,
        headers: {
          'X-InjOffice-Revision': `rev-${'b'.repeat(64)}`,
          'X-InjOffice-Package-SHA256': `sha256:${'b'.repeat(64)}`,
          'X-InjOffice-Artifact-Id': 'artifact-1',
        },
      }))
      .mockResolvedValueOnce(new Response(JSON.stringify(deck), {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'artifact-1' },
      }))
    const runtime = createServerPptxRoundTripRuntime('https://api.example.test/', fetcher as typeof fetch)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original)).resolves.toEqual({ deck, artifactId: 'artifact-1' })
    await expect(runtime.apply({ original, deck, mutation, artifactId: 'artifact-1' })).resolves.toMatchObject({ artifactId: 'artifact-1' })

    expect(fetcher).toHaveBeenNthCalledWith(1, 'https://api.example.test/v1/pptx/extract', expect.objectContaining({ method: 'POST' }))
    expect(new Uint8Array(fetcher.mock.calls[0]![1]!.body as ArrayBuffer)).toEqual(original)
    const form = fetcher.mock.calls[1]![1]!.body as FormData
    expect(form.get('artifact_id')).toBe('artifact-1')
    expect(form.get('original')).toBeNull()
    expect(form.get('payload')).toBe(JSON.stringify(mutation))
    expect(form.get('expected_revision')).toBe(`sha256:${'a'.repeat(64)}`)

    await expect(runtime.extract(new Uint8Array([7, 8, 9]), 'artifact-1')).resolves.toEqual({ deck, artifactId: 'artifact-1' })
    expect(fetcher).toHaveBeenNthCalledWith(3, 'https://api.example.test/v1/pptx/extract', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ artifact_id: 'artifact-1' }),
    })
  })

  it('uploads original bytes only after server mode is explicitly constructed', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response(new Uint8Array([4]), { status: 200 }))
    const runtime = createServerPptxRoundTripRuntime('https://api.example.test', fetcher as typeof fetch)
    await runtime.apply({ original: new Uint8Array([1, 2]), deck, mutation, sourceName: 'input.pptx' })

    const form = fetcher.mock.calls[0]![1]!.body as FormData
    expect(form.get('artifact_id')).toBeNull()
    expect(form.get('original')).toBeInstanceOf(Blob)
    expect((form.get('original') as File).name).toBe('input.pptx')
  })

  it('refuses to create a same-origin server fallback without explicit configuration', () => {
    expect(() => createServerPptxRoundTripRuntime('')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
    expect(() => createServerPptxRoundTripRuntime('   ')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
  })
})
