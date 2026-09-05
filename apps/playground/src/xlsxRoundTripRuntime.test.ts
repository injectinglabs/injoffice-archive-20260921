import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { adaptWorkbookMutationBatchV1, type XlsxWasmClient } from '@injoffice/xlsx-wasm'
import { decodeNativeWorkbook, editableTargets, buildCellMutation } from './nativeRoundTrip'
import {
  createBrowserXlsxRoundTripRuntime,
  createServerXlsxRoundTripRuntime,
} from './xlsxRoundTripRuntime'

const fixtureSource = readFileSync(
  new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json', import.meta.url),
  'utf8',
)
const workbook = decodeNativeWorkbook(JSON.parse(fixtureSource) as unknown)
const target = editableTargets(workbook)[0]!
const transaction = adaptWorkbookMutationBatchV1(workbook, buildCellMutation(workbook, target, 'after', 'edit-1'))

describe('XLSX round-trip runtimes', () => {
  it('keeps the browser path entirely inside the injected WASM client', async () => {
    const client: XlsxWasmClient = {
      extract: vi.fn(async () => workbook),
      apply: vi.fn(async () => new Uint8Array([4, 5, 6])),
      terminate: vi.fn(),
    }
    const runtime = createBrowserXlsxRoundTripRuntime(client)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original)).resolves.toEqual({ workbook, artifactId: '' })
    await expect(runtime.apply({ original, workbook, transaction })).resolves.toEqual({ bytes: new Uint8Array([4, 5, 6]) })
    expect(client.extract).toHaveBeenCalledWith(original)
    expect(client.apply).toHaveBeenCalledWith(original, workbook, transaction)
    runtime.terminate()
    expect(client.terminate).toHaveBeenCalledOnce()
  })

  it('does not introduce a remote retry when browser extraction fails', async () => {
    const refusal = new Error('native refusal')
    const client: XlsxWasmClient = {
      extract: vi.fn(async () => { throw refusal }),
      apply: vi.fn(),
      terminate: vi.fn(),
    }
    const runtime = createBrowserXlsxRoundTripRuntime(client)

    await expect(runtime.extract(new Uint8Array([1]))).rejects.toBe(refusal)
    expect(client.extract).toHaveBeenCalledOnce()
  })

  it('uses the HTTP API only for the explicit server runtime', async () => {
    const fetcher = vi.fn()
      .mockResolvedValueOnce(new Response(fixtureSource, {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'artifact-1' },
      }))
      .mockResolvedValueOnce(new Response(new Uint8Array([7, 8, 9]), {
        status: 200,
        headers: {
          'X-InjOffice-Revision': 'rev:saved',
          'X-InjOffice-Package-SHA256': 'sha256:saved',
        },
      }))
    const runtime = createServerXlsxRoundTripRuntime('https://api.example.test/', fetcher as typeof fetch)
    const original = new Uint8Array([1, 2, 3])

    await expect(runtime.extract(original)).resolves.toEqual({ workbook, artifactId: 'artifact-1' })
    await expect(runtime.apply({ original, workbook, transaction, artifactId: 'artifact-1' })).resolves.toEqual({
      bytes: new Uint8Array([7, 8, 9]),
      revision: 'rev:saved',
      packageSHA256: 'sha256:saved',
    })

    expect(fetcher).toHaveBeenNthCalledWith(1, 'https://api.example.test/v1/xlsx/extract', expect.objectContaining({ method: 'POST' }))
    expect(new Uint8Array(fetcher.mock.calls[0]![1]!.body as ArrayBuffer)).toEqual(original)
    const mutationRequest = fetcher.mock.calls[1]![1] as RequestInit
    const form = mutationRequest.body as FormData
    expect(form.get('artifact_id')).toBe('artifact-1')
    expect(form.get('original')).toBeNull()
    expect(form.get('payload')).toBe(JSON.stringify(transaction))
    expect(form.get('expected_revision')).toBe(workbook.source.package_sha256)
  })

  it('uploads original bytes only after the server runtime is selected', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response(new Uint8Array([4]), { status: 200 }))
    const runtime = createServerXlsxRoundTripRuntime('https://api.example.test', fetcher as typeof fetch)
    await runtime.apply({ original: new Uint8Array([1, 2]), workbook, transaction, sourceName: 'input.xlsx' })

    const form = fetcher.mock.calls[0]![1]!.body as FormData
    expect(form.get('artifact_id')).toBeNull()
    expect(form.get('original')).toBeInstanceOf(Blob)
    expect((form.get('original') as File).name).toBe('input.xlsx')
  })

  it('refuses to create a same-origin server fallback without explicit configuration', () => {
    expect(() => createServerXlsxRoundTripRuntime('')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
    expect(() => createServerXlsxRoundTripRuntime('   ')).toThrow(/explicit nonempty VITE_INJOFFICE_API_BASE/)
  })
})
