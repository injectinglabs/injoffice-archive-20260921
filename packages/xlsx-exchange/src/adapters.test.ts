import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { NativeWorkbookV2 } from '@injoffice/sheets/browser'
import { createInjOfficeServerImportCodec, createXlsxWasmImportCodec } from './adapters'
import type { XlsxExchangeContext } from './types'

const fixtureSource = readFileSync(
  new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json', import.meta.url),
  'utf8',
)
const workbook = JSON.parse(fixtureSource) as NativeWorkbookV2

function context(): XlsxExchangeContext {
  return { jobId: 'job-1', signal: new AbortController().signal, report: vi.fn() }
}

describe('native XLSX import adapters', () => {
  it('delegates snapshot extraction to the existing WASM client boundary', async () => {
    const extract = vi.fn(async () => workbook)
    const codec = createXlsxWasmImportCodec({ extract })
    const ctx = context()
    const bytes = new Uint8Array([1, 2])
    await expect(codec.importSnapshot!({ bytes, name: 'book.xlsx' }, ctx)).resolves.toMatchObject({
      snapshot: workbook,
      revision: workbook.source.package_sha256,
      metadata: { engine: 'xlsx-wasm' },
    })
    expect(extract).toHaveBeenCalledWith(bytes, { signal: ctx.signal })
    expect(codec.exportSnapshot).toBeUndefined()
  })

  it('uses only an explicitly configured native server import endpoint', async () => {
    let sentInit: RequestInit | undefined
    const fetcher = vi.fn(async (_input: string | URL | Request, init?: RequestInit) => {
      sentInit = init
      return new Response(fixtureSource, {
        status: 200,
        headers: { 'X-InjOffice-Artifact-Id': 'art_123' },
      })
    })
    const codec = createInjOfficeServerImportCodec({ apiBase: 'https://api.example.test/', fetch: fetcher as typeof fetch })
    const ctx = context()
    const bytes = new Uint8Array([3, 4])
    await expect(codec.importSnapshot!({ bytes }, ctx)).resolves.toMatchObject({ artifactId: 'art_123', snapshot: workbook })
    expect(fetcher).toHaveBeenCalledWith('https://api.example.test/v1/xlsx/extract', expect.objectContaining({
      method: 'POST',
      signal: ctx.signal,
      headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    }))
    const sent = sentInit!.body as ArrayBuffer
    expect(new Uint8Array(sent)).toEqual(bytes)

    await expect(codec.importServerUnit!({ bytes }, ctx)).resolves.toMatchObject({
      unitId: 'art_123',
      revision: workbook.source.package_sha256,
    })
  })

  it('refuses implicit same-origin server use and structures server errors', async () => {
    expect(() => createInjOfficeServerImportCodec({ apiBase: '' })).toThrow(/explicit nonempty apiBase/)
    const codec = createInjOfficeServerImportCodec({
      apiBase: 'https://api.example.test',
      fetch: vi.fn(async () => new Response('{"error":"artifact not found"}', { status: 404 })) as typeof fetch,
    })
    await expect(codec.importSnapshot!({ bytes: new Uint8Array([1]) }, context())).rejects.toThrow('artifact not found')
  })

  it('refuses server-unit import when the configured store does not mint an artifact ID', async () => {
    const codec = createInjOfficeServerImportCodec({
      apiBase: 'https://api.example.test',
      fetch: vi.fn(async () => new Response(fixtureSource, { status: 200 })) as typeof fetch,
    })
    await expect(codec.importServerUnit!({ bytes: new Uint8Array([1]) }, context())).rejects.toThrow(/did not return an artifact ID/)
  })
})
