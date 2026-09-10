import { createHash } from 'node:crypto'
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it, vi } from 'vitest'
import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'
import { docxImageDimensions, extractDocxPreviewImages } from './docxPreviewImages'

const png = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=', 'base64'))
const hash = (bytes: Uint8Array) => `sha256:${createHash('sha256').update(bytes).digest('hex')}`
function fixture(image = png, method = 0) {
  const name = new TextEncoder().encode('word/media/image1.png')
  const compressed = method === 8 ? deflateRawSync(image) : image
  const start = 30 + name.length + compressed.length
  const bytes = new Uint8Array(start + 46 + name.length + 22)
  const view = new DataView(bytes.buffer)
  view.setUint32(0, 0x04034b50, true)
  view.setUint16(8, method, true)
  view.setUint32(18, compressed.length, true)
  view.setUint32(22, image.length, true)
  view.setUint16(26, name.length, true)
  bytes.set(name, 30)
  bytes.set(compressed, 30 + name.length)
  view.setUint32(start, 0x02014b50, true)
  view.setUint16(start + 10, method, true)
  view.setUint32(start + 20, compressed.length, true)
  view.setUint32(start + 24, image.length, true)
  view.setUint16(start + 28, name.length, true)
  bytes.set(name, start + 46)
  const end = start + 46 + name.length
  view.setUint32(end, 0x06054b50, true)
  view.setUint16(end + 8, 1, true)
  view.setUint16(end + 10, 1, true)
  view.setUint32(end + 12, 46 + name.length, true)
  view.setUint32(end + 16, start, true)
  const document = { source: { package_sha256: hash(bytes) }, passthrough_parts: [{ part_name: 'word/media/image1.png', content_type: 'image/png', byte_length: image.length, sha256: hash(image) }] } as NativeDocxDocumentV1
  return { bytes, document }
}

describe('source-bound DOCX preview images', () => {
  it('returns no media when cancelled before or during package hashing', async () => {
    const { bytes, document } = fixture()
    const before = new AbortController()
    before.abort()
    expect((await extractDocxPreviewImages(bytes, document, before.signal)).size).toBe(0)
    const during = new AbortController()
    const pending = extractDocxPreviewImages(bytes, document, during.signal)
    during.abort()
    expect((await pending).size).toBe(0)
  })
  it('stops cancelled inflation without rejecting a late viewer effect', async () => {
    const { bytes, document } = fixture(png, 8)
    const controller = new AbortController()
    const NativeStream = DecompressionStream
    vi.stubGlobal('DecompressionStream', class {
      constructor(format: CompressionFormat) { controller.abort(); return new NativeStream(format) }
    })
    try { expect((await extractDocxPreviewImages(bytes, document, controller.signal)).size).toBe(0) }
    finally { vi.unstubAllGlobals() }
  })
  it('reserves cumulative work budget even when every inflation attempt fails', async () => {
    const { bytes, document } = fixture(png, 8)
    const view = new DataView(bytes.buffer)
    const start = view.getUint32(bytes.length - 6, true)
    view.setUint32(start + 24, 5 * 1024 * 1024, true)
    document.source.package_sha256 = hash(bytes)
    const part = { ...document.passthrough_parts[0], byte_length: 5 * 1024 * 1024 }
    document.passthrough_parts = Array.from({ length: 8 }, () => ({ ...part }))
    let attempts = 0
    vi.stubGlobal('DecompressionStream', class {
      constructor() { attempts++; throw new Error('corrupt compressed stream') }
    })
    try {
      expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
      expect(attempts).toBe(4)
    } finally { vi.unstubAllGlobals() }
  })
  it.each([0, 8])('extracts verified embedded PNG using ZIP method %i', async (method) => {
    const { bytes, document } = fixture(png, method)
    const result = await extractDocxPreviewImages(bytes, document)
    expect(result.get('word/media/image1.png')).toEqual({ bytes: png, mime: 'image/png', width: 1, height: 1 })
  })
  it('refuses package or media digest mismatch', async () => {
    const { bytes, document } = fixture()
    document.source.package_sha256 = 'sha256:wrong'
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
    document.source.package_sha256 = hash(bytes)
    document.passthrough_parts[0].sha256 = 'sha256:wrong'
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
  })
  it('requires exact part names and refuses SVG or MIME spoofing', async () => {
    const { bytes, document } = fixture()
    document.passthrough_parts[0].part_name = 'https://example.com/image.png'
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
    document.passthrough_parts[0].part_name = 'word/media/image1.png'
    document.passthrough_parts[0].content_type = 'image/svg+xml'
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
    document.passthrough_parts[0].content_type = 'image/jpeg'
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
  })
  it('rejects dangerous image dimensions before creating a browser image', async () => {
    const large = png.slice()
    new DataView(large.buffer).setUint32(16, 0x7fffffff)
    const { bytes, document } = fixture(large)
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
  })
  it('rejects oversized inflation declarations and corrupt directories', async () => {
    const { bytes, document } = fixture()
    const view = new DataView(bytes.buffer)
    const start = view.getUint32(bytes.length - 6, true)
    view.setUint32(start + 24, 6 * 1024 * 1024, true)
    document.passthrough_parts[0].byte_length = 6 * 1024 * 1024
    document.source.package_sha256 = hash(bytes)
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
    view.setUint32(start, 0, true)
    document.source.package_sha256 = hash(bytes)
    expect((await extractDocxPreviewImages(bytes, document)).size).toBe(0)
  })
  it('reads JPEG dimensions and rejects truncated segments', () => {
    const jpeg = Uint8Array.from([255, 216, 255, 192, 0, 8, 8, 0, 10, 0, 20, 1])
    expect(docxImageDimensions(jpeg, 'image/jpeg')).toEqual({ width: 20, height: 10 })
    expect(docxImageDimensions(jpeg.slice(0, 7), 'image/jpeg')).toBeUndefined()
  })
})
