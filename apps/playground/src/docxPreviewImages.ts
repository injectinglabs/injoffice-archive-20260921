import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'

export const DOCX_PREVIEW_IMAGE_LIMITS = { packageBytes: 64 * 1024 * 1024, imageBytes: 5 * 1024 * 1024, totalBytes: 20 * 1024 * 1024, images: 50, pixels: 16_000_000, totalPixels: 40_000_000 } as const
export type DocxPreviewImage = { bytes: Uint8Array; mime: 'image/png' | 'image/jpeg'; width: number; height: number }
type ZipEntry = { method: number; flags: number; size: number; compressedSize: number; offset: number }

async function digest(bytes: Uint8Array, signal?: AbortSignal): Promise<string> {
  signal?.throwIfAborted()
  // WebCrypto cannot cancel an in-flight digest; do not continue after it ends.
  const result = await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)
  signal?.throwIfAborted()
  return `sha256:${Array.from(new Uint8Array(result), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

/** Bounded, single-disk ZIP reader. No path extraction, ZIP64, encryption or
 * remote resources. Exact part names and native SHA-256 inventories bind bytes. */
function zipEntries(bytes: Uint8Array): Map<string, ZipEntry> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let end = bytes.length - 22
  for (; end >= Math.max(0, bytes.length - 65_557); end--) {
    if (view.getUint32(end, true) === 0x06054b50 && end + 22 + view.getUint16(end + 20, true) === bytes.length) break
  }
  if (end < 0 || end < bytes.length - 65_557 || view.getUint16(end + 4, true) || view.getUint16(end + 6, true)) throw new Error('Unsupported ZIP directory')
  const count = view.getUint16(end + 10, true)
  const directorySize = view.getUint32(end + 12, true)
  const start = view.getUint32(end + 16, true)
  if (count === 0xffff || count !== view.getUint16(end + 8, true) || start + directorySize !== end) throw new Error('Unsupported ZIP directory')
  let offset = start
  const entries = new Map<string, ZipEntry>()
  for (let index = 0; index < count; index++) {
    if (offset + 46 > end || view.getUint32(offset, true) !== 0x02014b50) throw new Error('Invalid ZIP entry')
    const nameLength = view.getUint16(offset + 28, true)
    const next = offset + 46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true)
    if (next > end || view.getUint16(offset + 34, true)) throw new Error('Invalid ZIP entry')
    const name = new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(offset + 46, offset + 46 + nameLength))
    if (entries.has(name)) throw new Error('Ambiguous ZIP part')
    const local = view.getUint32(offset + 42, true)
    if (local + 30 > start || view.getUint32(local, true) !== 0x04034b50) throw new Error('Invalid local ZIP entry')
    const localNameLength = view.getUint16(local + 26, true)
    const dataOffset = local + 30 + localNameLength + view.getUint16(local + 28, true)
    const localName = new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(local + 30, local + 30 + localNameLength))
    const method = view.getUint16(offset + 10, true)
    const flags = view.getUint16(offset + 8, true)
    const compressedSize = view.getUint32(offset + 20, true)
    if (localName !== name || method !== view.getUint16(local + 8, true) || flags !== view.getUint16(local + 6, true) || dataOffset + compressedSize > start) throw new Error('Conflicting ZIP entry')
    entries.set(name, { method, flags, compressedSize, size: view.getUint32(offset + 24, true), offset: dataOffset })
    offset = next
  }
  if (offset !== end) throw new Error('Invalid ZIP directory size')
  return entries
}

async function entryBytes(source: Uint8Array, entry: ZipEntry, signal?: AbortSignal): Promise<Uint8Array> {
  signal?.throwIfAborted()
  if (entry.flags & ~0x808 || entry.size > DOCX_PREVIEW_IMAGE_LIMITS.imageBytes || entry.compressedSize > DOCX_PREVIEW_IMAGE_LIMITS.imageBytes) throw new Error('Image exceeds preview limits')
  const compressed = source.slice(entry.offset, entry.offset + entry.compressedSize)
  if (entry.method === 0) {
    if (compressed.length !== entry.size) throw new Error('Invalid stored image size')
    return compressed
  }
  if (entry.method !== 8) throw new Error('Unsupported image compression')
  const reader = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader()
  const cancel = () => { void reader.cancel().catch(() => {}) }
  signal?.addEventListener('abort', cancel, { once: true })
  const chunks: Uint8Array[] = []
  let length = 0
  try {
    while (true) {
      signal?.throwIfAborted()
      const next = await reader.read()
      signal?.throwIfAborted()
      if (next.done) break
      length += next.value.length
      if (length > entry.size || length > DOCX_PREVIEW_IMAGE_LIMITS.imageBytes) throw new Error('Inflated image exceeds preview limits')
      chunks.push(next.value)
    }
  } finally {
    signal?.removeEventListener('abort', cancel)
    await reader.cancel().catch(() => {})
  }
  if (length !== entry.size) throw new Error('Invalid inflated image size')
  const bytes = new Uint8Array(length)
  let offset = 0
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
  return bytes
}

export function docxImageDimensions(bytes: Uint8Array, mime: string): { width: number; height: number } | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (mime === 'image/png' && bytes.length >= 33 && [137, 80, 78, 71, 13, 10, 26, 10].every((byte, index) => bytes[index] === byte) && view.getUint32(8) === 13 && view.getUint32(12) === 0x49484452) {
    // Animated PNG can amplify decode memory beyond the single-frame budget.
    let chunk = 8
    while (chunk + 12 <= bytes.length) {
      const length = view.getUint32(chunk)
      if (chunk + 12 + length > bytes.length || view.getUint32(chunk + 4) === 0x6163544c) return undefined
      chunk += length + 12
    }
    if (chunk !== bytes.length) return undefined
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (mime !== 'image/jpeg' || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined
  let offset = 2
  while (offset + 4 <= bytes.length) {
    if (bytes[offset++] !== 0xff) return undefined
    while (bytes[offset] === 0xff) offset++
    const marker = bytes[offset++]
    if (marker === 0xda || marker === 0xd9 || offset + 2 > bytes.length) return undefined
    const length = view.getUint16(offset)
    if (length < 2 || offset + length > bytes.length) return undefined
    if ([0xc0, 0xc1, 0xc2].includes(marker) && length >= 8) return { height: view.getUint16(offset + 3), width: view.getUint16(offset + 5) }
    offset += length
  }
  return undefined
}

export async function extractDocxPreviewImages(source: Uint8Array, document: NativeDocxDocumentV1, signal?: AbortSignal): Promise<Map<string, DocxPreviewImage>> {
  const output = new Map<string, DocxPreviewImage>()
  if (signal?.aborted || source.length > DOCX_PREVIEW_IMAGE_LIMITS.packageBytes || source.length < 22) return output
  try { if (await digest(source, signal) !== document.source.package_sha256) return output } catch { return output }
  let entries: Map<string, ZipEntry>
  try { entries = zipEntries(source) } catch { return output }
  let totalBytes = 0
  let totalPixels = 0
  let attempts = 0
  for (const part of document.passthrough_parts) {
    if (signal?.aborted) return new Map()
    if (output.size >= DOCX_PREVIEW_IMAGE_LIMITS.images) break
    if (part.content_type !== 'image/png' && part.content_type !== 'image/jpeg') continue
    if (++attempts > DOCX_PREVIEW_IMAGE_LIMITS.images) break
    const entry = entries.get(part.part_name)
    if (!entry || entry.size !== part.byte_length || totalBytes + entry.size > DOCX_PREVIEW_IMAGE_LIMITS.totalBytes) continue
    // Reserve before attempting inflation: corrupt/truncated/forged entries
    // consume the same work budget as successfully decoded media.
    totalBytes += entry.size
    try {
      const bytes = await entryBytes(source, entry, signal)
      if (await digest(bytes, signal) !== part.sha256) continue
      const size = docxImageDimensions(bytes, part.content_type)
      if (!size || !size.width || !size.height || size.width * size.height > DOCX_PREVIEW_IMAGE_LIMITS.pixels || totalPixels + size.width * size.height > DOCX_PREVIEW_IMAGE_LIMITS.totalPixels) continue
      totalPixels += size.width * size.height
      output.set(part.part_name, { bytes, mime: part.content_type, ...size })
    } catch { /* Unsupported or malformed media stays an explicit placeholder. */ }
  }
  return signal?.aborted ? new Map() : output
}
