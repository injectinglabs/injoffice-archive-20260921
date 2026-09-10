/**
 * Exact, bounded media boundary for native DOCX page-paint v1.
 *
 * Embedded, identity-transformed inline PNG and baseline JFIF JPEG pictures
 * are qualified. The
 * package extractor remains the relationship authority; this module exact-joins
 * its drawing projection to preserved package-part fingerprints and caller-
 * supplied bytes. It never fetches, decodes for layout, or accepts a renderer's
 * interpretation as authority.
 */

import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import type { NativeDocxDocumentV1, NativeDocxDrawingV1 } from './nativeContract.js'
import { asciiLowerNative as asciiLower, asciiUpperNative, compareNativeCodeUnits } from './nativeDeterminism.js'
import { nativeBaselineJpegDimensions } from './nativeJpegV1.js'

export const DOCX_INLINE_IMAGE_LIMITS = Object.freeze({
  maxAssets: 256,
  maxAssetBytes: 16 * 1024 * 1024,
  maxTotalBytes: 64 * 1024 * 1024,
  maxPixelDimension: 32_768,
  maxPixels: 100_000_000,
  maxGeometryMilliPoints: 1_000_000_000,
})

export interface NativeDocxAuthoritativeMediaAssetV1 {
  part_name: string
  content_type: string
  content_digest: `sha256:${string}`
  bytes: Uint8Array
}

export interface NativeDocxPagePaintMediaAssetV1 {
  id: string
  part_name: string
  content_type: 'image/png' | 'image/jpeg'
  content_digest: `sha256:${string}`
  byte_length: number
  width_px: number
  height_px: number
  bytes_base64: string
}

export interface NativeDocxQualifiedInlineImageV1 {
  drawing_id: string
  run_id: string
  asset_id: string
  part_name: string
  relationship_id: string
  relationship_part: string
  relationship_sha256: `sha256:${string}`
  content_type: 'image/png' | 'image/jpeg'
  content_digest: `sha256:${string}`
  byte_length: number
  width_emu: number
  height_emu: number
  width_millipoints: number
  height_millipoints: number
  source_crop: { left: 0; top: 0; right: 0; bottom: 0; unit: 'one-hundred-thousandth' }
  transform: { rotation_degrees: 0; flip_horizontal: false; flip_vertical: false }
}

export type NativeDocxInlineImageQualificationV1 =
  | { ok: true; value: NativeDocxQualifiedInlineImageV1 }
  | { ok: false; code: 'unsupported-image' | 'invalid-image' | 'resource-limit'; message: string }

const SHA256 = /^sha256:[0-9a-f]{64}$/
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const

function canonicalPart(value: string): string {
  try { return asciiLower(value.split('/').map((segment) => decodeURIComponent(segment)).join('/')) } catch { return value }
}

function validPartName(value: string): boolean {
  if (value.length === 0 || value.length > 1024 || value.startsWith('/') || value.endsWith('/') || value.includes('\\') || value.includes('?') || value.includes('#') || value.includes('\0')) return false
  try {
    return value.split('/').every((segment) => {
      const decoded = decodeURIComponent(segment)
      return PART_SEGMENT.test(segment) && decoded.length > 0 && decoded !== '.' && decoded !== '..' && !decoded.endsWith('.') && !/[\\/?#%]/.test(decoded)
        && ![...decoded].some((character) => {
          const code = character.codePointAt(0) ?? 0
          return code < 0x20 || code === 0x7f
        })
    })
  } catch {
    return false
  }
}

function digest(bytes: Uint8Array): `sha256:${string}` {
  return `sha256:${bytesToHex(sha256(bytes))}`
}

function base64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  let output = ''
  for (let index = 0; index < bytes.length; index += 3) {
    const a = bytes[index]!
    const hasB = index + 1 < bytes.length
    const hasC = index + 2 < bytes.length
    const b = hasB ? bytes[index + 1]! : 0
    const c = hasC ? bytes[index + 2]! : 0
    output += alphabet[a >>> 2]
      + alphabet[((a & 3) << 4) | (b >>> 4)]
      + (hasB ? alphabet[((b & 15) << 2) | (c >>> 6)] : '=')
      + (hasC ? alphabet[c & 63] : '=')
  }
  return output
}

function unbase64(value: string): Uint8Array | undefined {
  if (!BASE64.test(value)) return undefined
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  const output = new Uint8Array((value.length / 4) * 3 - (value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0))
  let cursor = 0
  for (let index = 0; index < value.length; index += 4) {
    const a = alphabet.indexOf(value[index]!)
    const b = alphabet.indexOf(value[index + 1]!)
    const c = value[index + 2] === '=' ? 0 : alphabet.indexOf(value[index + 2]!)
    const d = value[index + 3] === '=' ? 0 : alphabet.indexOf(value[index + 3]!)
    if (a < 0 || b < 0 || c < 0 || d < 0) return undefined
    if (cursor < output.length) output[cursor++] = (a << 2) | (b >>> 4)
    if (cursor < output.length) output[cursor++] = ((b & 15) << 4) | (c >>> 2)
    if (cursor < output.length) output[cursor++] = ((c & 3) << 6) | d
  }
  return base64(output) === value ? output : undefined
}

function crc32(bytes: Uint8Array, start: number, end: number): number {
  let crc = 0xffffffff
  for (let index = start; index < end; index += 1) {
    crc ^= bytes[index]!
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1))
  }
  return (crc ^ 0xffffffff) >>> 0
}

function pngDimensions(bytes: Uint8Array): { width: number; height: number } | undefined {
  if (bytes.byteLength < 33 || PNG_SIGNATURE.some((value, index) => bytes[index] !== value)) return undefined
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(8, false) !== 13 || String.fromCharCode(...bytes.slice(12, 16)) !== 'IHDR') return undefined
  const width = view.getUint32(16, false)
  const height = view.getUint32(20, false)
  const bitDepth = bytes[24]
  const colorType = bytes[25]
  const compression = bytes[26]
  const filter = bytes[27]
  const interlace = bytes[28]
  const validDepths = new Map<number, readonly number[]>([[0, [1, 2, 4, 8, 16]], [2, [8, 16]], [3, [1, 2, 4, 8]], [4, [8, 16]], [6, [8, 16]]])
  if (width === 0 || height === 0 || !validDepths.get(colorType ?? -1)?.includes(bitDepth ?? 0) || compression !== 0 || filter !== 0 || ![0, 1].includes(interlace ?? -1)) return undefined
  // Animated PNG changes the visual result over time and is outside a static page-paint contract.
  let chunkIndex = 0
  let sawPalette = false
  let sawImageData = false
  let leftImageData = false
  for (let offset = 8; offset + 12 <= bytes.byteLength;) {
    const length = view.getUint32(offset, false)
    if (length > bytes.byteLength - offset - 12) return undefined
    const type = String.fromCharCode(...bytes.slice(offset + 4, offset + 8))
    const dataStart = offset + 8
    const dataEnd = dataStart + length
    if (!/^[A-Za-z]{4}$/.test(type) || view.getUint32(dataEnd, false) !== crc32(bytes, offset + 4, dataEnd)) return undefined
    if (chunkIndex === 0 && type !== 'IHDR' || chunkIndex > 0 && type === 'IHDR') return undefined
    if (type === 'acTL') return undefined
    if (type === 'PLTE') {
      if (sawPalette || sawImageData || colorType === 0 || colorType === 4 || length === 0 || length % 3 !== 0 || length > 768 || colorType === 3 && length / 3 > 2 ** (bitDepth ?? 0)) return undefined
      sawPalette = true
    }
    if (type === 'IDAT') {
      if (leftImageData || colorType === 3 && !sawPalette) return undefined
      sawImageData = true
    } else if (sawImageData && type !== 'IEND') {
      leftImageData = true
    }
    // Unknown critical chunks change decoding semantics; preserve-only rather than guessing.
    if (type[0] === asciiUpperNative(type[0] ?? '') && !['IHDR', 'PLTE', 'IDAT', 'IEND'].includes(type)) return undefined
    offset += 12 + length
    chunkIndex += 1
    if (type === 'IEND') return length === 0 && sawImageData && offset === bytes.byteLength ? { width, height } : undefined
  }
  return undefined
}

function emuToMilliPoints(value: number): number | undefined {
  // 914400 EMU = 72000 milli-points, so the exact reduced ratio is 10/127.
  if (!Number.isSafeInteger(value) || value <= 0 || value % 127 !== 0) return undefined
  const output = (value / 127) * 10
  return Number.isSafeInteger(output) && output <= DOCX_INLINE_IMAGE_LIMITS.maxGeometryMilliPoints ? output : undefined
}

function imageAssetID(contentDigest: string, partName: string): string {
  const partDigest = bytesToHex(sha256(new TextEncoder().encode(canonicalPart(partName))))
  return `image:${contentDigest.slice('sha256:'.length)}:${partDigest}`
}

function uniquePreservedPart(document: NativeDocxDocumentV1, partName: string) {
  const key = canonicalPart(partName)
  const matches = document.passthrough_parts.filter((part) => canonicalPart(part.part_name) === key)
  return matches.length === 1 ? matches[0] : undefined
}

function relationshipPart(ownerPart: string): string {
  const slash = ownerPart.lastIndexOf('/')
  const directory = slash < 0 ? '' : ownerPart.slice(0, slash + 1)
  const base = slash < 0 ? ownerPart : ownerPart.slice(slash + 1)
  return `${directory}_rels/${base}.rels`
}

export function qualifyNativeDocxInlineImageV1(document: NativeDocxDocumentV1, runID: string, drawing: NativeDocxDrawingV1): NativeDocxInlineImageQualificationV1 {
  if (drawing.placement !== 'inline' || drawing.x_emu !== undefined || drawing.y_emu !== undefined || drawing.wrap !== undefined || drawing.horizontal_relative_from !== undefined || drawing.vertical_relative_from !== undefined) {
    return { ok: false, code: 'unsupported-image', message: 'Only bounded inline pictures without anchor, wrap, or floating offsets are supported' }
  }
  if (!drawing.relationship_id || !drawing.media_part || !drawing.content_type) return { ok: false, code: 'invalid-image', message: 'Inline picture lacks an exact embedded relationship/media identity' }
  const contentType = asciiLower(drawing.content_type)
  if (contentType !== 'image/png' && contentType !== 'image/jpeg') return { ok: false, code: 'unsupported-image', message: 'Only embedded static PNG or baseline JFIF JPEG media is supported' }
  const part = uniquePreservedPart(document, drawing.media_part)
  if (!part || canonicalPart(part.part_name) !== canonicalPart(drawing.media_part) || asciiLower(part.content_type) !== contentType || !SHA256.test(part.sha256) || !Number.isSafeInteger(part.byte_length) || part.byte_length <= 0) {
    return { ok: false, code: 'invalid-image', message: 'Inline picture does not exact-join one preserved content-addressed raster part' }
  }
  const relPartName = relationshipPart(drawing.anchor.part_name)
  const relPart = uniquePreservedPart(document, relPartName)
  if (!relPart || asciiLower(relPart.content_type) !== 'application/vnd.openxmlformats-package.relationships+xml' || !SHA256.test(relPart.sha256)) return { ok: false, code: 'invalid-image', message: 'Inline picture owning relationship part is not uniquely preserved and digest-bound' }
  const width = emuToMilliPoints(drawing.width_emu)
  const height = emuToMilliPoints(drawing.height_emu)
  if (width === undefined || height === undefined) return { ok: false, code: 'unsupported-image', message: 'Picture EMU extent is not exactly representable in integer milli-points within the geometry bound' }
  return {
    ok: true,
    value: {
      drawing_id: drawing.id,
      run_id: runID,
      asset_id: imageAssetID(part.sha256, part.part_name),
      part_name: part.part_name,
      relationship_id: drawing.relationship_id,
      relationship_part: relPart.part_name,
      relationship_sha256: relPart.sha256 as `sha256:${string}`,
      content_type: contentType,
      content_digest: part.sha256 as `sha256:${string}`,
      byte_length: part.byte_length,
      width_emu: drawing.width_emu,
      height_emu: drawing.height_emu,
      width_millipoints: width,
      height_millipoints: height,
      source_crop: { left: 0, top: 0, right: 0, bottom: 0, unit: 'one-hundred-thousandth' },
      transform: { rotation_degrees: 0, flip_horizontal: false, flip_vertical: false },
    },
  }
}

export function collectNativeDocxQualifiedInlineImagesV1(document: NativeDocxDocumentV1): NativeDocxInlineImageQualificationV1[] {
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  return stories.flatMap((story) => story.blocks.flatMap((block) => {
    const paragraphs = block.paragraph ? [block.paragraph] : block.table ? block.table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) : []
    return paragraphs.flatMap((paragraph) => paragraph.runs.flatMap((run) => run.drawing ? [qualifyNativeDocxInlineImageV1(document, run.id, run.drawing)] : []))
  }))
}

export function prepareNativeDocxPagePaintMediaAssetsV1(document: NativeDocxDocumentV1, values: readonly NativeDocxAuthoritativeMediaAssetV1[]): NativeDocxPagePaintMediaAssetV1[] {
  if (!Array.isArray(values) || values.length > DOCX_INLINE_IMAGE_LIMITS.maxAssets) throw new RangeError(`authoritative media assets exceed ${DOCX_INLINE_IMAGE_LIMITS.maxAssets} entries`)
  const qualified = collectNativeDocxQualifiedInlineImagesV1(document).flatMap((entry) => entry.ok ? [entry.value] : [])
  const required = new Map(qualified.map((entry) => [canonicalPart(entry.part_name), entry]))
  if (values.length !== required.size) throw new TypeError('authoritative media assets must exactly cover every qualified unique inline picture part')
  const output: NativeDocxPagePaintMediaAssetV1[] = []
  const seen = new Set<string>()
  let total = 0
  for (const value of values) {
    if (!value || typeof value !== 'object' || typeof value.part_name !== 'string' || typeof value.content_type !== 'string' || !(value.bytes instanceof Uint8Array) || !SHA256.test(value.content_digest)) throw new TypeError('authoritative media asset is malformed')
    const key = canonicalPart(value.part_name)
    const image = required.get(key)
    if (!image || seen.has(key) || value.part_name !== image.part_name || asciiLower(value.content_type) !== image.content_type || value.content_digest !== image.content_digest || value.bytes.byteLength !== image.byte_length) throw new TypeError('authoritative media asset does not exact-join one qualified native picture')
    if (value.bytes.byteLength === 0 || value.bytes.byteLength > DOCX_INLINE_IMAGE_LIMITS.maxAssetBytes || total + value.bytes.byteLength > DOCX_INLINE_IMAGE_LIMITS.maxTotalBytes) throw new RangeError('authoritative media bytes exceed the bounded page-paint budget')
    const owned = Uint8Array.from(value.bytes)
    if (digest(owned) !== value.content_digest) throw new TypeError('authoritative media bytes do not match their content digest')
    const dimensions = image.content_type === 'image/png' ? pngDimensions(owned) : nativeBaselineJpegDimensions(owned)
    if (!dimensions) throw new TypeError('authoritative media is not a supported static, structurally complete PNG or baseline JFIF JPEG')
    if (dimensions.width > DOCX_INLINE_IMAGE_LIMITS.maxPixelDimension || dimensions.height > DOCX_INLINE_IMAGE_LIMITS.maxPixelDimension || dimensions.width * dimensions.height > DOCX_INLINE_IMAGE_LIMITS.maxPixels) throw new RangeError('authoritative raster dimensions exceed the bounded decode budget')
    output.push({
      id: image.asset_id,
      part_name: image.part_name,
      content_type: image.content_type,
      content_digest: image.content_digest,
      byte_length: owned.byteLength,
      width_px: dimensions.width,
      height_px: dimensions.height,
      bytes_base64: base64(owned),
    })
    seen.add(key)
    total += owned.byteLength
  }
  return output.sort((left, right) => compareNativeCodeUnits(canonicalPart(left.part_name), canonicalPart(right.part_name)))
}

export function decodeNativeDocxPagePaintMediaAssetsV1(document: NativeDocxDocumentV1, value: unknown): NativeDocxPagePaintMediaAssetV1[] {
  const resources = decodeNativeDocxPagePaintResourceListV1(value)
  const authoritative: NativeDocxAuthoritativeMediaAssetV1[] = resources.map((asset) => ({
    part_name: asset.part_name,
    content_type: asset.content_type,
    content_digest: asset.content_digest,
    bytes: unbase64(asset.bytes_base64)!,
  }))
  const canonical = prepareNativeDocxPagePaintMediaAssetsV1(document, authoritative)
  if (JSON.stringify(canonical) !== JSON.stringify(value)) throw new TypeError('page-paint media asset inventory is not the canonical document-bound projection')
  return canonical
}

export function decodeNativeDocxPagePaintResourceListV1(value: unknown): NativeDocxPagePaintMediaAssetV1[] {
  if (!Array.isArray(value) || value.length > DOCX_INLINE_IMAGE_LIMITS.maxAssets) throw new TypeError('page-paint media asset inventory is malformed or unbounded')
  let total = 0
  const seenParts = new Set<string>()
  const seenIDs = new Set<string>()
  const resources: NativeDocxPagePaintMediaAssetV1[] = value.map((entry) => {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) throw new TypeError('page-paint media asset is not an object')
    const asset = entry as Record<string, unknown>
    const fields = ['id', 'part_name', 'content_type', 'content_digest', 'byte_length', 'width_px', 'height_px', 'bytes_base64']
    if (Object.keys(asset).sort().join(',') !== fields.sort().join(',')) throw new TypeError('page-paint media asset contains unknown or missing fields')
    if (typeof asset.part_name !== 'string' || !validPartName(asset.part_name) || typeof asset.content_type !== 'string' || typeof asset.content_digest !== 'string' || typeof asset.bytes_base64 !== 'string' || asset.bytes_base64.length > DOCX_INLINE_IMAGE_LIMITS.maxAssetBytes * 2) throw new TypeError('page-paint media asset identity is malformed')
    const bytes = unbase64(asset.bytes_base64)
    if (!bytes || asset.byte_length !== bytes.byteLength) throw new TypeError('page-paint media asset bytes are not canonical or length-bound')
    if (!['image/png', 'image/jpeg'].includes(asset.content_type) || !SHA256.test(asset.content_digest) || digest(bytes) !== asset.content_digest || asset.id !== imageAssetID(asset.content_digest, asset.part_name)) throw new TypeError('page-paint media asset content identity is invalid')
    const partKey = canonicalPart(asset.part_name)
    if (seenParts.has(partKey) || seenIDs.has(asset.id as string)) throw new TypeError('page-paint media asset identity is duplicated')
    const dimensions = asset.content_type === 'image/png' ? pngDimensions(bytes) : nativeBaselineJpegDimensions(bytes)
    if (!dimensions || asset.width_px !== dimensions.width || asset.height_px !== dimensions.height) throw new TypeError('page-paint media asset dimensions do not match its raster bytes')
    if (bytes.byteLength > DOCX_INLINE_IMAGE_LIMITS.maxAssetBytes || total + bytes.byteLength > DOCX_INLINE_IMAGE_LIMITS.maxTotalBytes || dimensions.width > DOCX_INLINE_IMAGE_LIMITS.maxPixelDimension || dimensions.height > DOCX_INLINE_IMAGE_LIMITS.maxPixelDimension || dimensions.width * dimensions.height > DOCX_INLINE_IMAGE_LIMITS.maxPixels) throw new RangeError('page-paint media resource exceeds its byte or pixel budget')
    total += bytes.byteLength
    seenParts.add(partKey)
    seenIDs.add(asset.id as string)
    return asset as unknown as NativeDocxPagePaintMediaAssetV1
  })
  const sorted = [...resources].sort((left, right) => compareNativeCodeUnits(canonicalPart(left.part_name), canonicalPart(right.part_name)))
  if (JSON.stringify(sorted) !== JSON.stringify(resources)) throw new TypeError('page-paint media assets must be in canonical part-name order')
  return resources
}
