import { isProjectedNativeWorkbookV2 } from './nativeRenderModelV2.js'
import type { NativeWorkbookRenderModelV2 } from './nativeRenderModelV2.js'
import { NativeSheetGeometryV2Error } from './nativeSheetGeometryV2.js'
import type { NativeMaximumDigitWidthAuthorityV2 } from './nativeSheetGeometryV2.js'
import { sha256Hex, sha256HexBytes } from './nativeSha256.js'

export const NATIVE_XLSX_MDW_PROVIDER_ID = 'injoffice.sfnt-maximum-digit-width'
export const NATIVE_XLSX_MDW_PROVIDER_REVISION = `sha256:${sha256Hex('injoffice.sfnt-maximum-digit-width.v1\0fixed-truetype;unicode-cmap-4-12;integer-css-pixel-round;96dpi')}`
const maximumFontBytes = 64 * 1024 * 1024
const metricAuthorities = new WeakSet<object>()

export function createNativeMaximumDigitWidthAuthorityV2(workbook: NativeWorkbookRenderModelV2, fontBytes: Uint8Array): NativeMaximumDigitWidthAuthorityV2 {
  if (!isProjectedNativeWorkbookV2(workbook)) throw metricError('$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV2')
  const normal = workbook.normal_style
  if (!normal) throw metricError('$.normal_style', 'the built-in Normal-style font identity is unavailable')
  if (!(fontBytes instanceof Uint8Array) || Object.getPrototypeOf(fontBytes) !== Uint8Array.prototype || fontBytes.byteLength < 12 || fontBytes.byteLength > maximumFontBytes) throw metricError('$.font_bytes', 'font bytes must be a bounded direct Uint8Array')
  let bytes: Uint8Array
  try { bytes = fontBytes.slice() } catch { throw metricError('$.font_bytes', 'Proxy font bytes are refused') }
  const parsed = parseSfntDigitMetrics(bytes)
  const normalizedExpected = normalizeFontName(normal.font_name)
  if (!parsed.names.some((name) => normalizeFontName(name) === normalizedExpected)) throw metricError('$.font_bytes', 'font name table does not match the projected Normal font')
  if (parsed.bold !== normal.font_bold || parsed.italic !== normal.font_italic) throw metricError('$.font_bytes', 'font style bits do not match the projected Normal font')
  const maximumDigitWidthPixels = Math.round(parsed.maximumAdvance * normal.font_size_points * 96 / (72 * parsed.unitsPerEm))
  if (!Number.isSafeInteger(maximumDigitWidthPixels) || maximumDigitWidthPixels < 1 || maximumDigitWidthPixels > 512) throw metricError('$.font_bytes', 'computed maximum digit width is outside the qualified integer-pixel range')
  const authority: NativeMaximumDigitWidthAuthorityV2 = {
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    normal_style_xf_id: normal.style_xf_id,
    normal_style_font_id: normal.font_id,
    font_name: normal.font_name,
    font_size_points: normal.font_size_points,
    font_bold: normal.font_bold,
    font_italic: normal.font_italic,
    normal_font_record_sha256: normal.font_record_sha256 as `sha256:${string}`,
    font_sha256: `sha256:${sha256HexBytes(bytes)}`,
    provider_id: NATIVE_XLSX_MDW_PROVIDER_ID,
    provider_revision: NATIVE_XLSX_MDW_PROVIDER_REVISION,
    measurement_dpi: 96,
    maximum_digit_width_pixels: maximumDigitWidthPixels,
  }
  const frozen = Object.freeze(authority)
  metricAuthorities.add(frozen)
  return frozen
}

export function isNativeMaximumDigitWidthAuthorityV2(value: unknown): value is NativeMaximumDigitWidthAuthorityV2 {
  return typeof value === 'object' && value !== null && metricAuthorities.has(value)
}

type Table = { offset: number; length: number; checksum: number }

function parseSfntDigitMetrics(bytes: Uint8Array): { unitsPerEm: number; maximumAdvance: number; names: string[]; bold: boolean; italic: boolean } {
  if (u32(bytes, 0) !== 0x00010000 && u32(bytes, 0) !== 0x74727565) throw metricError('$.font_bytes', 'only standalone fixed TrueType sfnt fonts are qualified')
  const count = u16(bytes, 4)
  if (count < 1 || count > 64 || !range(bytes, 12, count * 16)) throw metricError('$.font_bytes', 'sfnt table directory is malformed or oversized')
  const tables = new Map<string, Table>()
  for (let index = 0; index < count; index++) {
    const at = 12 + index * 16, name = ascii(bytes, at, 4), checksum = u32(bytes, at + 4), offset = u32(bytes, at + 8), length = u32(bytes, at + 12)
    if (tables.has(name) || length < 1 || offset % 4 !== 0 || !range(bytes, offset, length)) throw metricError('$.font_bytes', 'sfnt table records are duplicate, misaligned, or out of range')
    const table = { offset, length, checksum }
    if (tableChecksum(bytes, name, table) !== checksum) throw metricError('$.font_bytes', `sfnt ${name} checksum is invalid`)
    tables.set(name, table)
  }
  const head = requiredTable(tables, 'head', 54), hhea = requiredTable(tables, 'hhea', 36), hmtx = requiredTable(tables, 'hmtx', 4), maxp = requiredTable(tables, 'maxp', 6), cmap = requiredTable(tables, 'cmap', 12), name = requiredTable(tables, 'name', 6)
  if (u32(bytes, head.offset + 12) !== 0x5f0f3cf5) throw metricError('$.font_bytes', 'sfnt head magic is invalid')
  const unitsPerEm = u16(bytes, head.offset + 18), glyphCount = u16(bytes, maxp.offset + 4), metricCount = u16(bytes, hhea.offset + 34)
  if (unitsPerEm < 16 || unitsPerEm > 16_384 || glyphCount < 1 || metricCount < 1 || metricCount > glyphCount || hmtx.length < metricCount * 4 + (glyphCount - metricCount) * 2) throw metricError('$.font_bytes', 'sfnt horizontal metrics are inconsistent')
  const advances: number[] = []
  for (let codePoint = 0x30; codePoint <= 0x39; codePoint++) {
    const glyph = cmapGlyph(bytes, cmap, codePoint)
    if (glyph === undefined || glyph >= glyphCount) throw metricError('$.font_bytes', 'font lacks a canonical Unicode decimal digit glyph')
    const metricIndex = Math.min(glyph, metricCount - 1)
    advances.push(u16(bytes, hmtx.offset + metricIndex * 4))
  }
  const names = fontNames(bytes, name)
  if (names.length === 0) throw metricError('$.font_bytes', 'font has no qualified family name')
  const os2 = tables.get('OS/2'), selection = os2 && os2.length >= 64 ? u16(bytes, os2.offset + 62) : undefined
  const macStyle = u16(bytes, head.offset + 44)
  return { unitsPerEm, maximumAdvance: Math.max(...advances), names, bold: selection === undefined ? (macStyle & 1) !== 0 : (selection & 0x20) !== 0, italic: selection === undefined ? (macStyle & 2) !== 0 : (selection & 1) !== 0 }
}

function cmapGlyph(bytes: Uint8Array, cmap: Table, codePoint: number): number | undefined {
  const count = u16(bytes, cmap.offset + 2)
  if (count < 1 || count > 64 || cmap.length < 4 + count * 8) throw metricError('$.font_bytes', 'cmap encoding inventory is malformed')
  for (let index = 0; index < count; index++) {
    const record = cmap.offset + 4 + index * 8, platform = u16(bytes, record), encoding = u16(bytes, record + 2), relative = u32(bytes, record + 4)
    if (relative > cmap.length - 2 || !(platform === 0 || platform === 3 && (encoding === 1 || encoding === 10))) continue
    const start = cmap.offset + relative, format = u16(bytes, start)
    if (format === 12) {
      if (!range(bytes, start, 16) || u32(bytes, start + 4) > cmap.length - relative) continue
      const groups = u32(bytes, start + 12)
      if (groups > 1_000_000 || !range(bytes, start + 16, groups * 12)) continue
      let low = 0, high = groups - 1
      while (low <= high) { const middle = (low + high) >>> 1, at = start + 16 + middle * 12, first = u32(bytes, at), last = u32(bytes, at + 4); if (codePoint < first) high = middle - 1; else if (codePoint > last) low = middle + 1; else return u32(bytes, at + 8) + codePoint - first }
    } else if (format === 4) {
      const length = u16(bytes, start + 2), segmentCount = u16(bytes, start + 6) / 2
      if (segmentCount < 1 || !Number.isInteger(segmentCount) || length > cmap.length - relative || length < 16 + segmentCount * 8) continue
      const endCodes = start + 14, startCodes = endCodes + segmentCount * 2 + 2, deltas = startCodes + segmentCount * 2, offsets = deltas + segmentCount * 2
      for (let segment = 0; segment < segmentCount; segment++) if (codePoint <= u16(bytes, endCodes + segment * 2) && codePoint >= u16(bytes, startCodes + segment * 2)) {
        const delta = i16(bytes, deltas + segment * 2), offset = u16(bytes, offsets + segment * 2)
        if (offset === 0) return (codePoint + delta) & 0xffff
        const glyphAt = offsets + segment * 2 + offset + (codePoint - u16(bytes, startCodes + segment * 2)) * 2
        if (!range(bytes, glyphAt, 2)) return undefined
        const glyph = u16(bytes, glyphAt); return glyph === 0 ? 0 : (glyph + delta) & 0xffff
      }
    }
  }
  return undefined
}

function fontNames(bytes: Uint8Array, table: Table): string[] {
  const count = u16(bytes, table.offset + 2), strings = u16(bytes, table.offset + 4), result = new Set<string>()
  if (count > 16_384 || table.length < 6 + count * 12 || strings > table.length) return []
  for (let index = 0; index < count; index++) {
    const at = table.offset + 6 + index * 12, platform = u16(bytes, at), nameID = u16(bytes, at + 6), length = u16(bytes, at + 8), offset = u16(bytes, at + 10)
    if (![1, 4, 6, 16].includes(nameID) || offset + length > table.length - strings) continue
    const valueAt = table.offset + strings + offset
    let value = ''
    if (platform === 0 || platform === 3) { if (length % 2 !== 0) continue; for (let unit = 0; unit < length; unit += 2) value += String.fromCharCode(u16(bytes, valueAt + unit)) }
    else if (platform === 1) value = ascii(bytes, valueAt, length)
    if (value.length > 0 && value.length <= 255) result.add(value)
  }
  return [...result]
}

function requiredTable(tables: Map<string, Table>, name: string, minimum: number): Table { const value = tables.get(name); if (!value || value.length < minimum) throw metricError('$.font_bytes', `required sfnt ${name} table is missing or truncated`); return value }
function tableChecksum(bytes: Uint8Array, name: string, table: Table): number { let sum = 0; for (let offset = 0; offset < table.length; offset += 4) { let word = 0; for (let byte = 0; byte < 4; byte++) { const index = offset + byte, value = index < table.length && !(name === 'head' && index >= 8 && index < 12) ? bytes[table.offset + index]! : 0; word = (word * 256 + value) >>> 0 } sum = (sum + word) >>> 0 } return sum }
function normalizeFontName(value: string): string { return value.normalize('NFKC').toLowerCase().replace(/[\s_-]/g, '') }
function range(bytes: Uint8Array, offset: number, length: number): boolean { return Number.isSafeInteger(offset) && Number.isSafeInteger(length) && offset >= 0 && length >= 0 && offset <= bytes.length && length <= bytes.length - offset }
function u16(bytes: Uint8Array, offset: number): number { if (!range(bytes, offset, 2)) throw metricError('$.font_bytes', 'sfnt read exceeds font bytes'); return bytes[offset]! * 256 + bytes[offset + 1]! }
function i16(bytes: Uint8Array, offset: number): number { const value = u16(bytes, offset); return value >= 0x8000 ? value - 0x10000 : value }
function u32(bytes: Uint8Array, offset: number): number { if (!range(bytes, offset, 4)) throw metricError('$.font_bytes', 'sfnt read exceeds font bytes'); return ((bytes[offset]! * 0x1000000) + (bytes[offset + 1]! << 16) + (bytes[offset + 2]! << 8) + bytes[offset + 3]!) >>> 0 }
function ascii(bytes: Uint8Array, offset: number, length: number): string { if (!range(bytes, offset, length)) throw metricError('$.font_bytes', 'sfnt read exceeds font bytes'); let result = ''; for (let index = 0; index < length; index++) result += String.fromCharCode(bytes[offset + index]!); return result }
function metricError(path: string, message: string): NativeSheetGeometryV2Error { return new NativeSheetGeometryV2Error('geometry.metricAuthority', path, message) }
