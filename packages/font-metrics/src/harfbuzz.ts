/**
 * Canonical Node 22 HarfBuzz shaper for the qualified native page-paint slice.
 *
 * This is deliberately resolver-neutral: callers supply a resolved face and
 * its content-addressed bytes. It performs no font discovery or substitution.
 * Document-level bidi and script itemization remain upstream responsibilities;
 * this v1 provider accepts explicit, complete horizontal LTR and RTL runs.
 */

import { readFileSync } from 'node:fs'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import * as hb from 'harfbuzzjs'
import {
  MAX_TEXT_RUN_UTF16,
  scaleFontUnits,
  scaleLineMetrics,
  validateTextRunInput,
  type FontDesignMetrics,
  type FontResource,
  type NativeTextDecision,
  type NativeTextRefusal,
  type NativeTextShaper,
  type ResolvedFontFace,
  type ShapeProviderRequest,
  type ShapedCluster,
  type ShapedGlyph,
  type ShapedSegment,
} from './layout.js'
import {
  UNICODE_13_GENERATOR_SHA256,
  UNICODE_13_PROJECTION_SHA256,
  UNICODE_13_TABLES_ENCODING_SHA256,
  UNICODE_13_TABLES_RUNTIME_MATCH,
  UNICODE_13_VERSION,
  isUnicode13Control,
  isUnicode13DefaultIgnorable,
  isUnicode13WhiteSpace,
  unicode13Script,
} from './unicode13.js'

export const HARFBUZZJS_PACKAGE_VERSION = '1.6.0' as const
export const HARFBUZZ_RUNTIME_VERSION = '14.3.0' as const
export const HARFBUZZ_WASM_SHA256 = 'sha256:66f25d50cdf9942366498d0504d13096d567946cb1f6b5af6cdd87f392be99da' as const
export const HARFBUZZJS_ENTRY_SHA256 = 'sha256:04ece1914c8720ad96257728119c1bf24088ee4ef3ab2a2ca688c9a4895d4e7e' as const
export const HARFBUZZJS_LOADER_SHA256 = 'sha256:ba43463df44851fd58211ef8c19aa7956965390857727ae867b70dc7b3682de1' as const
export const HARFBUZZJS_MANIFEST_SHA256 = 'sha256:aa0da40ea2373b213631e1d73f09ddac08dfa3a6f8f5b52be97797c902f443ab' as const
export const HARFBUZZ_UNICODE_DATA_VERSION = UNICODE_13_VERSION
export const HARFBUZZ_SHAPER_CONFIG_REVISION = 'injoffice.hb-horizontal-bidi-unicode13-tables.v3' as const

export const HARFBUZZ_SHAPER_LIMITS = Object.freeze({
  maxFontBytes: 64 * 1024 * 1024,
  maxCachedFontBytes: 128 * 1024 * 1024,
  maxCachedFaces: 256,
  maxTables: 256,
  maxCollectionFaces: 256,
  maxGlyphs: MAX_TEXT_RUN_UTF16,
  maxClusters: MAX_TEXT_RUN_UTF16,
})

const SOURCE_REVISION_RE = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,63}$/
const ID_RE = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const DIGEST_RE = /^sha256:[a-f0-9]{64}$/
const SUPPORTED_SCRIPTS = Object.freeze(['Latn', 'Arab', 'Hebr', 'Zyyy'] as const)
const SUPPORTED_SCRIPT_SET = new Set<string>(SUPPORTED_SCRIPTS)
const QUALIFIED_FEATURES = Object.freeze(['kern', 'liga'] as const)
const QUALIFIED_FEATURE_SET = new Set<string>(QUALIFIED_FEATURES)
const BUFFER_FLAG_NAMES = Object.freeze(['bot', 'eot', 'produce-unsafe-to-concat'] as const)
const BUFFER_FLAGS = hb.BufferFlag.BOT | hb.BufferFlag.EOT | hb.BufferFlag.PRODUCE_UNSAFE_TO_CONCAT
const REQUIRED_TABLES = ['cmap', 'head', 'hhea', 'hmtx', 'maxp'] as const
const PROVIDER_ID = 'injoffice.harfbuzzjs'
const EXPECTED_WASM_BYTES = 421_964
const MAX_ABS_DESIGN_VALUE = 1_000_000_000
const MAX_OUTPUT_MILLIPOINTS = 1_000_000_000
const BIDI_CONTROL_RE = /^[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]$/u

function compareCodeUnits(left: string, right: string): number {
  if (left === right) return 0
  const length = Math.min(left.length, right.length)
  for (let index = 0; index < length; index += 1) {
    const difference = left.charCodeAt(index) - right.charCodeAt(index)
    if (difference !== 0) return difference
  }
  return left.length - right.length
}

export interface HarfBuzzShaperOptionsV1 {
  /** Deployment/source revision included in the immutable provider revision. */
  sourceRevision: string
}

export interface HarfBuzzShaperProvenanceV1 {
  protocol: 'injoffice.text.harfbuzz-shaper-provenance'
  version: 1
  provider_id: typeof PROVIDER_ID
  provider_revision: string
  source_revision: string
  runtime_package: `harfbuzzjs@${typeof HARFBUZZJS_PACKAGE_VERSION}`
  runtime_version: typeof HARFBUZZ_RUNTIME_VERSION
  runtime_wasm_sha256: typeof HARFBUZZ_WASM_SHA256
  runtime_entry_sha256: typeof HARFBUZZJS_ENTRY_SHA256
  runtime_loader_sha256: typeof HARFBUZZJS_LOADER_SHA256
  runtime_manifest_sha256: typeof HARFBUZZJS_MANIFEST_SHA256
  unicode_data_version: string
  unicode_tables_sha256: typeof UNICODE_13_TABLES_ENCODING_SHA256
  unicode_projection_sha256: typeof UNICODE_13_PROJECTION_SHA256
  unicode_generator_sha256: typeof UNICODE_13_GENERATOR_SHA256
  configuration_revision: typeof HARFBUZZ_SHAPER_CONFIG_REVISION
  direction_policy: 'explicit-complete-horizontal-ltr-rtl'
  script_policy: readonly string[]
  cluster_level: 'monotone-graphemes'
  buffer_flags: readonly ['bot', 'eot', 'produce-unsafe-to-concat']
  default_feature_policy: 'harfbuzz-14.3.0-shape-defaults'
  feature_policy: 'explicit-qualified-and-font-advertised-kern-liga-only'
  variation_policy: 'refuse'
  spacing_policy: 'zero-only'
  font_policy: 'bounded-fixed-truetype-sfnt-or-ttc-preflight-v1'
  scale_policy: 'per-glyph-scaleFontUnits-then-sum-v1'
}

export interface HarfBuzzTextShaperV1 extends NativeTextShaper {
  readonly provenance: HarfBuzzShaperProvenanceV1
  shape(request: ShapeProviderRequest): ShapedSegment | NativeTextRefusal
}

export interface HarfBuzzFontMetricsRequestV1 {
  /** Immutable standalone sfnt bytes, or the complete TTC/OTC collection. */
  bytes: Uint8Array
  /** Content address supplied by the authoritative resolver. */
  contentDigest: `sha256:${string}`
  /** Required for a collection and forbidden for a standalone sfnt. */
  collectionIndex?: number
}

interface TableRecord {
  checksum: number
  offset: number
  length: number
}

interface SfntPreflight {
  collectionIndex: number
  unitsPerEm: number
  numGlyphs: number
  tables: ReadonlyMap<string, TableRecord>
  requiredMetrics: Pick<FontDesignMetrics, 'unitsPerEm' | 'ascender' | 'descender' | 'lineGap'>
  optionalMetrics: Pick<FontDesignMetrics, 'capHeight' | 'xHeight' | 'underlinePosition' | 'underlineThickness'>
}

interface CachedFace {
  blob: hb.Blob
  face: hb.Face
  font: hb.Font
  preflight: SfntPreflight
  bytes: number
  features: ReadonlySet<string>
}

const canonicalHarfBuzzShapers = new WeakSet<object>()

function digestBytes(bytes: Uint8Array): `sha256:${string}` {
  return `sha256:${bytesToHex(sha256(bytes))}`
}

function refusal(code: NativeTextDecision['code'], message: string, faceId?: string, startUtf16?: number, endUtf16?: number): NativeTextRefusal {
  const decision: NativeTextDecision = Object.freeze({
    code,
    message,
    recoverable: false,
    ...(faceId !== undefined ? { faceId } : {}),
    ...(startUtf16 !== undefined && endUtf16 !== undefined ? { startUtf16, endUtf16 } : {}),
  })
  return Object.freeze({ status: 'refused', decisions: Object.freeze([decision]), attemptedFaceIds: Object.freeze(faceId === undefined ? [] : [faceId]) })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function exactKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const allowed = new Set(keys)
  return Object.keys(value).every((key) => allowed.has(key))
}

function safeInteger(value: unknown, min: number, max: number): value is number {
  return Number.isSafeInteger(value) && !Object.is(value, -0) && (value as number) >= min && (value as number) <= max
}

function validResolvedFace(face: unknown): face is ResolvedFontFace {
  if (!isRecord(face) || !exactKeys(face, ['faceId', 'family', 'postscriptName', 'weight', 'style', 'stretch', 'sourceKind', 'resourceId', 'contentDigest', 'collectionIndex', 'resolution', 'matchedFamily', 'fallbackChainId'])) return false
  return typeof face.faceId === 'string' && ID_RE.test(face.faceId)
    && typeof face.family === 'string' && face.family.length > 0 && face.family.length <= 1_024
    && (face.postscriptName === undefined || (typeof face.postscriptName === 'string' && face.postscriptName.length > 0 && face.postscriptName.length <= 1_024))
    && safeInteger(face.weight, 1, 1_000)
    && ['normal', 'italic', 'oblique'].includes(face.style as string)
    && safeInteger(face.stretch, 50, 200)
    && ['document', 'bundled', 'system', 'host'].includes(face.sourceKind as string)
    && typeof face.resourceId === 'string' && ID_RE.test(face.resourceId)
    && typeof face.contentDigest === 'string' && DIGEST_RE.test(face.contentDigest)
    && (face.collectionIndex === undefined || safeInteger(face.collectionIndex, 0, 65_535))
    && ['exact', 'substitute', 'fallback'].includes(face.resolution as string)
    && typeof face.matchedFamily === 'string' && face.matchedFamily.length > 0 && face.matchedFamily.length <= 1_024
    && (face.fallbackChainId === undefined || (typeof face.fallbackChainId === 'string' && ID_RE.test(face.fallbackChainId)))
}

function validDesignMetrics(metrics: unknown): metrics is FontDesignMetrics {
  if (!isRecord(metrics) || !exactKeys(metrics, ['unitsPerEm', 'ascender', 'descender', 'lineGap', 'capHeight', 'xHeight', 'underlinePosition', 'underlineThickness'])) return false
  return safeInteger(metrics.unitsPerEm, 16, 16_384)
    && safeInteger(metrics.ascender, 0, MAX_ABS_DESIGN_VALUE)
    && safeInteger(metrics.descender, -MAX_ABS_DESIGN_VALUE, 0)
    && safeInteger(metrics.lineGap, 0, MAX_ABS_DESIGN_VALUE)
    && [metrics.capHeight, metrics.xHeight].every((value) => value === undefined || safeInteger(value, 0, MAX_ABS_DESIGN_VALUE))
    && (metrics.underlinePosition === undefined || safeInteger(metrics.underlinePosition, -MAX_ABS_DESIGN_VALUE, MAX_ABS_DESIGN_VALUE))
    && (metrics.underlineThickness === undefined || safeInteger(metrics.underlineThickness, 0, MAX_ABS_DESIGN_VALUE))
}

function u16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! << 8) | bytes[offset + 1]!
}

function i16(bytes: Uint8Array, offset: number): number {
  const value = u16(bytes, offset)
  return value >= 0x8000 ? value - 0x1_0000 : value
}

function u32(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset]! * 0x1_000000) + (bytes[offset + 1]! << 16) + (bytes[offset + 2]! << 8) + bytes[offset + 3]!) >>> 0
}

function tag(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(bytes[offset]!, bytes[offset + 1]!, bytes[offset + 2]!, bytes[offset + 3]!)
}

function checkedRange(bytes: Uint8Array, offset: number, length: number): boolean {
  return Number.isSafeInteger(offset) && Number.isSafeInteger(length) && offset >= 0 && length >= 0 && offset <= bytes.byteLength && length <= bytes.byteLength - offset
}

function tableChecksum(bytes: Uint8Array, tableTag: string, record: TableRecord): number {
  let checksum = 0
  for (let relative = 0; relative < record.length; relative += 4) {
    let word = 0
    for (let byte = 0; byte < 4; byte++) {
      const tableIndex = relative + byte
      const value = tableIndex < record.length && !(tableTag === 'head' && tableIndex >= 8 && tableIndex < 12)
        ? bytes[record.offset + tableIndex]!
        : 0
      word = (word * 256 + value) >>> 0
    }
    checksum = (checksum + word) >>> 0
  }
  return checksum
}

function preflightSfnt(bytes: Uint8Array, requestedCollectionIndex: number | undefined): SfntPreflight | string {
  if (bytes.byteLength < 12) return 'font bytes are shorter than an sfnt header'
  const flavor = u32(bytes, 0)
  let collectionIndex = 0
  let sfntOffset = 0
  if (flavor === 0x74746366) { // ttcf
    if (bytes.byteLength < 16) return 'font collection header is truncated'
    const version = u32(bytes, 4)
    if (version !== 0x00010000 && version !== 0x00020000) return 'font collection version is unsupported'
    const faceCount = u32(bytes, 8)
    if (faceCount < 1 || faceCount > HARFBUZZ_SHAPER_LIMITS.maxCollectionFaces || !checkedRange(bytes, 12, faceCount * 4)) return 'font collection face table is invalid or exceeds limits'
    if (requestedCollectionIndex === undefined) return 'collectionIndex is required for TTC/OTC bytes'
    if (requestedCollectionIndex >= faceCount) return 'collectionIndex does not identify a face in the font collection'
    collectionIndex = requestedCollectionIndex
    sfntOffset = u32(bytes, 12 + collectionIndex * 4)
  } else {
    if (requestedCollectionIndex !== undefined) return 'collectionIndex must be omitted for standalone sfnt bytes'
  }
  if (!checkedRange(bytes, sfntOffset, 12)) return 'sfnt offset table is outside the font bytes'
  const sfntFlavor = u32(bytes, sfntOffset)
  if (sfntFlavor !== 0x00010000 && sfntFlavor !== 0x74727565) return 'only fixed TrueType glyf/loca sfnt fonts are qualified'
  const numTables = u16(bytes, sfntOffset + 4)
  if (numTables < 1 || numTables > HARFBUZZ_SHAPER_LIMITS.maxTables || !checkedRange(bytes, sfntOffset + 12, numTables * 16)) return 'sfnt table directory is invalid or exceeds limits'
  const tables = new Map<string, TableRecord>()
  for (let index = 0; index < numTables; index++) {
    const recordOffset = sfntOffset + 12 + index * 16
    const tableTag = tag(bytes, recordOffset)
    const checksum = u32(bytes, recordOffset + 4)
    const offset = u32(bytes, recordOffset + 8)
    const length = u32(bytes, recordOffset + 12)
    if (tables.has(tableTag)) return `sfnt table ${JSON.stringify(tableTag)} is duplicated`
    if (offset % 4 !== 0) return `sfnt table ${JSON.stringify(tableTag)} is not four-byte aligned`
    if (length === 0 || !checkedRange(bytes, offset, length)) return `sfnt table ${JSON.stringify(tableTag)} is empty or outside the font bytes`
    tables.set(tableTag, { checksum, offset, length })
  }
  const directoryEnd = sfntOffset + 12 + numTables * 16
  for (const [tableTag, record] of tables) {
    if (record.offset < directoryEnd && record.offset + record.length > sfntOffset) return `sfnt table ${JSON.stringify(tableTag)} overlaps its face directory`
    if (tableChecksum(bytes, tableTag, record) !== record.checksum) return `sfnt table ${JSON.stringify(tableTag)} checksum is invalid`
  }
  const orderedTables = [...tables.entries()].sort((left, right) => left[1].offset - right[1].offset || compareCodeUnits(left[0], right[0]))
  for (let index = 1; index < orderedTables.length; index++) {
    const previous = orderedTables[index - 1]![1]
    const current = orderedTables[index]![1]
    if (current.offset < previous.offset + previous.length) return 'sfnt tables overlap'
  }
  for (const required of REQUIRED_TABLES) if (!tables.has(required)) return `required sfnt table ${required} is missing`
  const hasGlyf = tables.has('glyf') && tables.has('loca')
  if (!hasGlyf || tables.has('CFF ') || tables.has('CFF2')) return 'font does not contain the qualified TrueType glyf/loca outline flavor'
  for (const variationTable of ['avar', 'cvar', 'fvar', 'gvar', 'HVAR', 'MVAR', 'STAT', 'VVAR']) {
    if (tables.has(variationTable)) return `variable-font table ${variationTable} is outside the fixed-font v1 contract`
  }
  const head = tables.get('head')!
  const hhea = tables.get('hhea')!
  const hmtx = tables.get('hmtx')!
  const maxp = tables.get('maxp')!
  const cmap = tables.get('cmap')!
  if (head.length < 54 || hhea.length < 36 || maxp.length < 6) return 'required sfnt metric table is truncated'
  if (u32(bytes, head.offset + 12) !== 0x5f0f3cf5) return 'sfnt head magic number is invalid'
  const unitsPerEm = u16(bytes, head.offset + 18)
  const numGlyphs = u16(bytes, maxp.offset + 4)
  if (unitsPerEm < 16 || unitsPerEm > 16_384 || numGlyphs === 0) return 'font unitsPerEm or glyph count is invalid'
  const numberOfHMetrics = u16(bytes, hhea.offset + 34)
  const requiredHmtxLength = numberOfHMetrics * 4 + (numGlyphs - numberOfHMetrics) * 2
  if (numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs || hmtx.length < requiredHmtxLength) return 'sfnt horizontal metrics table is inconsistent with the glyph count'
  if (cmap.length < 12 || u16(bytes, cmap.offset) !== 0) return 'sfnt cmap header is invalid'
  const cmapCount = u16(bytes, cmap.offset + 2)
  if (cmapCount < 1 || cmapCount > HARFBUZZ_SHAPER_LIMITS.maxTables || cmap.length < 4 + cmapCount * 8) return 'sfnt cmap encoding table is invalid or exceeds limits'
  let hasQualifiedCmap = false
  for (let index = 0; index < cmapCount; index++) {
    const encodingRecord = cmap.offset + 4 + index * 8
    const platformId = u16(bytes, encodingRecord)
    const encodingId = u16(bytes, encodingRecord + 2)
    const subtableOffset = u32(bytes, encodingRecord + 4)
    if (subtableOffset > cmap.length - 2) return 'sfnt cmap subtable points outside the cmap table'
    const absoluteSubtable = cmap.offset + subtableOffset
    const format = u16(bytes, absoluteSubtable)
    const isUnicodeEncoding = platformId === 0 || (platformId === 3 && (encodingId === 1 || encodingId === 10))
    if (!isUnicodeEncoding || (format !== 4 && format !== 12)) continue
    if (format === 4) {
      if (subtableOffset > cmap.length - 14) return 'sfnt cmap format 4 header is truncated'
      const length = u16(bytes, absoluteSubtable + 2)
      const segmentCountX2 = u16(bytes, absoluteSubtable + 6)
      if (length < 24 || length > cmap.length - subtableOffset || segmentCountX2 < 2 || segmentCountX2 % 2 !== 0 || length < 16 + segmentCountX2 * 4) return 'sfnt cmap format 4 subtable is malformed'
      hasQualifiedCmap = true
    } else {
      if (subtableOffset > cmap.length - 16) return 'sfnt cmap format 12 header is truncated'
      const length = u32(bytes, absoluteSubtable + 4)
      const groupCount = u32(bytes, absoluteSubtable + 12)
      if (length < 16 || length > cmap.length - subtableOffset || groupCount > MAX_TEXT_RUN_UTF16 || groupCount > Math.floor((length - 16) / 12)) return 'sfnt cmap format 12 subtable is malformed'
      let previousEnd = -1
      for (let group = 0; group < groupCount; group++) {
        const groupOffset = absoluteSubtable + 16 + group * 12
        const start = u32(bytes, groupOffset)
        const end = u32(bytes, groupOffset + 4)
        const startGlyph = u32(bytes, groupOffset + 8)
        if (start > end || start <= previousEnd || end > 0x10ffff || startGlyph >= numGlyphs || end - start > numGlyphs - 1 - startGlyph) return 'sfnt cmap format 12 groups are invalid'
        previousEnd = end
      }
      hasQualifiedCmap = true
    }
  }
  if (!hasQualifiedCmap) return 'font has no qualified Unicode cmap format 4 or 12 subtable'
  if (hasGlyf) {
    const loca = tables.get('loca')!
    const glyf = tables.get('glyf')!
    const indexToLocFormat = i16(bytes, head.offset + 50)
    if (![0, 1].includes(indexToLocFormat)) return 'sfnt loca offset format is invalid'
    const requiredLocaLength = (numGlyphs + 1) * (indexToLocFormat === 0 ? 2 : 4)
    if (loca.length < requiredLocaLength) return 'sfnt loca table is inconsistent with the glyph count'
    let previousGlyphOffset = 0
    for (let glyphIndex = 0; glyphIndex <= numGlyphs; glyphIndex++) {
      const glyphOffset = indexToLocFormat === 0 ? u16(bytes, loca.offset + glyphIndex * 2) * 2 : u32(bytes, loca.offset + glyphIndex * 4)
      if (glyphOffset < previousGlyphOffset || glyphOffset > glyf.length) return 'sfnt loca offsets are descending or outside the glyf table'
      if (glyphIndex > 0 && glyphOffset !== previousGlyphOffset && glyphOffset - previousGlyphOffset < 10) return 'sfnt glyf record is shorter than its required header'
      previousGlyphOffset = glyphOffset
    }
  }
  const ascender = i16(bytes, hhea.offset + 4)
  const descender = i16(bytes, hhea.offset + 6)
  const lineGap = i16(bytes, hhea.offset + 8)
  if (ascender < 0 || descender > 0 || lineGap < 0) return 'font horizontal line metrics are outside the qualified contract'
  const os2 = tables.get('OS/2')
  const os2Version = os2 && os2.length >= 2 ? u16(bytes, os2.offset) : undefined
  const post = tables.get('post')
  return {
    collectionIndex,
    unitsPerEm,
    numGlyphs,
    tables,
    requiredMetrics: { unitsPerEm, ascender, descender, lineGap },
    optionalMetrics: {
      ...(os2 && os2Version !== undefined && os2Version >= 2 && os2.length >= 90 ? { xHeight: i16(bytes, os2.offset + 86), capHeight: i16(bytes, os2.offset + 88) } : {}),
      ...(post && post.length >= 12 ? { underlinePosition: i16(bytes, post.offset + 8), underlineThickness: i16(bytes, post.offset + 10) } : {}),
    },
  }
}

function metricsMatchFont(metrics: FontDesignMetrics, preflight: SfntPreflight): boolean {
  if (metrics.unitsPerEm !== preflight.requiredMetrics.unitsPerEm || metrics.ascender !== preflight.requiredMetrics.ascender || metrics.descender !== preflight.requiredMetrics.descender || metrics.lineGap !== preflight.requiredMetrics.lineGap) return false
  for (const key of ['capHeight', 'xHeight', 'underlinePosition', 'underlineThickness'] as const) {
    if (metrics[key] !== preflight.optionalMetrics[key]) return false
  }
  return true
}

function snapshotFace(face: ResolvedFontFace): ResolvedFontFace {
  return Object.freeze({
    faceId: face.faceId,
    family: face.family,
    ...(face.postscriptName !== undefined ? { postscriptName: face.postscriptName } : {}),
    weight: face.weight,
    style: face.style,
    stretch: face.stretch,
    sourceKind: face.sourceKind,
    resourceId: face.resourceId,
    contentDigest: face.contentDigest,
    ...(face.collectionIndex !== undefined ? { collectionIndex: face.collectionIndex } : {}),
    resolution: face.resolution,
    matchedFamily: face.matchedFamily,
    ...(face.fallbackChainId !== undefined ? { fallbackChainId: face.fallbackChainId } : {}),
  })
}

function snapshotMetrics(metrics: FontDesignMetrics): FontDesignMetrics {
  return Object.freeze({
    unitsPerEm: metrics.unitsPerEm,
    ascender: metrics.ascender,
    descender: metrics.descender,
    lineGap: metrics.lineGap,
    ...(metrics.capHeight !== undefined ? { capHeight: metrics.capHeight } : {}),
    ...(metrics.xHeight !== undefined ? { xHeight: metrics.xHeight } : {}),
    ...(metrics.underlinePosition !== undefined ? { underlinePosition: metrics.underlinePosition } : {}),
    ...(metrics.underlineThickness !== undefined ? { underlineThickness: metrics.underlineThickness } : {}),
  })
}

/**
 * Extracts the exact design metrics used by the pinned HarfBuzz shaper.
 *
 * Resolver integrations use this instead of inventing metrics or maintaining a
 * second sfnt parser. The same bounded preflight and table policy is applied
 * again when the returned resource is shaped, closing time-of-check/time-of-use
 * and fabricated-metric seams.
 */
export function inspectHarfBuzzFontMetricsV1(request: HarfBuzzFontMetricsRequestV1): FontDesignMetrics {
  if (!isRecord(request) || !exactKeys(request, ['bytes', 'contentDigest', 'collectionIndex']) || !(request.bytes instanceof Uint8Array)) throw new TypeError('font metric request must contain exact immutable bytes, digest, and optional collection index')
  if (typeof request.contentDigest !== 'string' || !DIGEST_RE.test(request.contentDigest) || digestBytes(request.bytes) !== request.contentDigest) throw new TypeError('font metric bytes do not match the authoritative content digest')
  if (request.bytes.byteLength === 0 || request.bytes.byteLength > HARFBUZZ_SHAPER_LIMITS.maxFontBytes) throw new RangeError(`font metric bytes must contain 1 through ${HARFBUZZ_SHAPER_LIMITS.maxFontBytes} bytes`)
  if (request.collectionIndex !== undefined && !safeInteger(request.collectionIndex, 0, 65_535)) throw new RangeError('font collection index must be a bounded non-negative integer')
  const ownedBytes = Uint8Array.from(request.bytes)
  const preflight = preflightSfnt(ownedBytes, request.collectionIndex)
  if (typeof preflight === 'string') throw new TypeError(preflight)
  if (digestBytes(request.bytes) !== request.contentDigest) throw new TypeError('font metric bytes changed while being inspected')
  return Object.freeze(snapshotMetrics({ ...preflight.requiredMetrics, ...preflight.optionalMetrics }))
}

export type HarfBuzzOutlineCommandV1 =
  | { kind: 'move_to' | 'line_to'; x: number; y: number }
  | { kind: 'quadratic_to'; control_x: number; control_y: number; x: number; y: number }
  | { kind: 'cubic_to'; control_1_x: number; control_1_y: number; control_2_x: number; control_2_y: number; x: number; y: number }
  | { kind: 'close_path' }

/** Content-addressed, bounded outlines from the same pinned runtime as shaping.
 * Construct once per face, not once per glyph. No platform font substitution.
 */
export function createHarfBuzzOutlineProviderV1(request: HarfBuzzFontMetricsRequestV1) {
  const bytes = Uint8Array.from(request.bytes)
  const metrics = inspectHarfBuzzFontMetricsV1({ ...request, bytes })
  const face = new hb.Face(new hb.Blob(bytes), request.collectionIndex ?? 0)
  const font = new hb.Font(face)
  font.setScale(metrics.unitsPerEm, metrics.unitsPerEm)
  const outlineScale = 2 ** Math.floor(Math.log2(1_000_000 / metrics.unitsPerEm))
  let path: HarfBuzzOutlineCommandV1[] = []
  const append = (command: HarfBuzzOutlineCommandV1) => {
    if (path.length >= 65_536) throw new RangeError('glyph outline exceeds the command budget')
    for (const value of Object.values(command)) if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > 1_000_000_000)) throw new RangeError('glyph outline has invalid coordinates')
    path.push(command)
  }
  const draw = new hb.DrawFuncs()
  draw.setMoveToFunc((x, y) => append({ kind: 'move_to', x, y }))
  draw.setLineToFunc((x, y) => append({ kind: 'line_to', x, y }))
  draw.setQuadraticToFunc((control_x, control_y, x, y) => append({ kind: 'quadratic_to', control_x, control_y, x, y }))
  draw.setCubicToFunc((control_1_x, control_1_y, control_2_x, control_2_y, x, y) => append({ kind: 'cubic_to', control_1_x, control_1_y, control_2_x, control_2_y, x, y }))
  draw.setClosePathFunc(() => append({ kind: 'close_path' }))
  return {
    outline(glyphId: number): { units_per_em: number; path: HarfBuzzOutlineCommandV1[] } {
      if (!Number.isSafeInteger(glyphId) || glyphId < 0 || glyphId > 65_535) throw new RangeError('glyph id is outside the bounded outline range')
      path = []
      if (!font.drawGlyphOrFail(glyphId, draw)) throw new TypeError('font refused the requested glyph outline')
      // TrueType implied on-curve midpoints can be half design units. Preserve
      // those exactly by scaling BOTH the integer wire grid and units-per-em;
      // rounding coordinates here would silently alter the font's contours.
      const scale = outlineScale
      const coordinates = path.flatMap((command) => Object.values(command).filter((value): value is number => typeof value === 'number'))
      if (coordinates.some((value) => !Number.isSafeInteger(value * scale) || Math.abs(value * scale) > 1_000_000_000)) throw new RangeError('glyph coordinates cannot be represented on the bounded exact integer grid')
      const normalized = path.map((command) => Object.fromEntries(Object.entries(command).map(([key, value]) => [key, typeof value === 'number' ? value * scale || 0 : value])) as HarfBuzzOutlineCommandV1)
      return { units_per_em: metrics.unitsPerEm * scale, path: normalized }
    },
  }
}

function isFixedWhitespace(text: string, start: number, end: number): boolean {
  for (let offset = start; offset < end;) {
    const codePoint = text.codePointAt(offset)!
    if (!isUnicode13WhiteSpace(codePoint)) return false
    offset += codePoint > 0xffff ? 2 : 1
  }
  return end > start
}

function isQualifiedIgnorableClusterScalar(codePoint: number): boolean {
  return codePoint === 0x200d || (codePoint >= 0xfe00 && codePoint <= 0xfe0f) || (codePoint >= 0xe0100 && codePoint <= 0xe01ef)
}

function validateQualifiedScalars(text: string, script: string): string | undefined {
  if (!UNICODE_13_TABLES_RUNTIME_MATCH) return 'pinned Unicode 13 classification tables failed their runtime digest'
  const scriptScalar = script === 'Latn' || script === 'Arab' || script === 'Hebr' ? script : undefined
  let canInherit = false
  for (const character of text) {
    const codePoint = character.codePointAt(0)!
    const actualScript = unicode13Script(codePoint)
    if ((isUnicode13DefaultIgnorable(codePoint) || isUnicode13Control(codePoint)) && !isQualifiedIgnorableClusterScalar(codePoint)) {
      return 'default-ignorable, control, line-separator, and paragraph-separator scalars are not qualified'
    }
    if (scriptScalar) {
      if (actualScript === scriptScalar) {
        canInherit = true
        continue
      }
      if (actualScript === 'Zinh') {
        if (!canInherit) return 'leading or standalone inherited marks are not qualified'
        continue
      }
      if (actualScript === 'Zyyy') {
        canInherit = false
        continue
      }
      return `a scalar does not match the declared ${script} script run`
    }
    if (script === 'Zyyy' && actualScript !== 'Zyyy') return 'a scalar classified upstream as Common has an unqualified or unknown script'
  }
  return undefined
}

function containsBidiControl(text: string): boolean {
  return [...text].some((character) => BIDI_CONTROL_RE.test(character))
}

function scaleDesignValue(value: number, unitsPerEm: number, fontSizeMilliPoints: number): number {
  const scaled = scaleFontUnits(value, unitsPerEm, fontSizeMilliPoints)
  if (Math.abs(scaled) > MAX_OUTPUT_MILLIPOINTS) throw new RangeError('scaled design value exceeds the page-paint provider bound')
  return Object.is(scaled, -0) ? 0 : scaled
}

function qualifiedLineMetrics(metrics: FontDesignMetrics, fontSizeMilliPoints: number): ReturnType<typeof scaleLineMetrics> {
  const scaled = scaleLineMetrics(metrics, fontSizeMilliPoints)
  const required = [scaled.ascentMilliPoints, scaled.descentMilliPoints, scaled.lineGapMilliPoints, scaled.lineHeightMilliPoints]
  const optional = [scaled.capHeightMilliPoints, scaled.xHeightMilliPoints, scaled.underlinePositionMilliPoints, scaled.underlineThicknessMilliPoints]
  if (required.some((value) => Math.abs(value) > MAX_OUTPUT_MILLIPOINTS)
    || optional.some((value) => value !== undefined && Math.abs(value) > MAX_OUTPUT_MILLIPOINTS)) throw new RangeError('scaled line metrics exceed the page-paint provider bound')
  return scaled
}

function canonicalFeatures(run: ShapeProviderRequest['run'], advertised: ReadonlySet<string>, faceId: string): hb.Feature[] | NativeTextRefusal {
  const features = [...(run.features ?? [])].sort((left, right) => (left.tag < right.tag ? -1 : left.tag > right.tag ? 1 : 0)
    || (left.startUtf16 ?? -1) - (right.startUtf16 ?? -1)
    || (left.endUtf16 ?? -1) - (right.endUtf16 ?? -1)
    || left.value - right.value)
  for (const feature of features) {
    if (!QUALIFIED_FEATURE_SET.has(feature.tag)) return refusal('unsupported-feature', `OpenType feature ${JSON.stringify(feature.tag)} is outside the qualified explicit feature set`, faceId, feature.startUtf16, feature.endUtf16)
    if (!advertised.has(feature.tag)) return refusal('unsupported-feature', `OpenType feature ${JSON.stringify(feature.tag)} is not advertised by this exact font face`, faceId, feature.startUtf16, feature.endUtf16)
  }
  return features.map((feature) => new hb.Feature(feature.tag, feature.value, feature.startUtf16 ?? hb.Feature.GLOBAL_START, feature.endUtf16 ?? hb.Feature.GLOBAL_END))
}

function runPolicyRefusal(run: ShapeProviderRequest['run'], faceId: string): NativeTextRefusal | undefined {
  if (run.direction !== 'ltr' && run.direction !== 'rtl') return refusal('unsupported-direction', 'canonical HarfBuzz provider v1 supports only explicit horizontal LTR/RTL runs', faceId)
  if (containsBidiControl(run.text)) return refusal('unsupported-direction', 'bidi controls require the deferred authoritative bidi/itemization stage', faceId, 0, run.text.length)
  if (!SUPPORTED_SCRIPT_SET.has(run.script)) return refusal('unsupported-script', `script ${JSON.stringify(run.script)} is outside the qualified canonical shaper v1 set`, faceId)
  const scalarIssue = validateQualifiedScalars(run.text, run.script)
  if (scalarIssue) return refusal('unsupported-script', scalarIssue, faceId, 0, run.text.length)
  if ((run.letterSpacingMilliPoints ?? 0) !== 0 || (run.wordSpacingMilliPoints ?? 0) !== 0) return refusal('unsupported-feature', 'nonzero letter/word spacing is not qualified for canonical cluster-edge placement', faceId)
  if ((run.variations?.length ?? 0) !== 0) return refusal('unsupported-feature', 'font variations require variation-qualified authoritative line metrics', faceId)
  for (const feature of run.features ?? []) {
    if (!QUALIFIED_FEATURE_SET.has(feature.tag)) return refusal('unsupported-feature', `OpenType feature ${JSON.stringify(feature.tag)} is outside the qualified explicit feature set`, faceId, feature.startUtf16, feature.endUtf16)
  }
  return undefined
}

function shapeWithCachedFace(request: ShapeProviderRequest, cached: CachedFace, buffer: hb.Buffer): ShapedSegment | NativeTextRefusal {
  const { run } = request
  const policyRefusal = runPolicyRefusal(run, request.font.face.faceId)
  if (policyRefusal) return policyRefusal
  const features = canonicalFeatures(run, cached.features, request.font.face.faceId)
  if (!Array.isArray(features)) return features

  buffer.reset()
  buffer.addText(run.text)
  buffer.setDirection(run.direction === 'rtl' ? hb.Direction.RTL : hb.Direction.LTR)
  buffer.setScript(run.script)
  buffer.setLanguage(run.language)
  buffer.setClusterLevel(hb.ClusterLevel.MONOTONE_GRAPHEMES)
  buffer.setFlags(BUFFER_FLAGS)
  cached.font.setScale(cached.preflight.unitsPerEm, cached.preflight.unitsPerEm)
  cached.font.setVariations([])
  hb.shape(cached.font, buffer, features)
  const values = buffer.getGlyphInfosAndPositions()
  if (values.length > HARFBUZZ_SHAPER_LIMITS.maxGlyphs) return refusal('provider-failure', 'HarfBuzz glyph output exceeds the bounded v1 contract', request.font.face.faceId)
  if (run.text.length === 0) {
    if (values.length !== 0) return refusal('provider-failure', 'HarfBuzz returned glyphs for an empty run', request.font.face.faceId)
    const metrics = Object.freeze(qualifiedLineMetrics(request.font.metrics, run.fontSizeMilliPoints))
    return Object.freeze({ startUtf16: 0, endUtf16: 0, face: snapshotFace(request.font.face), glyphs: Object.freeze([]), clusters: Object.freeze([]), metrics, advanceInlineMilliPoints: 0, advanceBlockMilliPoints: 0 })
  }
  if (values.length === 0 || !values.some((value) => value.cluster === 0)) return refusal('missing-glyph', 'HarfBuzz did not cover the complete UTF-16 run', request.font.face.faceId, 0, run.text.length)

  const groupStarts: number[] = []
  let previousCluster = run.direction === 'rtl' ? run.text.length : -1
  for (let index = 0; index < values.length; index++) {
    const glyph = values[index]!
    if (!safeInteger(glyph.codepoint, 0, cached.preflight.numGlyphs - 1)) return refusal('provider-failure', 'HarfBuzz returned an out-of-range glyph identifier', request.font.face.faceId)
    const descending = run.direction === 'rtl' ? glyph.cluster > previousCluster : glyph.cluster < previousCluster
    if (!safeInteger(glyph.cluster, 0, run.text.length - 1) || descending) return refusal('provider-failure', `HarfBuzz returned non-monotone or out-of-range ${run.direction === 'rtl' ? 'RTL' : 'LTR'} clusters`, request.font.face.faceId)
    if (glyph.cluster !== previousCluster) groupStarts.push(index)
    previousCluster = glyph.cluster
  }
  if (groupStarts.length > HARFBUZZ_SHAPER_LIMITS.maxClusters) return refusal('provider-failure', 'HarfBuzz cluster output exceeds the bounded v1 contract', request.font.face.faceId)

  const visualGroups = groupStarts.map((glyphStart, index) => ({ glyphStart, glyphEnd: groupStarts[index + 1] ?? values.length, startUtf16: values[glyphStart]!.cluster }))
  const logicalGroups = [...visualGroups].sort((left, right) => left.startUtf16 - right.startUtf16)
  const glyphs: ShapedGlyph[] = []
  const clusters: ShapedCluster[] = []
  let advanceInline = 0
  for (let clusterIndex = 0; clusterIndex < logicalGroups.length; clusterIndex++) {
    const group = logicalGroups[clusterIndex]!
    const startUtf16 = group.startUtf16
    const endUtf16 = logicalGroups[clusterIndex + 1]?.startUtf16 ?? run.text.length
    if (endUtf16 <= startUtf16) return refusal('provider-failure', 'HarfBuzz returned an empty or descending UTF-16 cluster', request.font.face.faceId)
    if (startUtf16 > 0 && run.text.charCodeAt(startUtf16 - 1) >= 0xd800 && run.text.charCodeAt(startUtf16 - 1) <= 0xdbff && run.text.charCodeAt(startUtf16) >= 0xdc00 && run.text.charCodeAt(startUtf16) <= 0xdfff) return refusal('provider-failure', 'HarfBuzz cluster split a UTF-16 surrogate pair', request.font.face.faceId)
    if (values.slice(group.glyphStart, group.glyphEnd).some((value) => value.codepoint === 0)) return refusal('missing-glyph', 'font produced the .notdef glyph for canonical shaping', request.font.face.faceId, startUtf16, endUtf16)
    let clusterAdvance = 0
    let unsafeToBreak = false
    const glyphStart = glyphs.length
    for (let index = group.glyphStart; index < group.glyphEnd; index++) {
      const value = values[index]!
      if (![value.xAdvance, value.yAdvance, value.xOffset, value.yOffset].every((entry) => safeInteger(entry, -MAX_ABS_DESIGN_VALUE, MAX_ABS_DESIGN_VALUE))) return refusal('provider-failure', 'HarfBuzz returned an unbounded design-unit position', request.font.face.faceId)
      if (value.yAdvance !== 0) return refusal('unsupported-direction', 'horizontal LTR shaping unexpectedly returned vertical advance', request.font.face.faceId)
      const advanceX = scaleDesignValue(value.xAdvance!, cached.preflight.unitsPerEm, run.fontSizeMilliPoints)
      if (advanceX < 0) return refusal('provider-failure', 'scaled horizontal glyph advance is negative', request.font.face.faceId)
      const glyph: ShapedGlyph = Object.freeze({
        glyphId: value.codepoint,
        clusterIndex,
        advanceXMilliPoints: advanceX,
        advanceYMilliPoints: 0,
        offsetXMilliPoints: scaleDesignValue(value.xOffset!, cached.preflight.unitsPerEm, run.fontSizeMilliPoints),
        offsetYMilliPoints: scaleDesignValue(value.yOffset!, cached.preflight.unitsPerEm, run.fontSizeMilliPoints),
      })
      glyphs.push(glyph)
      clusterAdvance += advanceX
      if (!Number.isSafeInteger(clusterAdvance) || clusterAdvance > MAX_OUTPUT_MILLIPOINTS) return refusal('provider-failure', 'scaled cluster advance exceeds the page-paint provider bound', request.font.face.faceId)
      unsafeToBreak ||= (value.flags & hb.GlyphFlag.UNSAFE_TO_BREAK) !== 0
    }
    if (clusterAdvance < 0) return refusal('provider-failure', 'scaled horizontal cluster advance is negative', request.font.face.faceId)
    clusters.push(Object.freeze({
      startUtf16,
      endUtf16,
      glyphStart,
      glyphEnd: glyphs.length,
      advanceInlineMilliPoints: clusterAdvance,
      ...(unsafeToBreak ? { unsafeToBreak: true } : {}),
      ...(isFixedWhitespace(run.text, startUtf16, endUtf16) ? { whitespace: true } : {}),
    }))
    advanceInline += clusterAdvance
    if (!Number.isSafeInteger(advanceInline)) return refusal('provider-failure', 'scaled run advance exceeds deterministic integer precision', request.font.face.faceId)
  }
  const metrics = Object.freeze(qualifiedLineMetrics(request.font.metrics, run.fontSizeMilliPoints))
  return Object.freeze({
    startUtf16: 0,
    endUtf16: run.text.length,
    face: snapshotFace(request.font.face),
    glyphs: Object.freeze(glyphs),
    clusters: Object.freeze(clusters),
    metrics,
    advanceInlineMilliPoints: advanceInline,
    advanceBlockMilliPoints: 0,
  })
}

function runtimeArtifactDigests(): {
  wasm: `sha256:${string}`
  entry: `sha256:${string}`
  loader: `sha256:${string}`
  manifest: `sha256:${string}`
} {
  const runtimeEntry = import.meta.resolve('harfbuzzjs')
  const entryUrl = new URL(runtimeEntry)
  const wasm = readFileSync(new URL('./harfbuzz.wasm', entryUrl))
  if (wasm.byteLength !== EXPECTED_WASM_BYTES) throw new Error(`harfbuzzjs shaping WASM byte length ${wasm.byteLength} does not match the pinned runtime`)
  return {
    wasm: digestBytes(wasm),
    entry: digestBytes(readFileSync(entryUrl)),
    loader: digestBytes(readFileSync(new URL('./harfbuzz.js', entryUrl))),
    manifest: digestBytes(readFileSync(new URL('../package.json', entryUrl))),
  }
}

const ACTUAL_RUNTIME_VERSION = hb.versionString()
const ACTUAL_RUNTIME_ARTIFACTS = runtimeArtifactDigests()
if (ACTUAL_RUNTIME_VERSION !== HARFBUZZ_RUNTIME_VERSION) throw new Error(`harfbuzzjs runtime ${ACTUAL_RUNTIME_VERSION} does not match pinned HarfBuzz ${HARFBUZZ_RUNTIME_VERSION}`)
if (ACTUAL_RUNTIME_ARTIFACTS.wasm !== HARFBUZZ_WASM_SHA256) throw new Error(`harfbuzzjs shaping WASM ${ACTUAL_RUNTIME_ARTIFACTS.wasm} does not match the pinned digest`)
if (ACTUAL_RUNTIME_ARTIFACTS.entry !== HARFBUZZJS_ENTRY_SHA256) throw new Error(`harfbuzzjs entry ${ACTUAL_RUNTIME_ARTIFACTS.entry} does not match the pinned digest`)
if (ACTUAL_RUNTIME_ARTIFACTS.loader !== HARFBUZZJS_LOADER_SHA256) throw new Error(`harfbuzzjs loader ${ACTUAL_RUNTIME_ARTIFACTS.loader} does not match the pinned digest`)
if (ACTUAL_RUNTIME_ARTIFACTS.manifest !== HARFBUZZJS_MANIFEST_SHA256) throw new Error(`harfbuzzjs manifest ${ACTUAL_RUNTIME_ARTIFACTS.manifest} does not match the pinned digest`)

/**
 * Creates a bounded NativeTextShaper qualification boundary over the
 * repository-pinned HarfBuzzJS runtime. The source revision is mandatory and
 * becomes part of the provider revision consumed by shaped-lines and
 * page-paint provenance.
 */
export function createHarfBuzzTextShaperV1(options: HarfBuzzShaperOptionsV1): HarfBuzzTextShaperV1 {
  if (!isRecord(options) || !exactKeys(options, ['sourceRevision']) || typeof options.sourceRevision !== 'string' || !SOURCE_REVISION_RE.test(options.sourceRevision)) throw new TypeError('sourceRevision must be a stable 1-64 character InjOffice identifier')
  const revisionInput = Object.freeze({
    provider_id: PROVIDER_ID,
    source_revision: options.sourceRevision,
    runtime_package: `harfbuzzjs@${HARFBUZZJS_PACKAGE_VERSION}`,
    runtime_version: HARFBUZZ_RUNTIME_VERSION,
    runtime_wasm_sha256: HARFBUZZ_WASM_SHA256,
    runtime_entry_sha256: HARFBUZZJS_ENTRY_SHA256,
    runtime_loader_sha256: HARFBUZZJS_LOADER_SHA256,
    runtime_manifest_sha256: HARFBUZZJS_MANIFEST_SHA256,
    unicode_data_version: HARFBUZZ_UNICODE_DATA_VERSION,
    unicode_tables_sha256: UNICODE_13_TABLES_ENCODING_SHA256,
    unicode_projection_sha256: UNICODE_13_PROJECTION_SHA256,
    unicode_generator_sha256: UNICODE_13_GENERATOR_SHA256,
    configuration_revision: HARFBUZZ_SHAPER_CONFIG_REVISION,
    direction_policy: 'explicit-complete-horizontal-ltr-rtl',
    script_policy: SUPPORTED_SCRIPTS,
    cluster_level: 'monotone-graphemes',
    buffer_flags: BUFFER_FLAG_NAMES,
    default_feature_policy: 'harfbuzz-14.3.0-shape-defaults',
    feature_policy: 'explicit-qualified-and-font-advertised-kern-liga-only',
    variation_policy: 'refuse',
    spacing_policy: 'zero-only',
    font_policy: 'bounded-fixed-truetype-sfnt-or-ttc-preflight-v1',
    scale_policy: 'per-glyph-scaleFontUnits-then-sum-v1',
  })
  const providerRevision = `sha256:${bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(revisionInput))))}`
  const provenance: HarfBuzzShaperProvenanceV1 = Object.freeze({
    protocol: 'injoffice.text.harfbuzz-shaper-provenance',
    version: 1,
    provider_revision: providerRevision,
    ...revisionInput,
  })
  const cache = new Map<string, CachedFace>()
  const buffer = new hb.Buffer()
  let cachedBytes = 0

  const shapeUnsafe = (request: ShapeProviderRequest): ShapedSegment | NativeTextRefusal => {
    if (!isRecord(request) || !exactKeys(request, ['run', 'startUtf16', 'endUtf16', 'font'])) return refusal('invalid-contract', 'shape request must contain exactly run, startUtf16, endUtf16, and font')
    const runValidation = validateTextRunInput(request.run)
    if (!runValidation.ok) return refusal('invalid-contract', `invalid text run: ${runValidation.issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}`)
    const run = runValidation.value
    if (request.startUtf16 !== 0 || request.endUtf16 !== run.text.length) return refusal('invalid-contract', 'canonical HarfBuzz provider v1 accepts complete runs only')
    if (!isRecord(request.font) || !exactKeys(request.font, ['face', 'bytes', 'metrics']) || !validResolvedFace(request.font.face) || !(request.font.bytes instanceof Uint8Array) || !validDesignMetrics(request.font.metrics)) return refusal('invalid-contract', 'font resource is malformed or contains unknown fields')
    const font = request.font as FontResource
    const policyRefusal = runPolicyRefusal(run, font.face.faceId)
    if (policyRefusal) return policyRefusal
    if (font.face.sourceKind === 'system') return refusal('unsupported-font-format', 'system faces are outside canonical page-paint v1 even when bytes are content-addressed', font.face.faceId)
    if (font.bytes.byteLength === 0 || font.bytes.byteLength > HARFBUZZ_SHAPER_LIMITS.maxFontBytes) return refusal('font-bytes-unavailable', `font bytes must contain 1 through ${HARFBUZZ_SHAPER_LIMITS.maxFontBytes} bytes`, font.face.faceId)
    if ((typeof SharedArrayBuffer !== 'undefined' && font.bytes.buffer instanceof SharedArrayBuffer)
      || ('resizable' in font.bytes.buffer && font.bytes.buffer.resizable === true)) return refusal('invalid-contract', 'shared or resizable font byte storage is outside the immutable snapshot contract', font.face.faceId)
    const ownedBytes = Uint8Array.from(font.bytes)
    const beforeDigest = digestBytes(ownedBytes)
    if (beforeDigest !== font.face.contentDigest) return refusal('font-digest-mismatch', 'font bytes do not match the resolved content digest', font.face.faceId)
    const fontSnapshot: FontResource = Object.freeze({ face: snapshotFace(font.face), bytes: ownedBytes, metrics: snapshotMetrics(font.metrics) })
    const cacheKey = `${fontSnapshot.face.contentDigest}:${fontSnapshot.face.collectionIndex ?? 'standalone'}`
    let cached = cache.get(cacheKey)
    if (!cached) {
      if (cache.size >= HARFBUZZ_SHAPER_LIMITS.maxCachedFaces || cachedBytes + fontSnapshot.bytes.byteLength > HARFBUZZ_SHAPER_LIMITS.maxCachedFontBytes) return refusal('provider-failure', 'HarfBuzz face cache exceeds its bounded lifetime budget', fontSnapshot.face.faceId)
      const preflight = preflightSfnt(ownedBytes, fontSnapshot.face.collectionIndex)
      if (typeof preflight === 'string') return refusal('unsupported-font-format', preflight, font.face.faceId)
      if (!metricsMatchFont(fontSnapshot.metrics, preflight)) return refusal('font-metrics-unavailable', 'resolver metrics do not exactly match the digest-bound sfnt metric tables', fontSnapshot.face.faceId)
      try {
        const blob = new hb.Blob(ownedBytes)
        const face = new hb.Face(blob, preflight.collectionIndex)
        if (face.upem !== preflight.unitsPerEm) return refusal('unsupported-font-format', 'HarfBuzz face unitsPerEm disagree with the preflighted sfnt', font.face.faceId)
        const hbFont = new hb.Font(face)
        const features = new Set([...face.getTableFeatureTags('GSUB'), ...face.getTableFeatureTags('GPOS')])
        cached = { blob, face, font: hbFont, preflight, bytes: ownedBytes.byteLength, features }
        cache.set(cacheKey, cached)
        cachedBytes += ownedBytes.byteLength
      } catch {
        return refusal('unsupported-font-format', 'HarfBuzz could not construct the preflighted font face', font.face.faceId)
      }
    } else if (!metricsMatchFont(fontSnapshot.metrics, cached.preflight)) {
      return refusal('font-metrics-unavailable', 'font metrics changed for an existing content-addressed face', fontSnapshot.face.faceId)
    }
    try {
      const result = shapeWithCachedFace({ run, startUtf16: 0, endUtf16: run.text.length, font: fontSnapshot }, cached, buffer)
      if (digestBytes(font.bytes) !== beforeDigest) return refusal('font-digest-mismatch', 'font bytes changed during shaping', font.face.faceId)
      return result
    } catch {
      return refusal('provider-failure', 'pinned HarfBuzz runtime failed while shaping the qualified run', font.face.faceId)
    } finally {
      buffer.reset()
    }
  }

  const shape = (request: ShapeProviderRequest): ShapedSegment | NativeTextRefusal => {
    try {
      return shapeUnsafe(request)
    } catch {
      return refusal('provider-failure', 'canonical HarfBuzz provider rejected unreadable or hostile input')
    }
  }

  const shaper = Object.freeze({ providerId: PROVIDER_ID, providerRevision, provenance, shape })
  canonicalHarfBuzzShapers.add(shaper)
  return shaper
}

/** Proves that a shaper was constructed by this exact pinned runtime module. */
export function isCanonicalHarfBuzzTextShaperV1(value: unknown, sourceRevision: string): value is HarfBuzzTextShaperV1 {
  if (!isRecord(value) || !canonicalHarfBuzzShapers.has(value) || value.providerId !== PROVIDER_ID || typeof value.providerRevision !== 'string') return false
  const provenance = value.provenance
  return isRecord(provenance)
    && provenance.protocol === 'injoffice.text.harfbuzz-shaper-provenance'
    && provenance.version === 1
    && provenance.provider_id === PROVIDER_ID
    && provenance.provider_revision === value.providerRevision
    && provenance.source_revision === sourceRevision
    && provenance.runtime_package === `harfbuzzjs@${HARFBUZZJS_PACKAGE_VERSION}`
    && provenance.runtime_version === HARFBUZZ_RUNTIME_VERSION
    && provenance.runtime_wasm_sha256 === HARFBUZZ_WASM_SHA256
    && provenance.runtime_entry_sha256 === HARFBUZZJS_ENTRY_SHA256
    && provenance.runtime_loader_sha256 === HARFBUZZJS_LOADER_SHA256
    && provenance.runtime_manifest_sha256 === HARFBUZZJS_MANIFEST_SHA256
    && provenance.unicode_data_version === HARFBUZZ_UNICODE_DATA_VERSION
    && provenance.unicode_tables_sha256 === UNICODE_13_TABLES_ENCODING_SHA256
    && provenance.unicode_projection_sha256 === UNICODE_13_PROJECTION_SHA256
    && provenance.unicode_generator_sha256 === UNICODE_13_GENERATOR_SHA256
    && provenance.configuration_revision === HARFBUZZ_SHAPER_CONFIG_REVISION
    && provenance.direction_policy === 'explicit-complete-horizontal-ltr-rtl'
    && Array.isArray(provenance.script_policy) && JSON.stringify(provenance.script_policy) === JSON.stringify(SUPPORTED_SCRIPTS)
    && provenance.cluster_level === 'monotone-graphemes'
    && Array.isArray(provenance.buffer_flags) && JSON.stringify(provenance.buffer_flags) === JSON.stringify(BUFFER_FLAG_NAMES)
    && provenance.default_feature_policy === 'harfbuzz-14.3.0-shape-defaults'
    && provenance.feature_policy === 'explicit-qualified-and-font-advertised-kern-liga-only'
    && provenance.variation_policy === 'refuse'
    && provenance.spacing_policy === 'zero-only'
    && provenance.font_policy === 'bounded-fixed-truetype-sfnt-or-ttc-preflight-v1'
    && provenance.scale_policy === 'per-glyph-scaleFontUnits-then-sum-v1'
}
