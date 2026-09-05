/**
 * Bounded, deterministic Unicode Bidirectional Algorithm boundary.
 *
 * bidi-js@1.0.3 implements UAX #9 Unicode 13.0.0 conformance clause UAX-C1.
 * The exact factory source is verified before use so a modified runtime fails
 * closed instead of silently changing visual order.
 */

import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import bidiFactory from 'bidi-js'
import { readFileSync } from 'node:fs'
import { MAX_TEXT_RUN_UTF16 } from './layout.js'
import {
  UNICODE_13_GENERATOR_SHA256,
  UNICODE_13_PROJECTION_SHA256,
  UNICODE_13_TABLES_ENCODING_SHA256,
  UNICODE_13_TABLES_RUNTIME_MATCH,
  UNICODE_13_UCD_SOURCE_SHA256,
  isUnicode13Assigned,
} from './unicode13.js'

export const BIDI_JS_PACKAGE_VERSION = '1.0.3' as const
export const BIDI_UNICODE_VERSION = '13.0.0' as const
export const BIDI_JS_FACTORY_SHA256 = 'sha256:ce1928a26521e7eca2dde603a76f2332f7a1ab85977a4a10651af3f31f952a26' as const
export const BIDI_JS_ENTRY_SHA256 = 'sha256:5b58433ed951be70376bee55cb9a98cdce788e00cd2fa5f881cbcb1dcad03102' as const
export const BIDI_JS_MANIFEST_SHA256 = 'sha256:c2579f7705ab96dcc6192ab5a317d87760d2be7c67d72976d85edb4181bec5de' as const
export const NATIVE_BIDI_PROVIDER_ID = 'injoffice.bidi-js' as const
export const NATIVE_BIDI_CONFIGURATION_REVISION = 'injoffice.uax9-explicit-isolates-unicode13-bmp-only.v3' as const
export const UNICODE_13_DERIVED_AGE_SHA256 = 'sha256:e779a443d3aa2a3166a15becaa2b737c922480e32c0453d5956093633555078f' as const
export const NATIVE_BIDI_LIMITS = Object.freeze({
  maxUtf16: MAX_TEXT_RUN_UTF16,
  maxExplicitRanges: 16_384,
  maxClustersPerLine: MAX_TEXT_RUN_UTF16,
})

export interface NativeBidiExplicitRangeV1 {
  startUtf16: number
  endUtf16: number
  direction: 'ltr' | 'rtl'
}

export interface NativeBidiParagraphInputV1 {
  text: string
  baseDirection: 'ltr' | 'rtl'
  /** Non-overlapping higher-level-protocol runs, represented as UAX #9 isolates. */
  explicitRanges?: readonly NativeBidiExplicitRangeV1[]
}

export interface NativeBidiParagraphV1 {
  providerId: typeof NATIVE_BIDI_PROVIDER_ID
  providerRevision: `sha256:${string}`
  unicodeVersion: typeof BIDI_UNICODE_VERSION
  textLengthUtf16: number
  baseLevel: 0 | 1
  /** Resolved embedding level for every UTF-16 code unit in the original text. */
  levels: readonly number[]
}

export interface NativeBidiVisualOrderV1 {
  visualToLogical: readonly number[]
  logicalToVisual: readonly number[]
}

export type NativeBidiResult<T> =
  | { ok: true; value: T }
  | { ok: false; code: 'invalid-input' | 'unsupported-control' | 'unsupported-scalar' | 'resource-limit' | 'runtime-mismatch' | 'resolution-failure'; message: string }

const BIDI_CONTROL_RE = /[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/u

function digestText(value: string): `sha256:${string}` {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(value)))}`
}

function wellFormedUtf16(text: string): boolean {
  for (let index = 0; index < text.length; index++) {
    const value = text.charCodeAt(index)
    if (value >= 0xd800 && value <= 0xdbff) {
      const next = text.charCodeAt(++index)
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false
    } else if (value >= 0xdc00 && value <= 0xdfff) return false
  }
  return true
}

function unicode13Repertoire(text: string): boolean {
  for (let index = 0; index < text.length;) {
    const codePoint = text.codePointAt(index)
    if (codePoint === undefined || !isUnicode13Assigned(codePoint)) return false
    index += codePoint > 0xffff ? 2 : 1
  }
  return true
}

function codePointBoundary(text: string, index: number): boolean {
  if (index <= 0 || index >= text.length) return true
  const before = text.charCodeAt(index - 1)
  const after = text.charCodeAt(index)
  return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff)
}

const factoryDigest = digestText(String(bidiFactory))
const runtimeEntry = new URL(import.meta.resolve('bidi-js'))
const entryDigest = `sha256:${bytesToHex(sha256(readFileSync(runtimeEntry)))}`
const manifestDigest = `sha256:${bytesToHex(sha256(readFileSync(new URL('../package.json', runtimeEntry))))}`
const repertoireEncodingDigest = UNICODE_13_TABLES_ENCODING_SHA256
const runtimeMatches = factoryDigest === BIDI_JS_FACTORY_SHA256 && entryDigest === BIDI_JS_ENTRY_SHA256 && manifestDigest === BIDI_JS_MANIFEST_SHA256 && UNICODE_13_TABLES_RUNTIME_MATCH
const bidi = runtimeMatches ? bidiFactory() : undefined
const revisionInput = JSON.stringify({
  provider: NATIVE_BIDI_PROVIDER_ID,
  package: `bidi-js@${BIDI_JS_PACKAGE_VERSION}`,
  unicode: BIDI_UNICODE_VERSION,
  repertoire: UNICODE_13_DERIVED_AGE_SHA256,
  ucd_sources: UNICODE_13_UCD_SOURCE_SHA256,
  unicode_projection: UNICODE_13_PROJECTION_SHA256,
  unicode_generator: UNICODE_13_GENERATOR_SHA256,
  unicode_tables: UNICODE_13_TABLES_ENCODING_SHA256,
  factory: BIDI_JS_FACTORY_SHA256,
  runtime_entry: BIDI_JS_ENTRY_SHA256,
  runtime_manifest: BIDI_JS_MANIFEST_SHA256,
  configuration: NATIVE_BIDI_CONFIGURATION_REVISION,
})
export const NATIVE_BIDI_PROVIDER_REVISION = digestText(revisionInput)

/** Resolve one bounded paragraph, applying explicit DOCX run direction as isolates. */
export function resolveNativeBidiParagraphV1(input: NativeBidiParagraphInputV1): NativeBidiResult<NativeBidiParagraphV1> {
  if (!runtimeMatches || !bidi) return { ok: false, code: 'runtime-mismatch', message: `bidi-js factory ${factoryDigest}, entry ${entryDigest}, manifest ${manifestDigest}, or Unicode table projection ${repertoireEncodingDigest} does not match the pinned runtime` }
  let text: string
  let baseDirection: 'ltr' | 'rtl'
  let ranges: NativeBidiExplicitRangeV1[]
  try {
    if (!input || typeof input !== 'object') return { ok: false, code: 'invalid-input', message: 'bidi input must contain text and an explicit horizontal base direction' }
    const liveText = input.text
    const liveDirection = input.baseDirection
    const liveRanges = input.explicitRanges
    if (typeof liveText !== 'string' || (liveDirection !== 'ltr' && liveDirection !== 'rtl')) return { ok: false, code: 'invalid-input', message: 'bidi input must contain text and an explicit horizontal base direction' }
    text = liveText
    baseDirection = liveDirection
    if (text.length > NATIVE_BIDI_LIMITS.maxUtf16) return { ok: false, code: 'resource-limit', message: `bidi text exceeds ${NATIVE_BIDI_LIMITS.maxUtf16} UTF-16 code units` }
    if (liveRanges === undefined) ranges = []
    else {
      if (!Array.isArray(liveRanges)) return { ok: false, code: 'invalid-input', message: 'explicit bidi ranges must be an array' }
      const rangeCount = liveRanges.length
      if (rangeCount > NATIVE_BIDI_LIMITS.maxExplicitRanges) return { ok: false, code: 'resource-limit', message: `explicit bidi ranges exceed ${NATIVE_BIDI_LIMITS.maxExplicitRanges}` }
      ranges = new Array<NativeBidiExplicitRangeV1>(rangeCount)
      for (let index = 0; index < rangeCount; index++) {
        const range = liveRanges[index]
        if (!range || typeof range !== 'object') return { ok: false, code: 'invalid-input', message: 'explicit bidi range is malformed' }
        ranges[index] = { startUtf16: range.startUtf16, endUtf16: range.endUtf16, direction: range.direction }
      }
    }
  } catch {
    return { ok: false, code: 'invalid-input', message: 'bidi input could not be snapshotted as bounded plain data' }
  }
  if (!wellFormedUtf16(text)) return { ok: false, code: 'invalid-input', message: 'bidi text contains malformed UTF-16' }
  if (/[^\u0000-\uFFFF]/u.test(text)) return { ok: false, code: 'unsupported-scalar', message: 'bidi-js@1.0.3 is qualified only for BMP scalars; astral scalars are atomically refused before resolution' }
  if (!unicode13Repertoire(text)) return { ok: false, code: 'unsupported-scalar', message: 'bidi text contains a scalar outside the assigned Unicode 13.0.0 repertoire' }
  if (BIDI_CONTROL_RE.test(text)) return { ok: false, code: 'unsupported-control', message: 'authored Unicode bidi controls are outside the DOCX higher-level-protocol slice' }
  let previousEnd = 0
  let synthetic = ''
  const originalIndices: number[] = []
  const appendOriginal = (start: number, end: number): void => {
    for (let index = start; index < end; index++) {
      synthetic += text[index]!
      originalIndices.push(index)
    }
  }
  for (const range of ranges) {
    if (!Number.isSafeInteger(range.startUtf16) || !Number.isSafeInteger(range.endUtf16) || range.startUtf16 < previousEnd || range.endUtf16 <= range.startUtf16 || range.endUtf16 > text.length || !codePointBoundary(text, range.startUtf16) || !codePointBoundary(text, range.endUtf16) || (range.direction !== 'ltr' && range.direction !== 'rtl')) return { ok: false, code: 'invalid-input', message: 'explicit bidi ranges must be ordered, non-overlapping, non-empty code-point ranges' }
    appendOriginal(previousEnd, range.startUtf16)
    synthetic += range.direction === 'rtl' ? '\u2067' : '\u2066'
    originalIndices.push(-1)
    appendOriginal(range.startUtf16, range.endUtf16)
    synthetic += '\u2069'
    originalIndices.push(-1)
    previousEnd = range.endUtf16
  }
  appendOriginal(previousEnd, text.length)
  if (synthetic.length > NATIVE_BIDI_LIMITS.maxUtf16 + NATIVE_BIDI_LIMITS.maxExplicitRanges * 2) return { ok: false, code: 'resource-limit', message: 'synthetic bidi isolate input exceeds its bounded expansion' }
  try {
    const resolved = bidi.getEmbeddingLevels(synthetic, baseDirection)
    if (!(resolved.levels instanceof Uint8Array) || resolved.levels.length !== synthetic.length) return { ok: false, code: 'resolution-failure', message: 'bidi runtime returned malformed embedding levels' }
    const levels = new Array<number>(text.length)
    for (let syntheticIndex = 0; syntheticIndex < originalIndices.length; syntheticIndex++) {
      const originalIndex = originalIndices[syntheticIndex]!
      if (originalIndex >= 0) levels[originalIndex] = resolved.levels[syntheticIndex]!
    }
    if (levels.some((level) => !Number.isSafeInteger(level) || level < 0 || level > 125)) return { ok: false, code: 'resolution-failure', message: 'bidi runtime did not cover the original UTF-16 input with bounded levels' }
    return { ok: true, value: Object.freeze({ providerId: NATIVE_BIDI_PROVIDER_ID, providerRevision: NATIVE_BIDI_PROVIDER_REVISION, unicodeVersion: BIDI_UNICODE_VERSION, textLengthUtf16: text.length, baseLevel: baseDirection === 'rtl' ? 1 : 0, levels: Object.freeze(levels) }) }
  } catch {
    return { ok: false, code: 'resolution-failure', message: 'pinned bidi runtime failed while resolving the bounded paragraph' }
  }
}

/** Apply UAX #9 L1/L2 to already cluster-aligned levels for one soft line. */
export function reorderNativeBidiLineV1(levelValues: readonly number[], baseLevel: 0 | 1, trailingWhitespace: readonly boolean[]): NativeBidiResult<NativeBidiVisualOrderV1> {
  let levels: number[]
  let whitespace: boolean[]
  try {
    if (!Array.isArray(levelValues) || !Array.isArray(trailingWhitespace) || (baseLevel !== 0 && baseLevel !== 1)) return { ok: false, code: 'invalid-input', message: 'line bidi levels and trailing-whitespace flags must be equal bounded arrays' }
    const levelCount = levelValues.length
    const whitespaceCount = trailingWhitespace.length
    if (levelCount !== whitespaceCount || levelCount > NATIVE_BIDI_LIMITS.maxClustersPerLine) return { ok: false, code: 'invalid-input', message: 'line bidi levels and trailing-whitespace flags must be equal bounded arrays' }
    levels = new Array<number>(levelCount)
    whitespace = new Array<boolean>(levelCount)
    for (let index = 0; index < levelCount; index++) {
      levels[index] = levelValues[index]!
      whitespace[index] = trailingWhitespace[index]!
    }
  } catch {
    return { ok: false, code: 'invalid-input', message: 'line bidi input could not be snapshotted as bounded plain data' }
  }
  if (levels.some((level) => !Number.isSafeInteger(level) || level < 0 || level > 125) || whitespace.some((value) => typeof value !== 'boolean')) return { ok: false, code: 'invalid-input', message: 'line bidi levels or whitespace flags are malformed' }
  for (let index = levels.length - 1; index >= 0 && whitespace[index]; index--) levels[index] = baseLevel
  const visualToLogical = levels.map((_, index) => index)
  let maxLevel: number = baseLevel
  let minOdd = 126
  for (const level of levels) {
    maxLevel = Math.max(maxLevel, level)
    minOdd = Math.min(minOdd, level | 1)
  }
  for (let level = maxLevel; level >= minOdd; level--) {
    for (let start = 0; start < levels.length; start++) {
      if (levels[start]! < level) continue
      let end = start
      while (end + 1 < levels.length && levels[end + 1]! >= level) end++
      for (let left = start, right = end; left < right; left++, right--) [visualToLogical[left], visualToLogical[right]] = [visualToLogical[right]!, visualToLogical[left]!]
      start = end
    }
  }
  const logicalToVisual = new Array<number>(visualToLogical.length)
  visualToLogical.forEach((logical, visual) => { logicalToVisual[logical] = visual })
  return { ok: true, value: Object.freeze({ visualToLogical: Object.freeze(visualToLogical), logicalToVisual: Object.freeze(logicalToVisual) }) }
}
