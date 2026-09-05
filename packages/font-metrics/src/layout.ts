/**
 * Renderer-neutral native Office text contract.
 *
 * This module deliberately contains no DOM, canvas, React, Konva, Node, Electron,
 * or shaping implementation. PPTX RenderTree and DOCX pagination can depend on
 * the same integer-unit contract while hosts inject font and shaping providers.
 */

export const NATIVE_TEXT_LAYOUT_VERSION = 1 as const
export const MAX_TEXT_RUN_UTF16 = 262_144
export const MAX_FONT_FACES = 4_096
export const MAX_FALLBACK_CHAINS = 256

export type NativeTextLayoutVersion = typeof NATIVE_TEXT_LAYOUT_VERSION
export type MaybePromise<T> = T | Promise<T>
export type FontStyle = 'normal' | 'italic' | 'oblique'
export type TextDirection = 'ltr' | 'rtl' | 'ttb' | 'btt'
export type FontSourceKind = 'document' | 'bundled' | 'system' | 'host'
export type FontResolutionKind = 'exact' | 'substitute' | 'fallback'
export type Sha256Digest = `sha256:${string}`

/**
 * Conservative, deterministic line-break boundary shared by native Office
 * layout consumers. `unsupported` means the caller must preserve/refuse when a
 * break decision is required; it must never fall back to a browser or provider
 * whitespace guess.
 */
export type NativeOfficeLineBreakDecision = 'allowed' | 'prohibited' | 'unsupported'

type NativeOfficeLineBreakClass =
  | 'word'
  | 'space'
  | 'ideographicSpace'
  | 'glue'
  | 'wordJoiner'
  | 'zeroWidthSpace'
  | 'breakAfter'
  | 'cjk'
  | 'openPunctuation'
  | 'closePunctuation'
  | 'mark'

const OFFICE_OPEN_PUNCTUATION = new Set(Array.from('([{\uFF08\uFF3B\uFF5B\u3008\u300A\u300C\u300E\u3010\u3014\u3016\u3018\u301A', (character) => character.codePointAt(0)!))
const OFFICE_CLOSE_PUNCTUATION = new Set(Array.from(')]}\uFF09\uFF3D\uFF5D\u3009\u300B\u300D\u300F\u3011\u3015\u3017\u3019\u301B\u3001\u3002\uFF0C\uFF0E\uFF01\uFF1F\uFF1A\uFF1B,.!?:;', (character) => character.codePointAt(0)!))

function nativeOfficeLineBreakClass(codePoint: number): NativeOfficeLineBreakClass | undefined {
  if (codePoint === 0x20) return 'space'
  if (codePoint === 0x3000) return 'ideographicSpace'
  if (codePoint === 0xa0 || codePoint === 0x2007 || codePoint === 0x202f || codePoint === 0x2011) return 'glue'
  if (codePoint === 0x2060 || codePoint === 0xfeff) return 'wordJoiner'
  if (codePoint === 0x200b) return 'zeroWidthSpace'
  if (codePoint === 0x2d || codePoint === 0x2f || codePoint === 0x2010) return 'breakAfter'
  if (OFFICE_OPEN_PUNCTUATION.has(codePoint)) return 'openPunctuation'
  if (OFFICE_CLOSE_PUNCTUATION.has(codePoint)) return 'closePunctuation'
  if (
    (codePoint >= 0x30 && codePoint <= 0x39) ||
    (codePoint >= 0x41 && codePoint <= 0x5a) ||
    (codePoint >= 0x61 && codePoint <= 0x7a) ||
    (codePoint >= 0xc0 && codePoint <= 0x2af) ||
    (codePoint >= 0x370 && codePoint <= 0x52f)
  ) return 'word'
  // v1 intentionally does not claim Japanese kana kinsoku or jamo behavior:
  // those classes depend on language and line-breaking settings which are not
  // carried by this boundary classifier. Han ideographs and ordinary Hangul
  // syllables remain the bounded inter-character subset.
  if (
    (codePoint >= 0x3400 && codePoint <= 0x4dbf) ||
    (codePoint >= 0x4e00 && codePoint <= 0x9fff) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7a3) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff)
  ) return 'cjk'
  if (
    (codePoint >= 0x300 && codePoint <= 0x36f) ||
    (codePoint >= 0x1ab0 && codePoint <= 0x1aff) ||
    (codePoint >= 0x1dc0 && codePoint <= 0x1dff) ||
    (codePoint >= 0x20d0 && codePoint <= 0x20ff) ||
    (codePoint >= 0xfe00 && codePoint <= 0xfe0f) ||
    (codePoint >= 0xfe20 && codePoint <= 0xfe2f)
  ) return 'mark'
  return undefined
}

function nativeOfficeClusterEdges(text: string, startUtf16: number, endUtf16: number): { first: NativeOfficeLineBreakClass; last: NativeOfficeLineBreakClass } | undefined {
  if (!Number.isSafeInteger(startUtf16) || !Number.isSafeInteger(endUtf16) || startUtf16 < 0 || endUtf16 <= startUtf16 || endUtf16 > text.length) return undefined
  let first: NativeOfficeLineBreakClass | undefined
  let last: NativeOfficeLineBreakClass | undefined
  for (let index = startUtf16; index < endUtf16;) {
    const codePoint = text.codePointAt(index)
    if (codePoint === undefined || index + (codePoint > 0xffff ? 2 : 1) > endUtf16) return undefined
    const classification = nativeOfficeLineBreakClass(codePoint)
    index += codePoint > 0xffff ? 2 : 1
    if (classification === undefined || (classification === 'mark' && first === undefined)) return undefined
    if (classification === 'mark') continue
    first ??= classification
    last = classification
  }
  return first && last ? { first, last } : undefined
}

/**
 * Range form used by bounded layout engines so complete shaped-cluster slices can
 * be classified without allocating one substring per cluster.
 */
export function classifyNativeOfficeLineBreakRanges(
  leftText: string,
  leftStartUtf16: number,
  leftEndUtf16: number,
  rightText: string,
  rightStartUtf16: number,
  rightEndUtf16: number,
): NativeOfficeLineBreakDecision {
  const left = nativeOfficeClusterEdges(leftText, leftStartUtf16, leftEndUtf16)
  const right = nativeOfficeClusterEdges(rightText, rightStartUtf16, rightEndUtf16)
  if (!left || !right) return 'unsupported'
  // U+3000 is deliberately not folded into ordinary U+0020 separator behavior:
  // Office-visible ideographic spacing at a soft line edge needs its own proven
  // paint/advance rule before a consumer can claim native fidelity.
  if (left.last === 'ideographicSpace' || right.first === 'ideographicSpace') return 'unsupported'
  if (left.last === 'wordJoiner' || right.first === 'wordJoiner' || left.last === 'glue' || right.first === 'glue') return 'prohibited'
  if (left.last === 'zeroWidthSpace') return 'allowed'
  if (right.first === 'zeroWidthSpace') return 'prohibited'
  if (left.last === 'openPunctuation' || right.first === 'closePunctuation') return 'prohibited'
  if (left.last === 'space' || left.last === 'breakAfter') return 'allowed'
  if (left.last === 'cjk' && (right.first === 'cjk' || right.first === 'openPunctuation')) return 'allowed'
  if (left.last === 'closePunctuation' && right.first === 'cjk') return 'allowed'
  return 'prohibited'
}

/**
 * Classifies the boundary between two complete shaped-cluster text slices.
 * The deliberately bounded v1 subset models words, glue/WJ/ZWSP, explicit
 * hyphen/slash opportunities, and common East-Asian punctuation. Unknown
 * classes are refused instead of approximated.
 */
export function classifyNativeOfficeLineBreak(leftClusterText: string, rightClusterText: string): NativeOfficeLineBreakDecision {
  return classifyNativeOfficeLineBreakRanges(leftClusterText, 0, leftClusterText.length, rightClusterText, 0, rightClusterText.length)
}

/** OpenType script/feature/variation tag. Validation requires four printable ASCII characters. */
export type OpenTypeTag = string

/** Stable identity within a versioned InjOffice font manifest. */
export interface NativeFontFaceIdentity {
  faceId: string
  family: string
  postscriptName?: string
  weight: number
  style: FontStyle
  /** CSS-compatible percentage, 100 = normal. */
  stretch: number
}

/** Where a host resolver can obtain a face. Resource ids are opaque to layout code. */
export interface NativeFontSource {
  kind: FontSourceKind
  resourceId: string
  /** Optional before resolution; required on every ResolvedFontFace and in shaping cache keys. */
  contentDigest?: Sha256Digest
  /** Face index in TTC/OTC collections. Omit for standalone sfnt data. */
  collectionIndex?: number
}

export interface NativeFontFaceManifest extends NativeFontFaceIdentity {
  aliases?: readonly string[]
  source: NativeFontSource
  /** OpenType script tags known to be supported; absence means not pre-declared. */
  scripts?: readonly OpenTypeTag[]
  /** BCP-47 language tags known to be supported; absence means not pre-declared. */
  languages?: readonly string[]
}

/** Ordered face ids. Earlier faces must be attempted before later faces. */
export interface FontFallbackChain {
  chainId: string
  faceIds: readonly string[]
  scripts?: readonly OpenTypeTag[]
  languages?: readonly string[]
}

export interface NativeFontManifest {
  version: NativeTextLayoutVersion
  manifestId: string
  /** Changes whenever face metadata, bytes, or fallback ordering changes. */
  revision: string
  faces: readonly NativeFontFaceManifest[]
  fallbackChains: readonly FontFallbackChain[]
}

export interface NativeFontRequest {
  /** Office-authored family preference, ordered from strongest to weakest. */
  families: readonly string[]
  postscriptName?: string
  weight: number
  style: FontStyle
  stretch: number
  /** Ordered chain ids from NativeFontManifest. */
  fallbackChainIds?: readonly string[]
}

export interface OpenTypeFeature {
  tag: OpenTypeTag
  value: number
  /** UTF-16 offsets into TextRunInput.text. Both or neither must be present. */
  startUtf16?: number
  endUtf16?: number
}

export interface FontVariation {
  tag: OpenTypeTag
  /** Fixed numeric coordinate; NaN and infinity are invalid. */
  value: number
}

/** One native rich-text run before fallback resolution and shaping. */
export interface TextRunInput {
  version: NativeTextLayoutVersion
  text: string
  /** Integer 1/1000 point. Avoids browser/device-pixel-dependent input. */
  fontSizeMilliPoints: number
  font: NativeFontRequest
  /** Four-character OpenType script tag, e.g. Latn or Arab. */
  script: OpenTypeTag
  /** BCP-47 language tag used by the shaping provider. */
  language: string
  direction: TextDirection
  features?: readonly OpenTypeFeature[]
  variations?: readonly FontVariation[]
  letterSpacingMilliPoints?: number
  wordSpacingMilliPoints?: number
}

export interface FontResolutionRequest {
  manifest: NativeFontManifest
  run: TextRunInput
}

/** Content-addressed face chosen by an injected resolver. */
export interface ResolvedFontFace extends NativeFontFaceIdentity {
  sourceKind: FontSourceKind
  resourceId: string
  contentDigest: Sha256Digest
  collectionIndex?: number
  resolution: FontResolutionKind
  matchedFamily: string
  fallbackChainId?: string
}

/** Raw sfnt metrics. Values are signed integer design units unless noted. */
export interface FontDesignMetrics {
  unitsPerEm: number
  ascender: number
  descender: number
  lineGap: number
  capHeight?: number
  xHeight?: number
  underlinePosition?: number
  underlineThickness?: number
}

export interface FontResource {
  face: ResolvedFontFace
  /** Standalone sfnt bytes, or the original collection when collectionIndex is set. */
  bytes: Uint8Array
  metrics: FontDesignMetrics
}

export type NativeTextIssueCode =
  | 'invalid-contract'
  | 'font-not-found'
  | 'font-bytes-unavailable'
  | 'font-digest-mismatch'
  | 'font-metrics-unavailable'
  | 'unsupported-font-format'
  | 'unsupported-script'
  | 'unsupported-direction'
  | 'unsupported-feature'
  | 'missing-glyph'
  | 'provider-failure'

export interface NativeTextDecision {
  code: NativeTextIssueCode
  message: string
  /** True only when a caller may safely try a later fallback candidate. */
  recoverable: boolean
  faceId?: string
  startUtf16?: number
  endUtf16?: number
}

export interface FontResolutionSuccess {
  status: 'resolved'
  face: ResolvedFontFace
  attemptedFaceIds: readonly string[]
  decisions: readonly NativeTextDecision[]
}

export interface NativeTextRefusal {
  status: 'refused'
  decisions: readonly NativeTextDecision[]
  attemptedFaceIds: readonly string[]
}

export type FontResolutionResult = FontResolutionSuccess | NativeTextRefusal

/**
 * Host boundary for browser workers and servers. A resolver owns system/font-
 * asset lookup and metrics extraction; contract consumers never read paths or URLs.
 */
export interface NativeFontResolver {
  readonly providerId: string
  readonly providerRevision: string
  resolve(request: FontResolutionRequest): MaybePromise<FontResolutionResult>
  load(face: ResolvedFontFace): MaybePromise<FontResource | NativeTextRefusal>
}

/** Integer 1/1000-point line metrics shared by slide and document layout. */
export interface ScaledLineMetrics {
  fontSizeMilliPoints: number
  ascentMilliPoints: number
  descentMilliPoints: number
  lineGapMilliPoints: number
  lineHeightMilliPoints: number
  capHeightMilliPoints?: number
  xHeightMilliPoints?: number
  underlinePositionMilliPoints?: number
  underlineThicknessMilliPoints?: number
}

export interface ShapedGlyph {
  glyphId: number
  /** Index into the containing segment's clusters array. */
  clusterIndex: number
  advanceXMilliPoints: number
  advanceYMilliPoints: number
  offsetXMilliPoints: number
  offsetYMilliPoints: number
}

export interface ShapedCluster {
  /** UTF-16 offsets into the original TextRunInput.text. */
  startUtf16: number
  endUtf16: number
  /** Half-open range in ShapedSegment.glyphs. */
  glyphStart: number
  glyphEnd: number
  advanceInlineMilliPoints: number
  unsafeToBreak?: boolean
  whitespace?: boolean
}

/** One contiguous range shaped with one resolved font face. */
export interface ShapedSegment {
  startUtf16: number
  endUtf16: number
  face: ResolvedFontFace
  glyphs: readonly ShapedGlyph[]
  clusters: readonly ShapedCluster[]
  metrics: ScaledLineMetrics
  advanceInlineMilliPoints: number
  advanceBlockMilliPoints: number
}

export interface ShapedText {
  version: NativeTextLayoutVersion
  text: string
  direction: TextDirection
  segments: readonly ShapedSegment[]
  advanceInlineMilliPoints: number
  /** Max ascent/descent/line-gap across all segments, not a sum. */
  lineMetrics: ScaledLineMetrics
}

export interface ShapedTextSuccess {
  status: 'shaped'
  shaped: ShapedText
  decisions: readonly NativeTextDecision[]
}

export type ShapeTextResult = ShapedTextSuccess | NativeTextRefusal

export interface ShapeProviderRequest {
  run: TextRunInput
  /** A resolver may split the original run; these offsets select the provider's segment. */
  startUtf16: number
  endUtf16: number
  font: FontResource
}

/** Future HarfBuzz/OpenType implementations plug in here without changing consumers. */
export interface NativeTextShaper {
  readonly providerId: string
  readonly providerRevision: string
  shape(request: ShapeProviderRequest): MaybePromise<ShapedSegment | NativeTextRefusal>
}

export interface NativeTextValidationIssue {
  path: string
  code: 'type' | 'unknown-field' | 'required' | 'range' | 'format' | 'duplicate' | 'reference'
  message: string
}

export type NativeTextValidationResult<T> =
  | { ok: true; value: T }
  | { ok: false; issues: readonly NativeTextValidationIssue[] }

const ID_RE = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const TAG_RE = /^[\x20-\x7e]{4}$/
const LANGUAGE_RE = /^[A-Za-z0-9]{1,8}(?:-[A-Za-z0-9]{1,8})*$/
const DIGEST_RE = /^sha256:[a-f0-9]{64}$/
const MAX_METADATA_STRING = 1_024

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function exactKeys(value: Record<string, unknown>, allowed: readonly string[], path: string, issues: NativeTextValidationIssue[]): void {
  const set = new Set(allowed)
  for (const key of Object.keys(value)) {
    if (!set.has(key)) issues.push({ path: `${path}.${key}`, code: 'unknown-field', message: 'field is not part of native text contract v1' })
  }
}

function requiredString(value: unknown, path: string, issues: NativeTextValidationIssue[], format?: RegExp): value is string {
  if (typeof value !== 'string' || value.length === 0) {
    issues.push({ path, code: 'required', message: 'must be a non-empty string' })
    return false
  }
  if (value.length > MAX_METADATA_STRING) {
    issues.push({ path, code: 'range', message: `must not exceed ${MAX_METADATA_STRING} UTF-16 code units` })
    return false
  }
  if (format && !format.test(value)) {
    issues.push({ path, code: 'format', message: 'has an invalid format' })
    return false
  }
  return true
}

function optionalString(value: unknown, path: string, issues: NativeTextValidationIssue[], format?: RegExp): void {
  if (value === undefined) return
  requiredString(value, path, issues, format)
}

function integerInRange(value: unknown, min: number, max: number, path: string, issues: NativeTextValidationIssue[]): value is number {
  if (!Number.isSafeInteger(value) || (value as number) < min || (value as number) > max) {
    issues.push({ path, code: 'range', message: `must be a safe integer from ${min} through ${max}` })
    return false
  }
  return true
}

function finiteNumber(value: unknown, path: string, issues: NativeTextValidationIssue[]): value is number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    issues.push({ path, code: 'range', message: 'must be a finite number' })
    return false
  }
  return true
}

function validateStringArray(
  value: unknown,
  path: string,
  issues: NativeTextValidationIssue[],
  max: number,
  format?: RegExp,
): value is readonly string[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > max) {
    issues.push({ path, code: 'range', message: `must be a non-empty array with at most ${max} entries` })
    return false
  }
  const seen = new Set<string>()
  for (let i = 0; i < value.length; i++) {
    if (!requiredString(value[i], `${path}[${i}]`, issues, format)) continue
    const key = value[i] as string
    if (seen.has(key)) issues.push({ path: `${path}[${i}]`, code: 'duplicate', message: 'entry must be unique' })
    seen.add(key)
  }
  return true
}

function isWellFormedUtf16(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i)
    if (code >= 0xd800 && code <= 0xdbff) {
      if (i + 1 >= text.length) return false
      const next = text.charCodeAt(i + 1)
      if (next < 0xdc00 || next > 0xdfff) return false
      i++
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false
    }
  }
  return true
}

function validateFontRequest(value: unknown, path: string, issues: NativeTextValidationIssue[]): void {
  if (!isRecord(value)) {
    issues.push({ path, code: 'type', message: 'must be an object' })
    return
  }
  exactKeys(value, ['families', 'postscriptName', 'weight', 'style', 'stretch', 'fallbackChainIds'], path, issues)
  validateStringArray(value.families, `${path}.families`, issues, 32)
  optionalString(value.postscriptName, `${path}.postscriptName`, issues)
  integerInRange(value.weight, 1, 1_000, `${path}.weight`, issues)
  if (!['normal', 'italic', 'oblique'].includes(value.style as string)) {
    issues.push({ path: `${path}.style`, code: 'format', message: 'must be normal, italic, or oblique' })
  }
  integerInRange(value.stretch, 50, 200, `${path}.stretch`, issues)
  if (value.fallbackChainIds !== undefined) validateStringArray(value.fallbackChainIds, `${path}.fallbackChainIds`, issues, 32, ID_RE)
}

function isUtf16Boundary(text: string, index: number): boolean {
  if (index <= 0 || index >= text.length) return true
  const before = text.charCodeAt(index - 1)
  const after = text.charCodeAt(index)
  return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff)
}

function validateFeature(value: unknown, index: number, text: string, issues: NativeTextValidationIssue[]): void {
  const path = `$.features[${index}]`
  if (!isRecord(value)) {
    issues.push({ path, code: 'type', message: 'must be an object' })
    return
  }
  exactKeys(value, ['tag', 'value', 'startUtf16', 'endUtf16'], path, issues)
  requiredString(value.tag, `${path}.tag`, issues, TAG_RE)
  integerInRange(value.value, 0, 65_535, `${path}.value`, issues)
  const hasStart = value.startUtf16 !== undefined
  const hasEnd = value.endUtf16 !== undefined
  if (hasStart !== hasEnd) {
    issues.push({ path, code: 'required', message: 'startUtf16 and endUtf16 must be supplied together' })
  } else if (hasStart && hasEnd) {
    const startOk = integerInRange(value.startUtf16, 0, text.length, `${path}.startUtf16`, issues)
    const endOk = integerInRange(value.endUtf16, 0, text.length, `${path}.endUtf16`, issues)
    if (startOk && endOk && (value.startUtf16 as number) >= (value.endUtf16 as number)) {
      issues.push({ path, code: 'range', message: 'feature range must be non-empty and ascending' })
    }
    if (startOk && !isUtf16Boundary(text, value.startUtf16 as number)) issues.push({ path: `${path}.startUtf16`, code: 'range', message: 'must be on a Unicode code-point boundary' })
    if (endOk && !isUtf16Boundary(text, value.endUtf16 as number)) issues.push({ path: `${path}.endUtf16`, code: 'range', message: 'must be on a Unicode code-point boundary' })
  }
}

export function validateTextRunInput(value: unknown): NativeTextValidationResult<TextRunInput> {
  const issues: NativeTextValidationIssue[] = []
  if (!isRecord(value)) return { ok: false, issues: [{ path: '$', code: 'type', message: 'must be an object' }] }
  exactKeys(
    value,
    [
      'version',
      'text',
      'fontSizeMilliPoints',
      'font',
      'script',
      'language',
      'direction',
      'features',
      'variations',
      'letterSpacingMilliPoints',
      'wordSpacingMilliPoints',
    ],
    '$',
    issues,
  )
  if (value.version !== NATIVE_TEXT_LAYOUT_VERSION) issues.push({ path: '$.version', code: 'format', message: 'must be native text contract version 1' })
  if (typeof value.text !== 'string') {
    issues.push({ path: '$.text', code: 'type', message: 'must be a string' })
  } else {
    if (value.text.length > MAX_TEXT_RUN_UTF16) issues.push({ path: '$.text', code: 'range', message: `must not exceed ${MAX_TEXT_RUN_UTF16} UTF-16 code units` })
    if (!isWellFormedUtf16(value.text)) issues.push({ path: '$.text', code: 'format', message: 'must not contain unpaired UTF-16 surrogates' })
  }
  integerInRange(value.fontSizeMilliPoints, 1, 10_000_000, '$.fontSizeMilliPoints', issues)
  validateFontRequest(value.font, '$.font', issues)
  requiredString(value.script, '$.script', issues, TAG_RE)
  requiredString(value.language, '$.language', issues, LANGUAGE_RE)
  if (!['ltr', 'rtl', 'ttb', 'btt'].includes(value.direction as string)) {
    issues.push({ path: '$.direction', code: 'format', message: 'must be ltr, rtl, ttb, or btt' })
  }
  if (value.letterSpacingMilliPoints !== undefined) integerInRange(value.letterSpacingMilliPoints, -10_000_000, 10_000_000, '$.letterSpacingMilliPoints', issues)
  if (value.wordSpacingMilliPoints !== undefined) integerInRange(value.wordSpacingMilliPoints, -10_000_000, 10_000_000, '$.wordSpacingMilliPoints', issues)

  if (value.features !== undefined) {
    if (!Array.isArray(value.features) || value.features.length > 256) {
      issues.push({ path: '$.features', code: 'range', message: 'must be an array with at most 256 entries' })
    } else {
      const seen = new Set<string>()
      for (let i = 0; i < value.features.length; i++) {
        validateFeature(value.features[i], i, typeof value.text === 'string' ? value.text : '', issues)
        const feature = value.features[i]
        if (isRecord(feature) && typeof feature.tag === 'string') {
          const key = `${feature.tag}:${String(feature.startUtf16 ?? '')}:${String(feature.endUtf16 ?? '')}`
          if (seen.has(key)) issues.push({ path: `$.features[${i}]`, code: 'duplicate', message: 'feature tag and range must be unique' })
          seen.add(key)
        }
      }
    }
  }

  if (value.variations !== undefined) {
    if (!Array.isArray(value.variations) || value.variations.length > 64) {
      issues.push({ path: '$.variations', code: 'range', message: 'must be an array with at most 64 entries' })
    } else {
      const seen = new Set<string>()
      for (let i = 0; i < value.variations.length; i++) {
        const path = `$.variations[${i}]`
        const variation = value.variations[i]
        if (!isRecord(variation)) {
          issues.push({ path, code: 'type', message: 'must be an object' })
          continue
        }
        exactKeys(variation, ['tag', 'value'], path, issues)
        if (requiredString(variation.tag, `${path}.tag`, issues, TAG_RE)) {
          if (seen.has(variation.tag)) issues.push({ path: `${path}.tag`, code: 'duplicate', message: 'variation tag must be unique' })
          seen.add(variation.tag)
        }
        finiteNumber(variation.value, `${path}.value`, issues)
      }
    }
  }
  return issues.length === 0 ? { ok: true, value: value as unknown as TextRunInput } : { ok: false, issues }
}

export function validateFontManifest(value: unknown): NativeTextValidationResult<NativeFontManifest> {
  const issues: NativeTextValidationIssue[] = []
  if (!isRecord(value)) return { ok: false, issues: [{ path: '$', code: 'type', message: 'must be an object' }] }
  exactKeys(value, ['version', 'manifestId', 'revision', 'faces', 'fallbackChains'], '$', issues)
  if (value.version !== NATIVE_TEXT_LAYOUT_VERSION) issues.push({ path: '$.version', code: 'format', message: 'must be native text contract version 1' })
  requiredString(value.manifestId, '$.manifestId', issues, ID_RE)
  requiredString(value.revision, '$.revision', issues, ID_RE)

  const faceIds = new Set<string>()
  if (!Array.isArray(value.faces) || value.faces.length === 0 || value.faces.length > MAX_FONT_FACES) {
    issues.push({ path: '$.faces', code: 'range', message: `must contain 1 through ${MAX_FONT_FACES} faces` })
  } else {
    for (let i = 0; i < value.faces.length; i++) {
      const path = `$.faces[${i}]`
      const face = value.faces[i]
      if (!isRecord(face)) {
        issues.push({ path, code: 'type', message: 'must be an object' })
        continue
      }
      exactKeys(face, ['faceId', 'family', 'postscriptName', 'weight', 'style', 'stretch', 'aliases', 'source', 'scripts', 'languages'], path, issues)
      if (requiredString(face.faceId, `${path}.faceId`, issues, ID_RE)) {
        if (faceIds.has(face.faceId)) issues.push({ path: `${path}.faceId`, code: 'duplicate', message: 'faceId must be unique' })
        faceIds.add(face.faceId)
      }
      requiredString(face.family, `${path}.family`, issues)
      optionalString(face.postscriptName, `${path}.postscriptName`, issues)
      integerInRange(face.weight, 1, 1_000, `${path}.weight`, issues)
      if (!['normal', 'italic', 'oblique'].includes(face.style as string)) issues.push({ path: `${path}.style`, code: 'format', message: 'must be normal, italic, or oblique' })
      integerInRange(face.stretch, 50, 200, `${path}.stretch`, issues)
      if (face.aliases !== undefined) validateStringArray(face.aliases, `${path}.aliases`, issues, 64)
      if (face.scripts !== undefined) validateStringArray(face.scripts, `${path}.scripts`, issues, 256, TAG_RE)
      if (face.languages !== undefined) validateStringArray(face.languages, `${path}.languages`, issues, 256, LANGUAGE_RE)
      if (!isRecord(face.source)) {
        issues.push({ path: `${path}.source`, code: 'type', message: 'must be an object' })
      } else {
        exactKeys(face.source, ['kind', 'resourceId', 'contentDigest', 'collectionIndex'], `${path}.source`, issues)
        if (!['document', 'bundled', 'system', 'host'].includes(face.source.kind as string)) issues.push({ path: `${path}.source.kind`, code: 'format', message: 'must be document, bundled, system, or host' })
        requiredString(face.source.resourceId, `${path}.source.resourceId`, issues, ID_RE)
        optionalString(face.source.contentDigest, `${path}.source.contentDigest`, issues, DIGEST_RE)
        if (face.source.collectionIndex !== undefined) integerInRange(face.source.collectionIndex, 0, 65_535, `${path}.source.collectionIndex`, issues)
      }
    }
  }

  const chainIds = new Set<string>()
  if (!Array.isArray(value.fallbackChains) || value.fallbackChains.length > MAX_FALLBACK_CHAINS) {
    issues.push({ path: '$.fallbackChains', code: 'range', message: `must be an array with at most ${MAX_FALLBACK_CHAINS} chains` })
  } else {
    for (let i = 0; i < value.fallbackChains.length; i++) {
      const path = `$.fallbackChains[${i}]`
      const chain = value.fallbackChains[i]
      if (!isRecord(chain)) {
        issues.push({ path, code: 'type', message: 'must be an object' })
        continue
      }
      exactKeys(chain, ['chainId', 'faceIds', 'scripts', 'languages'], path, issues)
      if (requiredString(chain.chainId, `${path}.chainId`, issues, ID_RE)) {
        if (chainIds.has(chain.chainId)) issues.push({ path: `${path}.chainId`, code: 'duplicate', message: 'chainId must be unique' })
        chainIds.add(chain.chainId)
      }
      if (validateStringArray(chain.faceIds, `${path}.faceIds`, issues, 128, ID_RE)) {
        for (let j = 0; j < chain.faceIds.length; j++) {
          if (!faceIds.has(chain.faceIds[j])) issues.push({ path: `${path}.faceIds[${j}]`, code: 'reference', message: 'must reference a faceId in this manifest' })
        }
      }
      if (chain.scripts !== undefined) validateStringArray(chain.scripts, `${path}.scripts`, issues, 256, TAG_RE)
      if (chain.languages !== undefined) validateStringArray(chain.languages, `${path}.languages`, issues, 256, LANGUAGE_RE)
    }
  }
  return issues.length === 0 ? { ok: true, value: value as unknown as NativeFontManifest } : { ok: false, issues }
}

/** Scale a signed design-unit metric to integer 1/1000 point, with deterministic rounding. */
export function scaleFontUnits(value: number, unitsPerEm: number, fontSizeMilliPoints: number): number {
  if (!Number.isSafeInteger(value) || !Number.isSafeInteger(unitsPerEm) || unitsPerEm <= 0 || !Number.isSafeInteger(fontSizeMilliPoints) || fontSizeMilliPoints <= 0) {
    throw new RangeError('font metrics and size must be safe integers, and unitsPerEm/size must be positive')
  }
  const product = value * fontSizeMilliPoints
  if (!Number.isSafeInteger(product)) throw new RangeError('font metric scaling exceeds deterministic integer precision')
  const scaled = product / unitsPerEm
  if (!Number.isSafeInteger(Math.round(scaled))) throw new RangeError('scaled font metric exceeds the safe integer range')
  return Math.round(scaled)
}

export function scaleLineMetrics(metrics: FontDesignMetrics, fontSizeMilliPoints: number): ScaledLineMetrics {
  const ascent = scaleFontUnits(metrics.ascender, metrics.unitsPerEm, fontSizeMilliPoints)
  const descent = scaleFontUnits(metrics.descender, metrics.unitsPerEm, fontSizeMilliPoints)
  const lineGap = scaleFontUnits(metrics.lineGap, metrics.unitsPerEm, fontSizeMilliPoints)
  const lineHeight = ascent - descent + lineGap
  if (!Number.isSafeInteger(lineHeight)) throw new RangeError('scaled line height exceeds the safe integer range')
  const optional = (value: number | undefined): number | undefined =>
    value === undefined ? undefined : scaleFontUnits(value, metrics.unitsPerEm, fontSizeMilliPoints)
  return {
    fontSizeMilliPoints,
    ascentMilliPoints: ascent,
    descentMilliPoints: descent,
    lineGapMilliPoints: lineGap,
    lineHeightMilliPoints: lineHeight,
    capHeightMilliPoints: optional(metrics.capHeight),
    xHeightMilliPoints: optional(metrics.xHeight),
    underlinePositionMilliPoints: optional(metrics.underlinePosition),
    underlineThicknessMilliPoints: optional(metrics.underlineThickness),
  }
}

/** Unicode-normalized, case-insensitive family key for matching only; display names remain untouched. */
export function asciiLower(value: string): string {
  return value.replace(/[A-Z]/g, (character) => String.fromCharCode(character.charCodeAt(0) + 0x20))
}

export function hasAsciiEdgeWhitespace(value: string): boolean {
  return /^[\u0009-\u000D\u0020]|[\u0009-\u000D\u0020]$/.test(value)
}

export function normalizeFontFamilyName(family: string): string {
  return asciiLower(family.replace(/^[\u0009-\u000D\u0020]+|[\u0009-\u000D\u0020]+$/g, '').replace(/[\u0009-\u000D\u0020]+/g, ' '))
}

function compareCodeUnits(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

function canonicalFeatures(features: readonly OpenTypeFeature[] | undefined): readonly (readonly [string, number, number | null, number | null])[] {
  return [...(features ?? [])]
    .sort((a, b) => compareCodeUnits(a.tag, b.tag) || (a.startUtf16 ?? -1) - (b.startUtf16 ?? -1) || (a.endUtf16 ?? -1) - (b.endUtf16 ?? -1))
    .map((feature) => [feature.tag, feature.value, feature.startUtf16 ?? null, feature.endUtf16 ?? null] as const)
}

function canonicalVariations(variations: readonly FontVariation[] | undefined): readonly (readonly [string, number])[] {
  return [...(variations ?? [])].sort((a, b) => compareCodeUnits(a.tag, b.tag)).map((variation) => [variation.tag, variation.value] as const)
}

/** Deterministic key for resolver results. Validate the manifest and run before calling. */
export function fontResolutionCacheKey(request: FontResolutionRequest, resolver: Pick<NativeFontResolver, 'providerId' | 'providerRevision'>): string {
  const { manifest, run } = request
  return `native-font-resolution:v1:${JSON.stringify({
    provider: [resolver.providerId, resolver.providerRevision],
    manifest: [manifest.manifestId, manifest.revision],
    font: {
      families: run.font.families.map(normalizeFontFamilyName),
      postscriptName: run.font.postscriptName ?? null,
      weight: run.font.weight,
      style: run.font.style,
      stretch: run.font.stretch,
      fallbackChainIds: run.font.fallbackChainIds ?? [],
    },
    script: run.script,
    language: asciiLower(run.language),
    direction: run.direction,
    text: run.text,
    size: run.fontSizeMilliPoints,
    features: canonicalFeatures(run.features),
    variations: canonicalVariations(run.variations),
    letterSpacing: run.letterSpacingMilliPoints ?? 0,
    wordSpacing: run.wordSpacingMilliPoints ?? 0,
  })}`
}

/**
 * Content-addressed shaping key. It changes with bytes, face index, provider,
 * direction, language, features, variations, spacing, size, or text.
 */
export function shapingCacheKey(run: TextRunInput, face: ResolvedFontFace, shaper: Pick<NativeTextShaper, 'providerId' | 'providerRevision'>): string {
  return `native-text-shape:v1:${JSON.stringify({
    provider: [shaper.providerId, shaper.providerRevision],
    face: [face.contentDigest, face.collectionIndex ?? 0, face.faceId],
    text: run.text,
    size: run.fontSizeMilliPoints,
    script: run.script,
    language: asciiLower(run.language),
    direction: run.direction,
    features: canonicalFeatures(run.features),
    variations: canonicalVariations(run.variations),
    letterSpacing: run.letterSpacingMilliPoints ?? 0,
    wordSpacing: run.wordSpacingMilliPoints ?? 0,
  })}`
}
