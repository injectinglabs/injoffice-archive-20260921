import { qualifySymbolBullet } from './symbolBullet.js'
import {
  NATIVE_TEXT_LAYOUT_VERSION,
  classifyNativeOfficeLineBreakRanges,
  scaleLineMetrics,
  scaleFontUnits,
  validateFontManifest,
  validateTextRunInput,
  type NativeTextDecision,
  type NativeTextRefusal,
  type FontResource,
  type NativeFontManifest,
  type NativeFontResolver,
  type NativeTextShaper,
  type ResolvedFontFace,
  type ScaledLineMetrics,
  type ShapedCluster,
  type ShapedGlyph,
  type ShapedSegment,
  type TextDirection,
  type TextRunInput,
} from '@injoffice/font-metrics/layout'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  assertNativePptx,
  type NativeAsset,
  type NativeCompatibility,
  type NativeElement,
  type NativeParagraph,
  type NativePptxDeck,
  type NativeSlide,
  type NativeTextBodyLayout,
  type NativeTextRun,
} from '@injoffice/pptx-native'
import { connectorPath, defaultPentagonTextRect, localBounds, presetPath, translationTransform,quarterTurnTransform } from './geometry.js'
import {
  PPTX_RENDER_LIMITS,
  PPTX_RENDER_TREE_VERSION,
  RenderCompileError,
  type CompileSlideOptions,
  type NativePptxTextDefaults,
  type NativePptxTextOverride,
  type NativePptxTextRunContext,
  type RenderAsset,
  type RenderDiagnostic,
  type RenderCluster,
  type RenderGlyph,
  type RenderGroupNode,
  type RenderImageNode,
  type RenderNode,
  type RenderParagraphNode,
  type RenderPathCommand,
  type RenderPlaceholderNode,
  type RenderRect,
  type RenderStroke,
  type RenderTableCellNode,
  type RenderTextRunNode,
  type RenderTextBodyNode,
  type SlideRenderTree,
} from './types.js'

interface Budget {
  nodes: number
  glyphs: number
  clusters: number
  providerDecisions: number
  providerAttemptedFaceIds: number
  textLines: number
  textFragments: number
  readonly maxNodes: number
  readonly maxGlyphs: number
  readonly maxClusters: number
  readonly maxDepth: number
  readonly maxCoordinateEmu: number
}

interface CompileState {
  readonly lineLayoutPolicy?: 'max-run-natural-v1'
  readonly sourceFrameAutoFitPreview: boolean
  readonly deck: NativePptxDeck
  readonly slide: NativeSlide
  readonly options: CompileSlideOptions
  readonly assets: ReadonlyMap<string, NativeAsset>
  readonly referencedAssets: Map<string, RenderAsset>
  readonly hostResolutionDiagnosticElements: Set<string>
  readonly diagnostics: RenderDiagnostic[]
  readonly budget: Budget
  readonly providers: ProviderSnapshot
  readonly fontManifest: NativeFontManifest
  readonly fontResources: Map<string, FontResource>
  readonly providerFontResources: Map<string, FontResource>
  readonly nativeTextInheritanceUnresolved: boolean
  readonly textDefaults: NativePptxTextDefaults
  readonly resolveRun?: (context: NativePptxTextRunContext) => NativePptxTextOverride
  providerResolveCalls: number
  providerLoadCalls: number
  providerShapeCalls: number
  uniqueFontBytes: number
  providerFacingFontBytes: number
}

interface ProviderSnapshot {
  readonly resolver: {
    readonly providerId: string
    readonly providerRevision: string
    readonly live: NativeFontResolver
    readonly resolve: NativeFontResolver['resolve']
    readonly load: NativeFontResolver['load']
  }
  readonly shaper: {
    readonly providerId: string
    readonly providerRevision: string
    readonly live: NativeTextShaper
    readonly shape: NativeTextShaper['shape']
  }
}

interface ExactRational {
  readonly numerator: bigint
  readonly denominator: bigint
}

interface WorldAffine {
  readonly a: ExactRational
  readonly b: ExactRational
  readonly c: ExactRational
  readonly d: ExactRational
  readonly tx: ExactRational
  readonly ty: ExactRational
}

const AFFINE_PPM = 1_000_000n
const RATIONAL_ZERO: ExactRational = { numerator: 0n, denominator: 1n }
const RATIONAL_ONE: ExactRational = { numerator: 1n, denominator: 1n }
const IDENTITY_WORLD_AFFINE: WorldAffine = { a: RATIONAL_ONE, b:RATIONAL_ZERO, c:RATIONAL_ZERO, d: RATIONAL_ONE, tx: RATIONAL_ZERO, ty: RATIONAL_ZERO }

function affineGcd(left: bigint, right: bigint): bigint {
  let a = left < 0n ? -left : left
  let b = right < 0n ? -right : right
  while (b !== 0n) {
    const remainder = a % b
    a = b
    b = remainder
  }
  return a
}

function exactRational(numerator: bigint, denominator: bigint): ExactRational {
  if (denominator <= 0n) throw new RenderCompileError('render.worldTransform', '$.transform', 'internal affine denominator must be positive')
  if (numerator === 0n) return RATIONAL_ZERO
  const divisor = affineGcd(numerator, denominator)
  return { numerator: numerator / divisor, denominator: denominator / divisor }
}

function multiplyRational(left: ExactRational, right: ExactRational): ExactRational {
  const leftDivisor = affineGcd(left.numerator, right.denominator)
  const rightDivisor = affineGcd(right.numerator, left.denominator)
  return exactRational(
    (left.numerator / leftDivisor) * (right.numerator / rightDivisor),
    (left.denominator / rightDivisor) * (right.denominator / leftDivisor),
  )
}

function multiplyRationalInteger(value: ExactRational, integer: number): ExactRational {
  return exactRational(value.numerator * BigInt(integer), value.denominator)
}

function addRational(left: ExactRational, right: ExactRational): ExactRational {
  const divisor = affineGcd(left.denominator, right.denominator)
  const leftFactor = right.denominator / divisor
  const rightFactor = left.denominator / divisor
  return exactRational(left.numerator * leftFactor + right.numerator * rightFactor, left.denominator * leftFactor)
}

interface TextContainerContext {
  readonly elementId: string
  readonly elementKind: 'text' | 'shape' | 'table'
  readonly bounds: RenderRect
  readonly layout?: NativeTextBodyLayout
}

interface WrapPoint {
  runIndex: number
  clusterIndex: number
}

interface WrapLineRange {
  readonly startRunIndex: number
  readonly startClusterIndex: number
  readonly endRunIndex: number
  readonly endClusterIndex: number
  readonly consumedSeparatorRunIndex?: number
  readonly consumedSeparatorClusterIndex?: number
}

interface FragmentProgress {
  clusterIndex: number
  glyphIndex: number
  glyphPenXEmu: number
  glyphPenYEmu: number
}

class TextBodyLayoutRefusal extends Error {
  constructor(readonly code: string, message: string) {
    super(message)
  }
}

interface ShapedRunResult {
  readonly run: RenderTextRunNode
  readonly direction: TextDirection
  readonly lineHeightEmu: number
  readonly ascentEmu: number
  readonly ascentMilliPoints: number
  readonly descentMilliPoints: number
  readonly lineGapMilliPoints: number
}

const DEFAULT_COLOR = '000000'
const DEFAULT_BACKGROUND = 'FFFFFF'

function configuredLimit(value: number | undefined, maximum: number, name: string): number {
  const result = value ?? maximum
  if (!Number.isSafeInteger(result) || result < 1 || result > maximum) {
    throw new RenderCompileError('render.invalidLimit', '$.options', `${name} must be an integer from 1 through ${maximum}`)
  }
  return result
}

function checkCoordinate(value: number, path: string, budget: Budget, positive = false): void {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || Math.abs(value) > budget.maxCoordinateEmu || (positive && value < 1)) {
    throw new RenderCompileError('render.coordinateBudget', path, `coordinate must be ${positive ? 'positive and ' : ''}within ±${budget.maxCoordinateEmu} EMU`)
  }
}

function boundedStroke(
  stroke: Readonly<{ color: string; widthEmu: number; cap?: 'flat' | 'round' | 'square'; join?: 'round' | 'bevel' | 'miter'; dash?: 'solid'; miterLimit?: number }>,
  path: string,
  budget: Budget,
): RenderStroke {
  checkCoordinate(stroke.widthEmu, `${path}.widthEmu`, budget)
  return { color: stroke.color, widthEmu: stroke.widthEmu, cap: stroke.cap, join: stroke.join, dash: stroke.dash, miterLimit: stroke.miterLimit }
}

function takeNode(state: CompileState, path: string): void {
  state.budget.nodes++
  if (state.budget.nodes > state.budget.maxNodes) {
    throw new RenderCompileError('render.nodeBudget', path, `RenderTree exceeds ${state.budget.maxNodes} nodes`)
  }
}

function takeGlyphs(state: CompileState, count: number, path: string): void {
  state.budget.glyphs += count
  if (state.budget.glyphs > state.budget.maxGlyphs) {
    throw new RenderCompileError('render.glyphBudget', path, `RenderTree exceeds ${state.budget.maxGlyphs} glyphs`)
  }
}

function takeClusters(state: CompileState, count: number, path: string): void {
  state.budget.clusters += count
  if (state.budget.clusters > state.budget.maxClusters) {
    throw new RenderCompileError('render.clusterBudget', path, `RenderTree exceeds ${state.budget.maxClusters} shaped clusters`)
  }
}

const PROVIDER_ID_RE = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const PROVIDER_DIGEST_RE = /^sha256:[a-f0-9]{64}$/
const MAX_PROVIDER_OBJECT_KEYS = 16
const PROVIDER_DECISION_CODES = new Set(['invalid-contract', 'font-not-found', 'font-bytes-unavailable', 'font-digest-mismatch', 'font-metrics-unavailable', 'unsupported-font-format', 'unsupported-script', 'unsupported-direction', 'unsupported-feature', 'missing-glyph', 'provider-failure'])

function providerRecord(value: unknown, path: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new RenderCompileError('render.invalidProviderOutput', path, 'provider output must be an object')
  }
  let prototype: object | null
  try { prototype = Object.getPrototypeOf(value) } catch { throw new RenderCompileError('render.invalidProviderOutput', path, 'provider output prototype is unreadable') }
  if (prototype !== Object.prototype && prototype !== null) throw new RenderCompileError('render.invalidProviderOutput', path, 'provider output must be a plain bounded record')
  const captured: Record<string, unknown> = {}
  let count = 0
  try {
    for (const key in value as Record<string, unknown>) {
      if (!Object.hasOwn(value, key)) continue
      count++
      if (count > MAX_PROVIDER_OBJECT_KEYS) throw new RenderCompileError('render.invalidProviderOutput', path, `provider record exceeds ${MAX_PROVIDER_OBJECT_KEYS} fields`)
      captured[key] = (value as Record<string, unknown>)[key]
    }
  } catch (error) {
    if (error instanceof RenderCompileError) throw error
    throw new RenderCompileError('render.invalidProviderOutput', path, 'provider output fields are unreadable or unstable')
  }
  return captured
}

function exactProviderKeys(value: Record<string, unknown>, allowed: readonly string[], path: string): void {
  const keys = new Set(allowed)
  let count = 0
  try {
    for (const key in value) {
      if (!Object.hasOwn(value, key)) continue
      count++
      if (count > allowed.length || !keys.has(key)) throw new RenderCompileError('render.invalidProviderOutput', `${path}.${key}`, 'unknown or excessive provider output field')
    }
  } catch (error) {
    if (error instanceof RenderCompileError) throw error
    throw new RenderCompileError('render.invalidProviderOutput', path, 'provider output keys are unreadable')
  }
}

function providerString(value: unknown, path: string, max = 1_024, format?: RegExp): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > max || (format && !format.test(value))) {
    throw new RenderCompileError('render.invalidProviderOutput', path, `must be a non-empty bounded string${format ? ' with the required format' : ''}`)
  }
  return value
}

function providerInteger(value: unknown, path: string, minimum = Number.MIN_SAFE_INTEGER, maximum = Number.MAX_SAFE_INTEGER): number {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < minimum || (value as number) > maximum) {
    throw new RenderCompileError('render.invalidProviderOutput', path, `must be a safe integer from ${minimum} through ${maximum}`)
  }
  return value as number
}

function normalizeDecisions(value: unknown, state: CompileState, path: string, textLength: number): readonly NativeTextDecision[] {
  if (!Array.isArray(value)) throw new RenderCompileError('render.invalidProviderOutput', path, 'decisions must be an array')
  state.budget.providerDecisions += value.length
  if (state.budget.providerDecisions > PPTX_RENDER_LIMITS.maxProviderDecisions) {
    throw new RenderCompileError('render.providerDecisionBudget', path, `text providers exceed ${PPTX_RENDER_LIMITS.maxProviderDecisions} decisions`)
  }
  return value.map((item, index) => {
    const decisionPath = `${path}[${index}]`
    const decision = providerRecord(item, decisionPath)
    exactProviderKeys(decision, ['code', 'message', 'recoverable', 'faceId', 'startUtf16', 'endUtf16'], decisionPath)
    const code = providerString(decision.code, `${decisionPath}.code`, 64)
    if (!PROVIDER_DECISION_CODES.has(code)) throw new RenderCompileError('render.invalidProviderOutput', `${decisionPath}.code`, 'unknown native text decision code')
    if (typeof decision.recoverable !== 'boolean') throw new RenderCompileError('render.invalidProviderOutput', `${decisionPath}.recoverable`, 'must be a boolean')
    const normalized: NativeTextDecision = {
      code: code as NativeTextDecision['code'],
      message: providerString(decision.message, `${decisionPath}.message`, 4_096),
      recoverable: decision.recoverable,
    }
    if (decision.faceId !== undefined) normalized.faceId = providerString(decision.faceId, `${decisionPath}.faceId`, 128, PROVIDER_ID_RE)
    if (decision.startUtf16 !== undefined) normalized.startUtf16 = providerInteger(decision.startUtf16, `${decisionPath}.startUtf16`, 0, textLength)
    if (decision.endUtf16 !== undefined) normalized.endUtf16 = providerInteger(decision.endUtf16, `${decisionPath}.endUtf16`, 0, textLength)
    if ((normalized.startUtf16 === undefined) !== (normalized.endUtf16 === undefined) || (normalized.startUtf16 !== undefined && normalized.endUtf16 !== undefined && normalized.startUtf16 >= normalized.endUtf16)) {
      throw new RenderCompileError('render.invalidProviderOutput', decisionPath, 'decision text offsets must be supplied together as a non-empty ascending range')
    }
    return normalized
  })
}

function normalizeAttemptedFaceIds(value: unknown, state: CompileState, path: string): readonly string[] {
  if (!Array.isArray(value) || value.length > 4_096) throw new RenderCompileError('render.invalidProviderOutput', path, 'attemptedFaceIds must be an array with at most 4096 entries')
  state.budget.providerAttemptedFaceIds += value.length
  if (state.budget.providerAttemptedFaceIds > PPTX_RENDER_LIMITS.maxProviderAttemptedFaceIds) {
    throw new RenderCompileError('render.providerFaceBudget', path, `text providers exceed ${PPTX_RENDER_LIMITS.maxProviderAttemptedFaceIds} attempted face IDs`)
  }
  const normalized = value.map((item, index) => providerString(item, `${path}[${index}]`, 128, PROVIDER_ID_RE))
  if (new Set(normalized).size !== normalized.length) throw new RenderCompileError('render.invalidProviderOutput', path, 'attemptedFaceIds must not contain duplicates')
  return Object.freeze(normalized)
}

function normalizeFace(value: unknown, path: string): ResolvedFontFace {
  const face = providerRecord(value, path)
  exactProviderKeys(face, ['faceId', 'family', 'postscriptName', 'weight', 'style', 'stretch', 'sourceKind', 'resourceId', 'contentDigest', 'collectionIndex', 'resolution', 'matchedFamily', 'fallbackChainId'], path)
  const style = providerString(face.style, `${path}.style`, 16)
  if (!['normal', 'italic', 'oblique'].includes(style)) throw new RenderCompileError('render.invalidProviderOutput', `${path}.style`, 'unsupported font style')
  const sourceKind = providerString(face.sourceKind, `${path}.sourceKind`, 16)
  if (!['document', 'bundled', 'system', 'host'].includes(sourceKind)) throw new RenderCompileError('render.invalidProviderOutput', `${path}.sourceKind`, 'unsupported font source kind')
  const resolution = providerString(face.resolution, `${path}.resolution`, 16)
  if (!['exact', 'substitute', 'fallback'].includes(resolution)) throw new RenderCompileError('render.invalidProviderOutput', `${path}.resolution`, 'unsupported font resolution kind')
  const normalized: ResolvedFontFace = {
    faceId: providerString(face.faceId, `${path}.faceId`, 128, PROVIDER_ID_RE),
    family: providerString(face.family, `${path}.family`),
    weight: providerInteger(face.weight, `${path}.weight`, 1, 1_000),
    style: style as ResolvedFontFace['style'],
    stretch: providerInteger(face.stretch, `${path}.stretch`, 50, 200),
    sourceKind: sourceKind as ResolvedFontFace['sourceKind'],
    resourceId: providerString(face.resourceId, `${path}.resourceId`, 128, PROVIDER_ID_RE),
    contentDigest: providerString(face.contentDigest, `${path}.contentDigest`, 71, PROVIDER_DIGEST_RE) as ResolvedFontFace['contentDigest'],
    resolution: resolution as ResolvedFontFace['resolution'],
    matchedFamily: providerString(face.matchedFamily, `${path}.matchedFamily`),
  }
  if (face.postscriptName !== undefined) normalized.postscriptName = providerString(face.postscriptName, `${path}.postscriptName`)
  if (face.collectionIndex !== undefined) normalized.collectionIndex = providerInteger(face.collectionIndex, `${path}.collectionIndex`, 0, 65_535)
  if (face.fallbackChainId !== undefined) normalized.fallbackChainId = providerString(face.fallbackChainId, `${path}.fallbackChainId`, 128, PROVIDER_ID_RE)
  return Object.freeze(normalized)
}

function sameFace(left: ResolvedFontFace, right: ResolvedFontFace): boolean {
  return left.faceId === right.faceId
    && left.family === right.family
    && left.postscriptName === right.postscriptName
    && left.weight === right.weight
    && left.style === right.style
    && left.stretch === right.stretch
    && left.sourceKind === right.sourceKind
    && left.resourceId === right.resourceId
    && left.contentDigest === right.contentDigest
    && left.collectionIndex === right.collectionIndex
    && left.resolution === right.resolution
    && left.matchedFamily === right.matchedFamily
    && left.fallbackChainId === right.fallbackChainId
}

function normalizeRefusal(value: unknown, state: CompileState, path: string, textLength: number): NativeTextRefusal {
  const refusal = providerRecord(value, path)
  exactProviderKeys(refusal, ['status', 'decisions', 'attemptedFaceIds'], path)
  if (refusal.status !== 'refused') throw new RenderCompileError('render.invalidProviderOutput', `${path}.status`, 'must equal refused')
  return {
    status: 'refused',
    decisions: normalizeDecisions(refusal.decisions, state, `${path}.decisions`, textLength),
    attemptedFaceIds: normalizeAttemptedFaceIds(refusal.attemptedFaceIds, state, `${path}.attemptedFaceIds`),
  }
}

type NormalizedResolution =
  | { readonly status: 'resolved'; readonly face: ResolvedFontFace; readonly attemptedFaceIds: readonly string[]; readonly decisions: readonly NativeTextDecision[] }
  | NativeTextRefusal

function normalizeResolution(value: unknown, state: CompileState, path: string, textLength: number): NormalizedResolution {
  const resolution = providerRecord(value, path)
  if (resolution.status === 'refused') return normalizeRefusal(resolution, state, path, textLength)
  exactProviderKeys(resolution, ['status', 'face', 'attemptedFaceIds', 'decisions'], path)
  if (resolution.status !== 'resolved') throw new RenderCompileError('render.invalidProviderOutput', `${path}.status`, 'must equal resolved or refused')
  return {
    status: 'resolved',
    face: normalizeFace(resolution.face, `${path}.face`),
    attemptedFaceIds: normalizeAttemptedFaceIds(resolution.attemptedFaceIds, state, `${path}.attemptedFaceIds`),
    decisions: normalizeDecisions(resolution.decisions, state, `${path}.decisions`, textLength),
  }
}

function normalizeFontMetrics(value: unknown, path: string): FontResource['metrics'] {
  const metrics = providerRecord(value, path)
  exactProviderKeys(metrics, ['unitsPerEm', 'ascender', 'descender', 'lineGap', 'capHeight', 'xHeight', 'underlinePosition', 'underlineThickness'], path)
  const normalized: FontResource['metrics'] = {
    unitsPerEm: providerInteger(metrics.unitsPerEm, `${path}.unitsPerEm`, 1, 1_000_000),
    ascender: providerInteger(metrics.ascender, `${path}.ascender`, 0, 1_000_000_000),
    descender: providerInteger(metrics.descender, `${path}.descender`, -1_000_000_000, 0),
    lineGap: providerInteger(metrics.lineGap, `${path}.lineGap`, 0, 1_000_000_000),
  }
  if (metrics.capHeight !== undefined) normalized.capHeight = providerInteger(metrics.capHeight, `${path}.capHeight`, 0, 1_000_000_000)
  if (metrics.xHeight !== undefined) normalized.xHeight = providerInteger(metrics.xHeight, `${path}.xHeight`, 0, 1_000_000_000)
  if (metrics.underlinePosition !== undefined) normalized.underlinePosition = providerInteger(metrics.underlinePosition, `${path}.underlinePosition`, -1_000_000_000, 1_000_000_000)
  if (metrics.underlineThickness !== undefined) normalized.underlineThickness = providerInteger(metrics.underlineThickness, `${path}.underlineThickness`, 0, 1_000_000_000)
  return normalized
}

function normalizeFontLoad(value: unknown, state: CompileState, path: string, textLength: number): FontResource | NativeTextRefusal {
  const resource = providerRecord(value, path)
  if ('status' in resource) return normalizeRefusal(resource, state, path, textLength)
  exactProviderKeys(resource, ['face', 'bytes', 'metrics'], path)
  const bytes = resource.bytes
  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0 || bytes.byteLength > PPTX_RENDER_LIMITS.maxFontResourceBytes) {
    throw new RenderCompileError('render.invalidProviderOutput', `${path}.bytes`, `must be a non-empty Uint8Array no larger than ${PPTX_RENDER_LIMITS.maxFontResourceBytes} bytes`)
  }
  return Object.freeze({ face: normalizeFace(resource.face, `${path}.face`), bytes: Uint8Array.prototype.slice.call(bytes), metrics: Object.freeze(normalizeFontMetrics(resource.metrics, `${path}.metrics`)) })
}

function normalizeLineMetrics(value: unknown, path: string): ScaledLineMetrics {
  const metrics = providerRecord(value, path)
  exactProviderKeys(metrics, ['fontSizeMilliPoints', 'ascentMilliPoints', 'descentMilliPoints', 'lineGapMilliPoints', 'lineHeightMilliPoints', 'capHeightMilliPoints', 'xHeightMilliPoints', 'underlinePositionMilliPoints', 'underlineThicknessMilliPoints'], path)
  const normalized: ScaledLineMetrics = {
    fontSizeMilliPoints: providerInteger(metrics.fontSizeMilliPoints, `${path}.fontSizeMilliPoints`, 1, 10_000_000),
    ascentMilliPoints: providerInteger(metrics.ascentMilliPoints, `${path}.ascentMilliPoints`),
    descentMilliPoints: providerInteger(metrics.descentMilliPoints, `${path}.descentMilliPoints`),
    lineGapMilliPoints: providerInteger(metrics.lineGapMilliPoints, `${path}.lineGapMilliPoints`),
    lineHeightMilliPoints: providerInteger(metrics.lineHeightMilliPoints, `${path}.lineHeightMilliPoints`, 1),
  }
  for (const key of ['capHeightMilliPoints', 'xHeightMilliPoints', 'underlinePositionMilliPoints', 'underlineThicknessMilliPoints'] as const) {
    if (metrics[key] !== undefined) normalized[key] = providerInteger(metrics[key], `${path}.${key}`)
  }
  return normalized
}

function normalizeSegment(value: unknown, state: CompileState, path: string): ShapedSegment {
  const segment = providerRecord(value, path)
  exactProviderKeys(segment, ['startUtf16', 'endUtf16', 'face', 'glyphs', 'clusters', 'metrics', 'advanceInlineMilliPoints', 'advanceBlockMilliPoints'], path)
  if (!Array.isArray(segment.glyphs) || segment.glyphs.length > state.budget.maxGlyphs - state.budget.glyphs) throw new RenderCompileError('render.glyphBudget', `${path}.glyphs`, 'provider glyph array exceeds remaining RenderTree budget')
  if (!Array.isArray(segment.clusters) || segment.clusters.length > state.budget.maxClusters - state.budget.clusters) throw new RenderCompileError('render.clusterBudget', `${path}.clusters`, 'provider cluster array exceeds remaining RenderTree budget')
  const glyphs: ShapedGlyph[] = segment.glyphs.map((item, index) => {
    const glyphPath = `${path}.glyphs[${index}]`
    const glyph = providerRecord(item, glyphPath)
    exactProviderKeys(glyph, ['glyphId', 'clusterIndex', 'advanceXMilliPoints', 'advanceYMilliPoints', 'offsetXMilliPoints', 'offsetYMilliPoints'], glyphPath)
    return {
      glyphId: providerInteger(glyph.glyphId, `${glyphPath}.glyphId`, 0),
      clusterIndex: providerInteger(glyph.clusterIndex, `${glyphPath}.clusterIndex`, 0),
      advanceXMilliPoints: providerInteger(glyph.advanceXMilliPoints, `${glyphPath}.advanceXMilliPoints`),
      advanceYMilliPoints: providerInteger(glyph.advanceYMilliPoints, `${glyphPath}.advanceYMilliPoints`),
      offsetXMilliPoints: providerInteger(glyph.offsetXMilliPoints, `${glyphPath}.offsetXMilliPoints`),
      offsetYMilliPoints: providerInteger(glyph.offsetYMilliPoints, `${glyphPath}.offsetYMilliPoints`),
    }
  })
  const clusters: ShapedCluster[] = segment.clusters.map((item, index) => {
    const clusterPath = `${path}.clusters[${index}]`
    const cluster = providerRecord(item, clusterPath)
    exactProviderKeys(cluster, ['startUtf16', 'endUtf16', 'glyphStart', 'glyphEnd', 'advanceInlineMilliPoints', 'unsafeToBreak', 'whitespace'], clusterPath)
    const normalized: ShapedCluster = {
      startUtf16: providerInteger(cluster.startUtf16, `${clusterPath}.startUtf16`, 0),
      endUtf16: providerInteger(cluster.endUtf16, `${clusterPath}.endUtf16`, 0),
      glyphStart: providerInteger(cluster.glyphStart, `${clusterPath}.glyphStart`, 0),
      glyphEnd: providerInteger(cluster.glyphEnd, `${clusterPath}.glyphEnd`, 0),
      advanceInlineMilliPoints: providerInteger(cluster.advanceInlineMilliPoints, `${clusterPath}.advanceInlineMilliPoints`, 0),
    }
    if (cluster.unsafeToBreak !== undefined) {
      if (typeof cluster.unsafeToBreak !== 'boolean') throw new RenderCompileError('render.invalidProviderOutput', `${clusterPath}.unsafeToBreak`, 'must be a boolean')
      normalized.unsafeToBreak = cluster.unsafeToBreak
    }
    if (cluster.whitespace !== undefined) {
      if (typeof cluster.whitespace !== 'boolean') throw new RenderCompileError('render.invalidProviderOutput', `${clusterPath}.whitespace`, 'must be a boolean')
      normalized.whitespace = cluster.whitespace
    }
    return normalized
  })
  return {
    startUtf16: providerInteger(segment.startUtf16, `${path}.startUtf16`, 0),
    endUtf16: providerInteger(segment.endUtf16, `${path}.endUtf16`, 0),
    face: normalizeFace(segment.face, `${path}.face`),
    glyphs,
    clusters,
    metrics: normalizeLineMetrics(segment.metrics, `${path}.metrics`),
    advanceInlineMilliPoints: providerInteger(segment.advanceInlineMilliPoints, `${path}.advanceInlineMilliPoints`, 0),
    advanceBlockMilliPoints: providerInteger(segment.advanceBlockMilliPoints, `${path}.advanceBlockMilliPoints`),
  }
}

function normalizeShaperOutput(value: unknown, state: CompileState, path: string, textLength: number): ShapedSegment | NativeTextRefusal {
  const output = providerRecord(value, path)
  if ('status' in output) return normalizeRefusal(output, state, path, textLength)
  return normalizeSegment(output, state, path)
}

function snapshotFontManifest(manifest: NativeFontManifest): NativeFontManifest {
  const faces = manifest.faces.map((face) => Object.freeze({
    faceId: face.faceId,
    family: face.family,
    ...(face.postscriptName === undefined ? {} : { postscriptName: face.postscriptName }),
    weight: face.weight,
    style: face.style,
    stretch: face.stretch,
    ...(face.aliases ? { aliases: Object.freeze([...face.aliases]) } : {}),
    ...(face.scripts ? { scripts: Object.freeze([...face.scripts]) } : {}),
    ...(face.languages ? { languages: Object.freeze([...face.languages]) } : {}),
    source: Object.freeze({
      kind: face.source.kind,
      resourceId: face.source.resourceId,
      ...(face.source.contentDigest === undefined ? {} : { contentDigest: face.source.contentDigest }),
      ...(face.source.collectionIndex === undefined ? {} : { collectionIndex: face.source.collectionIndex }),
    }),
  }))
  const fallbackChains = manifest.fallbackChains.map((chain) => Object.freeze({
    ...chain,
    faceIds: Object.freeze([...chain.faceIds]),
    ...(chain.scripts ? { scripts: Object.freeze([...chain.scripts]) } : {}),
    ...(chain.languages ? { languages: Object.freeze([...chain.languages]) } : {}),
  }))
  return Object.freeze({ ...manifest, faces: Object.freeze(faces), fallbackChains: Object.freeze(fallbackChains) })
}

function snapshotProviders(layout: CompileSlideOptions['textLayout']): ProviderSnapshot {
  try {
    const resolverId = layout.resolver.providerId
    const resolverRevision = layout.resolver.providerRevision
    const shaperId = layout.shaper.providerId
    const shaperRevision = layout.shaper.providerRevision
    const resolve = layout.resolver.resolve
    const load = layout.resolver.load
    const shape = layout.shaper.shape
    for (const [path, value] of [['resolver.providerId', resolverId], ['resolver.providerRevision', resolverRevision], ['shaper.providerId', shaperId], ['shaper.providerRevision', shaperRevision]] as const) {
      if (typeof value !== 'string' || !PROVIDER_ID_RE.test(value)) throw new RenderCompileError('render.invalidProvider', `$.options.textLayout.${path}`, 'provider identity must be a bounded stable identifier')
    }
    if (typeof resolve !== 'function' || typeof load !== 'function' || typeof shape !== 'function') throw new RenderCompileError('render.invalidProvider', '$.options.textLayout', 'providers must expose callable resolve/load/shape functions')
    if (layout.resolver.providerId !== resolverId || layout.resolver.providerRevision !== resolverRevision || layout.shaper.providerId !== shaperId || layout.shaper.providerRevision !== shaperRevision || layout.resolver.resolve !== resolve || layout.resolver.load !== load || layout.shaper.shape !== shape) throw new RenderCompileError('render.invalidProvider', '$.options.textLayout', 'provider identity and callable access must be stable')
    return {
      resolver: { providerId: resolverId, providerRevision: resolverRevision, live: layout.resolver, resolve: resolve.bind(layout.resolver), load: load.bind(layout.resolver) },
      shaper: { providerId: shaperId, providerRevision: shaperRevision, live: layout.shaper, shape: shape.bind(layout.shaper) },
    }
  } catch (error) {
    if (error instanceof RenderCompileError) throw error
    throw new RenderCompileError('render.invalidProvider', '$.options.textLayout', 'provider identity access failed')
  }
}

function providerIdentityStable(providers: ProviderSnapshot): boolean {
  try {
    return providers.resolver.live.providerId === providers.resolver.providerId
      && providers.resolver.live.providerRevision === providers.resolver.providerRevision
      && providers.shaper.live.providerId === providers.shaper.providerId
      && providers.shaper.live.providerRevision === providers.shaper.providerRevision
  } catch { return false }
}

function normalizedFamily(value: string): string {
  return value.normalize('NFKC').trim().toLowerCase()
}

function faceBackedByManifest(face: ResolvedFontFace, manifest: NativeFontManifest, run: TextRunInput): boolean {
  const declared = manifest.faces.find((candidate) => candidate.faceId === face.faceId)
  if (!declared) return false
  const matchedFamily = run.font.families.some((family) => normalizedFamily(family) === normalizedFamily(face.matchedFamily))
  const fallbackValid = face.fallbackChainId === undefined || manifest.fallbackChains.some((chain) => chain.chainId === face.fallbackChainId && chain.faceIds.includes(face.faceId))
  return matchedFamily && fallbackValid
    && declared.family === face.family
    && declared.postscriptName === face.postscriptName
    && declared.weight === face.weight
    && declared.style === face.style
    && declared.stretch === face.stretch
    && declared.source.kind === face.sourceKind
    && declared.source.resourceId === face.resourceId
    && declared.source.collectionIndex === face.collectionIndex
    // A stable manifest revision is not a byte attestation. Native layout only
    // accepts faces whose expected digest is explicit in the snapshotted
    // manifest and then independently verified against the loaded bytes.
    && declared.source.contentDigest !== undefined
    && declared.source.contentDigest === face.contentDigest
}

function resolvedFaceCacheKey(face: ResolvedFontFace): string {
  return JSON.stringify([face.faceId, face.family, face.postscriptName ?? null, face.weight, face.style, face.stretch, face.sourceKind, face.resourceId, face.contentDigest, face.collectionIndex ?? null, face.resolution, face.matchedFamily, face.fallbackChainId ?? null])
}

function fontDigest(bytes: Uint8Array): string {
  return `sha256:${bytesToHex(sha256(bytes))}`
}

function sameMetrics(left: ScaledLineMetrics, right: ScaledLineMetrics): boolean {
  return left.fontSizeMilliPoints === right.fontSizeMilliPoints
    && left.ascentMilliPoints === right.ascentMilliPoints
    && left.descentMilliPoints === right.descentMilliPoints
    && left.lineGapMilliPoints === right.lineGapMilliPoints
    && left.lineHeightMilliPoints === right.lineHeightMilliPoints
    && left.capHeightMilliPoints === right.capHeightMilliPoints
    && left.xHeightMilliPoints === right.xHeightMilliPoints
    && left.underlinePositionMilliPoints === right.underlinePositionMilliPoints
    && left.underlineThicknessMilliPoints === right.underlineThicknessMilliPoints
}

function decisionsBlockLayout(decisions: readonly NativeTextDecision[]): boolean {
  return decisions.some((decision) => !decision.recoverable || decision.code === 'missing-glyph' || decision.code.startsWith('unsupported-') || decision.code === 'provider-failure' || decision.code === 'font-digest-mismatch' || decision.code === 'font-metrics-unavailable')
}

function boundedProviderError(error: unknown): string {
  let message = 'unknown provider error'
  try {
    if (typeof error === 'string') message = error
    else if (error !== null && (typeof error === 'object' || typeof error === 'function')) {
      const value = Reflect.get(error, 'message')
      if (typeof value === 'string') message = value
    }
  } catch { message = 'unreadable provider error' }
  return message.length <= 2_048 ? message : `${message.slice(0, 2_047)}…`
}

function milliPointsToEmu(value: number, path: string): number {
  if (!Number.isSafeInteger(value)) throw new RenderCompileError('render.textMetric', path, 'text metric must be an integer')
  const numerator = value * 127
  if (!Number.isSafeInteger(numerator)) throw new RenderCompileError('render.textMetric', path, 'text metric exceeds integer precision')
  const result = Math.round(numerator / 10)
  return Object.is(result, -0) ? 0 : result
}

function boundedPath(path: readonly RenderPathCommand[], sourcePath: string) {
  if (path.length > PPTX_RENDER_LIMITS.maxPathCommands) {
    throw new RenderCompileError('render.pathBudget', sourcePath, `path exceeds ${PPTX_RENDER_LIMITS.maxPathCommands} commands`)
  }
  return path
}

function elementBase(element: NativeElement, zIndex: number, budget: Budget, clipToElement = true) {
  const { x, y, cx, cy } = element.transform
  checkCoordinate(x, `$.elements.${element.id}.transform.x`, budget)
  checkCoordinate(y, `$.elements.${element.id}.transform.y`, budget)
  checkCoordinate(cx, `$.elements.${element.id}.transform.cx`, budget, true)
  checkCoordinate(cy, `$.elements.${element.id}.transform.cy`, budget, true)
  return {
    sourceElementId: element.id,
    sourceKind: element.kind,
    zIndex,
    transform: quarterTurnTransform(element.transform),
    bounds: localBounds(cx, cy),
    ...(clipToElement ? { clip: { kind: 'rect' as const, rect: localBounds(cx, cy) } } : {}),
    compatibility: element.compatibility.status,
  }
}

function nativeTextBodyBounds(element: Extract<NativeElement, { kind: 'text' | 'shape' }>, state: CompileState): RenderRect {
  const layout = element.textBody
  if (!layout) return localBounds(element.transform.cx, element.transform.cy)
  const region = element.kind === 'shape' && element.preset === 'pentagon'
    ? defaultPentagonTextRect(element.transform.cx, element.transform.cy)
    : localBounds(element.transform.cx, element.transform.cy)
  const bounds = {
    x: region.x + layout.leftInsetEmu,
    y: region.y + layout.topInsetEmu,
    cx: region.cx - layout.leftInsetEmu - layout.rightInsetEmu,
    cy: region.cy - layout.topInsetEmu - layout.bottomInsetEmu,
  }
  checkCoordinate(bounds.x, `$.elements.${element.id}.textBody.leftInsetEmu`, state.budget)
  checkCoordinate(bounds.y, `$.elements.${element.id}.textBody.topInsetEmu`, state.budget)
  checkCoordinate(bounds.cx, `$.elements.${element.id}.textBody`, state.budget, true)
  checkCoordinate(bounds.cy, `$.elements.${element.id}.textBody`, state.budget, true)
  return bounds
}

function nativeTableCellTextBodyBounds(elementId: string, rowIndex: number, columnIndex: number, width: number, height: number, layout: NativeTextBodyLayout, state: CompileState): RenderRect {
  const path = `$.elements.${elementId}.table.rows[${rowIndex}][${columnIndex}].textBody`
  const bounds = {
    x: layout.leftInsetEmu,
    y: layout.topInsetEmu,
    cx: width - layout.leftInsetEmu - layout.rightInsetEmu,
    cy: height - layout.topInsetEmu - layout.bottomInsetEmu,
  }
  checkCoordinate(bounds.x, `${path}.leftInsetEmu`, state.budget)
  checkCoordinate(bounds.y, `${path}.topInsetEmu`, state.budget)
  checkCoordinate(bounds.cx, path, state.budget, true)
  checkCoordinate(bounds.cy, path, state.budget, true)
  return bounds
}

function exactGroupBase(element: Extract<NativeElement, { kind: 'group' }>, zIndex: number, budget: Budget) {
  const { x, y, cx, cy } = element.transform
  checkCoordinate(x, `$.elements.${element.id}.transform.x`, budget)
  checkCoordinate(y, `$.elements.${element.id}.transform.y`, budget)
  checkCoordinate(cx, `$.elements.${element.id}.transform.cx`, budget, true)
  checkCoordinate(cy, `$.elements.${element.id}.transform.cy`, budget, true)
  if (!element.childTransform) {
    // Preserve the original v1 authored-group identity semantics, including
    // its legacy clip. Parsed DrawingML groups always carry childTransform.
    return elementBase(element, zIndex, budget)
  }
  const child = element.childTransform
  checkCoordinate(child.x, `$.elements.${element.id}.childTransform.x`, budget)
  checkCoordinate(child.y, `$.elements.${element.id}.childTransform.y`, budget)
  checkCoordinate(child.cx, `$.elements.${element.id}.childTransform.cx`, budget, true)
  checkCoordinate(child.cy, `$.elements.${element.id}.childTransform.cy`, budget, true)
  const ppm = 1_000_000n
  const exactScale = (extent: number, childExtent: number, path: string): bigint => {
    const numerator = BigInt(extent) * ppm
    const divisor = BigInt(childExtent)
    if (numerator % divisor !== 0n) throw new RenderCompileError('render.groupTransform', path, 'group scale is not exact integer PPM')
    const result = numerator / divisor
    if (result <= 0n || result > BigInt(PPTX_RENDER_LIMITS.maxAffinePpm)) throw new RenderCompileError('render.groupTransform', path, 'group scale exceeds the bounded affine coefficient range')
    return result
  }
  const exactTranslation = (offset: number, childOffset: number, scale: bigint, path: string): bigint => {
    const product = BigInt(childOffset) * scale
    if (product % ppm !== 0n) throw new RenderCompileError('render.groupTransform', path, 'group translation is not exact integer EMU')
    const result = BigInt(offset) - product / ppm
    if (result < BigInt(-budget.maxCoordinateEmu) || result > BigInt(budget.maxCoordinateEmu)) throw new RenderCompileError('render.groupTransform', path, 'group translation exceeds the render coordinate budget')
    return result
  }
  const scaleX = exactScale(cx, child.cx, `$.elements.${element.id}.childTransform.cx`)
  const scaleY = exactScale(cy, child.cy, `$.elements.${element.id}.childTransform.cy`)
  const tx = exactTranslation(x, child.x, scaleX, `$.elements.${element.id}.childTransform.x`)
  const ty = exactTranslation(y, child.y, scaleY, `$.elements.${element.id}.childTransform.y`)
  return {
    sourceElementId: element.id,
    sourceKind: element.kind,
    zIndex,
    transform: { aPpm: Number(scaleX), bPpm: 0, cPpm: 0, dPpm: Number(scaleY), txEmu: Number(tx), tyEmu: Number(ty) },
    bounds: { x: child.x, y: child.y, cx: child.cx, cy: child.cy },
    compatibility: element.compatibility.status,
  }
}

function checkedWorldAffine(
  parent: WorldAffine,
  local: Readonly<{ aPpm: number; bPpm: number; cPpm: number; dPpm: number; txEmu: number; tyEmu: number }>,
  bounds: Readonly<{ x: number; y: number; cx: number; cy: number }>,
  path: string,
  budget: Budget,
): WorldAffine {
  const scale=local.bPpm===0&&local.cPpm===0&&local.aPpm>0&&local.dPpm>0
  const half=local.bPpm===0&&local.cPpm===0&&local.aPpm===-1000000&&local.dPpm===-1000000
  const quarter=local.aPpm===0&&local.dPpm===0&&local.bPpm===-local.cPpm&&Math.abs(local.bPpm)===1000000
  if (!scale&&!half&&!quarter) {
    throw new RenderCompileError('render.worldTransform', `${path}.transform`, 'only exact scale/translation and source quarter turns are supported')
  }
  const localScale = (value: number, componentPath: string): ExactRational => {
    if (!Number.isSafeInteger(value) || Math.abs(value) > PPTX_RENDER_LIMITS.maxAffinePpm) {
      throw new RenderCompileError('render.worldTransform', componentPath, 'affine coefficient exceeds the safe integer range')
    }
    return exactRational(BigInt(value), AFFINE_PPM)
  }
  const withinCoordinateBudget = (value: ExactRational): boolean => {
    const limit = BigInt(budget.maxCoordinateEmu) * value.denominator
    return value.numerator >= -limit && value.numerator <= limit
  }
  const a=localScale(local.aPpm,`${path}.transform.aPpm`),b=localScale(local.bPpm,`${path}.transform.bPpm`),c=localScale(local.cPpm,`${path}.transform.cPpm`),d=localScale(local.dPpm,`${path}.transform.dPpm`)
  const pair=(a:ExactRational,b:ExactRational,c:ExactRational,d:ExactRational)=>addRational(multiplyRational(a,b),multiplyRational(c,d))
  const world: WorldAffine = {
    a:pair(parent.a,a,parent.c,b), b:pair(parent.b,a,parent.d,b),
    c:pair(parent.a,c,parent.c,d), d:pair(parent.b,c,parent.d,d),
    tx:addRational(parent.tx,addRational(multiplyRationalInteger(parent.a,local.txEmu),multiplyRationalInteger(parent.c,local.tyEmu))),
    ty:addRational(parent.ty,addRational(multiplyRationalInteger(parent.b,local.txEmu),multiplyRationalInteger(parent.d,local.tyEmu))),
  }
  if (!withinCoordinateBudget(world.tx)) throw new RenderCompileError('render.worldTransform', `${path}.transform.txEmu`, 'cumulative translation exceeds the render coordinate budget')
  if (!withinCoordinateBudget(world.ty)) throw new RenderCompileError('render.worldTransform', `${path}.transform.tyEmu`, 'cumulative translation exceeds the render coordinate budget')
  const coordinate = (a:ExactRational,b:ExactRational,translation:ExactRational,x:number,y:number,componentPath:string): void => {
    const result=addRational(translation,addRational(multiplyRationalInteger(a,x),multiplyRationalInteger(b,y)))
    if (!withinCoordinateBudget(result)) throw new RenderCompileError('render.worldTransform', componentPath, 'world-space bound exceeds the render coordinate budget')
  }
  for(const x of [bounds.x,bounds.x+bounds.cx])for(const y of [bounds.y,bounds.y+bounds.cy]){
    coordinate(world.a,world.c,world.tx,x,y,`${path}.bounds.x`)
    coordinate(world.b,world.d,world.ty,x,y,`${path}.bounds.y`)
  }
  return world
}

function copyNativeDiagnostics(state: CompileState, compatibility: NativeCompatibility, elementId?: string): void {
  for (const diagnostic of compatibility.diagnostics) {
    state.diagnostics.push({
      severity: diagnostic.severity,
      code: 'native.compatibility',
      sourceCode: diagnostic.code,
      message: diagnostic.message,
      slideId: state.slide.id,
      elementId,
    })
  }
}

function placeholder(element: NativeElement, zIndex: number, reason: RenderPlaceholderNode['reason'], label: string, state: CompileState): RenderPlaceholderNode {
  takeNode(state, `$.elements.${element.id}`)
  return { kind: 'placeholder', ...elementBase(element, zIndex, state.budget), reason, label }
}

function referenceAsset(assetId: string, elementId: string, state: CompileState): NativeAsset {
  const asset = state.assets.get(assetId)
  if (!asset) throw new RenderCompileError('render.assetReference', `$.elements.${elementId}`, `asset ${assetId} is missing after native validation`)
  if (asset.byteLength > PPTX_RENDER_LIMITS.maxReferencedAssetBytes) {
    throw new RenderCompileError('render.assetBudget', `$.assets.${asset.id}`, `asset exceeds ${PPTX_RENDER_LIMITS.maxReferencedAssetBytes} bytes`)
  }
  if (!state.referencedAssets.has(asset.id)) {
    state.referencedAssets.set(asset.id, {
      id: asset.id,
      contentType: asset.contentType,
      sha256: asset.sha256,
      byteLength: asset.byteLength,
      resolutionSource: asset.dataBase64 === undefined ? 'host' : 'sourceDeck',
    })
  }
  if (asset.dataBase64 === undefined && !state.hostResolutionDiagnosticElements.has(elementId)) {
    state.hostResolutionDiagnosticElements.add(elementId)
    state.diagnostics.push({
      severity: 'info',
      code: 'asset.hostResolutionRequired',
      message: `asset ${asset.id} requires host resolution by stable ID and digest`,
      slideId: state.slide.id,
      elementId,
    })
  }
  return asset
}

function validateSegment(segment: ShapedSegment, run: TextRunInput, expectedFace: ResolvedFontFace, path: string): void {
  if (segment.startUtf16 !== 0 || segment.endUtf16 !== run.text.length || !sameFace(segment.face, expectedFace)) {
    throw new RenderCompileError('render.invalidShaping', path, 'shaper returned a segment for a different range or face')
  }
  for (const [name, value] of Object.entries({
    advanceInlineMilliPoints: segment.advanceInlineMilliPoints,
    advanceBlockMilliPoints: segment.advanceBlockMilliPoints,
    ascentMilliPoints: segment.metrics.ascentMilliPoints,
    descentMilliPoints: segment.metrics.descentMilliPoints,
    lineGapMilliPoints: segment.metrics.lineGapMilliPoints,
    lineHeightMilliPoints: segment.metrics.lineHeightMilliPoints,
  })) {
    if (!Number.isSafeInteger(value)) throw new RenderCompileError('render.invalidShaping', path, `${name} must be a safe integer`)
  }
  if (segment.advanceInlineMilliPoints < 0 || segment.metrics.lineHeightMilliPoints < 1) {
    throw new RenderCompileError('render.invalidShaping', path, 'inline advance and line height must be non-negative/positive')
  }
  const naturalLineHeight = segment.metrics.ascentMilliPoints - segment.metrics.descentMilliPoints + segment.metrics.lineGapMilliPoints
  if (
    segment.metrics.ascentMilliPoints < 0 ||
    segment.metrics.descentMilliPoints > 0 ||
    segment.metrics.lineGapMilliPoints < 0 ||
    !Number.isSafeInteger(naturalLineHeight) ||
    naturalLineHeight !== segment.metrics.lineHeightMilliPoints
  ) {
    throw new RenderCompileError('render.invalidShaping', `${path}.metrics`, 'line metrics must be non-negative ascent/gap, non-positive descent, and an exact natural line height')
  }
  if (segment.metrics.fontSizeMilliPoints !== run.fontSizeMilliPoints) {
    throw new RenderCompileError('render.invalidShaping', path, 'shaper returned metrics for a different font size')
  }
  for (let index = 0; index < segment.glyphs.length; index++) {
    const glyph = segment.glyphs[index]!
    if (![glyph.glyphId, glyph.clusterIndex, glyph.advanceXMilliPoints, glyph.advanceYMilliPoints, glyph.offsetXMilliPoints, glyph.offsetYMilliPoints].every(Number.isSafeInteger)) {
      throw new RenderCompileError('render.invalidShaping', `${path}.glyphs[${index}]`, 'glyph fields must be safe integers')
    }
    if (glyph.glyphId < 0 || glyph.clusterIndex < 0 || glyph.clusterIndex >= segment.clusters.length) {
      throw new RenderCompileError('render.invalidShaping', `${path}.glyphs[${index}]`, 'glyph identity or cluster reference is invalid')
    }
  }
  if (run.text.length > 0 && segment.clusters.length === 0) {
    throw new RenderCompileError('render.invalidShaping', `${path}.clusters`, 'non-empty text must have complete cluster coverage')
  }
  if (run.text.length === 0 && (segment.clusters.length !== 0 || segment.glyphs.length !== 0)) {
    throw new RenderCompileError('render.invalidShaping', path, 'empty text cannot return clusters or glyphs')
  }
  const descendingText = run.direction === 'rtl' || run.direction === 'btt'
  let textBoundary = descendingText ? run.text.length : 0
  let previousGlyphEnd = 0
  let clusterAdvance = 0
  for (let index = 0; index < segment.clusters.length; index++) {
    const cluster = segment.clusters[index]!
    if (![cluster.startUtf16, cluster.endUtf16, cluster.glyphStart, cluster.glyphEnd, cluster.advanceInlineMilliPoints].every(Number.isSafeInteger)) {
      throw new RenderCompileError('render.invalidShaping', `${path}.clusters[${index}]`, 'cluster fields must be safe integers')
    }
    const textContiguous = descendingText ? cluster.endUtf16 === textBoundary : cluster.startUtf16 === textBoundary
    if (!textContiguous || cluster.endUtf16 <= cluster.startUtf16 || cluster.endUtf16 > run.text.length || cluster.glyphStart !== previousGlyphEnd || cluster.glyphEnd < cluster.glyphStart || cluster.glyphEnd > segment.glyphs.length || cluster.advanceInlineMilliPoints < 0) {
      throw new RenderCompileError('render.invalidShaping', `${path}.clusters[${index}]`, 'cluster ranges must contiguously cover directional text and paint-order glyphs with non-negative advances')
    }
    let hasInvisibleControl = false
    for (let textIndex = cluster.startUtf16; textIndex < cluster.endUtf16; textIndex++) {
      const unit = run.text.charCodeAt(textIndex)
      if (unit === 0x200b || unit === 0x2060 || unit === 0xfeff) hasInvisibleControl = true
    }
    const isolatedInvisibleControl = cluster.endUtf16 - cluster.startUtf16 === 1 && hasInvisibleControl
    if (hasInvisibleControl && (!isolatedInvisibleControl || cluster.glyphStart !== cluster.glyphEnd || cluster.advanceInlineMilliPoints !== 0)) {
      throw new RenderCompileError('render.invalidShaping', `${path}.clusters[${index}]`, 'ZWSP, WJ, and FEFF must be isolated glyphless zero-advance clusters')
    }
    if (cluster.glyphStart === cluster.glyphEnd && cluster.advanceInlineMilliPoints !== 0) {
      throw new RenderCompileError('render.invalidShaping', `${path}.clusters[${index}]`, 'a glyphless cluster must have zero inline advance')
    }
    if (!isUtf16Boundary(run.text, cluster.startUtf16) || !isUtf16Boundary(run.text, cluster.endUtf16)) {
      throw new RenderCompileError('render.invalidShaping', `${path}.clusters[${index}]`, 'cluster text range must use Unicode code-point boundaries')
    }
    for (let glyphIndex = cluster.glyphStart; glyphIndex < cluster.glyphEnd; glyphIndex++) {
      if (segment.glyphs[glyphIndex]!.clusterIndex !== index) {
        throw new RenderCompileError('render.invalidShaping', `${path}.glyphs[${glyphIndex}].clusterIndex`, 'glyph must reference the cluster whose glyph range owns it')
      }
    }
    textBoundary = descendingText ? cluster.startUtf16 : cluster.endUtf16
    previousGlyphEnd = cluster.glyphEnd
    clusterAdvance += cluster.advanceInlineMilliPoints
    if (!Number.isSafeInteger(clusterAdvance)) throw new RenderCompileError('render.invalidShaping', `${path}.clusters`, 'cluster advances exceed integer precision')
  }
  const completeText = descendingText ? textBoundary === 0 : textBoundary === run.text.length
  if (!completeText || previousGlyphEnd !== segment.glyphs.length) {
    throw new RenderCompileError('render.invalidShaping', `${path}.clusters`, 'clusters must completely cover returned text and glyphs')
  }
  if (clusterAdvance !== segment.advanceInlineMilliPoints) {
    throw new RenderCompileError('render.invalidShaping', `${path}.advanceInlineMilliPoints`, 'segment advance must equal the sum of cluster advances')
  }
}

function isUtf16Boundary(text: string, index: number): boolean {
  if (index <= 0 || index >= text.length) return true
  const before = text.charCodeAt(index - 1)
  const after = text.charCodeAt(index)
  return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff)
}

function refusedRun(
  nativeRun: NativeTextRun,
  input: TextRunInput,
  context: NativePptxTextRunContext,
  decisions: readonly NativeTextDecision[],
  attemptedFaceIds: readonly string[],
  color: string,
): ShapedRunResult {
  const lineHeightEmu = milliPointsToEmu(input.fontSizeMilliPoints, '$.text.refused.lineHeight')
  return {
    direction: input.direction,
    lineHeightEmu,
    ascentEmu: lineHeightEmu,
    ascentMilliPoints: input.fontSizeMilliPoints,
    descentMilliPoints: 0,
    lineGapMilliPoints: 0,
    run: {
      kind: 'textRun',
      sourceElementId: context.elementId,
      paragraphIndex: context.paragraphIndex,
      runIndex: context.runIndex,
      startUtf16: 0,
      endUtf16: nativeRun.text.length,
      text: nativeRun.text,
      direction: input.direction,
      fontSizeMilliPoints: input.fontSizeMilliPoints,
      color,
      bold: nativeRun.bold ?? false,
      italic: nativeRun.italic ?? false,
      x: 0,
      baselineY: lineHeightEmu,
      advanceInlineEmu: 0,
      lineHeightEmu,
      glyphs: [],
      clusters: [],
      decisions,
      attemptedFaceIds,
      status: 'refused',
    },
  }
}

async function shapeRun(nativeRun: NativeTextRun, context: NativePptxTextRunContext, state: CompileState, exactFamily?: string, symbolEncoding?: 'windows-symbol-byte-v1'): Promise<ShapedRunResult> {
  const override = state.resolveRun?.(context) ?? {}
  const families = exactFamily ? [exactFamily] : [...(override.fontFamilies ?? (nativeRun.fontFamily ? [nativeRun.fontFamily] : state.textDefaults.fontFamilies))]
  const fallbackChainIds = exactFamily ? undefined : override.fallbackChainIds ?? state.textDefaults.fallbackChainIds
  const input: TextRunInput = deepFreeze({
    version: NATIVE_TEXT_LAYOUT_VERSION,
    text: nativeRun.text,
    fontSizeMilliPoints: (nativeRun.fontSizeHundredthPt ?? state.textDefaults.fontSizeHundredthPt) * 10,
    font: {
      families,
      weight: nativeRun.bold ? 700 : 400,
      style: nativeRun.italic ? 'italic' : 'normal',
      stretch: 100,
      ...(fallbackChainIds === undefined ? {} : { fallbackChainIds: [...fallbackChainIds] }),
    },
    script: override.script ?? state.textDefaults.script,
    language: override.language ?? nativeRun.language ?? state.textDefaults.language,
    direction: override.direction ?? state.textDefaults.direction,
    features: override.features?.map((feature) => ({ ...feature })),
    variations: override.variations?.map((variation) => ({ ...variation })),
    letterSpacingMilliPoints: override.letterSpacingMilliPoints,
    wordSpacingMilliPoints: override.wordSpacingMilliPoints,
  })
  const inputValidation = validateTextRunInput(input)
  if (!inputValidation.ok) {
    throw new RenderCompileError('render.invalidTextInput', `$.elements.${context.elementId}.paragraphs[${context.paragraphIndex}].runs[${context.runIndex}]`, inputValidation.issues.map((issue) => `${issue.path} ${issue.message}`).join('; '))
  }
  const path = `$.elements.${context.elementId}.paragraphs[${context.paragraphIndex}].runs[${context.runIndex}]`
  if (input.direction === 'ttb' || input.direction === 'btt') {
    return refusedRun(nativeRun, input, context, [{
      code: 'unsupported-direction',
      message: 'vertical PPTX text layout is not modeled by the native render tree',
      recoverable: false,
    }], [], nativeRun.color ?? DEFAULT_COLOR)
  }
  const color = nativeRun.color ?? DEFAULT_COLOR
  let resolution: Exclude<NormalizedResolution, NativeTextRefusal>
  let loaded: FontResource
  let shaped: ShapedSegment
  let symbolEvidence: ReturnType<typeof qualifySymbolBullet> | undefined
  try {
    if (state.providerResolveCalls >= PPTX_RENDER_LIMITS.maxProviderResolveCalls) throw new TextBodyLayoutRefusal('text.providerBudget', `resolver calls exceed ${PPTX_RENDER_LIMITS.maxProviderResolveCalls}`)
    state.providerResolveCalls++
    const resolvedLive = await state.providers.resolver.resolve(Object.freeze({ manifest: state.fontManifest, run: input }))
    if (!providerIdentityStable(state.providers)) throw new TextBodyLayoutRefusal('text.refused', 'the injected provider changed its snapshotted identity during resolution')
    const normalizedResolution = normalizeResolution(resolvedLive, state, `${path}.resolution`, input.text.length)
    if (normalizedResolution.status === 'refused') throw new TextBodyLayoutRefusal('text.refused', 'the injected font resolver refused the complete native paragraph')
    if (exactFamily && (normalizedResolution.face.family !== exactFamily || normalizedResolution.face.resolution !== 'exact')) throw new TextBodyLayoutRefusal('text.refused', 'authored bullet requires its exact font face without substitution')
    if (!faceBackedByManifest(normalizedResolution.face, state.fontManifest, input)) throw new TextBodyLayoutRefusal('text.refused', 'the injected resolver returned a face that is not exactly backed by the snapshotted font manifest')
    if (decisionsBlockLayout(normalizedResolution.decisions)) throw new TextBodyLayoutRefusal('text.refused', 'the injected resolver reported a blocking shaping decision')
    resolution = normalizedResolution

    const cacheKey = resolvedFaceCacheKey(resolution.face)
    const cached = state.fontResources.get(cacheKey)
    if (cached) {
      if (!sameFace(cached.face, resolution.face)) throw new TextBodyLayoutRefusal('text.refused', 'the authoritative font cache identity collided')
      loaded = cached
    } else {
      if (state.providerLoadCalls >= PPTX_RENDER_LIMITS.maxProviderLoadCalls || state.fontResources.size >= PPTX_RENDER_LIMITS.maxCachedFontResources) throw new TextBodyLayoutRefusal('text.providerBudget', 'unique font loads exceed the bounded native font cache')
      state.providerLoadCalls++
      const loadedLive = await state.providers.resolver.load(resolution.face)
      if (!providerIdentityStable(state.providers)) throw new TextBodyLayoutRefusal('text.refused', 'the injected provider changed its snapshotted identity during font loading')
      const normalizedLoad = normalizeFontLoad(loadedLive, state, `${path}.fontResource`, input.text.length)
      if ('status' in normalizedLoad) throw new TextBodyLayoutRefusal('text.refused', 'the injected font resolver refused the complete native paragraph while loading bytes')
      if (!sameFace(normalizedLoad.face, resolution.face) || fontDigest(normalizedLoad.bytes) !== resolution.face.contentDigest) throw new TextBodyLayoutRefusal('text.refused', 'loaded font bytes, digest, or face do not match the resolved manifest face')
      if (state.uniqueFontBytes + normalizedLoad.bytes.byteLength > PPTX_RENDER_LIMITS.maxUniqueFontBytes) throw new TextBodyLayoutRefusal('text.providerBudget', `unique authoritative font bytes exceed ${PPTX_RENDER_LIMITS.maxUniqueFontBytes}`)
      state.uniqueFontBytes += normalizedLoad.bytes.byteLength
      state.fontResources.set(cacheKey, normalizedLoad)
      loaded = normalizedLoad
    }

    let providerResource = state.providerFontResources.get(cacheKey)
    if (providerResource) {
      if (!sameFace(providerResource.face, loaded.face)) throw new TextBodyLayoutRefusal('text.refused', 'the provider-facing font cache identity collided')
    } else {
      if (state.uniqueFontBytes + state.providerFacingFontBytes + loaded.bytes.byteLength > PPTX_RENDER_LIMITS.maxUniqueFontBytes) throw new TextBodyLayoutRefusal('text.providerBudget', `authoritative and provider-facing font bytes exceed ${PPTX_RENDER_LIMITS.maxUniqueFontBytes}`)
      providerResource = Object.freeze({ face: loaded.face, metrics: loaded.metrics, bytes: loaded.bytes.slice() })
      state.providerFacingFontBytes += providerResource.bytes.byteLength
      state.providerFontResources.set(cacheKey, providerResource)
    }
    if (state.providerShapeCalls >= PPTX_RENDER_LIMITS.maxProviderShapeCalls) throw new TextBodyLayoutRefusal('text.providerBudget', `shaper calls exceed ${PPTX_RENDER_LIMITS.maxProviderShapeCalls}`)
    if (fontDigest(providerResource.bytes) !== resolution.face.contentDigest) throw new TextBodyLayoutRefusal('text.refused', 'the isolated provider-facing font bytes were mutated before shaping')
    state.providerShapeCalls++
    if (symbolEncoding) {
      if (!exactFamily) throw new TextBodyLayoutRefusal('text.refused', 'symbol encoding requires an exact authored family')
      if(input.direction!=='ltr' || input.features?.length || input.variations?.length || input.letterSpacingMilliPoints!==undefined || input.wordSpacingMilliPoints!==undefined) throw new TextBodyLayoutRefusal('text.refused','symbol glyph placement does not implement shaping features, variations, spacing or non-LTR direction')
      symbolEvidence = qualifySymbolBullet(loaded.bytes, nativeRun.text)
      if(symbolEvidence.unitsPerEm!==loaded.metrics.unitsPerEm || symbolEvidence.ascender!==loaded.metrics.ascender || symbolEvidence.descender!==loaded.metrics.descender || symbolEvidence.lineGap!==loaded.metrics.lineGap) throw new TextBodyLayoutRefusal('text.refused','symbol font units or line metrics disagree with source metrics')
    }
    let shapedLive: unknown
    if(symbolEvidence) {
      const advance=scaleFontUnits(symbolEvidence.advanceWidth,symbolEvidence.unitsPerEm,input.fontSizeMilliPoints)
      shapedLive={startUtf16:0,endUtf16:1,face:resolution.face,glyphs:[{glyphId:symbolEvidence.glyphId,clusterIndex:0,advanceXMilliPoints:advance,advanceYMilliPoints:0,offsetXMilliPoints:0,offsetYMilliPoints:0}],clusters:[{startUtf16:0,endUtf16:1,glyphStart:0,glyphEnd:1,advanceInlineMilliPoints:advance}],metrics:scaleLineMetrics(loaded.metrics,input.fontSizeMilliPoints),advanceInlineMilliPoints:advance,advanceBlockMilliPoints:0}
    } else shapedLive = await state.providers.shaper.shape(Object.freeze({ run: input, startUtf16: 0, endUtf16: input.text.length, font: providerResource }))
    if (!providerIdentityStable(state.providers)) throw new TextBodyLayoutRefusal('text.refused', 'the injected provider changed its snapshotted identity during shaping')
    if (fontDigest(providerResource.bytes) !== resolution.face.contentDigest) throw new TextBodyLayoutRefusal('text.refused', 'the injected shaper mutated its isolated font bytes during shaping')
    const normalizedShaped = normalizeShaperOutput(shapedLive, state, `${path}.shaping`, input.text.length)
    if ('status' in normalizedShaped) throw new TextBodyLayoutRefusal('text.refused', 'the injected shaper refused the complete native paragraph')
    shaped = normalizedShaped
    if (symbolEvidence && (shaped.glyphs.length !== 1 || shaped.glyphs[0]!.glyphId !== symbolEvidence.glyphId || shaped.clusters.length !== 1)) throw new TextBodyLayoutRefusal('text.refused', 'symbol transport glyph disagrees with source cmap evidence')
    validateSegment(shaped, input, resolution.face, path)
    const expectedMetrics = scaleLineMetrics(loaded.metrics, input.fontSizeMilliPoints)
    if (!sameMetrics(shaped.metrics, expectedMetrics)) throw new TextBodyLayoutRefusal('text.refused', 'shaper line metrics are not exactly derived from the digest-bound font resource')
  } catch (error) {
    if (error instanceof TextBodyLayoutRefusal) throw error
    if (error instanceof RenderCompileError) {
      if (['render.glyphBudget', 'render.clusterBudget', 'render.providerDecisionBudget', 'render.providerFaceBudget'].includes(error.code)) throw error
      throw new TextBodyLayoutRefusal('text.refused', `invalid provider output: ${error.message.slice(0, 2_048)}`)
    }
    throw new TextBodyLayoutRefusal('text.refused', `injected native text provider failed: ${boundedProviderError(error)}`)
  }
  const resolutionDecisions = resolution.decisions
  takeGlyphs(state, shaped.glyphs.length, path)
  takeClusters(state, shaped.clusters.length, path)
  let cursorX = 0
  let cursorY = 0
  const glyphs: RenderGlyph[] = shaped.glyphs.map((glyph) => {
    const rendered: RenderGlyph = {
      glyphId: glyph.glyphId,
      clusterIndex: glyph.clusterIndex,
      advanceXMilliPoints: glyph.advanceXMilliPoints,
      advanceYMilliPoints: glyph.advanceYMilliPoints,
      offsetXMilliPoints: glyph.offsetXMilliPoints,
      offsetYMilliPoints: glyph.offsetYMilliPoints,
      xEmu: cursorX + milliPointsToEmu(glyph.offsetXMilliPoints, path),
      yEmu: cursorY + milliPointsToEmu(glyph.offsetYMilliPoints, path),
      advanceXEmu: milliPointsToEmu(glyph.advanceXMilliPoints, path),
      advanceYEmu: milliPointsToEmu(glyph.advanceYMilliPoints, path),
      offsetXEmu: milliPointsToEmu(glyph.offsetXMilliPoints, path),
      offsetYEmu: milliPointsToEmu(glyph.offsetYMilliPoints, path),
    }
    cursorX += rendered.advanceXEmu
    cursorY += rendered.advanceYEmu
    if (!Number.isSafeInteger(cursorX) || !Number.isSafeInteger(cursorY)) {
      throw new RenderCompileError('render.coordinateBudget', path, 'glyph placement exceeds integer precision')
    }
    for (const coordinate of [rendered.xEmu, rendered.yEmu, cursorX, cursorY]) checkCoordinate(coordinate, path, state.budget)
    return rendered
  })
  const lineHeightEmu = milliPointsToEmu(shaped.metrics.lineHeightMilliPoints, path)
  const ascentEmu = milliPointsToEmu(shaped.metrics.ascentMilliPoints, path)
  const descentEmu = milliPointsToEmu(shaped.metrics.descentMilliPoints, path)
  const lineGapEmu = milliPointsToEmu(shaped.metrics.lineGapMilliPoints, path)
  const advanceInlineEmu = milliPointsToEmu(shaped.advanceInlineMilliPoints, path)
  checkCoordinate(advanceInlineEmu, path, state.budget)
  checkCoordinate(lineHeightEmu, path, state.budget, true)
  checkCoordinate(ascentEmu, path, state.budget)
  checkCoordinate(descentEmu, path, state.budget)
  checkCoordinate(lineGapEmu, path, state.budget)
  const clusters: RenderCluster[] = shaped.clusters.map((cluster, index) => {
    const advanceInlineEmu = milliPointsToEmu(cluster.advanceInlineMilliPoints, `${path}.clusters[${index}]`)
    checkCoordinate(advanceInlineEmu, `${path}.clusters[${index}]`, state.budget)
    const rendered: RenderCluster = {
      startUtf16: cluster.startUtf16,
      endUtf16: cluster.endUtf16,
      glyphStart: cluster.glyphStart,
      glyphEnd: cluster.glyphEnd,
      advanceInlineMilliPoints: cluster.advanceInlineMilliPoints,
      advanceInlineEmu,
    }
    if (cluster.unsafeToBreak !== undefined) rendered.unsafeToBreak = cluster.unsafeToBreak
    if (cluster.whitespace !== undefined) rendered.whitespace = cluster.whitespace
    return rendered
  })
  return {
    direction: input.direction,
    lineHeightEmu,
    ascentEmu,
    ascentMilliPoints: shaped.metrics.ascentMilliPoints,
    descentMilliPoints: shaped.metrics.descentMilliPoints,
    lineGapMilliPoints: shaped.metrics.lineGapMilliPoints,
    run: {
      kind: 'textRun',
      sourceElementId: context.elementId,
      paragraphIndex: context.paragraphIndex,
      runIndex: context.runIndex,
      startUtf16: 0,
      endUtf16: nativeRun.text.length,
      text: nativeRun.text,
      ...(symbolEvidence ? {symbolEncoding:symbolEvidence} : {}),
      direction: input.direction,
      fontSizeMilliPoints: input.fontSizeMilliPoints,
      faceId: resolution.face.faceId,
      contentDigest: resolution.face.contentDigest,
      color,
      bold: nativeRun.bold ?? false,
      italic: nativeRun.italic ?? false,
      x: 0,
      baselineY: ascentEmu,
      advanceInlineEmu,
      lineHeightEmu,
      glyphs,
      clusters,
      decisions: resolutionDecisions,
      attemptedFaceIds: resolution.attemptedFaceIds,
      status: 'shaped',
    },
  }
}

function alignOffset(align: RenderParagraphNode['align'], width: number, advance: number): number {
  if (align === 'center') return Math.round((width - advance) / 2)
  if (align === 'right') return width - advance
  return 0
}

function takeTextLine(state: CompileState, path: string): void {
  state.budget.textLines++
  if (state.budget.textLines > PPTX_RENDER_LIMITS.maxTextLines) throw new RenderCompileError('render.textLineBudget', path, `RenderTree exceeds ${PPTX_RENDER_LIMITS.maxTextLines} text lines`)
  takeNode(state, path)
}

function takeTextFragment(state: CompileState, path: string): void {
  state.budget.textFragments++
  if (state.budget.textFragments > PPTX_RENDER_LIMITS.maxTextFragments) throw new RenderCompileError('render.textFragmentBudget', path, `RenderTree exceeds ${PPTX_RENDER_LIMITS.maxTextFragments} text fragments`)
  takeNode(state, path)
}

function wrapAtomCount(item: ShapedRunResult): number {
  return Math.max(1, item.run.clusters.length)
}

function advanceWrapPoint(shaped: readonly ShapedRunResult[], point: WrapPoint): boolean {
  if (point.runIndex >= shaped.length) return false
  if (point.clusterIndex + 1 < wrapAtomCount(shaped[point.runIndex]!)) {
    point.clusterIndex++
    return true
  }
  point.runIndex++
  point.clusterIndex = 0
  return point.runIndex < shaped.length
}

function wrapAtomAdvance(shaped: readonly ShapedRunResult[], runIndex: number, clusterIndex: number): number {
  const run = shaped[runIndex]!.run
  return run.clusters[clusterIndex]?.advanceInlineEmu ?? run.advanceInlineEmu
}

function wrapAtomUnsafe(shaped: readonly ShapedRunResult[], runIndex: number, clusterIndex: number): boolean {
  return shaped[runIndex]!.run.clusters[clusterIndex]?.unsafeToBreak === true
}

function wrapAtomStartsWith(shaped: readonly ShapedRunResult[], runIndex: number, clusterIndex: number, codeUnit: number): boolean {
  const run = shaped[runIndex]!.run
  const cluster = run.clusters[clusterIndex]
  const startUtf16 = cluster?.startUtf16 ?? 0
  const endUtf16 = cluster?.endUtf16 ?? run.text.length
  return startUtf16 < endUtf16 && run.text.charCodeAt(startUtf16) === codeUnit
}

function wrapAtomEndsWith(shaped: readonly ShapedRunResult[], runIndex: number, clusterIndex: number, codeUnit: number): boolean {
  const run = shaped[runIndex]!.run
  const cluster = run.clusters[clusterIndex]
  const startUtf16 = cluster?.startUtf16 ?? 0
  const endUtf16 = cluster?.endUtf16 ?? run.text.length
  return startUtf16 < endUtf16 && run.text.charCodeAt(endUtf16 - 1) === codeUnit
}

function wrapAtomIsSingleU0020(shaped: readonly ShapedRunResult[], runIndex: number, clusterIndex: number): boolean {
  const run = shaped[runIndex]!.run
  const cluster = run.clusters[clusterIndex]
  const startUtf16 = cluster?.startUtf16 ?? 0
  const endUtf16 = cluster?.endUtf16 ?? run.text.length
  return endUtf16 - startUtf16 === 1 && run.text.charCodeAt(startUtf16) === 0x20
}

function wrapBoundaryDecision(
  shaped: readonly ShapedRunResult[],
  leftRunIndex: number,
  leftClusterIndex: number,
  rightRunIndex: number,
  rightClusterIndex: number,
) {
  const leftRun = shaped[leftRunIndex]!.run
  const leftCluster = leftRun.clusters[leftClusterIndex]
  const rightRun = shaped[rightRunIndex]!.run
  const rightCluster = rightRun.clusters[rightClusterIndex]
  return classifyNativeOfficeLineBreakRanges(
    leftRun.text,
    leftCluster?.startUtf16 ?? 0,
    leftCluster?.endUtf16 ?? leftRun.text.length,
    rightRun.text,
    rightCluster?.startUtf16 ?? 0,
    rightCluster?.endUtf16 ?? rightRun.text.length,
  )
}

function singleConsumableSeparator(
  shaped: readonly ShapedRunResult[],
  previousRunIndex: number,
  previousClusterIndex: number,
  separatorRunIndex: number,
  separatorClusterIndex: number,
  followingRunIndex: number,
  followingClusterIndex: number,
): boolean {
  if (!wrapAtomIsSingleU0020(shaped, separatorRunIndex, separatorClusterIndex)) return false
  const previous = shaped[previousRunIndex]!
  const separator = shaped[separatorRunIndex]!
  const following = shaped[followingRunIndex]!
  if (
    separator.ascentMilliPoints !== previous.ascentMilliPoints ||
    separator.descentMilliPoints !== previous.descentMilliPoints ||
    separator.lineGapMilliPoints !== previous.lineGapMilliPoints ||
    separator.ascentMilliPoints !== following.ascentMilliPoints ||
    separator.descentMilliPoints !== following.descentMilliPoints ||
    separator.lineGapMilliPoints !== following.lineGapMilliPoints
  ) return false
  if (
    wrapAtomEndsWith(shaped, previousRunIndex, previousClusterIndex, 0x20) ||
    wrapAtomEndsWith(shaped, previousRunIndex, previousClusterIndex, 0x3000) ||
    wrapAtomStartsWith(shaped, followingRunIndex, followingClusterIndex, 0x20) ||
    wrapAtomStartsWith(shaped, followingRunIndex, followingClusterIndex, 0x3000) ||
    wrapAtomUnsafe(shaped, previousRunIndex, previousClusterIndex) ||
    wrapAtomUnsafe(shaped, separatorRunIndex, separatorClusterIndex) ||
    wrapAtomUnsafe(shaped, followingRunIndex, followingClusterIndex)
  ) return false
  return (
    wrapBoundaryDecision(shaped, previousRunIndex, previousClusterIndex, separatorRunIndex, separatorClusterIndex) !== 'unsupported' &&
    wrapBoundaryDecision(shaped, separatorRunIndex, separatorClusterIndex, followingRunIndex, followingClusterIndex) === 'allowed'
  )
}

function* squareWrappedLineRanges(shaped: readonly ShapedRunResult[], width: number, continuationWidth = width): Generator<WrapLineRange> {
  if (shaped.length === 0) {
    yield { startRunIndex: 0, startClusterIndex: 0, endRunIndex: 0, endClusterIndex: 0 }
    return
  }
  let lineStartRunIndex = 0
  let lineStartClusterIndex = 0
  const cursor: WrapPoint = { runIndex: 0, clusterIndex: 0 }
  const next: WrapPoint = { runIndex: 0, clusterIndex: 0 }
  const after: WrapPoint = { runIndex: 0, clusterIndex: 0 }
  let advance = 0
  let hasBreakOpportunity = false
  let breakEndRunIndex = 0
  let breakEndClusterIndex = 0
  let breakNextRunIndex = 0
  let breakNextClusterIndex = 0
  let breakConsumedRunIndex: number | undefined
  let breakConsumedClusterIndex: number | undefined

  while (cursor.runIndex < shaped.length) {
    const nextAdvance = advance + wrapAtomAdvance(shaped, cursor.runIndex, cursor.clusterIndex)
    if (!Number.isSafeInteger(nextAdvance)) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'cluster advances exceed integer precision during text wrapping')
    if (nextAdvance > width) {
      if (hasBreakOpportunity && (breakEndRunIndex !== lineStartRunIndex || breakEndClusterIndex !== lineStartClusterIndex)) {
        yield {
          startRunIndex: lineStartRunIndex,
          startClusterIndex: lineStartClusterIndex,
          endRunIndex: breakEndRunIndex,
          endClusterIndex: breakEndClusterIndex,
          ...(breakConsumedRunIndex === undefined ? {} : { consumedSeparatorRunIndex: breakConsumedRunIndex, consumedSeparatorClusterIndex: breakConsumedClusterIndex }),
        }
        lineStartRunIndex = breakNextRunIndex
        width = continuationWidth
        lineStartClusterIndex = breakNextClusterIndex
        cursor.runIndex = lineStartRunIndex
        cursor.clusterIndex = lineStartClusterIndex
        advance = 0
        hasBreakOpportunity = false
        breakConsumedRunIndex = undefined
        breakConsumedClusterIndex = undefined
        continue
      }
      if (cursor.runIndex === lineStartRunIndex && cursor.clusterIndex === lineStartClusterIndex) {
        throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'one unbreakable shaped cluster exceeds the native text-body width')
      }
      throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'no modeled Unicode shaped-cluster boundary fits the native text-body width')
    }
    advance = nextAdvance

    next.runIndex = cursor.runIndex
    next.clusterIndex = cursor.clusterIndex
    if (advanceWrapPoint(shaped, next)) {
      after.runIndex = next.runIndex
      after.clusterIndex = next.clusterIndex
      const hasAfter = advanceWrapPoint(shaped, after)
      if (
        hasAfter &&
        singleConsumableSeparator(
          shaped,
          cursor.runIndex,
          cursor.clusterIndex,
          next.runIndex,
          next.clusterIndex,
          after.runIndex,
          after.clusterIndex,
        )
      ) {
        // The separator participates in the unwrapped line advance, but if this
        // opportunity is used it is consumed and excluded from both paint lines.
        breakEndRunIndex = next.runIndex
        breakEndClusterIndex = next.clusterIndex
        breakNextRunIndex = after.runIndex
        breakNextClusterIndex = after.clusterIndex
        breakConsumedRunIndex = next.runIndex
        breakConsumedClusterIndex = next.clusterIndex
        hasBreakOpportunity = true
      } else if (
        !wrapAtomEndsWith(shaped, cursor.runIndex, cursor.clusterIndex, 0x20) &&
        wrapBoundaryDecision(shaped, cursor.runIndex, cursor.clusterIndex, next.runIndex, next.clusterIndex) === 'allowed' &&
        !wrapAtomUnsafe(shaped, cursor.runIndex, cursor.clusterIndex) &&
        !wrapAtomUnsafe(shaped, next.runIndex, next.clusterIndex)
      ) {
        breakEndRunIndex = next.runIndex
        breakEndClusterIndex = next.clusterIndex
        breakNextRunIndex = next.runIndex
        breakNextClusterIndex = next.clusterIndex
        breakConsumedRunIndex = undefined
        breakConsumedClusterIndex = undefined
        hasBreakOpportunity = true
      }
      cursor.runIndex = next.runIndex
      cursor.clusterIndex = next.clusterIndex
    } else {
      cursor.runIndex = shaped.length
      cursor.clusterIndex = 0
    }
  }

  yield {
    startRunIndex: lineStartRunIndex,
    startClusterIndex: lineStartClusterIndex,
    endRunIndex: shaped.length,
    endClusterIndex: 0,
  }
}

function fragmentRun(source: RenderTextRunNode, clusterStart: number | undefined, clusterEnd: number | undefined, progress: FragmentProgress, state: CompileState, path: string): RenderTextRunNode {
  if (clusterStart === undefined || clusterEnd === undefined || source.clusters.length === 0) return { ...source }
  if (clusterStart < 0 || clusterEnd <= clusterStart || clusterEnd > source.clusters.length) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'fragment cluster range is invalid')
  if (progress.clusterIndex > clusterStart) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'fragment cluster ranges are not monotonic')
  if (clusterStart === 0 && clusterEnd === source.clusters.length && progress.clusterIndex === 0 && progress.glyphIndex === 0) {
    progress.clusterIndex = clusterEnd
    progress.glyphIndex = source.glyphs.length
    return { ...source }
  }
  const firstCluster = source.clusters[clusterStart]!
  const lastCluster = source.clusters[clusterEnd - 1]!
  const textStart = firstCluster.startUtf16
  const textEnd = lastCluster.endUtf16
  let advanceInlineEmu = 0
  for (let index = clusterStart; index < clusterEnd; index++) advanceInlineEmu += source.clusters[index]!.advanceInlineEmu
  const glyphStart = firstCluster.glyphStart
  const glyphEnd = lastCluster.glyphEnd
  if (progress.glyphIndex > glyphStart) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'fragment glyph ranges are not monotonic')
  while (progress.glyphIndex < glyphStart) {
    progress.glyphPenXEmu += source.glyphs[progress.glyphIndex]!.advanceXEmu
    progress.glyphPenYEmu += source.glyphs[progress.glyphIndex]!.advanceYEmu
    progress.glyphIndex++
  }
  const glyphPenXBefore = progress.glyphPenXEmu
  const glyphPenYBefore = progress.glyphPenYEmu
  if (![glyphPenXBefore, glyphPenYBefore, advanceInlineEmu].every(Number.isSafeInteger)) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'fragment advances exceed integer precision')
  const glyphs: RenderGlyph[] = []
  for (let sourceIndex = glyphStart; sourceIndex < glyphEnd; sourceIndex++) {
    const item = source.glyphs[sourceIndex]!
    const xEmu = item.xEmu - glyphPenXBefore
    const yEmu = item.yEmu - glyphPenYBefore
    checkCoordinate(xEmu, `${path}.glyphs[${sourceIndex - glyphStart}].xEmu`, state.budget)
    checkCoordinate(yEmu, `${path}.glyphs[${sourceIndex - glyphStart}].yEmu`, state.budget)
    glyphs.push({
      glyphId: item.glyphId,
      clusterIndex: item.clusterIndex - clusterStart,
      advanceXMilliPoints: item.advanceXMilliPoints,
      advanceYMilliPoints: item.advanceYMilliPoints,
      offsetXMilliPoints: item.offsetXMilliPoints,
      offsetYMilliPoints: item.offsetYMilliPoints,
      xEmu,
      yEmu,
      advanceXEmu: item.advanceXEmu,
      advanceYEmu: item.advanceYEmu,
      offsetXEmu: item.offsetXEmu,
      offsetYEmu: item.offsetYEmu,
    })
    progress.glyphPenXEmu += item.advanceXEmu
    progress.glyphPenYEmu += item.advanceYEmu
    progress.glyphIndex++
  }
  progress.clusterIndex = clusterEnd
  if (![progress.glyphPenXEmu, progress.glyphPenYEmu].every(Number.isSafeInteger)) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'fragment glyph advances exceed integer precision')
  const clusters: RenderCluster[] = []
  for (let sourceIndex = clusterStart; sourceIndex < clusterEnd; sourceIndex++) {
    const item = source.clusters[sourceIndex]!
    const fragment: RenderCluster = {
      startUtf16: item.startUtf16 - textStart,
      endUtf16: item.endUtf16 - textStart,
      glyphStart: item.glyphStart - glyphStart,
      glyphEnd: item.glyphEnd - glyphStart,
      advanceInlineMilliPoints: item.advanceInlineMilliPoints,
      advanceInlineEmu: item.advanceInlineEmu,
    }
    if (item.unsafeToBreak !== undefined) fragment.unsafeToBreak = item.unsafeToBreak
    if (item.whitespace !== undefined) fragment.whitespace = item.whitespace
    clusters.push(fragment)
  }
  return {
    ...source,
    startUtf16: source.startUtf16 + textStart,
    endUtf16: source.startUtf16 + textEnd,
    text: source.text.slice(textStart, textEnd),
    advanceInlineEmu,
    glyphs,
    clusters,
  }
}

function lineRunFragments(
  shaped: readonly ShapedRunResult[],
  range: WrapLineRange,
  progress: readonly FragmentProgress[],
  state: CompileState,
  path: string,
): readonly ShapedRunResult[] {
  const result: ShapedRunResult[] = []
  for (let runIndex = range.startRunIndex; runIndex < shaped.length && runIndex <= range.endRunIndex; runIndex++) {
    const first = shaped[runIndex]!
    const atomCount = wrapAtomCount(first)
    const atomStart = runIndex === range.startRunIndex ? range.startClusterIndex : 0
    const atomEnd = runIndex === range.endRunIndex ? range.endClusterIndex : atomCount
    if (atomEnd <= atomStart) continue
    const clusterStart = first.run.clusters.length === 0 ? undefined : atomStart
    const clusterEnd = clusterStart === undefined ? undefined : atomEnd
    result.push({
      direction: first.direction,
      lineHeightEmu: first.lineHeightEmu,
      ascentEmu: first.ascentEmu,
      ascentMilliPoints: first.ascentMilliPoints,
      descentMilliPoints: first.descentMilliPoints,
      lineGapMilliPoints: first.lineGapMilliPoints,
      run: fragmentRun(first.run, clusterStart, clusterEnd, progress[runIndex]!, state, path),
    })
  }
  return result
}

function nativeRunLacksExplicitFont(run: NativeTextRun): boolean {
  return !run.fontFamily || run.fontFamily.startsWith('+') || run.fontSizeHundredthPt === undefined
}

async function compileParagraphs(paragraphs: readonly NativeParagraph[], context: TextContainerContext, state: CompileState): Promise<readonly RenderParagraphNode[]> {
  if (context.layout && !state.lineLayoutPolicy && context.layout.verticalAnchor !== 'top') {
    throw new TextBodyLayoutRefusal('text.verticalAnchorUnavailable', 'native center/bottom text anchoring requires an Office-qualified line-box rule')
  }
  const result: RenderParagraphNode[] = []
  let y = 0
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
    const paragraph = paragraphs[paragraphIndex]!
    if (context.layout && !state.lineLayoutPolicy && (paragraph.align === undefined || paragraph.level !== 0 || paragraph.bullet !== false || (paragraph.marginLeftEmu ?? 0) !== 0 || (paragraph.indentEmu ?? 0) !== 0)) {
      throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable', 'native layout requires explicit alignment and refuses bullets, nonzero list levels, margins or indents until their line geometry is qualified')
    }
    const measuredParagraph = Boolean(context.layout && state.lineLayoutPolicy)
    const margin = measuredParagraph ? paragraph.marginLeftEmu ?? 0 : 0
    const indent = measuredParagraph ? paragraph.indentEmu ?? 0 : 0
    if (measuredParagraph && (paragraph.align===undefined || paragraph.bullet===undefined || paragraph.level===undefined || (paragraph.level!==0&&(paragraph.marginLeftEmu===undefined||paragraph.indentEmu===undefined)))) throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable','measured paragraphs need explicit alignment/list semantics and explicit offsets at nonzero levels')
    const firstTextOffset=margin+(paragraph.bullet?0:indent)
    const firstWidth=context.bounds.cx-firstTextOffset,continuationWidth=context.bounds.cx-margin
    if(measuredParagraph&&(!Number.isSafeInteger(firstWidth)||!Number.isSafeInteger(continuationWidth)||firstWidth<=0||continuationWidth<=0))throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable','paragraph margins leave no positive text line width')
    if (context.layout && state.nativeTextInheritanceUnresolved && paragraph.runs.some(nativeRunLacksExplicitFont)) {
      throw new TextBodyLayoutRefusal('text.inheritanceUnavailable', 'native layout refuses runs that still need unresolved presentation/layout/master/theme fonts')
    }
    if (context.layout && paragraph.runs.some((run) => run.fontFamily?.startsWith('+'))) {
      throw new TextBodyLayoutRefusal('text.inheritanceUnavailable', 'theme font token was not resolved to an exact typeface')
    }
    const shaped: ShapedRunResult[] = []
    for (let runIndex = 0; runIndex < paragraph.runs.length; runIndex++) {
      shaped.push(await shapeRun(paragraph.runs[runIndex]!, {
        slideId: state.slide.id, elementId: context.elementId, elementKind: context.elementKind,
        paragraphIndex, runIndex, text: paragraph.runs[runIndex]!.text,
      }, state))
    }
    const direction = shaped[0]?.direction ?? state.textDefaults.direction
    if(measuredParagraph&&(margin!==0||indent!==0||paragraph.bullet)&&(direction!=='ltr'||(paragraph.bullet&&paragraph.align!=='left')))throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable','measured marker/indent placement requires horizontal LTR and left-aligned bullets')
    let shapedMarker:ShapedRunResult|undefined
    if(measuredParagraph&&paragraph.bullet){
      if(!paragraph.bulletCharacter||!paragraph.runs[0])throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable','native bullet needs one explicit authored character and a source font run')
      const markerRun={...paragraph.runs[0],text:paragraph.bulletCharacter,...(paragraph.bulletFontFamily ? {fontFamily:paragraph.bulletFontFamily} : {})}
      shapedMarker=await shapeRun(markerRun,{slideId:state.slide.id,elementId:context.elementId,elementKind:context.elementKind,paragraphIndex,runIndex:0,text:paragraph.bulletCharacter},state,paragraph.bulletFontFamily,paragraph.bulletFontEncoding)
      if (shapedMarker.run.symbolEncoding) state.diagnostics.push({severity:'info',code:'text.symbolBulletEncoding',message:'Authored bullet byte selects one exact font glyph through qualified Windows symbol cmap and hmtx metrics; source character and cluster remain unchanged. No GSUB/GPOS shaping is applied to this isolated symbol.',slideId:state.slide.id,elementId:context.elementId})
      if(shapedMarker.direction!=='ltr'||shapedMarker.run.status!=='shaped'||indent+shapedMarker.run.advanceInlineEmu>0)throw new TextBodyLayoutRefusal('text.paragraphSemanticsUnavailable','authored hanging indent does not fit the exact shaped marker before the text origin')
    }
    const directions = new Set(shaped.map((item) => item.direction))
    if (directions.size > 1) {
      if (context.layout) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'native text-body layout refuses mixed-direction runs until paragraph bidi metadata is modeled')
      state.diagnostics.push({
        severity: 'warning', code: 'text.bidiUnavailable',
        message: 'paragraph runs request mixed directions, but native PPTX v1 has no paragraph bidi metadata; runs remain in contract order without claiming faithful bidi reordering',
        slideId: state.slide.id, elementId: context.elementId,
      })
    }
    if (context.layout && (direction === 'ttb' || direction === 'btt')) {
      throw new TextBodyLayoutRefusal('text.verticalUnsupported', 'native vertical text layout is unavailable until RenderTree models vertical line progression')
    }
    if (shaped.some((item) => item.run.decisions.some((decision) => decision.code === 'unsupported-direction'))) {
      throw new TextBodyLayoutRefusal('text.verticalUnsupported', 'vertical text is refused until native vertical line progression is modeled')
    }
    if (shaped.some((item) => item.run.status === 'refused')) {
      throw new TextBodyLayoutRefusal('text.refused', 'the injected font resolver or shaper refused the complete paragraph')
    }
    if (context.layout?.wrap === 'square' && direction !== 'ltr') {
      throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'native square wrapping currently requires an exact horizontal LTR paragraph direction')
    }
    const squareWrap = context.layout?.wrap === 'square'
    const lineRanges: Iterable<WrapLineRange> = squareWrap
      ? squareWrappedLineRanges(shaped, firstWidth, continuationWidth)
      : [{ startRunIndex: 0, startClusterIndex: 0, endRunIndex: shaped.length, endClusterIndex: 0 }]
    const fragmentProgress: FragmentProgress[] | undefined = squareWrap
      ? shaped.map(() => ({ clusterIndex: 0, glyphIndex: 0, glyphPenXEmu: 0, glyphPenYEmu: 0 }))
      : undefined
    let lineIndex = 0
    for (const lineRange of lineRanges) {
      const path = `$.elements.${context.elementId}.paragraphs[${paragraphIndex}].lines[${lineIndex}]`
      takeTextLine(state, path)
      const lineShaped = squareWrap ? lineRunFragments(shaped, lineRange, fragmentProgress!, state, path) : shaped
      const advance = lineShaped.reduce((sum, item) => sum + item.run.advanceInlineEmu, 0)
      if (!Number.isSafeInteger(advance)) throw new RenderCompileError('render.textMetric', path, 'line advance exceeds integer precision')
      const contentMetrics = lineShaped.length > 0 ? lineShaped : shaped
      const metricRuns = lineIndex===0&&shapedMarker ? [...contentMetrics,shapedMarker] : contentMetrics
      if (context.layout && metricRuns.length === 0) {
        throw new TextBodyLayoutRefusal('text.metricsUnavailable', 'an empty native paragraph has no digest-bound font metrics for its line box')
      }
      const firstMetric = metricRuns[0]
      if (context.layout && !state.lineLayoutPolicy && firstMetric && metricRuns.some((item) =>
        item.ascentMilliPoints !== firstMetric.ascentMilliPoints ||
        item.descentMilliPoints !== firstMetric.descentMilliPoints ||
        item.lineGapMilliPoints !== firstMetric.lineGapMilliPoints
      )) {
        throw new TextBodyLayoutRefusal('text.metricsUnavailable', 'mixed shaped line metrics require an Office-qualified leading and baseline aggregation rule')
      }
      let ascentMilliPoints = firstMetric?.ascentMilliPoints ?? 0
      let descentMilliPoints = firstMetric?.descentMilliPoints ?? 0
      let lineGapMilliPoints = firstMetric?.lineGapMilliPoints ?? 0
      if (!context.layout || state.lineLayoutPolicy) {
        for (let metricIndex = 1; metricIndex < metricRuns.length; metricIndex++) {
          const item = metricRuns[metricIndex]!
          if (item.ascentMilliPoints > ascentMilliPoints) ascentMilliPoints = item.ascentMilliPoints
          if (item.descentMilliPoints < descentMilliPoints) descentMilliPoints = item.descentMilliPoints
          if (item.lineGapMilliPoints > lineGapMilliPoints) lineGapMilliPoints = item.lineGapMilliPoints
        }
      }
      let naturalLineHeightMilliPoints = ascentMilliPoints - descentMilliPoints
      if (!Number.isSafeInteger(naturalLineHeightMilliPoints)) throw new RenderCompileError('render.textMetric', path, 'line ascent/descent exceed integer precision')
      naturalLineHeightMilliPoints += lineGapMilliPoints
      if (!Number.isSafeInteger(naturalLineHeightMilliPoints)) throw new RenderCompileError('render.textMetric', path, 'line metrics exceed integer precision')
      let ascent = milliPointsToEmu(ascentMilliPoints, path)
      let lineHeight = milliPointsToEmu(naturalLineHeightMilliPoints, path)
      if (lineHeight < 1) {
        lineHeight = milliPointsToEmu(state.textDefaults.fontSizeHundredthPt * 10, '$.text.defaults')
        ascent = lineHeight
      }
      checkCoordinate(ascent, path, state.budget)
      checkCoordinate(milliPointsToEmu(-descentMilliPoints, path), path, state.budget)
      checkCoordinate(milliPointsToEmu(lineGapMilliPoints, path), path, state.budget)
      checkCoordinate(lineHeight, path, state.budget, true)
      const align = paragraph.align ?? 'left'
      const textOffset=lineIndex===0?firstTextOffset:margin
      const lineStart = textOffset+alignOffset(align, context.bounds.cx-textOffset, advance)
      let cursor = direction === 'rtl' ? lineStart + advance : lineStart
      const runs = lineShaped.map((item, fragmentIndex) => {
        takeTextFragment(state, `${path}.runs[${fragmentIndex}]`)
        if (direction === 'rtl') cursor -= item.run.advanceInlineEmu
        const runX = cursor
        if (direction !== 'rtl') cursor += item.run.advanceInlineEmu
        checkCoordinate(runX, path, state.budget)
        checkCoordinate(y + ascent, path, state.budget)
        return { ...item.run, x: runX, baselineY: y + ascent }
      })
      let consumedSoftSeparators: RenderParagraphNode['consumedSoftSeparators']
      if (lineRange.consumedSeparatorRunIndex !== undefined && lineRange.consumedSeparatorClusterIndex !== undefined) {
        const source = shaped[lineRange.consumedSeparatorRunIndex]!.run
        const cluster = source.clusters[lineRange.consumedSeparatorClusterIndex]!
        if (cluster.endUtf16 - cluster.startUtf16 !== 1 || source.text.charCodeAt(cluster.startUtf16) !== 0x20) throw new TextBodyLayoutRefusal('text.wrapUnavailable', 'consumed separator source mapping is not one exact U+0020')
        consumedSoftSeparators = [{
          sourceElementId: source.sourceElementId,
          paragraphIndex: source.paragraphIndex,
          runIndex: source.runIndex,
          startUtf16: cluster.startUtf16,
          endUtf16: cluster.endUtf16,
        }]
      }
      let marker:RenderTextRunNode|undefined
      if(lineIndex===0&&shapedMarker){takeTextFragment(state,`${path}.marker`);checkCoordinate(margin+indent,path,state.budget);marker={...shapedMarker.run,sourceRole:'paragraphBullet',x:margin+indent,baselineY:y+ascent}}
      result.push({
        kind: 'paragraph', sourceElementId: context.elementId, paragraphIndex, lineIndex,
        align, direction, level: paragraph.level ?? 0, bullet: paragraph.bullet ?? false,
        x: lineStart, y, widthEmu: advance, heightEmu: lineHeight, runs,
        ...(marker?{marker}:{}),
        ...(consumedSoftSeparators === undefined ? {} : { consumedSoftSeparators }),
      })
      if (context.layout === undefined && (advance > context.bounds.cx || y + lineHeight > context.bounds.cy)) {
        state.diagnostics.push({
          severity: 'warning', code: 'text.overflow',
          message: 'exact shaped text exceeds legacy bounds without native body metadata; this compatibility preview remains clipped and does not claim PowerPoint-faithful reflow',
          slideId: state.slide.id, elementId: context.elementId,
        })
      }
      if (runs.some((run) => run.status === 'refused')) state.diagnostics.push({ severity: 'refusal', code: 'text.refused', message: 'the injected font resolver or shaper refused a rich-text run', slideId: state.slide.id, elementId: context.elementId })
      if (runs.some((run) => run.decisions.some((decision) => decision.code === 'unsupported-direction'))) state.diagnostics.push({ severity: 'refusal', code: 'text.verticalUnsupported', message: 'vertical text is represented by a refusal placeholder until native vertical layout is modeled', slideId: state.slide.id, elementId: context.elementId })
      y += lineHeight
      checkCoordinate(y, path, state.budget)
      lineIndex++
    }
  }
  const offsetX = context.bounds.x
  const remainder = context.bounds.cy - y
  if (!Number.isSafeInteger(remainder)) throw new RenderCompileError('render.textMetric', `$.elements.${context.elementId}.textBody`, 'text anchor remainder exceeds integer precision')
  // The named policy intentionally specifies floor for half-EMU centers and
  // signed offsets for overflowing blocks. No browser/Office heuristic enters.
  const anchorOffset = state.lineLayoutPolicy && context.layout?.verticalAnchor === 'center' ? Math.floor(remainder / 2)
    : state.lineLayoutPolicy && context.layout?.verticalAnchor === 'bottom' ? remainder : 0
  const offsetY = context.bounds.y + anchorOffset
  checkCoordinate(offsetX, `$.elements.${context.elementId}.textBody`, state.budget)
  checkCoordinate(offsetY, `$.elements.${context.elementId}.textBody`, state.budget)
  return result.map((paragraph) => {
    const paragraphX = paragraph.x + offsetX
    const paragraphY = paragraph.y + offsetY
    checkCoordinate(paragraphX, `$.elements.${context.elementId}.textBody`, state.budget)
    checkCoordinate(paragraphY, `$.elements.${context.elementId}.textBody`, state.budget)
    const runs = paragraph.runs.map((run) => {
      const x = run.x + offsetX
      const baselineY = run.baselineY + offsetY
      checkCoordinate(x, `$.elements.${context.elementId}.textBody`, state.budget)
      checkCoordinate(baselineY, `$.elements.${context.elementId}.textBody`, state.budget)
      return { ...run, x, baselineY }
    })
    let marker:RenderTextRunNode|undefined
    if(paragraph.marker){const x=paragraph.marker.x+offsetX,baselineY=paragraph.marker.baselineY+offsetY;checkCoordinate(x,`$.elements.${context.elementId}.marker`,state.budget);checkCoordinate(baselineY,`$.elements.${context.elementId}.marker`,state.budget);marker={...paragraph.marker,x,baselineY}}
    return { ...paragraph, x: paragraphX, y: paragraphY, runs, ...(marker?{marker}:{}) }
  })
}

async function compileTextBody(paragraphs: readonly NativeParagraph[], context: TextContainerContext, state: CompileState): Promise<RenderTextBodyNode> {
  const path = `$.elements.${context.elementId}.textBody`
  takeNode(state, path)
  try {
    const approximateSourceFrame = context.layout?.autoFit === 'shape-source-frame'
    if (approximateSourceFrame && !state.sourceFrameAutoFitPreview) throw new TextBodyLayoutRefusal('text.sourceFrameAutoFitRequiresOptIn', 'Source-frame autofit is approximate and requires explicit preview opt-in.')
    if (approximateSourceFrame) state.diagnostics.push({ severity: 'warning', code: 'text.sourceFrameAutoFitApproximate', message: 'Read-only approximate autofit preview uses the saved source frame without resizing; frame size, layout, and overflow or clipping may differ from PowerPoint.', slideId: state.slide.id, elementId: context.elementId })
    const vertical = context.layout?.writingMode === 'vertical-clockwise'
    if (vertical && paragraphs.some(paragraph=>paragraph.bullet!==false || paragraph.level!==0 || (paragraph.marginLeftEmu??0)!==0 || (paragraph.indentEmu??0)!==0 || paragraph.runs.some(run=>!run.text || !/^[\x20-\x7e]+$/.test(run.text)))) throw new TextBodyLayoutRefusal('text.verticalUnsupported','Clockwise vertical preview requires nonempty ASCII Latin text and no bullets or paragraph offsets.')
    const layoutContext = vertical ? {...context,bounds:{x:0,y:0,cx:context.bounds.cy,cy:context.bounds.cx}} : context
    const compiled = await compileParagraphs(paragraphs, layoutContext, state)
    if(vertical && compiled.some(paragraph=>paragraph.runs.some(run=>run.direction!=='ltr'))) throw new TextBodyLayoutRefusal('text.verticalUnsupported','Clockwise vertical preview requires qualified left-to-right Latin shaping.')
    const transform = vertical ? {aPpm:0,bPpm:1_000_000,cPpm:-1_000_000,dPpm:0,txEmu:context.bounds.x+context.bounds.cx,tyEmu:context.bounds.y} : undefined
    if(transform){checkCoordinate(transform.txEmu,path,state.budget);checkCoordinate(transform.tyEmu,path,state.budget)}
    const deterministic = Boolean(context.layout && state.lineLayoutPolicy)
    if (deterministic) state.diagnostics.push({severity:'warning',code:'text.deterministicLayout',message:'Measured native glyphs use InjOffice max-run-natural-v1 line boxes and anchor offsets; this policy is not an Office visual-equivalence claim.',slideId:state.slide.id,elementId:context.elementId})
    return {
      kind: 'textBody', sourceElementId: context.elementId, bounds: context.bounds,
      ...(transform?{transform}:{}),
      fidelity: approximateSourceFrame ? 'approximateSourceFrame' : deterministic ? 'deterministicNative' : context.layout ? 'native' : 'legacyUnavailable',
      ...(deterministic ? {lineLayoutPolicy: state.lineLayoutPolicy} : {}),
      wrap: context.layout?.wrap, verticalAnchor: context.layout?.verticalAnchor, autoFit: context.layout?.autoFit,
      horizontalOverflow: context.layout?.horizontalOverflow, verticalOverflow: context.layout?.verticalOverflow,
      status: 'laidOut', paragraphs: compiled,
    }
  } catch (error) {
    if (!(error instanceof TextBodyLayoutRefusal)) throw error
    state.diagnostics.push({ severity: 'refusal', code: error.code, message: error.message, slideId: state.slide.id, elementId: context.elementId })
    return {
      kind: 'textBody', sourceElementId: context.elementId, bounds: context.bounds,
      fidelity: context.layout ? 'nativeUnavailable' : 'legacyUnavailable',
      wrap: context.layout?.wrap, verticalAnchor: context.layout?.verticalAnchor, autoFit: context.layout?.autoFit,
      horizontalOverflow: context.layout?.horizontalOverflow, verticalOverflow: context.layout?.verticalOverflow,
      status: 'refused', paragraphs: [], refusalLabel: 'Exact text layout unavailable',
    }
  }
}

function scaleTracks(source: readonly number[], count: number, target: number): readonly number[] {
  if (count === 0) return []
  const weights = source.length === count ? source : Array.from({ length: count }, () => 1)
  const total = weights.reduce((sum, value) => sum + value, 0)
  if (!Number.isSafeInteger(total) || total <= 0) throw new RenderCompileError('render.tableGeometry', '$.table', 'table track total is invalid')
  const totalBig = BigInt(total)
  const targetBig = BigInt(target)
  const result = weights.map((weight) => Number(BigInt(weight) * targetBig / totalBig))
  let remainder = target - result.reduce((sum, value) => sum + value, 0)
  for (let index = 0; remainder > 0; index = (index + 1) % result.length) {
    result[index]!++
    remainder--
  }
  return result
}

async function compileTableCells(element: Extract<NativeElement, { kind: 'table' }>, state: CompileState): Promise<readonly RenderTableCellNode[]> {
  const rowCount = element.table.rows.length
  const columnCount = element.table.columnWidths.length
  const widths = scaleTracks(element.table.columnWidths, columnCount, element.transform.cx)
  const heights = scaleTracks(element.table.rowHeights, rowCount, element.transform.cy)
  const cells: RenderTableCellNode[] = []
  let y = 0
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    let x = 0
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      takeNode(state, `$.elements.${element.id}.table.rows[${rowIndex}][${columnIndex}]`)
      const cell = element.table.rows[rowIndex]![columnIndex]!
      const bounds = { x, y, cx: widths[columnIndex]!, cy: heights[rowIndex]! }
      const common = {
        kind: 'tableCell',
        sourceElementId: element.id,
        rowIndex,
        columnIndex,
        bounds,
        fill: cell.fill ? { color: cell.fill } : undefined,
        border: cell.border ? boundedStroke(cell.border, `$.elements.${element.id}.table.rows[${rowIndex}][${columnIndex}].border`, state.budget) : undefined,
      } as const
      if (cell.paragraphs && cell.textBody) {
        cells.push({
          ...common,
          textBody: await compileTextBody(cell.paragraphs, {
            elementId: element.id,
            elementKind: 'table',
            bounds: nativeTableCellTextBodyBounds(element.id, rowIndex, columnIndex, bounds.cx, bounds.cy, cell.textBody, state),
            layout: cell.textBody,
          }, state),
        })
      } else {
        const paragraph = (await compileParagraphs([{ runs: [{ text: cell.text }], align: cell.align }], {
          elementId: element.id,
          elementKind: 'table',
          bounds: { x: 0, y: 0, cx: bounds.cx, cy: bounds.cy },
        }, state))[0]!
        cells.push({ ...common, paragraph })
      }
      x += widths[columnIndex]!
    }
    y += heights[rowIndex]!
  }
  return cells
}

async function compileElement(element: NativeElement, zIndex: number, depth: number, state: CompileState, parentWorld: WorldAffine): Promise<RenderNode> {
  if (depth > state.budget.maxDepth) {
    throw new RenderCompileError('render.depthBudget', `$.elements.${element.id}`, `RenderTree nesting exceeds ${state.budget.maxDepth}`)
  }
  copyNativeDiagnostics(state, element.compatibility, element.id)
  if (element.compatibility.status === 'refused') {
    state.diagnostics.push({ severity: 'refusal', code: 'native.refused', message: 'native compatibility refused this element', slideId: state.slide.id, elementId: element.id })
    const refused = placeholder(element, zIndex, 'refused', `Unsupported ${element.kind}`, state)
    checkedWorldAffine(parentWorld, refused.transform, refused.bounds, `$.elements.${element.id}`, state.budget)
    return refused
  }
  if (element.compatibility.status === 'preserveOnly') {
    state.diagnostics.push({ severity: 'warning', code: 'render.preserveOnly', message: 'element is rendered for preview but remains preserve-only', slideId: state.slide.id, elementId: element.id })
  }
  takeNode(state, `$.elements.${element.id}`)
  const hasNativeTextBody = ((element.kind === 'text' || element.kind === 'shape') && element.textBody !== undefined) ||
    (element.kind === 'table' && element.table.rows.every((row) => row.every((cell) => cell.paragraphs !== undefined && cell.textBody !== undefined)))
  // DrawingML groups own a real child coordinate space. Preserve its affine
  // scale so strokes, shaped text, pictures, and descendants all inherit it.
  const base = element.kind === 'group'
    ? exactGroupBase(element, zIndex, state.budget)
    : elementBase(element, zIndex, state.budget, !hasNativeTextBody && element.kind !== 'connector')
  const world = checkedWorldAffine(parentWorld, base.transform, base.bounds, `$.elements.${element.id}`, state.budget)
  switch (element.kind) {
    case 'text': {
      if (!element.textBody) state.diagnostics.push({ severity: 'info', code: 'text.layoutMetadataUnavailable', message: 'legacy native PPTX text has no text-body layout; shaped compatibility preview remains clipped to element bounds', slideId: state.slide.id, elementId: element.id })
      const bounds = nativeTextBodyBounds(element, state)
      return { kind: 'text', ...base, textBody: await compileTextBody(element.paragraphs, { elementId: element.id, elementKind: 'text', bounds, layout: element.textBody }, state) }
    }
    case 'shape':
      if (element.paragraphs.length && !element.textBody) state.diagnostics.push({ severity: 'info', code: 'text.layoutMetadataUnavailable', message: 'legacy native PPTX shape text has no text-body layout; shaped compatibility preview remains clipped to element bounds', slideId: state.slide.id, elementId: element.id })
      if (!element.preset) throw new RenderCompileError('native.invalidShapePreset', `$.elements.${element.id}.preset`, 'non-refused shapes require a native preset')
      return {
        kind: 'shape', ...base, preset: element.preset, path: boundedPath(presetPath(element.preset, base.bounds.cx, base.bounds.cy), `$.elements.${element.id}.path`),
        fill: element.fill ? { color: element.fill } : undefined,
        stroke: element.stroke ? boundedStroke(element.stroke, `$.elements.${element.id}.stroke`, state.budget) : undefined,
        textBody: element.paragraphs.length || element.textBody ? await compileTextBody(element.paragraphs, { elementId: element.id, elementKind: 'shape', bounds: nativeTextBodyBounds(element, state), layout: element.textBody }, state) : undefined,
      }
    case 'connector':
      return {
        kind: 'connector', ...base, path: boundedPath(connectorPath(base.bounds.cx, base.bounds.cy, element.flipH ?? false), `$.elements.${element.id}.path`),
        stroke: element.stroke ? boundedStroke(element.stroke, `$.elements.${element.id}.stroke`, state.budget) : undefined,
        headArrow: element.headArrow ?? false, tailArrow: element.tailArrow ?? false,
        ...(element.headEnd?{headEnd:{...element.headEnd}}:{}),...(element.tailEnd?{tailEnd:{...element.tailEnd}}:{}),
      }
    case 'picture': {
      if (element.compatibility.diagnostics.some((diagnostic) => diagnostic.code === 'pptx.picture-geometry-unavailable')) {
        return { kind: 'placeholder', ...base, reason: 'preserveOnly', label: 'Unsupported picture geometry preserved' }
      }
      if (element.compatibility.diagnostics.some((diagnostic) => diagnostic.code === 'pptx.picture-crop-unavailable')) {
        return { kind: 'placeholder', ...base, reason: 'preserveOnly', label: 'Unsupported picture crop preserved' }
      }
      const asset = referenceAsset(element.assetId, element.id, state)
      const image: RenderImageNode = {
        kind: 'image', ...base, role: 'picture', assetId: asset.id, contentType: asset.contentType, sha256: asset.sha256,
        byteLength: asset.byteLength, resolutionSource: asset.dataBase64 === undefined ? 'host' : 'sourceDeck',
        ...(element.crop ? { crop: { ...element.crop } } : {}),
        ...(element.clip === 'roundRect' ? { clip: { kind: 'roundRect' as const, rect: base.bounds, radiusEmu: Math.round(Math.min(base.bounds.cx, base.bounds.cy) * 16667 / 100000) } } : {}),
      }
      return image
    }
    case 'chart': {
      if (!element.chart.previewAssetId) {
        state.diagnostics.push({ severity: 'refusal', code: 'chart.missingPreview', message: 'opaque chart has no preview asset to paint; no chart renderer is invented', slideId: state.slide.id, elementId: element.id })
        return { kind: 'placeholder', ...base, reason: 'missingPreview', label: 'Chart preview unavailable' }
      }
      const asset = referenceAsset(element.chart.previewAssetId, element.id, state)
      return {
        kind: 'image', ...base, role: 'chartPreview', assetId: asset.id, contentType: asset.contentType, sha256: asset.sha256,
        byteLength: asset.byteLength, resolutionSource: asset.dataBase64 === undefined ? 'host' : 'sourceDeck',
      }
    }
    case 'table':
      if (!hasNativeTextBody) state.diagnostics.push({ severity: 'info', code: 'text.layoutMetadataUnavailable', message: 'legacy native PPTX table cells lack full body/wrap metadata; shaped compatibility preview remains clipped to cell bounds', slideId: state.slide.id, elementId: element.id })
      return { kind: 'table', ...base, rows: element.table.rows.length, columns: element.table.columnWidths.length, cells: await compileTableCells(element, state) }
    case 'group': {
      const children: RenderNode[] = []
      for (let childIndex = 0; childIndex < element.children.length; childIndex++) {
        children.push(await compileElement(element.children[childIndex]!, childIndex, depth + 1, state, world))
      }
      const group: RenderGroupNode = { kind: 'group', ...base, children }
      return group
    }
  }
}

function deepFreeze<T>(value: T, seen = new Set<object>()): T {
  if (typeof value !== 'object' || value === null || seen.has(value)) return value
  seen.add(value)
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen)
  return Object.freeze(value)
}

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalize)
  if (typeof value !== 'object' || value === null) return value
  return Object.fromEntries(Object.entries(value as Record<string, unknown>).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0).map(([key, child]) => [key, canonicalize(child)]))
}

/** Compile one validated native slide into a deterministic, immutable integer-EMU tree. */
export async function compileNativePptxSlide(deckInput: NativePptxDeck, slide: number | string, options: CompileSlideOptions): Promise<SlideRenderTree> {
  const lineLayoutPolicy = options.lineLayoutPolicy
  if (options.sourceFrameAutoFitPreview !== undefined && typeof options.sourceFrameAutoFitPreview !== 'boolean') throw new RenderCompileError('render.invalidContract', '$.options.sourceFrameAutoFitPreview', 'source-frame autofit opt-in must be boolean')
  if (lineLayoutPolicy !== undefined && lineLayoutPolicy !== 'max-run-natural-v1') throw new RenderCompileError('render.invalidContract', '$.options.lineLayoutPolicy', 'unknown native line layout policy')
  assertNativePptx(deckInput)
  let deck: NativePptxDeck
  try {
    // Providers are asynchronous injected code. Compile from an owned frozen
    // snapshot so a provider closure cannot mutate the already-validated deck
    // between runs and create a validation/use time-of-check gap.
    deck = deepFreeze(structuredClone(deckInput))
  } catch {
    throw new RenderCompileError('render.invalidContract', '$', 'validated native deck could not be snapshotted for deterministic compilation')
  }
  const manifest = validateFontManifest(options.textLayout.manifest)
  if (!manifest.ok) throw new RenderCompileError('render.invalidFontManifest', '$.options.textLayout.manifest', manifest.issues.map((issue) => `${issue.path} ${issue.message}`).join('; '))
  const fontManifest = snapshotFontManifest(options.textLayout.manifest)
  const providers = snapshotProviders(options.textLayout)
  const liveDefaults = options.textLayout.defaults
  const textDefaults: NativePptxTextDefaults = deepFreeze({
    fontFamilies: [...liveDefaults.fontFamilies],
    fontSizeHundredthPt: liveDefaults.fontSizeHundredthPt,
    script: liveDefaults.script,
    language: liveDefaults.language,
    direction: liveDefaults.direction,
    ...(liveDefaults.fallbackChainIds === undefined ? {} : { fallbackChainIds: [...liveDefaults.fallbackChainIds] }),
  })
  const resolveRun = options.textLayout.resolveRun
  if (resolveRun !== undefined && typeof resolveRun !== 'function') throw new RenderCompileError('render.invalidTextInput', '$.options.textLayout.resolveRun', 'resolveRun must be callable')
  const slideIndex = typeof slide === 'number' ? slide : deck.slides.findIndex((candidate) => candidate.id === slide)
  if (!Number.isSafeInteger(slideIndex) || slideIndex < 0 || slideIndex >= deck.slides.length) {
    throw new RenderCompileError('render.slideReference', '$.slide', `slide ${String(slide)} does not exist`)
  }
  const nativeSlide = deck.slides[slideIndex]!
  const budget: Budget = {
    nodes: 0,
    glyphs: 0,
    clusters: 0,
    providerDecisions: 0,
    providerAttemptedFaceIds: 0,
    textLines: 0,
    textFragments: 0,
    maxDepth: configuredLimit(options.maxDepth, PPTX_RENDER_LIMITS.maxDepth, 'maxDepth'),
    maxNodes: configuredLimit(options.maxNodes, PPTX_RENDER_LIMITS.maxNodes, 'maxNodes'),
    maxGlyphs: configuredLimit(options.maxGlyphs, PPTX_RENDER_LIMITS.maxGlyphs, 'maxGlyphs'),
    maxClusters: configuredLimit(options.maxClusters, PPTX_RENDER_LIMITS.maxClusters, 'maxClusters'),
    maxCoordinateEmu: configuredLimit(options.maxCoordinateEmu, PPTX_RENDER_LIMITS.maxCoordinateEmu, 'maxCoordinateEmu'),
  }
  checkCoordinate(deck.size.cx, '$.size.cx', budget, true)
  checkCoordinate(deck.size.cy, '$.size.cy', budget, true)
  const state: CompileState = {
    lineLayoutPolicy,
    sourceFrameAutoFitPreview: options.sourceFrameAutoFitPreview === true,
    deck,
    slide: nativeSlide,
    options,
    assets: new Map(deck.assets.map((asset) => [asset.id, asset])),
    referencedAssets: new Map(),
    hostResolutionDiagnosticElements: new Set(),
    diagnostics: [],
    budget,
    providers,
    fontManifest,
    fontResources: new Map(),
    providerFontResources: new Map(),
    nativeTextInheritanceUnresolved: nativeSlide.compatibility.diagnostics.some((diagnostic) =>
      diagnostic.code === 'pptx.unsupported-presentation-child' ||
      diagnostic.code === 'pptx.unsupported-layout-dependency' ||
      diagnostic.code === 'pptx.unsupported-master-dependency' ||
      diagnostic.code === 'pptx.unsupported-theme-dependency'),
    textDefaults,
    ...(resolveRun === undefined ? {} : { resolveRun }),
    providerResolveCalls: 0,
    providerLoadCalls: 0,
    providerShapeCalls: 0,
    uniqueFontBytes: 0,
    providerFacingFontBytes: 0,
  }
  copyNativeDiagnostics(state, nativeSlide.compatibility)
  const nodes: RenderNode[] = []
  const hasSlideLevelRefusal = nativeSlide.compatibility.status === 'refused' && nativeSlide.compatibility.diagnostics.some((diagnostic) => diagnostic.severity === 'refusal' && diagnostic.scope?.elementId === undefined)
  if (hasSlideLevelRefusal) {
    state.diagnostics.push({ severity: 'refusal', code: 'native.refused', message: 'native compatibility refused this slide', slideId: nativeSlide.id })
    takeNode(state, '$.slide')
    nodes.push({
      kind: 'placeholder', sourceElementId: nativeSlide.id, sourceKind: 'slide', zIndex: 0,
      transform: translationTransform(0, 0), bounds: localBounds(deck.size.cx, deck.size.cy),
      clip: { kind: 'rect', rect: localBounds(deck.size.cx, deck.size.cy) }, compatibility: 'refused',
      reason: 'refused', label: 'Slide rendering refused',
    })
  } else {
    if (nativeSlide.compatibility.status === 'preserveOnly') {
      state.diagnostics.push({ severity: 'warning', code: 'render.preserveOnly', message: 'slide is rendered for preview but remains preserve-only', slideId: nativeSlide.id })
    }
    for (let index = 0; index < nativeSlide.elements.length; index++) nodes.push(await compileElement(nativeSlide.elements[index]!, index, 1, state, IDENTITY_WORLD_AFFINE))
  }
  const tree: SlideRenderTree = {
    version: PPTX_RENDER_TREE_VERSION,
    documentId: deck.documentId,
    slideId: nativeSlide.id,
    slideIndex,
    size: { ...deck.size },
    background: { color: nativeSlide.background ?? DEFAULT_BACKGROUND },
    clip: { kind: 'rect', rect: localBounds(deck.size.cx, deck.size.cy) },
    nodes,
    assets: [...state.referencedAssets.values()].sort((a, b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0),
    diagnostics: state.diagnostics,
  }
  return deepFreeze(tree)
}

/** Canonical snapshots sort map-like object keys while preserving every ordered array. */
export function stringifySlideRenderTree(tree: SlideRenderTree): string {
  return JSON.stringify(canonicalize(tree))
}
