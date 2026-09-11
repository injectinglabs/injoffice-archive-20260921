/**
 * Renderer-neutral DOCX shaping and Unicode line-breaking core.
 *
 * This module consumes the native DOCX parse/style projections, delegates font
 * resolution and shaping to injected providers, and emits deterministic lines
 * in integer milli-points. It deliberately has no DOM, HTML, canvas, painting,
 * pagination, or package-mutation dependency.
 */

import {
  NATIVE_TEXT_LAYOUT_VERSION,
  MAX_TEXT_RUN_UTF16,
  validateFontManifest,
  validateTextRunInput,
  scaleLineMetrics,
  asciiLower,
  normalizeFontFamilyName,
  type NativeFontManifest,
  type NativeFontResolver,
  type FontResource,
  type NativeTextDecision,
  type NativeTextRefusal,
  type NativeTextShaper,
  type ResolvedFontFace,
  type ScaledLineMetrics,
  type ShapedCluster,
  type ShapedGlyph,
  type ShapedSegment,
  type TextRunInput,
} from '@injoffice/font-metrics/layout'
import {
  UNICODE_13_TABLES_RUNTIME_MATCH,
  UNICODE_13_CLASSIFIER_REVISION,
  isUnicode13BreakableWhiteSpace,
  isUnicode13TextWhiteSpace,
  unicode13Punctuation,
  unicode13Script,
} from '@injoffice/font-metrics/unicode13'
import {
  BIDI_UNICODE_VERSION,
  NATIVE_BIDI_PROVIDER_ID,
  NATIVE_BIDI_PROVIDER_REVISION,
  reorderNativeBidiLineV1,
  resolveNativeBidiParagraphV1,
} from '@injoffice/font-metrics/bidi'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  DOCX_NATIVE_LIMITS,
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxIssueCode,
  type NativeDocxParagraphV1,
  type NativeDocxRunV1,
  type NativeDocxStoryKind,
  type NativeDocxStoryV1,
  type NativeDocxTableV1,
  type NativeDocxValidationIssue,
} from './nativeContract.js'
import {
  decodeNativeDocxResolvedLayout,
  type NativeDocxResolvedLayoutInputV1,
  type NativeDocxResolvedNumberingSourceV1,
  type NativeDocxResolvedParagraphPropertiesV1,
  type NativeDocxResolvedParagraphV1,
  type NativeDocxResolvedRunPropertiesV1,
  type NativeDocxResolvedRunV1,
} from './nativeResolvedLayout.js'
import { qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { resolveNativeDocxParagraphBidiPlanV1, type NativeDocxParagraphBidiPlanV1 } from './nativeBidiPlanV1.js'
import { nativeDocxListSuffixTabTargetV1, positionNativeDocxListMarkerV1 } from './nativeNumberingV1.js'
import { compareNativeValidationIssues } from './nativeDeterminism.js'
import { readNativeDocxScriptTransformV1, nativeDocxScriptScaleV1, nativeDocxScriptShiftV1, type NativeDocxScriptTransformV1 } from './nativeScriptLayoutV1.js'

export const DOCX_SHAPED_LINES_PROTOCOL = 'injoffice.docx.shaped-lines'
export const DOCX_SHAPED_LINES_VERSION = 1 as const
export const DOCX_SHAPING_REQUEST_PROTOCOL = 'injoffice.docx.shaping-request'
export const DOCX_SHAPING_REQUEST_VERSION = 1 as const
export const DOCX_SHAPED_LINES_LIMITS = {
  maxParagraphs: 10_000,
  maxLines: 100_000,
  maxFragments: 500_000,
  maxDiagnostics: 1_000,
  maxWidthMilliPoints: 1_000_000_000,
  maxFontResourceBytes: 64 * 1024 * 1024,
  maxUniqueFontBytes: 128 * 1024 * 1024,
  maxCachedFontResources: 256,
  maxProviderResolveCalls: 50_000,
  maxProviderLoadCalls: 256,
  maxProviderShapeCalls: 50_000,
} as const

export interface NativeDocxShapingRequestV1 {
  protocol: typeof DOCX_SHAPING_REQUEST_PROTOCOL
  version: typeof DOCX_SHAPING_REQUEST_VERSION
  document: NativeDocxDocumentV1
  resolved_layout: NativeDocxResolvedLayoutInputV1
  font_manifest: NativeFontManifest
  /** Width offered by the future paginator, in integer 1/1000 point. */
  available_width_millipoints: number
  /**
   * Explicit host policy for w:tab and list suffix tabs; never inferred from CSS.
   * Stops are relative to each line's paragraph-content start after indentation.
   */
  tab_interval_millipoints: number
}

export interface NativeDocxShapingProviders {
  resolver: NativeFontResolver
  shaper: NativeTextShaper
}

export type NativeDocxShapingDiagnosticCode =
  | 'provider-refusal'
  | 'provider-decision'
  | 'provider-failure'
  | 'invalid-provider-output'
  | 'missing-run-font'
  | 'missing-run-size'
  | 'unresolved-layout-diagnostic'
  | 'paint-diagnostic-preserved'
  | 'bidi-resolution-refusal'
  | 'justification-unsupported'
  | 'table-layout-unsupported'
  | 'drawing-layout-unsupported'
  | 'reference-layout-unsupported'
  | 'page-control-deferred'
  | 'soft-hyphen-deferred'
  | 'unsupported-numbering-text'
  | 'unsupported-numbering-format'
  | 'list-marker-alignment-deferred'
  | 'list-marker-tab-deferred'
  | 'numbering-state-invalidated'
  | 'empty-line-metrics-unresolved'
  | 'cluster-overflow'
  | 'resource-limit'
  | 'diagnostic-overflow'

export interface NativeDocxShapingDiagnosticV1 {
  code: NativeDocxShapingDiagnosticCode
  severity: 'unsupported' | 'deferred'
  scope_id: string
  source_id?: string
  source_diagnostic_code?: string
  source_diagnostic_message?: string
  message: string
}

export interface NativeDocxPositionedGlyphV1 {
  glyph_id: number
  advance_x_millipoints: number
  advance_y_millipoints: number
  offset_x_millipoints: number
  offset_y_millipoints: number
}

export interface NativeDocxLineFragmentV1 {
  script_transform?: NativeDocxScriptTransformV1
  id: string
  source_kind: 'run' | 'list-marker' | 'tab' | 'image'
  /** Native run id for authored text/controls; paragraph id for a list marker. */
  source_id: string
  start_utf16: number
  end_utf16: number
  text: string
  direction: 'ltr' | 'rtl'
  bidi_level: number
  logical_order: number
  script: string
  language: string
  face_id?: string
  whitespace: boolean
  advance_inline_millipoints: number
  justification_expansion_millipoints: number
  ascent_millipoints: number
  descent_millipoints: number
  line_gap_millipoints: number
  /** Scaled metrics from the resolved font's post table, only for decorated text. */
  underline_position_millipoints?: number
  underline_thickness_millipoints?: number
  glyphs: NativeDocxPositionedGlyphV1[]
}

export interface NativeDocxHardBreakV1 {
  source_run_id: string
  control: 'line-break'
}

export interface NativeDocxShapedLineV1 {
  id: string
  ordinal: number
  available_width_millipoints: number
  inline_offset_millipoints: number
  advance_inline_millipoints: number
  ascent_millipoints: number
  descent_millipoints: number
  line_gap_millipoints: number
  line_height_millipoints: number
  justified: boolean
  /** Index by logical cluster order; each value is its index in visual fragments. */
  logical_to_visual: number[]
  fragments: NativeDocxLineFragmentV1[]
  hard_break_after?: NativeDocxHardBreakV1
}

export interface NativeDocxShapedParagraphV1 {
  paragraph_id: string
  story_id: string
  story_kind: NativeDocxStoryKind
  direction: 'ltr' | 'rtl'
  alignment: 'left' | 'right' | 'center' | 'both' | 'distribute' | 'start' | 'end'
  spacing_before_millipoints: number
  spacing_after_millipoints: number
  indent_start_millipoints: number
  indent_end_millipoints: number
  first_line_delta_millipoints: number
  list_marker?: NativeDocxShapedListMarkerV1
  block_advance_millipoints: number
  lines: NativeDocxShapedLineV1[]
}

export interface NativeDocxShapedListMarkerV1 {
  marker_id: string
  definition_sha256: string
  numbering_part_sha256: string
  model_sha256: string
  num_id: string
  abstract_num_id: string
  level: number
  counter_value: number
  text: string
  suffix: 'tab' | 'space' | 'nothing'
  alignment: 'left' | 'right' | 'center' | 'start' | 'end'
  label_start_millipoints: number
  label_end_millipoints: number
  marker_start_millipoints: number
  marker_advance_millipoints: number
  text_start_millipoints: number
}

export interface NativeDocxShapedLinesV1 {
  protocol: typeof DOCX_SHAPED_LINES_PROTOCOL
  version: typeof DOCX_SHAPED_LINES_VERSION
  document_id: string
  revision: string
  numbering_source?: NativeDocxResolvedNumberingSourceV1
  available_width_millipoints: number
  tab_interval_millipoints: number
  font_manifest: { manifest_id: string; revision: string }
  providers: { resolver_id: string; resolver_revision: string; shaper_id: string; shaper_revision: string; bidi_id: string; bidi_revision: string; bidi_unicode_version: string; unicode13_revision: string }
  paragraphs: NativeDocxShapedParagraphV1[]
  diagnostics: NativeDocxShapingDiagnosticV1[]
}

export type ShapeNativeDocxLinesResult =
  | { ok: true; value: NativeDocxShapedLinesV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

const REQUEST_FIELDS = ['protocol', 'version', 'document', 'resolved_layout', 'font_manifest', 'available_width_millipoints', 'tab_interval_millipoints'] as const
const MAX_PROVIDER_METRIC_MILLIPOINTS = 1_000_000_000
const MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS = 10_000_000_000
const MAX_PARAGRAPH_BLOCK_MILLIPOINTS = 1_000_000_000_000
const MAX_PROVIDER_DECISIONS = 1_000
const MAX_PROVIDER_ATTEMPTED_FACES = 4_096
const MAX_PROVIDER_DECISIONS_TOTAL = 10_000
const MAX_PROVIDER_ATTEMPTED_FACES_TOTAL = 100_000
const MAX_PROVIDER_MESSAGE_LENGTH = 4_096
const MAX_PROVIDER_ERROR_LENGTH = 2_048
const PROVIDER_ID_RE = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const PROVIDER_DECISION_CODES = new Set(['invalid-contract', 'font-not-found', 'font-bytes-unavailable', 'font-digest-mismatch', 'font-metrics-unavailable', 'unsupported-font-format', 'unsupported-script', 'unsupported-direction', 'unsupported-feature', 'missing-glyph', 'provider-failure'])
const PAINT_ONLY_RESOLUTION_DIAGNOSTICS = new Set([
  'INVALID_COLOR',
  'THEME_COLOR_PRESERVED',
  'UNSUPPORTED_HIGHLIGHT',
  'UNSUPPORTED_UNDERLINE',
  'UNDERLINE_COLOR_PRESERVED',
  'THEME_UNDERLINE_COLOR_PRESERVED',
])
const TABLE_ONLY_RESOLUTION_DIAGNOSTICS = new Set([
  'MISSING_TABLE_STYLE',
  'TABLE_STYLE_EFFECTS_PRESERVED',
  'CONDITIONAL_TABLE_STYLE_PRESERVED',
])
const BIDI_TRAILING_RE = /^[\u0009-\u000d\u001c-\u001e\u0020\u0085\u2028\u2029]+$/u
const GLUE_RE = /[\u00A0\u202F\u2060]/u

interface NativeShapingContext {
  request: NativeDocxShapingRequestV1
  providers: NativeProviderSnapshot
  fontManifest: NativeFontManifest
  paragraphs: Map<string, NativeDocxResolvedParagraphV1>
  runs: Map<string, NativeDocxResolvedRunV1>
  noteNumbers: Map<string, string>
  fontAliases: Map<string, string>
  blockingDiagnostics: Map<string, NativeDocxResolvedLayoutInputV1['diagnostics']>
  diagnostics: NativeDocxShapingDiagnosticV1[]
  diagnosticKeys: Set<string>
  lineCount: number
  fragmentCount: number
  atomCount: number
  shapedCodeUnits: number
  resourceExceeded: boolean
  diagnosticOverflow: boolean
  documentBlocked: boolean
  providerDecisionCount: number
  providerAttemptedFaceCount: number
  providerResolveCalls: number
  providerLoadCalls: number
  providerShapeCalls: number
  uniqueFontBytes: number
  providerFacingFontBytes: number
  fontResources: Map<string, FontResource>
  providerFontResources: Map<string, FontResource>
  activeParagraphID?: string
  activeParagraphFailed: boolean
  activeAvailableWidthMilliPoints?: number
  numberingFailed: boolean
}

interface NativeProviderSnapshot {
  resolver: {
    providerId: string
    providerRevision: string
    live: NativeFontResolver
    resolve: NativeFontResolver['resolve']
    load: NativeFontResolver['load']
  }
  shaper: {
    providerId: string
    providerRevision: string
    live: NativeTextShaper
    shape: NativeTextShaper['shape']
  }
}

interface SourceSpan {
  text: string
  startUtf16: number
  endUtf16: number
  script: string
  direction: 'ltr' | 'rtl'
  bidiLevel: number
  language: string
}

interface FragmentAtom {
  scriptTransform?: NativeDocxScriptTransformV1
  sourceKind: 'run' | 'list-marker' | 'tab' | 'image'
  sourceID: string
  startUtf16: number
  endUtf16: number
  text: string
  direction: 'ltr' | 'rtl'
  bidiLevel: number
  script: string
  language: string
  faceID?: string
  whitespace: boolean
  unsafeToBreak: boolean
  breakAfter: boolean
  dynamicTab: boolean
  advance: number
  metrics: ScaledLineMetrics
  glyphs: NativeDocxPositionedGlyphV1[]
}

type ParagraphEvent = { kind: 'atom'; atom: FragmentAtom } | { kind: 'hard-break'; runID: string }

function validationIssue(code: NativeDocxIssueCode, path: string, message: string): NativeDocxValidationIssue {
  return { code, path, message }
}

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function deepFreezeWire<T>(value: T): T {
  const pending: object[] = []
  const seen = new Set<object>()
  if (value !== null && typeof value === 'object') pending.push(value)
  while (pending.length > 0) {
    const current = pending.pop()!
    if (seen.has(current)) continue
    seen.add(current)
    for (const child of Object.values(current)) if (child !== null && typeof child === 'object') pending.push(child)
    Object.freeze(current)
  }
  return value
}

function validateRequest(value: unknown): { ok: true; value: NativeDocxShapingRequestV1 } | { ok: false; issues: NativeDocxValidationIssue[] } {
  if (!isRecord(value)) return { ok: false, issues: [validationIssue('INVALID_TYPE', '', 'shaping request must be an object')] }
  const issues: NativeDocxValidationIssue[] = []
  const allowed = new Set<string>(REQUEST_FIELDS)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) issues.push(validationIssue('UNKNOWN_FIELD', `/${key}`, `unknown field ${JSON.stringify(key)}`))
  if (value.protocol !== DOCX_SHAPING_REQUEST_PROTOCOL) issues.push(validationIssue('UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_SHAPING_REQUEST_PROTOCOL}`))
  if (value.version !== DOCX_SHAPING_REQUEST_VERSION) issues.push(validationIssue('UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_SHAPING_REQUEST_VERSION}`))
  const document = decodeNativeDocxDocument(value.document)
  if (!document.ok) issues.push(...document.issues.map((issue) => ({ ...issue, path: `/document${issue.path}` })))
  const resolved = decodeNativeDocxResolvedLayout(value.resolved_layout)
  if (!resolved.ok) issues.push(...resolved.issues.map((issue) => ({ ...issue, path: `/resolved_layout${issue.path}` })))
  const manifest = validateFontManifest(value.font_manifest)
  if (!manifest.ok) {
    issues.push(...manifest.issues.map((issue) => validationIssue(issue.code === 'unknown-field' ? 'UNKNOWN_FIELD' : issue.code === 'reference' ? 'BROKEN_REFERENCE' : 'INVALID_VALUE', `/font_manifest${issue.path === '$' ? '' : issue.path.slice(1).replace(/\./g, '/')}`, issue.message)))
  }
  for (const [key, max] of [['available_width_millipoints', DOCX_SHAPED_LINES_LIMITS.maxWidthMilliPoints], ['tab_interval_millipoints', DOCX_SHAPED_LINES_LIMITS.maxWidthMilliPoints]] as const) {
    const number = value[key]
    if (!Number.isSafeInteger(number) || Object.is(number, -0) || (number as number) <= 0 || (number as number) > max) issues.push(validationIssue('OUT_OF_RANGE', `/${key}`, `must be a positive safe integer no greater than ${max}`))
  }
  if (document.ok && resolved.ok) validateJoins(document.value, resolved.value, issues)
  issues.sort(compareNativeValidationIssues)
  return issues.length > 0 ? { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) } : { ok: true, value: value as unknown as NativeDocxShapingRequestV1 }
}

function snapshotProviderBoundary(providers: NativeDocxShapingProviders): { ok: true; value: NativeProviderSnapshot } | { ok: false; issues: NativeDocxValidationIssue[] } {
  const issues: NativeDocxValidationIssue[] = []
  try {
    if (!isRecord(providers) || !isRecord(providers.resolver) || !isRecord(providers.shaper)) return { ok: false, issues: [validationIssue('INVALID_TYPE', '/providers', 'resolver and shaper providers must be objects')] }
    const resolver = providers.resolver
    const shaper = providers.shaper
    const resolverID = resolver.providerId
    const resolverRevision = resolver.providerRevision
    const shaperID = shaper.providerId
    const shaperRevision = shaper.providerRevision
    const resolve = resolver.resolve
    const load = resolver.load
    const shape = shaper.shape
    for (const [kind, id, revision] of [['resolver', resolverID, resolverRevision], ['shaper', shaperID, shaperRevision]] as const) {
      if (typeof id !== 'string' || !PROVIDER_ID_RE.test(id)) issues.push(validationIssue('INVALID_VALUE', `/providers/${kind}/provider_id`, 'must be a bounded stable provider identifier'))
      if (typeof revision !== 'string' || !PROVIDER_ID_RE.test(revision)) issues.push(validationIssue('INVALID_VALUE', `/providers/${kind}/provider_revision`, 'must be a bounded stable provider revision'))
    }
    if (typeof resolve !== 'function' || typeof load !== 'function') issues.push(validationIssue('INVALID_TYPE', '/providers/resolver', 'resolver must expose resolve and load functions'))
    if (typeof shape !== 'function') issues.push(validationIssue('INVALID_TYPE', '/providers/shaper', 'shaper must expose a shape function'))
    if (resolver.providerId !== resolverID || resolver.providerRevision !== resolverRevision || shaper.providerId !== shaperID || shaper.providerRevision !== shaperRevision || resolver.resolve !== resolve || resolver.load !== load || shaper.shape !== shape) issues.push(validationIssue('INVALID_VALUE', '/providers', 'provider identities and callable references must remain stable while snapshotted'))
    if (issues.length > 0 || typeof resolverID !== 'string' || typeof resolverRevision !== 'string' || typeof shaperID !== 'string' || typeof shaperRevision !== 'string' || typeof resolve !== 'function' || typeof load !== 'function' || typeof shape !== 'function') return { ok: false, issues }
    return {
      ok: true,
      value: {
        resolver: { providerId: resolverID, providerRevision: resolverRevision, live: resolver, resolve: resolve.bind(resolver), load: load.bind(resolver) },
        shaper: { providerId: shaperID, providerRevision: shaperRevision, live: shaper, shape: shape.bind(shaper) },
      },
    }
  } catch {
    issues.push(validationIssue('INVALID_VALUE', '/providers', 'provider identity access must be deterministic and side-effect free'))
    return { ok: false, issues }
  }
}

function snapshotFontManifest(manifest: NativeFontManifest): NativeFontManifest {
  const faces = manifest.faces.map((face) => Object.freeze({
    ...face,
    ...(face.aliases ? { aliases: Object.freeze([...face.aliases]) } : {}),
    ...(face.scripts ? { scripts: Object.freeze([...face.scripts]) } : {}),
    ...(face.languages ? { languages: Object.freeze([...face.languages]) } : {}),
    source: Object.freeze({ ...face.source }),
  }))
  const fallbackChains = manifest.fallbackChains.map((chain) => Object.freeze({
    ...chain,
    faceIds: Object.freeze([...chain.faceIds]),
    ...(chain.scripts ? { scripts: Object.freeze([...chain.scripts]) } : {}),
    ...(chain.languages ? { languages: Object.freeze([...chain.languages]) } : {}),
  }))
  return Object.freeze({ ...manifest, faces: Object.freeze(faces), fallbackChains: Object.freeze(fallbackChains) })
}

function providerIdentityStable(providers: NativeProviderSnapshot): boolean {
  try {
    return providers.resolver.live.providerId === providers.resolver.providerId
      && providers.resolver.live.providerRevision === providers.resolver.providerRevision
      && providers.shaper.live.providerId === providers.shaper.providerId
      && providers.shaper.live.providerRevision === providers.shaper.providerRevision
  } catch {
    return false
  }
}

interface NativeIdentityInventory {
  paragraphs: Map<string, { paragraph: NativeDocxParagraphV1; ownerStory: NativeDocxStoryV1 }>
  runs: Map<string, string>
  tables: Map<string, NativeDocxTableV1>
}

function nativeStories(document: NativeDocxDocumentV1): NativeDocxStoryV1[] {
  return [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
}

interface QualifiedNoteNumbers {
  numbers: Map<string, string>
  issues: Array<{ scope: string; source?: string; message: string }>
}

/** Admit only one acyclic, relationship-bound, source-ordered decimal note graph. */
function qualifyNoteNumbers(document: NativeDocxDocumentV1): QualifiedNoteNumbers {
  const numbers = new Map<string, string>()
  const issues: QualifiedNoteNumbers['issues'] = []
  const contentNotes = new Map(document.notes.filter((story) => (story.note_role ?? 'content') === 'content').map((story) => [story.id, story]))
  const references = new Map<string, string>()
  const counters = new Map<'footnote' | 'endnote', number>([['footnote', 0], ['endnote', 0]])
  const relationships = new Map<'footnote' | 'endnote', string>()
  const relationshipKinds = new Map<string, 'footnote' | 'endnote'>()
  for (const note of document.notes) {
    // Continuation separator stories are inert until continuation is needed.
    // Bounded v1 never shapes them because actual continuation is refused.
    if (note.note_role === 'continuation-separator') continue
    if (!note.relationship_id) issues.push({ scope: note.id, message: 'Note story is not bound to its resolved main-document relationship' })
    else {
      const known = relationships.get(note.kind as 'footnote' | 'endnote')
      if (known !== undefined && known !== note.relationship_id) issues.push({ scope: note.id, message: 'One note kind resolves through more than one relationship identity' })
      const knownKind = relationshipKinds.get(note.relationship_id)
      if (knownKind !== undefined && knownKind !== note.kind) issues.push({ scope: note.id, message: 'One relationship identity cannot resolve both footnotes and endnotes' })
      relationships.set(note.kind as 'footnote' | 'endnote', note.relationship_id)
      relationshipKinds.set(note.relationship_id, note.kind as 'footnote' | 'endnote')
    }
    if ((note.note_role ?? 'content') !== 'content') {
      for (const block of note.blocks) {
        if (block.kind !== 'paragraph' || !block.paragraph) issues.push({ scope: block.id, message: 'Special note stories must contain only exact paintable paragraphs' })
        else for (const run of block.paragraph.runs) if (run.reference) issues.push({ scope: note.id, source: run.id, message: 'Special note stories cannot contain note references or labels' })
      }
      continue
    }
    let labels = 0
    for (const block of note.blocks) {
      if (block.kind !== 'paragraph' || !block.paragraph) {
        issues.push({ scope: block.id, message: 'Nested note tables are outside exact note page paint' })
        continue
      }
      for (const run of block.paragraph.runs) if (run.reference) {
        if (run.reference.role === 'label' && run.reference.kind === note.kind && run.reference.target_id === note.id) labels += 1
        else issues.push({ scope: note.id, source: run.id, message: 'Note-to-note references, cycles, and mismatched labels are unsupported' })
      }
    }
    if (labels !== 1) issues.push({ scope: note.id, message: 'Each content note must contain exactly one exact owning note label' })
  }
  for (const block of document.body.blocks) {
    if (block.kind !== 'paragraph' || !block.paragraph) continue
    for (const run of block.paragraph.runs) {
      const ref = run.reference
      if (!ref || ref.role === 'label' || (ref.kind !== 'footnote' && ref.kind !== 'endnote')) continue
      const target = contentNotes.get(ref.target_id)
      if (!target || target.kind !== ref.kind) {
        issues.push({ scope: block.paragraph.id, source: run.id, message: 'Body note anchor does not resolve to an exact content note of the same kind' })
        continue
      }
      if (references.has(target.id)) {
        issues.push({ scope: target.id, source: run.id, message: 'Duplicate note references make numbering and placement ambiguous' })
        continue
      }
      const next = (counters.get(ref.kind) ?? 0) + 1
      if (!Number.isSafeInteger(next) || next > DOCX_SHAPED_LINES_LIMITS.maxParagraphs) {
        issues.push({ scope: target.id, source: run.id, message: 'Note numbering exceeds the bounded decimal counter' })
        continue
      }
      counters.set(ref.kind, next)
      references.set(target.id, run.id)
      numbers.set(target.id, String(next))
    }
  }
  for (const note of contentNotes.values()) if (!references.has(note.id)) issues.push({ scope: note.id, message: 'Unreferenced note content has ambiguous placement and numbering' })
  return { numbers, issues }
}

function inventoryDocument(document: NativeDocxDocumentV1): NativeIdentityInventory {
  const inventory: NativeIdentityInventory = { paragraphs: new Map(), runs: new Map(), tables: new Map() }
  const addParagraph = (paragraph: NativeDocxParagraphV1, story: NativeDocxStoryV1): void => {
    inventory.paragraphs.set(paragraph.id, { paragraph, ownerStory: story })
    for (const run of paragraph.runs) inventory.runs.set(run.id, paragraph.id)
  }
  for (const story of nativeStories(document)) {
    for (const block of story.blocks) {
      if (block.kind === 'paragraph' && block.paragraph) addParagraph(block.paragraph, story)
      if (block.kind === 'table' && block.table) {
        inventory.tables.set(block.table.id, block.table)
        for (const row of block.table.rows) for (const cell of row.cells) for (const paragraph of cell.paragraphs) addParagraph(paragraph, story)
      }
    }
  }
  return inventory
}

function validateJoins(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, issues: NativeDocxValidationIssue[]): void {
  if (document.document_id !== resolved.document_id) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/document_id', 'must equal document.document_id'))
  if (document.revision !== resolved.revision) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/revision', 'must equal document.revision'))
  if (canonicalOPCPartKey(document.source.main_part) !== canonicalOPCPartKey(resolved.source_parts.main_part)) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/source_parts/main_part', 'must equal document.source.main_part by OPC case-insensitive URI equivalence'))
  if (resolved.numbering_source) {
    const source = resolved.numbering_source
    const relationshipPart = document.passthrough_parts.find((part) => canonicalOPCPartKey(part.part_name) === canonicalOPCPartKey(source.relationships_part))
    if (!relationshipPart || asciiLower(relationshipPart.content_type) !== 'application/vnd.openxmlformats-package.relationships+xml' || relationshipPart.sha256 !== source.relationships_sha256) {
      issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/numbering_source/relationships_part', 'must bind the exact preserved owning relationship part, content type, and SHA-256'))
    }
    const numberingPart = document.passthrough_parts.find((part) => canonicalOPCPartKey(part.part_name) === canonicalOPCPartKey(source.part_name))
    if (!numberingPart || asciiLower(numberingPart.content_type) !== asciiLower(source.content_type) || numberingPart.sha256 !== source.part_sha256) {
      issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/numbering_source/part_name', 'must bind the exact preserved numbering part, content type, and SHA-256'))
    }
  } else if (resolved.paragraphs.some((paragraph) => paragraph.numbering !== undefined)) {
    issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/numbering_source', 'resolved markers require exact numbering source attestation'))
  }
  const inventory = inventoryDocument(document)
  const resolvedParagraphs = new Map(resolved.paragraphs.map((entry, index) => [entry.paragraph_id, index]))
  const resolvedRuns = new Map(resolved.runs.map((entry, index) => [entry.run_id, { entry, index }]))
  const resolvedTables = new Map(resolved.tables.map((entry, index) => [entry.table_id, index]))
  resolved.paragraphs.forEach((paragraph, index) => {
    const properties = paragraph.properties
    const boundedTwips = ['spacing_before_twips', 'spacing_after_twips', 'indent_left_twips', 'indent_right_twips', 'indent_start_twips', 'indent_end_twips', 'first_line_twips', 'hanging_twips'] as const
    for (const key of boundedTwips) {
      const value = properties[key]
      if (value !== undefined && Math.abs(value) > Math.floor(MAX_PROVIDER_METRIC_MILLIPOINTS / 50)) issues.push(validationIssue('LIMIT_EXCEEDED', `/resolved_layout/paragraphs/${index}/properties/${key}`, `converted measurement must not exceed ${MAX_PROVIDER_METRIC_MILLIPOINTS} milli-points`))
    }
    if (properties.line !== undefined && properties.line_rule === 'auto' && properties.line > 2_400) issues.push(validationIssue('LIMIT_EXCEEDED', `/resolved_layout/paragraphs/${index}/properties/line`, 'automatic line spacing must not exceed 10x within the bounded shaping core'))
    if (properties.line !== undefined && properties.line_rule !== 'auto' && properties.line > Math.floor(MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS / 50)) issues.push(validationIssue('LIMIT_EXCEEDED', `/resolved_layout/paragraphs/${index}/properties/line`, `converted line height must not exceed ${MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS} milli-points`))
  })
  for (const id of inventory.paragraphs.keys()) if (!resolvedParagraphs.has(id)) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/paragraphs', `missing resolved paragraph ${JSON.stringify(id)}`))
  for (const [id, index] of resolvedParagraphs) if (!inventory.paragraphs.has(id)) issues.push(validationIssue('BROKEN_REFERENCE', `/resolved_layout/paragraphs/${index}/paragraph_id`, 'does not reference a native paragraph'))
  for (const [id, paragraphID] of inventory.runs) {
    const resolvedRun = resolvedRuns.get(id)
    if (!resolvedRun) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/runs', `missing resolved run ${JSON.stringify(id)}`))
    else if (resolvedRun.entry.paragraph_id !== paragraphID) issues.push(validationIssue('BROKEN_REFERENCE', `/resolved_layout/runs/${resolvedRun.index}/paragraph_id`, `must equal owning paragraph ${JSON.stringify(paragraphID)}`))
  }
  for (const [id, resolvedRun] of resolvedRuns) if (!inventory.runs.has(id)) issues.push(validationIssue('BROKEN_REFERENCE', `/resolved_layout/runs/${resolvedRun.index}/run_id`, 'does not reference a native run'))
  for (const id of inventory.tables.keys()) if (!resolvedTables.has(id)) issues.push(validationIssue('BROKEN_REFERENCE', '/resolved_layout/tables', `missing resolved table ${JSON.stringify(id)}`))
  for (const [id, index] of resolvedTables) if (!inventory.tables.has(id)) issues.push(validationIssue('BROKEN_REFERENCE', `/resolved_layout/tables/${index}/table_id`, 'does not reference a native table'))
}

function canonicalOPCPartKey(value: string): string {
  try {
    return asciiLower(value.split('/').map((segment) => decodeURIComponent(segment)).join('/'))
  } catch {
    return asciiLower(value)
  }
}

function addDiagnostic(context: NativeShapingContext, diagnostic: NativeDocxShapingDiagnosticV1): void {
  if (context.activeParagraphID !== undefined && diagnostic.severity === 'unsupported') context.activeParagraphFailed = true
  const key = `${diagnostic.code}\u0000${diagnostic.scope_id}\u0000${diagnostic.source_id ?? ''}\u0000${diagnostic.message}`
  if (context.diagnosticKeys.has(key)) return
  context.diagnosticKeys.add(key)
  if (context.diagnostics.length >= DOCX_SHAPED_LINES_LIMITS.maxDiagnostics - 1) {
    if (!context.diagnosticOverflow) {
      context.diagnosticOverflow = true
      context.diagnostics.push({
        code: 'diagnostic-overflow',
        severity: 'unsupported',
        scope_id: context.request.document.document_id,
        message: `Shaping diagnostics exceeded ${DOCX_SHAPED_LINES_LIMITS.maxDiagnostics - 1}; remaining diagnostics were suppressed`,
      })
    }
    return
  }
  context.diagnostics.push(diagnostic)
}

/** Convert OOXML twips to the shared integer 1/1000-point unit. */
export function twipsToMilliPoints(twips: number): number {
  if (!Number.isSafeInteger(twips) || Object.is(twips, -0) || Math.abs(twips) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS) throw new RangeError('twips must be exactly convertible to bounded integer milli-points and not negative zero')
  return twips * 50
}

/** Convert Word half-points to the shared integer 1/1000-point unit. */
export function halfPointsToMilliPoints(halfPoints: number): number {
  if (!Number.isSafeInteger(halfPoints) || Object.is(halfPoints, -0) || halfPoints <= 0) throw new RangeError('half-points must be a positive safe integer')
  const value = halfPoints * 500
  if (!Number.isSafeInteger(value)) throw new RangeError('half-point conversion exceeds deterministic integer precision')
  return value
}

function scriptForCharacter(character: string): string {
  const script = unicode13Script(character.codePointAt(0)!)
  return script === 'Other' ? 'Zyyy' : script ?? ''
}

function shapingSpans(text: string, language: string, sourceOffset: number, paragraphOffset: number, levels: readonly number[]): SourceSpan[] {
  const spans: SourceSpan[] = []
  let current: SourceSpan | undefined
  for (let offset = 0; offset < text.length;) {
    const codePoint = text.codePointAt(offset)
    if (codePoint === undefined) break
    const character = String.fromCodePoint(codePoint)
    const next = offset + character.length
    let script = scriptForCharacter(character)
    const bidiLevel = levels[paragraphOffset + offset]
    if (!Number.isSafeInteger(bidiLevel) || bidiLevel! < 0 || bidiLevel! > 125) return []
    for (let unit = offset + 1; unit < next; unit++) if (levels[paragraphOffset + unit] !== bidiLevel) return []
    const direction = (bidiLevel! & 1) === 1 ? 'rtl' : 'ltr'
    if (script === 'Zinh' && current) script = current.script
    if (!current || current.script !== script || current.direction !== direction || current.bidiLevel !== bidiLevel) {
      current = { text: character, startUtf16: sourceOffset + offset, endUtf16: sourceOffset + next, script, direction, bidiLevel: bidiLevel!, language }
      spans.push(current)
    } else {
      current.text += character
      current.endUtf16 = sourceOffset + next
    }
    offset = next
  }
  return spans
}

function paragraphBidiPlan(context: NativeShapingContext, paragraph: NativeDocxParagraphV1, baseDirection: 'ltr' | 'rtl'): NativeDocxParagraphBidiPlanV1 | null {
  const referenceText = new Map<string, string>()
  for (const run of paragraph.runs) if (run.kind === 'reference' && run.reference && (run.reference.kind === 'footnote' || run.reference.kind === 'endnote')) {
    const marker = context.noteNumbers.get(run.reference.target_id)
    if (marker) referenceText.set(run.id, marker)
  }
  const result = resolveNativeDocxParagraphBidiPlanV1(paragraph, context.runs, baseDirection, referenceText)
  if (!result.ok) {
    addDiagnostic(context, { code: 'bidi-resolution-refusal', severity: 'unsupported', scope_id: paragraph.id, ...(result.sourceID ? { source_id: result.sourceID } : {}), message: `Pinned Unicode bidi resolution refused the paragraph: ${result.code}: ${result.message}` })
    return null
  }
  return result.value
}

function familyCandidates(family: string, aliases: Map<string, string>): string[] {
  const values = [family]
  const alias = aliases.get(normalizeFontFamilyName(family))
  if (alias && alias !== family) values.push(alias)
  return values
}

function textRunInput(span: SourceSpan, properties: NativeDocxResolvedRunPropertiesV1, aliases: Map<string, string>): TextRunInput | null {
  if (!properties.font_family || !properties.font_size_half_points) return null
  return deepFreezeWire({
    version: NATIVE_TEXT_LAYOUT_VERSION,
    text: span.text,
    fontSizeMilliPoints: halfPointsToMilliPoints(properties.font_size_half_points),
    font: {
      families: familyCandidates(properties.font_family, aliases),
      weight: properties.bold ? 700 : 400,
      style: properties.italic ? 'italic' : 'normal',
      stretch: 100,
    },
    script: span.script,
    language: span.language,
    direction: span.direction,
  })
}

function finiteSafeInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && !Object.is(value, -0)
}

function hasExactKeys(value: unknown, allowed: readonly string[]): value is Record<string, unknown> {
  if (!isRecord(value)) return false
  const keys = new Set(allowed)
  return Object.keys(value).every((key) => keys.has(key))
}

function validMetrics(metrics: ScaledLineMetrics): boolean {
  if (!hasExactKeys(metrics, ['fontSizeMilliPoints', 'ascentMilliPoints', 'descentMilliPoints', 'lineGapMilliPoints', 'lineHeightMilliPoints', 'capHeightMilliPoints', 'xHeightMilliPoints', 'underlinePositionMilliPoints', 'underlineThicknessMilliPoints'])) return false
  const values = [metrics.ascentMilliPoints, metrics.descentMilliPoints, metrics.lineGapMilliPoints, metrics.lineHeightMilliPoints]
  return finiteSafeInteger(metrics.fontSizeMilliPoints) && metrics.fontSizeMilliPoints > 0
    && finiteSafeInteger(metrics.ascentMilliPoints) && metrics.ascentMilliPoints >= 0
    && finiteSafeInteger(metrics.descentMilliPoints) && metrics.descentMilliPoints <= 0
    && finiteSafeInteger(metrics.lineGapMilliPoints) && metrics.lineGapMilliPoints >= 0
    && finiteSafeInteger(metrics.lineHeightMilliPoints) && metrics.lineHeightMilliPoints > 0
    && metrics.lineHeightMilliPoints === metrics.ascentMilliPoints - metrics.descentMilliPoints + metrics.lineGapMilliPoints
    && values.every((value) => Math.abs(value) <= MAX_PROVIDER_METRIC_MILLIPOINTS)
    && [metrics.capHeightMilliPoints, metrics.xHeightMilliPoints, metrics.underlinePositionMilliPoints, metrics.underlineThicknessMilliPoints].every((value) => value === undefined || (finiteSafeInteger(value) && Math.abs(value) <= MAX_PROVIDER_METRIC_MILLIPOINTS))
}

function validResolvedFaceShape(face: ResolvedFontFace): boolean {
  return hasExactKeys(face, ['faceId', 'family', 'postscriptName', 'weight', 'style', 'stretch', 'sourceKind', 'resourceId', 'contentDigest', 'collectionIndex', 'resolution', 'matchedFamily', 'fallbackChainId'])
    && typeof face.faceId === 'string' && PROVIDER_ID_RE.test(face.faceId)
    && typeof face.family === 'string' && face.family.length > 0 && face.family.length <= 1_024
    && (face.postscriptName === undefined || (typeof face.postscriptName === 'string' && face.postscriptName.length > 0 && face.postscriptName.length <= 1_024))
    && finiteSafeInteger(face.weight) && face.weight >= 1 && face.weight <= 1_000
    && ['normal', 'italic', 'oblique'].includes(face.style)
    && finiteSafeInteger(face.stretch) && face.stretch >= 50 && face.stretch <= 200
    && ['document', 'bundled', 'system', 'host'].includes(face.sourceKind)
    && typeof face.resourceId === 'string' && PROVIDER_ID_RE.test(face.resourceId)
    && typeof face.contentDigest === 'string' && /^sha256:[a-f0-9]{64}$/.test(face.contentDigest)
    && (face.collectionIndex === undefined || (finiteSafeInteger(face.collectionIndex) && face.collectionIndex >= 0 && face.collectionIndex <= 65_535))
    && ['exact', 'substitute', 'fallback'].includes(face.resolution)
    && typeof face.matchedFamily === 'string' && face.matchedFamily.length > 0 && face.matchedFamily.length <= 1_024
    && (face.fallbackChainId === undefined || (typeof face.fallbackChainId === 'string' && PROVIDER_ID_RE.test(face.fallbackChainId)))
}

function validProviderSegment(segment: ShapedSegment, run: TextRunInput, expectedFace: ResolvedFontFace): boolean {
  if (!hasExactKeys(segment, ['startUtf16', 'endUtf16', 'face', 'glyphs', 'clusters', 'metrics', 'advanceInlineMilliPoints', 'advanceBlockMilliPoints'])) return false
  if (segment.face.faceId !== expectedFace.faceId || segment.face.contentDigest !== expectedFace.contentDigest || segment.face.collectionIndex !== expectedFace.collectionIndex) return false
  if (segment.startUtf16 !== 0 || segment.endUtf16 !== run.text.length || !finiteSafeInteger(segment.advanceInlineMilliPoints) || segment.advanceInlineMilliPoints < 0 || segment.advanceInlineMilliPoints > Number.MAX_SAFE_INTEGER || !finiteSafeInteger(segment.advanceBlockMilliPoints) || Math.abs(segment.advanceBlockMilliPoints) > MAX_PROVIDER_METRIC_MILLIPOINTS || !validMetrics(segment.metrics) || segment.metrics.fontSizeMilliPoints !== run.fontSizeMilliPoints) return false
  if (!Array.isArray(segment.clusters) || !Array.isArray(segment.glyphs) || segment.clusters.length > MAX_TEXT_RUN_UTF16 || segment.glyphs.length > MAX_TEXT_RUN_UTF16) return false
  let previousText = 0
  let previousGlyph = 0
  let advance = 0
  for (let index = 0; index < segment.clusters.length; index++) {
    const cluster = segment.clusters[index]!
    if (!hasExactKeys(cluster, ['startUtf16', 'endUtf16', 'glyphStart', 'glyphEnd', 'advanceInlineMilliPoints', 'unsafeToBreak', 'whitespace'])) return false
    if (cluster.startUtf16 !== previousText || !finiteSafeInteger(cluster.endUtf16) || cluster.endUtf16 <= cluster.startUtf16 || cluster.endUtf16 > run.text.length || cluster.glyphStart !== previousGlyph || !finiteSafeInteger(cluster.glyphEnd) || cluster.glyphEnd < cluster.glyphStart || cluster.glyphEnd > segment.glyphs.length || !finiteSafeInteger(cluster.advanceInlineMilliPoints) || cluster.advanceInlineMilliPoints < 0 || cluster.advanceInlineMilliPoints > MAX_PROVIDER_METRIC_MILLIPOINTS || (cluster.unsafeToBreak !== undefined && typeof cluster.unsafeToBreak !== 'boolean') || (cluster.whitespace !== undefined && typeof cluster.whitespace !== 'boolean')) return false
    if (cluster.startUtf16 > 0) {
      const before = run.text.charCodeAt(cluster.startUtf16 - 1)
      const after = run.text.charCodeAt(cluster.startUtf16)
      if (before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff) return false
    }
    for (let glyphIndex = cluster.glyphStart; glyphIndex < cluster.glyphEnd; glyphIndex++) {
      const glyph = segment.glyphs[glyphIndex]!
      if (!hasExactKeys(glyph, ['glyphId', 'clusterIndex', 'advanceXMilliPoints', 'advanceYMilliPoints', 'offsetXMilliPoints', 'offsetYMilliPoints'])) return false
      if (!finiteSafeInteger(glyph.glyphId) || glyph.glyphId < 0 || glyph.glyphId > 0xffffffff || glyph.clusterIndex !== index || !finiteSafeInteger(glyph.advanceXMilliPoints) || !finiteSafeInteger(glyph.advanceYMilliPoints) || !finiteSafeInteger(glyph.offsetXMilliPoints) || !finiteSafeInteger(glyph.offsetYMilliPoints) || [glyph.advanceXMilliPoints, glyph.advanceYMilliPoints, glyph.offsetXMilliPoints, glyph.offsetYMilliPoints].some((value) => Math.abs(value) > MAX_PROVIDER_METRIC_MILLIPOINTS)) return false
    }
    previousText = cluster.endUtf16
    previousGlyph = cluster.glyphEnd
    advance += cluster.advanceInlineMilliPoints
    if (!Number.isSafeInteger(advance)) return false
  }
  return previousText === run.text.length && previousGlyph === segment.glyphs.length && advance === segment.advanceInlineMilliPoints
}

function sameResolvedFace(left: ResolvedFontFace, right: ResolvedFontFace): boolean {
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

function validResolvedFace(face: ResolvedFontFace, manifest: NativeFontManifest, run: TextRunInput): boolean {
  if (!validResolvedFaceShape(face)) return false
  const declared = manifest.faces.find((entry) => entry.faceId === face.faceId)
  if (!declared) return false
  const digestValid = /^sha256:[a-f0-9]{64}$/.test(face.contentDigest)
  const matchedFamily = run.font.families.some((family) => normalizeFontFamilyName(family) === normalizeFontFamilyName(face.matchedFamily))
  const fallbackValid = face.fallbackChainId === undefined || manifest.fallbackChains.some((chain) => chain.chainId === face.fallbackChainId && chain.faceIds.includes(face.faceId))
  return digestValid
    && matchedFamily
    && fallbackValid
    && ['exact', 'substitute', 'fallback'].includes(face.resolution)
    && declared.family === face.family
    && declared.postscriptName === face.postscriptName
    && declared.weight === face.weight
    && declared.style === face.style
    && declared.stretch === face.stretch
    && declared.source.kind === face.sourceKind
    && declared.source.resourceId === face.resourceId
    && declared.source.collectionIndex === face.collectionIndex
    && (declared.source.contentDigest === undefined || declared.source.contentDigest === face.contentDigest)
}

function validFontResource(resource: FontResource, expectedFace: ResolvedFontFace): boolean {
  if (!hasExactKeys(resource, ['face', 'bytes', 'metrics']) || !hasExactKeys(resource.metrics, ['unitsPerEm', 'ascender', 'descender', 'lineGap', 'capHeight', 'xHeight', 'underlinePosition', 'underlineThickness'])) return false
  const metrics = resource.metrics
  if (!(resource.bytes instanceof Uint8Array) || resource.bytes.byteLength === 0 || resource.bytes.byteLength > DOCX_SHAPED_LINES_LIMITS.maxFontResourceBytes) return false
  const actualDigest = `sha256:${bytesToHex(sha256(resource.bytes))}`
  return sameResolvedFace(resource.face, expectedFace)
    && actualDigest === expectedFace.contentDigest
    && validDesignMetrics(metrics)
}

function validDesignMetrics(metrics: FontResource['metrics']): boolean {
  return finiteSafeInteger(metrics.unitsPerEm) && metrics.unitsPerEm > 0 && metrics.unitsPerEm <= 1_000_000
    && finiteSafeInteger(metrics.ascender) && metrics.ascender >= 0 && metrics.ascender <= MAX_PROVIDER_METRIC_MILLIPOINTS
    && finiteSafeInteger(metrics.descender) && metrics.descender <= 0 && metrics.descender >= -MAX_PROVIDER_METRIC_MILLIPOINTS
    && finiteSafeInteger(metrics.lineGap) && metrics.lineGap >= 0 && metrics.lineGap <= MAX_PROVIDER_METRIC_MILLIPOINTS
    && [metrics.capHeight, metrics.xHeight].every((value) => value === undefined || (finiteSafeInteger(value) && value >= 0 && value <= MAX_PROVIDER_METRIC_MILLIPOINTS))
    && (metrics.underlinePosition === undefined || (finiteSafeInteger(metrics.underlinePosition) && Math.abs(metrics.underlinePosition) <= MAX_PROVIDER_METRIC_MILLIPOINTS))
    && (metrics.underlineThickness === undefined || (finiteSafeInteger(metrics.underlineThickness) && metrics.underlineThickness >= 0 && metrics.underlineThickness <= MAX_PROVIDER_METRIC_MILLIPOINTS))
}

function providerDecisionMessage(decisions: readonly NativeTextDecision[]): string {
  let message = ''
  for (const decision of decisions) {
    const next = `${message === '' ? '' : '; '}${decision.code}: ${decision.message}`
    if (message.length + next.length > MAX_PROVIDER_MESSAGE_LENGTH) return `${(message + next).slice(0, MAX_PROVIDER_MESSAGE_LENGTH - 1)}…`
    message += next
  }
  return message || 'provider refused native shaping'
}

function validProviderDecision(entry: unknown, textLength: number): entry is NativeTextDecision {
  if (!hasExactKeys(entry, ['code', 'message', 'recoverable', 'faceId', 'startUtf16', 'endUtf16'])) return false
  if (typeof entry.code !== 'string' || !PROVIDER_DECISION_CODES.has(entry.code) || typeof entry.message !== 'string' || entry.message.length === 0 || entry.message.length > MAX_PROVIDER_MESSAGE_LENGTH || typeof entry.recoverable !== 'boolean') return false
  if (entry.faceId !== undefined && (typeof entry.faceId !== 'string' || !PROVIDER_ID_RE.test(entry.faceId))) return false
  if ((entry.startUtf16 === undefined) !== (entry.endUtf16 === undefined)) return false
  return entry.startUtf16 === undefined || (finiteSafeInteger(entry.startUtf16) && finiteSafeInteger(entry.endUtf16) && entry.startUtf16 >= 0 && entry.endUtf16 > entry.startUtf16 && entry.endUtf16 <= textLength)
}

function validProviderMetadata(value: Record<string, unknown>, textLength: number): value is Record<string, unknown> & { decisions: NativeTextDecision[]; attemptedFaceIds: string[] } {
  if (!Array.isArray(value.decisions) || value.decisions.length > MAX_PROVIDER_DECISIONS || !Array.isArray(value.attemptedFaceIds) || value.attemptedFaceIds.length > MAX_PROVIDER_ATTEMPTED_FACES) return false
  const attempted = value.attemptedFaceIds
  if (!attempted.every((entry) => typeof entry === 'string' && PROVIDER_ID_RE.test(entry)) || new Set(attempted).size !== attempted.length) return false
  return value.decisions.every((entry) => validProviderDecision(entry, textLength))
}

function isNativeTextRefusal(value: unknown, textLength: number): value is NativeTextRefusal {
  return hasExactKeys(value, ['status', 'decisions', 'attemptedFaceIds']) && value.status === 'refused' && validProviderMetadata(value, textLength)
}

function validResolutionSuccess(value: unknown, textLength: number): value is { status: 'resolved'; face: ResolvedFontFace; decisions: NativeTextDecision[]; attemptedFaceIds: string[] } {
  return hasExactKeys(value, ['status', 'face', 'decisions', 'attemptedFaceIds']) && value.status === 'resolved' && validProviderMetadata(value, textLength) && isRecord(value.face)
}

function consumeProviderMetadata(context: NativeShapingContext, value: { decisions: readonly NativeTextDecision[]; attemptedFaceIds: readonly string[] }): boolean {
  const decisionCount = context.providerDecisionCount + value.decisions.length
  const attemptedCount = context.providerAttemptedFaceCount + value.attemptedFaceIds.length
  if (decisionCount > MAX_PROVIDER_DECISIONS_TOTAL || attemptedCount > MAX_PROVIDER_ATTEMPTED_FACES_TOTAL) return false
  context.providerDecisionCount = decisionCount
  context.providerAttemptedFaceCount = attemptedCount
  return true
}

function boundedProviderError(error: unknown): string {
  let raw = 'unknown provider error'
  try {
    if (typeof error === 'string') raw = error
    else if (error !== null && (typeof error === 'object' || typeof error === 'function')) {
      const message = Reflect.get(error, 'message')
      if (typeof message === 'string') raw = message
    }
  } catch {
    raw = 'unreadable provider error'
  }
  return raw.length <= MAX_PROVIDER_ERROR_LENGTH ? raw : `${raw.slice(0, MAX_PROVIDER_ERROR_LENGTH - 1)}…`
}

function resolvedFaceCacheKey(face: ResolvedFontFace): string {
  return JSON.stringify([face.faceId, face.family, face.postscriptName ?? null, face.weight, face.style, face.stretch, face.sourceKind, face.resourceId, face.contentDigest, face.collectionIndex ?? null, face.resolution, face.matchedFamily, face.fallbackChainId ?? null])
}

function snapshotFontResource(resource: FontResource): FontResource {
  const face = Object.freeze({ ...resource.face })
  const metrics = Object.freeze({ ...resource.metrics })
  return Object.freeze({ face, metrics, bytes: resource.bytes.slice() })
}

function captureExactProviderRecord(value: unknown, allowed: readonly string[], required: readonly string[]): Record<string, unknown> | null {
  try {
    if (!isRecord(value)) return null
    const captured: Record<string, unknown> = {}
    let count = 0
    for (const key in value) {
      if (!Object.hasOwn(value, key) || !allowed.includes(key) || count >= allowed.length) return null
      captured[key] = value[key]
      count += 1
    }
    for (const key of required) if (!Object.hasOwn(value, key)) return null
    for (const key of allowed) if (Object.hasOwn(value, key) && !Object.hasOwn(captured, key)) captured[key] = value[key]
    return captured
  } catch {
    return null
  }
}

function snapshotResolvedFace(value: unknown): ResolvedFontFace | null {
  const captured = captureExactProviderRecord(value, ['faceId', 'family', 'postscriptName', 'weight', 'style', 'stretch', 'sourceKind', 'resourceId', 'contentDigest', 'collectionIndex', 'resolution', 'matchedFamily', 'fallbackChainId'], ['faceId', 'family', 'weight', 'style', 'stretch', 'sourceKind', 'resourceId', 'contentDigest', 'resolution', 'matchedFamily'])
  if (!captured) return null
  const face = Object.freeze({
    faceId: captured.faceId,
    family: captured.family,
    ...(captured.postscriptName !== undefined ? { postscriptName: captured.postscriptName } : {}),
    weight: captured.weight,
    style: captured.style,
    stretch: captured.stretch,
    sourceKind: captured.sourceKind,
    resourceId: captured.resourceId,
    contentDigest: captured.contentDigest,
    ...(captured.collectionIndex !== undefined ? { collectionIndex: captured.collectionIndex } : {}),
    resolution: captured.resolution,
    matchedFamily: captured.matchedFamily,
    ...(captured.fallbackChainId !== undefined ? { fallbackChainId: captured.fallbackChainId } : {}),
  }) as ResolvedFontFace
  return validResolvedFaceShape(face) ? face : null
}

function snapshotProviderDecision(value: unknown, textLength: number): NativeTextDecision | null {
  const captured = captureExactProviderRecord(value, ['code', 'message', 'recoverable', 'faceId', 'startUtf16', 'endUtf16'], ['code', 'message', 'recoverable'])
  if (!captured) return null
  const decision = Object.freeze({
    code: captured.code,
    message: captured.message,
    recoverable: captured.recoverable,
    ...(captured.faceId !== undefined ? { faceId: captured.faceId } : {}),
    ...(captured.startUtf16 !== undefined ? { startUtf16: captured.startUtf16 } : {}),
    ...(captured.endUtf16 !== undefined ? { endUtf16: captured.endUtf16 } : {}),
  }) as NativeTextDecision
  return validProviderDecision(decision, textLength) ? decision : null
}

function snapshotProviderMetadata(decisionsValue: unknown, attemptedValue: unknown, textLength: number): { decisions: readonly NativeTextDecision[]; attemptedFaceIds: readonly string[] } | null {
  if (!Array.isArray(decisionsValue) || decisionsValue.length > MAX_PROVIDER_DECISIONS || !Array.isArray(attemptedValue) || attemptedValue.length > MAX_PROVIDER_ATTEMPTED_FACES) return null
  const decisions: NativeTextDecision[] = []
  for (let index = 0; index < decisionsValue.length; index++) {
    const decision = snapshotProviderDecision(decisionsValue[index], textLength)
    if (!decision) return null
    decisions.push(decision)
  }
  const attemptedFaceIds: string[] = []
  const attemptedFaceIDSet = new Set<string>()
  for (let index = 0; index < attemptedValue.length; index++) {
    const id = attemptedValue[index]
    if (typeof id !== 'string' || !PROVIDER_ID_RE.test(id) || attemptedFaceIDSet.has(id)) return null
    attemptedFaceIds.push(id)
    attemptedFaceIDSet.add(id)
  }
  return { decisions: Object.freeze(decisions), attemptedFaceIds: Object.freeze(attemptedFaceIds) }
}

type NativeResolutionSnapshot = NativeTextRefusal | { status: 'resolved'; face: ResolvedFontFace; decisions: readonly NativeTextDecision[]; attemptedFaceIds: readonly string[] }

function snapshotResolutionResult(value: unknown, textLength: number): NativeResolutionSnapshot | null {
  const captured = captureExactProviderRecord(value, ['status', 'face', 'decisions', 'attemptedFaceIds'], ['status', 'decisions', 'attemptedFaceIds'])
  if (!captured) return null
  const metadata = snapshotProviderMetadata(captured.decisions, captured.attemptedFaceIds, textLength)
  if (!metadata) return null
  if (captured.status === 'refused') {
    if (Object.hasOwn(captured, 'face')) return null
    return Object.freeze({ status: 'refused', ...metadata })
  }
  if (captured.status !== 'resolved' || !Object.hasOwn(captured, 'face')) return null
  const face = snapshotResolvedFace(captured.face)
  return face ? Object.freeze({ status: 'resolved', face, ...metadata }) : null
}

function snapshotDesignMetrics(value: unknown): FontResource['metrics'] | null {
  const captured = captureExactProviderRecord(value, ['unitsPerEm', 'ascender', 'descender', 'lineGap', 'capHeight', 'xHeight', 'underlinePosition', 'underlineThickness'], ['unitsPerEm', 'ascender', 'descender', 'lineGap'])
  if (!captured) return null
  const metrics = Object.freeze({
    unitsPerEm: captured.unitsPerEm,
    ascender: captured.ascender,
    descender: captured.descender,
    lineGap: captured.lineGap,
    ...(captured.capHeight !== undefined ? { capHeight: captured.capHeight } : {}),
    ...(captured.xHeight !== undefined ? { xHeight: captured.xHeight } : {}),
    ...(captured.underlinePosition !== undefined ? { underlinePosition: captured.underlinePosition } : {}),
    ...(captured.underlineThickness !== undefined ? { underlineThickness: captured.underlineThickness } : {}),
  }) as FontResource['metrics']
  return validDesignMetrics(metrics) ? metrics : null
}

type NativeLoadSnapshot = NativeTextRefusal | FontResource

function snapshotLoadResult(value: unknown, textLength: number): NativeLoadSnapshot | null {
  const captured = captureExactProviderRecord(value, ['status', 'decisions', 'attemptedFaceIds', 'face', 'bytes', 'metrics'], [])
  if (!captured) return null
  if (captured.status === 'refused') {
    if (Object.keys(captured).some((key) => key !== 'status' && key !== 'decisions' && key !== 'attemptedFaceIds')) return null
    const metadata = snapshotProviderMetadata(captured.decisions, captured.attemptedFaceIds, textLength)
    return metadata ? Object.freeze({ status: 'refused', ...metadata }) : null
  }
  if (Object.hasOwn(captured, 'status') || !Object.hasOwn(captured, 'face') || !Object.hasOwn(captured, 'bytes') || !Object.hasOwn(captured, 'metrics') || Object.keys(captured).length !== 3) return null
  const face = snapshotResolvedFace(captured.face)
  const metrics = snapshotDesignMetrics(captured.metrics)
  const bytes = captured.bytes
  if (!face || !metrics || !(bytes instanceof Uint8Array) || bytes.byteLength === 0 || bytes.byteLength > DOCX_SHAPED_LINES_LIMITS.maxFontResourceBytes) return null
  try {
    return Object.freeze({ face, metrics, bytes: Uint8Array.prototype.slice.call(bytes) as Uint8Array })
  } catch {
    return null
  }
}

function snapshotScaledMetrics(value: unknown): ScaledLineMetrics | null {
  const captured = captureExactProviderRecord(value, ['fontSizeMilliPoints', 'ascentMilliPoints', 'descentMilliPoints', 'lineGapMilliPoints', 'lineHeightMilliPoints', 'capHeightMilliPoints', 'xHeightMilliPoints', 'underlinePositionMilliPoints', 'underlineThicknessMilliPoints'], ['fontSizeMilliPoints', 'ascentMilliPoints', 'descentMilliPoints', 'lineGapMilliPoints', 'lineHeightMilliPoints'])
  if (!captured) return null
  const metrics = Object.freeze({
    fontSizeMilliPoints: captured.fontSizeMilliPoints,
    ascentMilliPoints: captured.ascentMilliPoints,
    descentMilliPoints: captured.descentMilliPoints,
    lineGapMilliPoints: captured.lineGapMilliPoints,
    lineHeightMilliPoints: captured.lineHeightMilliPoints,
    ...(captured.capHeightMilliPoints !== undefined ? { capHeightMilliPoints: captured.capHeightMilliPoints } : {}),
    ...(captured.xHeightMilliPoints !== undefined ? { xHeightMilliPoints: captured.xHeightMilliPoints } : {}),
    ...(captured.underlinePositionMilliPoints !== undefined ? { underlinePositionMilliPoints: captured.underlinePositionMilliPoints } : {}),
    ...(captured.underlineThicknessMilliPoints !== undefined ? { underlineThicknessMilliPoints: captured.underlineThicknessMilliPoints } : {}),
  }) as ScaledLineMetrics
  return validMetrics(metrics) ? metrics : null
}

function snapshotGlyph(value: unknown): ShapedGlyph | null {
  const captured = captureExactProviderRecord(value, ['glyphId', 'clusterIndex', 'advanceXMilliPoints', 'advanceYMilliPoints', 'offsetXMilliPoints', 'offsetYMilliPoints'], ['glyphId', 'clusterIndex', 'advanceXMilliPoints', 'advanceYMilliPoints', 'offsetXMilliPoints', 'offsetYMilliPoints'])
  if (!captured) return null
  return Object.freeze({ glyphId: captured.glyphId, clusterIndex: captured.clusterIndex, advanceXMilliPoints: captured.advanceXMilliPoints, advanceYMilliPoints: captured.advanceYMilliPoints, offsetXMilliPoints: captured.offsetXMilliPoints, offsetYMilliPoints: captured.offsetYMilliPoints }) as ShapedGlyph
}

function snapshotCluster(value: unknown): ShapedCluster | null {
  const captured = captureExactProviderRecord(value, ['startUtf16', 'endUtf16', 'glyphStart', 'glyphEnd', 'advanceInlineMilliPoints', 'unsafeToBreak', 'whitespace'], ['startUtf16', 'endUtf16', 'glyphStart', 'glyphEnd', 'advanceInlineMilliPoints'])
  if (!captured) return null
  return Object.freeze({ startUtf16: captured.startUtf16, endUtf16: captured.endUtf16, glyphStart: captured.glyphStart, glyphEnd: captured.glyphEnd, advanceInlineMilliPoints: captured.advanceInlineMilliPoints, ...(captured.unsafeToBreak !== undefined ? { unsafeToBreak: captured.unsafeToBreak } : {}), ...(captured.whitespace !== undefined ? { whitespace: captured.whitespace } : {}) }) as ShapedCluster
}

type NativeShaperSnapshot = NativeTextRefusal | ShapedSegment

function snapshotShaperResult(value: unknown, textLength: number): NativeShaperSnapshot | null {
  const captured = captureExactProviderRecord(value, ['status', 'decisions', 'attemptedFaceIds', 'startUtf16', 'endUtf16', 'face', 'glyphs', 'clusters', 'metrics', 'advanceInlineMilliPoints', 'advanceBlockMilliPoints'], [])
  if (!captured) return null
  if (captured.status === 'refused') {
    if (Object.keys(captured).some((key) => key !== 'status' && key !== 'decisions' && key !== 'attemptedFaceIds')) return null
    const metadata = snapshotProviderMetadata(captured.decisions, captured.attemptedFaceIds, textLength)
    return metadata ? Object.freeze({ status: 'refused', ...metadata }) : null
  }
  if (Object.hasOwn(captured, 'status') || Object.keys(captured).length !== 8 || !Array.isArray(captured.glyphs) || !Array.isArray(captured.clusters) || captured.glyphs.length > MAX_TEXT_RUN_UTF16 || captured.clusters.length > MAX_TEXT_RUN_UTF16) return null
  const face = snapshotResolvedFace(captured.face)
  const metrics = snapshotScaledMetrics(captured.metrics)
  if (!face || !metrics) return null
  const glyphs: ShapedGlyph[] = []
  for (let index = 0; index < captured.glyphs.length; index++) {
    const glyph = snapshotGlyph(captured.glyphs[index])
    if (!glyph) return null
    glyphs.push(glyph)
  }
  const clusters: ShapedCluster[] = []
  for (let index = 0; index < captured.clusters.length; index++) {
    const cluster = snapshotCluster(captured.clusters[index])
    if (!cluster) return null
    clusters.push(cluster)
  }
  return Object.freeze({
    startUtf16: captured.startUtf16,
    endUtf16: captured.endUtf16,
    face,
    glyphs: Object.freeze(glyphs),
    clusters: Object.freeze(clusters),
    metrics,
    advanceInlineMilliPoints: captured.advanceInlineMilliPoints,
    advanceBlockMilliPoints: captured.advanceBlockMilliPoints,
  }) as ShapedSegment
}

function providerFontResource(context: NativeShapingContext, cacheKey: string, authoritative: FontResource, sourceID: string): FontResource | null {
  const cached = context.providerFontResources.get(cacheKey)
  if (cached) return cached
  if (context.providerFacingFontBytes + authoritative.bytes.byteLength > DOCX_SHAPED_LINES_LIMITS.maxUniqueFontBytes) return exhaustProviderBudget(context, sourceID, `Provider-facing copied font bytes exceed ${DOCX_SHAPED_LINES_LIMITS.maxUniqueFontBytes}`)
  const copy = snapshotFontResource(authoritative)
  context.providerFacingFontBytes += copy.bytes.byteLength
  context.providerFontResources.set(cacheKey, copy)
  return copy
}

function exhaustProviderBudget(context: NativeShapingContext, sourceID: string, message: string): null {
  context.resourceExceeded = true
  addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message })
  return null
}

function providerIdentityRemainsStable(context: NativeShapingContext, sourceID: string): boolean {
  if (providerIdentityStable(context.providers)) return true
  addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected provider changed its snapshotted ID or revision during native shaping' })
  return false
}

async function resolveFontResource(context: NativeShapingContext, run: TextRunInput, sourceID: string): Promise<FontResource | null> {
  if (context.providerResolveCalls >= DOCX_SHAPED_LINES_LIMITS.maxProviderResolveCalls) return exhaustProviderBudget(context, sourceID, `Resolver calls exceed ${DOCX_SHAPED_LINES_LIMITS.maxProviderResolveCalls}`)
  context.providerResolveCalls += 1
  const liveResolution: unknown = await context.providers.resolver.resolve(Object.freeze({ manifest: context.fontManifest, run }))
  if (!providerIdentityRemainsStable(context, sourceID)) return null
  const resolution = snapshotResolutionResult(liveResolution, run.text.length)
  if (!resolution) {
    addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected resolver returned malformed, recursive, or unbounded provider metadata' })
    return null
  }
  if (isNativeTextRefusal(resolution, run.text.length)) {
    if (!consumeProviderMetadata(context, resolution)) return exhaustProviderBudget(context, sourceID, 'Native text provider metadata exceeded the document-wide decision/attempted-face budget')
    addDiagnostic(context, { code: 'provider-refusal', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: providerDecisionMessage(resolution.decisions) })
    return null
  }
  if (!validResolutionSuccess(resolution, run.text.length)) {
    addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected resolver returned malformed or unbounded status, decision, attempted-face, or face metadata' })
    return null
  }
  if (!consumeProviderMetadata(context, resolution)) return exhaustProviderBudget(context, sourceID, 'Native text provider metadata exceeded the document-wide decision/attempted-face budget')
  if (!validResolvedFace(resolution.face, context.fontManifest, run)) {
    addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected resolver returned a face that is invalid, not manifest-backed, or not matched to the authored family request' })
    return null
  }
  for (const decision of resolution.decisions) addDiagnostic(context, { code: 'provider-decision', severity: 'deferred', scope_id: sourceID, source_id: sourceID, message: `${decision.code}: ${decision.message}` })
  const cacheKey = resolvedFaceCacheKey(resolution.face)
  const cached = context.fontResources.get(cacheKey)
  if (cached) {
    if (!sameResolvedFace(cached.face, resolution.face)) {
      addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Resolved-face cache identity collision changed exact face metadata' })
      return null
    }
    return cached
  }
  if (context.providerLoadCalls >= DOCX_SHAPED_LINES_LIMITS.maxProviderLoadCalls || context.fontResources.size >= DOCX_SHAPED_LINES_LIMITS.maxCachedFontResources) return exhaustProviderBudget(context, sourceID, 'Unique font loads exceed the bounded native shaping cache')
  context.providerLoadCalls += 1
  const liveResource: unknown = await context.providers.resolver.load(resolution.face)
  if (!providerIdentityRemainsStable(context, sourceID)) return null
  const loadResult = snapshotLoadResult(liveResource, run.text.length)
  if (!loadResult) {
    addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: `Injected resolver returned an empty, malformed, recursive, or larger than ${DOCX_SHAPED_LINES_LIMITS.maxFontResourceBytes}-byte font resource` })
    return null
  }
  if (isNativeTextRefusal(loadResult, run.text.length)) {
    if (!consumeProviderMetadata(context, loadResult)) return exhaustProviderBudget(context, sourceID, 'Native text provider metadata exceeded the document-wide decision/attempted-face budget')
    addDiagnostic(context, { code: 'provider-refusal', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: providerDecisionMessage(loadResult.decisions) })
    return null
  }
  const resource = loadResult
  if (context.uniqueFontBytes + resource.bytes.byteLength > DOCX_SHAPED_LINES_LIMITS.maxUniqueFontBytes) return exhaustProviderBudget(context, sourceID, `Unique loaded font bytes exceed ${DOCX_SHAPED_LINES_LIMITS.maxUniqueFontBytes}`)
  if (!validFontResource(resource, resolution.face)) {
    addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected resolver returned mismatched, digest-invalid, or malformed font resources' })
    return null
  }
  context.uniqueFontBytes += resource.bytes.byteLength
  context.fontResources.set(cacheKey, resource)
  return resource
}

async function shapeSpan(context: NativeShapingContext, span: SourceSpan, properties: NativeDocxResolvedRunPropertiesV1, sourceKind: 'run' | 'list-marker', sourceID: string): Promise<FragmentAtom[]> {
  if (context.shapedCodeUnits + span.text.length > DOCX_SHAPED_LINES_LIMITS.maxFragments) {
    context.resourceExceeded = true
    addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: `Pre-shaping UTF-16/cluster budget exceeds ${DOCX_SHAPED_LINES_LIMITS.maxFragments}; the paragraph was refused before another provider allocation` })
    return []
  }
  context.shapedCodeUnits += span.text.length
  if (!properties.font_family) {
    addDiagnostic(context, { code: 'missing-run-font', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Resolved font family is absent; shaping was refused instead of guessing a platform font' })
    return []
  }
  if (!properties.font_size_half_points) {
    addDiagnostic(context, { code: 'missing-run-size', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Resolved font size is absent; shaping was refused instead of guessing a Word default' })
    return []
  }
  const run = textRunInput(span, properties, context.fontAliases)
  if (!run) return []
  const validation = validateTextRunInput(run)
  if (!validation.ok) {
    addDiagnostic(context, { code: 'unresolved-layout-diagnostic', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, source_diagnostic_code: 'INVALID_SHAPING_INPUT', source_diagnostic_message: validation.issues.map((issue) => `${issue.path} ${issue.message}`).join('; '), message: `Resolved run cannot be represented by the shared shaping contract: ${validation.issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}` })
    return []
  }
  try {
    const resource = await resolveFontResource(context, run, sourceID)
    if (!resource) return []
    const cacheKey = resolvedFaceCacheKey(resource.face)
    const shaperResource = providerFontResource(context, cacheKey, resource, sourceID)
    if (!shaperResource) return []
    if (context.providerShapeCalls >= DOCX_SHAPED_LINES_LIMITS.maxProviderShapeCalls) {
      exhaustProviderBudget(context, sourceID, `Shaper calls exceed ${DOCX_SHAPED_LINES_LIMITS.maxProviderShapeCalls}`)
      return []
    }
    context.providerShapeCalls += 1
    const liveShaped: unknown = await context.providers.shaper.shape(Object.freeze({ run, startUtf16: 0, endUtf16: run.text.length, font: shaperResource }))
    if (!providerIdentityRemainsStable(context, sourceID)) return []
    const shaped = snapshotShaperResult(liveShaped, run.text.length)
    if (!shaped) {
      addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected shaper returned malformed, recursive, or unbounded provider output' })
      return []
    }
    if (isNativeTextRefusal(shaped, run.text.length)) {
      if (!consumeProviderMetadata(context, shaped)) {
        addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Native text provider metadata exceeded the document-wide decision/attempted-face budget' })
        return []
      }
      addDiagnostic(context, { code: 'provider-refusal', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: providerDecisionMessage(shaped.decisions) })
      return []
    }
    if (!isRecord(shaped) || !isRecord(shaped.face) || !sameResolvedFace(shaped.face as unknown as ResolvedFontFace, resource.face) || !validProviderSegment(shaped as unknown as ShapedSegment, run, resource.face)) {
      addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: 'Injected shaper returned non-contiguous, unbounded, or internally inconsistent clusters/glyphs' })
      return []
    }
    const shapedSegment = shaped as unknown as ShapedSegment
    if (context.atomCount + shapedSegment.clusters.length > DOCX_SHAPED_LINES_LIMITS.maxFragments) {
      context.resourceExceeded = true
      addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: `Shaped clusters exceed ${DOCX_SHAPED_LINES_LIMITS.maxFragments}` })
      return []
    }
    context.atomCount += shapedSegment.clusters.length
    const atoms = shapedSegment.clusters.map((cluster, index) => clusterAtom(shapedSegment, cluster, index, span, sourceKind, sourceID))
    const alignment = properties.vertical_alignment
    if (alignment && alignment !== 'baseline') {
      if (sourceKind !== 'run' || properties.underline && properties.underline !== 'none' || properties.highlight && properties.highlight !== 'none') throw new TypeError('Script transforms currently require undecorated text runs, not markers, underline or highlight')
      const transform = readNativeDocxScriptTransformV1(resource, alignment)
      if (nativeDocxScriptScaleV1(run.fontSizeMilliPoints,transform,'x') < 1 || nativeDocxScriptScaleV1(run.fontSizeMilliPoints,transform,'y') < 1) throw new TypeError('Script font size is below the native integer metric resolution')
      const dx = nativeDocxScriptShiftV1(run.fontSizeMilliPoints,transform,'x'), dy = nativeDocxScriptShiftV1(run.fontSizeMilliPoints,transform,'y')
      for (const atom of atoms) {
        atom.scriptTransform = transform
        atom.glyphs = atom.glyphs.map((glyph) => ({ ...glyph, advance_x_millipoints: nativeDocxScriptScaleV1(glyph.advance_x_millipoints,transform,'x'), advance_y_millipoints: nativeDocxScriptScaleV1(glyph.advance_y_millipoints,transform,'y'), offset_x_millipoints: nativeDocxScriptScaleV1(glyph.offset_x_millipoints,transform,'x')+dx, offset_y_millipoints: nativeDocxScriptScaleV1(glyph.offset_y_millipoints,transform,'y')+dy }))
        atom.advance = atom.glyphs.reduce((sum,glyph)=>sum+glyph.advance_x_millipoints,0)
        const ascent=Math.max(0,nativeDocxScriptScaleV1(atom.metrics.ascentMilliPoints,transform,'y')+dy), descent=Math.min(0,nativeDocxScriptScaleV1(atom.metrics.descentMilliPoints,transform,'y')+dy), gap=nativeDocxScriptScaleV1(atom.metrics.lineGapMilliPoints,transform,'y')
        atom.metrics = { ...atom.metrics, ascentMilliPoints:ascent, descentMilliPoints:descent, lineGapMilliPoints:gap, lineHeightMilliPoints:ascent-descent+gap }
      }
    }
    return atoms
  } catch (error) {
    addDiagnostic(context, { code: 'provider-failure', severity: 'unsupported', scope_id: sourceID, source_id: sourceID, message: `Injected native text provider failed: ${boundedProviderError(error)}` })
    return []
  }
}

async function resolveParagraphMarkMetrics(context: NativeShapingContext, paragraph: NativeDocxResolvedParagraphV1, direction: 'ltr' | 'rtl'): Promise<ScaledLineMetrics | null> {
  const properties = paragraph.paragraph_mark_properties
  if (properties.vertical_alignment && properties.vertical_alignment !== 'baseline') {
    addDiagnostic(context, { code: 'unresolved-layout-diagnostic', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Script-sized paragraph marks require a separate blank-line metric policy' })
    return null
  }
  if (!properties.font_family) {
    addDiagnostic(context, { code: 'missing-run-font', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Resolved paragraph-mark font is absent; blank-line metrics were refused instead of guessing a Word default' })
    return null
  }
  if (!properties.font_size_half_points) {
    addDiagnostic(context, { code: 'missing-run-size', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Resolved paragraph-mark font size is absent; blank-line metrics were refused instead of guessing a Word default' })
    return null
  }
  const markDirection = properties.rtl ? 'rtl' : direction
  const span: SourceSpan = { text: '', startUtf16: 0, endUtf16: 0, script: 'Zyyy', direction: markDirection, bidiLevel: markDirection === 'rtl' ? 1 : 0, language: properties.language ?? 'und' }
  const run = textRunInput(span, properties, context.fontAliases)
  if (!run) return null
  const validation = validateTextRunInput(run)
  if (!validation.ok) {
    addDiagnostic(context, { code: 'unresolved-layout-diagnostic', severity: 'unsupported', scope_id: paragraph.paragraph_id, source_diagnostic_code: 'INVALID_PARAGRAPH_MARK_INPUT', source_diagnostic_message: validation.issues.map((issue) => `${issue.path} ${issue.message}`).join('; '), message: 'Resolved paragraph-mark properties cannot be represented by the shared font contract' })
    return null
  }
  try {
    const resource = await resolveFontResource(context, run, paragraph.paragraph_id)
    if (!resource) return null
    const metrics = scaleLineMetrics(resource.metrics, run.fontSizeMilliPoints)
    if (!validMetrics(metrics) || metrics.fontSizeMilliPoints !== run.fontSizeMilliPoints) {
      addDiagnostic(context, { code: 'invalid-provider-output', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Paragraph-mark font metrics cannot produce bounded blank-line metrics' })
      return null
    }
    return metrics
  } catch (error) {
    addDiagnostic(context, { code: 'provider-failure', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: `Injected native font resolver failed for paragraph-mark metrics: ${boundedProviderError(error)}` })
    return null
  }
}

function clusterAtom(segment: ShapedSegment, cluster: ShapedCluster, clusterIndex: number, span: SourceSpan, sourceKind: 'run' | 'list-marker', sourceID: string): FragmentAtom {
  const text = span.text.slice(cluster.startUtf16, cluster.endUtf16)
  const glyphs: NativeDocxPositionedGlyphV1[] = segment.glyphs.slice(cluster.glyphStart, cluster.glyphEnd).map((glyph: ShapedGlyph) => ({
    glyph_id: glyph.glyphId,
    advance_x_millipoints: glyph.advanceXMilliPoints,
    advance_y_millipoints: glyph.advanceYMilliPoints,
    offset_x_millipoints: glyph.offsetXMilliPoints,
    offset_y_millipoints: glyph.offsetYMilliPoints,
  }))
  const nextText = span.text.slice(cluster.endUtf16, segment.clusters[clusterIndex + 1]?.endUtf16 ?? cluster.endUtf16)
  return {
    sourceKind,
    sourceID,
    startUtf16: span.startUtf16 + cluster.startUtf16,
    endUtf16: span.startUtf16 + cluster.endUtf16,
    text,
    direction: span.direction,
    bidiLevel: span.bidiLevel,
    script: span.script,
    language: span.language,
    faceID: segment.face.faceId,
    whitespace: cluster.whitespace === true || isUnicode13TextWhiteSpace(text),
    unsafeToBreak: cluster.unsafeToBreak === true,
    breakAfter: unicodeBreakAfter(text, nextText),
    dynamicTab: false,
    advance: cluster.advanceInlineMilliPoints,
    metrics: segment.metrics,
    glyphs,
  }
}

function lastCharacter(value: string): string {
  const values = [...value]
  return values[values.length - 1] ?? ''
}

function firstCharacter(value: string): string {
  return [...value][0] ?? ''
}

function unicodeBreakAfter(currentText: string, nextText: string): boolean {
  if (isUnicode13BreakableWhiteSpace(currentText) || currentText.endsWith('\u200B')) return true
  const current = lastCharacter(currentText)
  const next = firstCharacter(nextText)
  if (current === '-' || current === '\u2010' || current === '\u2013' || current === '/') return true
  const currentScript = current === '' ? undefined : unicode13Script(current.codePointAt(0)!)
  const nextScript = next === '' ? undefined : unicode13Script(next.codePointAt(0)!)
  const currentCJK = currentScript === 'Hani' || currentScript === 'Kana' || currentScript === 'Hang'
  const nextCJK = nextScript === 'Hani' || nextScript === 'Kana' || nextScript === 'Hang'
  if (currentCJK && next !== '' && unicode13Punctuation(current.codePointAt(0)!) !== 'open' && unicode13Punctuation(next.codePointAt(0)!) !== 'close') return true
  if (nextCJK && current !== '' && unicode13Punctuation(current.codePointAt(0)!) !== 'open' && unicode13Punctuation(next.codePointAt(0)!) !== 'close') return true
  return false
}

function splitHardBreaks(text: string): Array<{ text?: string; start: number; end: number; hardBreak?: true }> {
  const result: Array<{ text?: string; start: number; end: number; hardBreak?: true }> = []
  let start = 0
  for (let index = 0; index < text.length;) {
    const code = text.charCodeAt(index)
    if (code !== 0x0a && code !== 0x0d) {
      index += code >= 0xd800 && code <= 0xdbff ? 2 : 1
      continue
    }
    if (index > start) result.push({ text: text.slice(start, index), start, end: index })
    const end = code === 0x0d && text.charCodeAt(index + 1) === 0x0a ? index + 2 : index + 1
    result.push({ start: index, end, hardBreak: true })
    start = end
    index = end
  }
  if (start < text.length) result.push({ text: text.slice(start), start, end: text.length })
  return result
}

function reportBlockingDiagnostics(context: NativeShapingContext, diagnosticScopeID: string, outputScopeID: string, sourceID?: string): boolean {
  const diagnostics = context.blockingDiagnostics.get(diagnosticScopeID) ?? []
  for (const diagnostic of diagnostics) addDiagnostic(context, {
    code: 'unresolved-layout-diagnostic',
    severity: 'unsupported',
    scope_id: outputScopeID,
    ...(sourceID ? { source_id: sourceID } : {}),
    source_diagnostic_code: diagnostic.code,
    source_diagnostic_message: diagnostic.message,
    message: `Resolved layout diagnostic ${diagnostic.code} blocks native shaping: ${diagnostic.message}`,
  })
  return diagnostics.length > 0
}

async function shapeAuthoredRun(context: NativeShapingContext, paragraphID: string, run: NativeDocxRunV1, resolved: NativeDocxResolvedRunV1, plan: NativeDocxParagraphBidiPlanV1): Promise<ParagraphEvent[]> {
  if (reportBlockingDiagnostics(context, run.id, paragraphID, run.id)) return []
  if (resolved.properties.hidden) return []
  if (run.kind === 'drawing') {
    const qualified = run.drawing ? qualifyNativeDocxInlineImageV1(context.request.document, run.id, run.drawing) : undefined
    if (!qualified?.ok) {
      addDiagnostic(context, { code: 'drawing-layout-unsupported', severity: 'unsupported', scope_id: paragraphID, source_id: run.id, message: qualified?.message ?? 'Drawing payload is missing' })
      return []
    }
    if (!reserveVirtualAtom(context, paragraphID, run.id)) return []
    const image = qualified.value
    const level = plan.paragraph.levels[plan.runStarts.get(run.id) ?? -1] ?? plan.paragraph.baseLevel
    const direction = (level & 1) === 1 ? 'rtl' : 'ltr'
    const metrics: ScaledLineMetrics = {
      fontSizeMilliPoints: image.height_millipoints,
      ascentMilliPoints: image.height_millipoints,
      descentMilliPoints: 0,
      lineGapMilliPoints: 0,
      lineHeightMilliPoints: image.height_millipoints,
    }
    return [{ kind: 'atom', atom: {
      sourceKind: 'image', sourceID: run.id, startUtf16: 0, endUtf16: 0, text: '',
      direction, bidiLevel: level, script: 'Zyyy', language: resolved.properties.language ?? 'und',
      whitespace: false, unsafeToBreak: false, breakAfter: false, dynamicTab: false,
      advance: image.floating ? 0 : image.width_millipoints, metrics: image.floating ? emptyMetrics() : metrics, glyphs: [],
    } }]
  }
  if (run.kind === 'reference') {
    const text = run.reference && (run.reference.kind === 'footnote' || run.reference.kind === 'endnote') ? context.noteNumbers.get(run.reference.target_id) : undefined
    if (!text) {
      addDiagnostic(context, { code: 'reference-layout-unsupported', severity: 'unsupported', scope_id: paragraphID, source_id: run.id, message: 'Reference display text is not in the exact uniquely ordered decimal note subset' })
      return []
    }
    const paragraphOffset = plan.runStarts.get(run.id)
    if (paragraphOffset === undefined) return []
    const events: ParagraphEvent[] = []
    for (const span of shapingSpans(text, resolved.properties.language ?? 'und', 0, paragraphOffset, plan.paragraph.levels)) {
      const atoms = await shapeSpan(context, span, resolved.properties, 'run', run.id)
      events.push(...atoms.map((atom) => ({ kind: 'atom', atom }) as const))
    }
    return events
  }
  if (run.kind === 'control') {
    if (run.control === 'line-break') return [{ kind: 'hard-break', runID: run.id }]
    const level = plan.paragraph.levels[plan.runStarts.get(run.id) ?? -1] ?? plan.paragraph.baseLevel
    const direction = (level & 1) === 1 ? 'rtl' : 'ltr'
    if (run.control === 'tab') return reserveVirtualAtom(context, paragraphID, run.id) ? [{ kind: 'atom', atom: tabAtom(run.id, 'run', direction, level, resolved.properties.language ?? 'und') }] : []
    if (run.control === 'soft-hyphen') {
      addDiagnostic(context, { code: 'soft-hyphen-deferred', severity: 'deferred', scope_id: paragraphID, source_id: run.id, message: 'Conditional soft-hyphen glyph insertion is deferred; the source control remains a zero-width break opportunity' })
      return reserveVirtualAtom(context, paragraphID, run.id) ? [{ kind: 'atom', atom: zeroWidthBreakAtom(run.id, direction, level, resolved.properties.language ?? 'und') }] : []
    }
    addDiagnostic(context, { code: 'page-control-deferred', severity: 'deferred', scope_id: paragraphID, source_id: run.id, message: `${run.control} is a pagination control and does not participate in line shaping` })
    return []
  }
  if (run.kind !== 'text' || run.text === undefined || run.text.length === 0) return []
  const events: ParagraphEvent[] = []
  for (const chunk of splitHardBreaks(run.text)) {
    if (context.activeParagraphFailed || context.resourceExceeded) break
    if (chunk.hardBreak) {
      events.push({ kind: 'hard-break', runID: run.id })
      continue
    }
    const paragraphOffset = (plan.runStarts.get(run.id) ?? 0) + chunk.start
    const spans = shapingSpans(chunk.text ?? '', resolved.properties.language ?? 'und', chunk.start, paragraphOffset, plan.paragraph.levels)
    if ((chunk.text?.length ?? 0) > 0 && spans.length === 0) {
      addDiagnostic(context, { code: 'bidi-resolution-refusal', severity: 'unsupported', scope_id: paragraphID, source_id: run.id, message: 'Resolved bidi levels split or failed to cover one Unicode scalar' })
      return []
    }
    for (const span of spans) {
      if (context.activeParagraphFailed || context.resourceExceeded) break
      const atoms = await shapeSpan(context, span, resolved.properties, 'run', run.id)
      events.push(...atoms.map((atom) => ({ kind: 'atom', atom }) as const))
    }
  }
  return events
}

function emptyMetrics(): ScaledLineMetrics {
  return { fontSizeMilliPoints: 1, ascentMilliPoints: 0, descentMilliPoints: 0, lineGapMilliPoints: 0, lineHeightMilliPoints: 0 }
}

function reserveVirtualAtom(context: NativeShapingContext, scopeID: string, sourceID: string): boolean {
  if (context.atomCount >= DOCX_SHAPED_LINES_LIMITS.maxFragments) {
    context.resourceExceeded = true
    addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: scopeID, source_id: sourceID, message: `Virtual shaping fragments exceed ${DOCX_SHAPED_LINES_LIMITS.maxFragments}` })
    return false
  }
  context.atomCount += 1
  return true
}

function tabAtom(sourceID: string, sourceKind: 'run' | 'list-marker', direction: 'ltr' | 'rtl', bidiLevel: number, language: string): FragmentAtom {
  return { sourceKind: sourceKind === 'run' ? 'tab' : 'list-marker', sourceID, startUtf16: 0, endUtf16: 0, text: '\t', direction, bidiLevel, script: 'Zyyy', language, whitespace: true, unsafeToBreak: false, breakAfter: true, dynamicTab: true, advance: 0, metrics: emptyMetrics(), glyphs: [] }
}

function zeroWidthBreakAtom(sourceID: string, direction: 'ltr' | 'rtl', bidiLevel: number, language: string): FragmentAtom {
  return { sourceKind: 'run', sourceID, startUtf16: 0, endUtf16: 0, text: '', direction, bidiLevel, script: 'Zyyy', language, whitespace: false, unsafeToBreak: false, breakAfter: true, dynamicTab: false, advance: 0, metrics: emptyMetrics(), glyphs: [] }
}

function fixedMarkerAtom(sourceID: string, text: string, advance: number, direction: 'ltr' | 'rtl', language: string): FragmentAtom {
  return { sourceKind: 'list-marker', sourceID, startUtf16: 0, endUtf16: 0, text, direction, bidiLevel: direction === 'rtl' ? 1 : 0, script: 'Zyyy', language, whitespace: text === '\t', unsafeToBreak: false, breakAfter: text === '\t', dynamicTab: false, advance, metrics: emptyMetrics(), glyphs: [] }
}

async function markerEvents(context: NativeShapingContext, paragraph: NativeDocxResolvedParagraphV1, direction: 'ltr' | 'rtl'): Promise<{ events: ParagraphEvent[]; marker?: NativeDocxShapedListMarkerV1; firstLineStart?: number }> {
  if (!paragraph.numbering) return { events: [] }
  const numbering = paragraph.numbering
  const source = context.request.resolved_layout.numbering_source
  if (!source) {
    addDiagnostic(context, { code: 'unsupported-numbering-text', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Resolved list marker is missing its exact numbering source attestation' })
    return { events: [] }
  }
  const text = numbering.resolved_text
  const properties = paragraph.numbering.marker_properties
  const language = properties.language ?? 'und'
  const explicitDirection = properties.rtl === undefined ? undefined : properties.rtl ? 'rtl' : 'ltr'
  const bidi = resolveNativeBidiParagraphV1({ text, baseDirection: direction, ...(explicitDirection && text.length > 0 ? { explicitRanges: [{ startUtf16: 0, endUtf16: text.length, direction: explicitDirection }] } : {}) })
  if (!bidi.ok) {
    addDiagnostic(context, { code: 'bidi-resolution-refusal', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: `Pinned Unicode bidi resolution refused the list marker: ${bidi.code}: ${bidi.message}` })
    return { events: [] }
  }
  const markerAtoms: FragmentAtom[] = []
  for (const span of shapingSpans(text, language, 0, 0, bidi.value.levels)) {
    if (context.activeParagraphFailed || context.resourceExceeded) break
    const atoms = await shapeSpan(context, span, properties, 'list-marker', paragraph.paragraph_id)
    markerAtoms.push(...atoms)
  }
  if (context.activeParagraphFailed || context.resourceExceeded) return { events: [] }
  const markerAdvance = markerAtoms.reduce((sum, atom) => sum + atom.advance, 0)
  const geometry = positionNativeDocxListMarkerV1(numbering, direction, markerAdvance)
  if (!geometry) {
    addDiagnostic(context, { code: 'list-marker-alignment-deferred', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Shaped marker does not fit the exact bounded hanging-indent label region' })
    return { events: [] }
  }
  const { label_start_millipoints: labelStart, label_end_millipoints: labelEnd, body_text_start_millipoints: bodyTextStart, marker_start_millipoints: markerStart } = geometry
  const atoms: FragmentAtom[] = []
  const firstLineStart = Math.min(labelStart, markerStart)
  if (markerStart > firstLineStart) {
    if (!reserveVirtualAtom(context, paragraph.paragraph_id, paragraph.paragraph_id)) return { events: [] }
    atoms.push(fixedMarkerAtom(paragraph.paragraph_id, '', markerStart - firstLineStart, direction, language))
  }
  atoms.push(...markerAtoms)
  let textStart = markerStart + markerAdvance
  if (paragraph.numbering.suffix === 'space') {
    const level = direction === 'rtl' ? 1 : 0
    const span: SourceSpan = { text: ' ', startUtf16: text.length, endUtf16: text.length + 1, script: 'Zyyy', direction, bidiLevel: level, language }
    const suffixAtoms = await shapeSpan(context, span, properties, 'list-marker', paragraph.paragraph_id)
    atoms.push(...suffixAtoms)
    textStart += suffixAtoms.reduce((sum, atom) => sum + atom.advance, 0)
  } else if (paragraph.numbering.suffix === 'tab') {
    const numberingTab = numbering.numbering_tab_twips === undefined ? undefined : twipsToMilliPoints(numbering.numbering_tab_twips)
    const target = nativeDocxListSuffixTabTargetV1(textStart, bodyTextStart, context.request.tab_interval_millipoints, numberingTab)
    if (target === undefined) {
      addDiagnostic(context, { code: 'list-marker-tab-deferred', severity: 'unsupported', scope_id: paragraph.paragraph_id, message: 'Explicit numbering tab must follow the complete shaped marker end; the numbered paragraph was refused atomically' })
      return { events: [] }
    }
    if (!reserveVirtualAtom(context, paragraph.paragraph_id, paragraph.paragraph_id)) return { events: [] }
    atoms.push(fixedMarkerAtom(paragraph.paragraph_id, '\t', target - textStart, direction, language))
    textStart = target
  }
  return {
    events: atoms.map((atom) => ({ kind: 'atom', atom }) as const),
    firstLineStart,
    marker: {
      marker_id: numbering.marker_id, definition_sha256: numbering.definition_sha256,
      numbering_part_sha256: source.part_sha256, model_sha256: source.model_sha256,
      num_id: numbering.num_id, abstract_num_id: numbering.abstract_num_id, level: numbering.level,
      counter_value: numbering.counter_value, text, suffix: numbering.suffix, alignment: numbering.alignment,
      label_start_millipoints: labelStart, label_end_millipoints: labelEnd,
      marker_start_millipoints: markerStart, marker_advance_millipoints: markerAdvance,
      text_start_millipoints: textStart,
    },
  }
}

function logicalIndents(properties: NativeDocxResolvedParagraphPropertiesV1, direction: 'ltr' | 'rtl'): { start: number; end: number; firstDelta: number } {
  const physicalStart = direction === 'rtl' ? properties.indent_right_twips : properties.indent_left_twips
  const physicalEnd = direction === 'rtl' ? properties.indent_left_twips : properties.indent_right_twips
  const start = twipsToMilliPoints(properties.indent_start_twips ?? physicalStart ?? 0)
  const end = twipsToMilliPoints(properties.indent_end_twips ?? physicalEnd ?? 0)
  const firstDelta = properties.first_line_twips !== undefined ? twipsToMilliPoints(properties.first_line_twips) : properties.hanging_twips !== undefined ? -twipsToMilliPoints(properties.hanging_twips) : 0
  return { start, end, firstDelta }
}

function atomAdvance(atom: FragmentAtom, currentAdvance: number, tabInterval: number): number {
  if (!atom.dynamicTab) return atom.advance
  const remainder = ((currentAdvance % tabInterval) + tabInterval) % tabInterval
  return remainder === 0 ? tabInterval : tabInterval - remainder
}

function safeBoundary(left: FragmentAtom, right: FragmentAtom | undefined): boolean {
  return !left.unsafeToBreak && (right === undefined || !right.unsafeToBreak) && !GLUE_RE.test(left.text) && (right === undefined || !GLUE_RE.test(right.text))
}

function chooseLineEnd(atoms: FragmentAtom[], start: number, width: number, tabInterval: number): number {
  let advance = 0
  let lastOpportunity = -1
  for (let index = start; index < atoms.length; index++) {
    const atom = atoms[index]!
    const nextAdvance = advance + atomAdvance(atom, advance, tabInterval)
    if (!Number.isSafeInteger(nextAdvance)) return Math.max(start + 1, index)
    if (nextAdvance > width) {
      if (lastOpportunity > start) return lastOpportunity
      if (index === start) return start + 1
      for (let boundary = index; boundary > start; boundary--) if (safeBoundary(atoms[boundary - 1]!, atoms[boundary])) return boundary
      for (let boundary = index + 1; boundary <= atoms.length; boundary++) if (safeBoundary(atoms[boundary - 1]!, atoms[boundary])) return boundary
      return atoms.length
    }
    advance = nextAdvance
    if ((atom.breakAfter || unicodeBreakAfter(atom.text, atoms[index + 1]?.text ?? '')) && safeBoundary(atom, atoms[index + 1])) lastOpportunity = index + 1
  }
  return atoms.length
}

function naturalLineMetrics(atoms: FragmentAtom[], paragraphMarkMetrics?: ScaledLineMetrics): { ascent: number; descent: number; gap: number; height: number } {
  let ascent = 0
  let descent = 0
  let gap = 0
  for (const atom of atoms) {
    ascent = Math.max(ascent, atom.metrics.ascentMilliPoints)
    descent = Math.min(descent, atom.metrics.descentMilliPoints)
    gap = Math.max(gap, atom.metrics.lineGapMilliPoints)
  }
  if (ascent === 0 && descent === 0 && gap === 0 && paragraphMarkMetrics) return { ascent: paragraphMarkMetrics.ascentMilliPoints, descent: paragraphMarkMetrics.descentMilliPoints, gap: paragraphMarkMetrics.lineGapMilliPoints, height: paragraphMarkMetrics.lineHeightMilliPoints }
  const height = ascent - descent + gap
  if (!Number.isSafeInteger(height) || height < 0 || height > MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS) throw new RangeError('combined line metrics exceed deterministic bounds')
  return { ascent, descent, gap, height }
}

function resolvedLineHeight(natural: number, properties: NativeDocxResolvedParagraphPropertiesV1): number | null {
  if (properties.line === undefined || properties.line_rule === undefined) return natural
  if (properties.line_rule === 'auto') {
    const scaled = Math.round((natural * properties.line) / 240)
    return Number.isSafeInteger(scaled) && scaled >= 0 && scaled <= MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS ? scaled : null
  }
  const specified = twipsToMilliPoints(properties.line)
  const resolved = properties.line_rule === 'exact' ? specified : Math.max(natural, specified)
  return Number.isSafeInteger(resolved) && resolved >= 0 && resolved <= MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS ? resolved : null
}

function alignmentOffset(alignment: NativeDocxShapedParagraphV1['alignment'], direction: 'ltr' | 'rtl', available: number, advance: number): number {
  const remaining = Math.max(0, available - advance)
  if (alignment === 'center') return Math.round(remaining / 2)
  if (alignment === 'left' || alignment === 'both' || alignment === 'distribute') return 0
  if (alignment === 'right') return remaining
  if (alignment === 'start') return direction === 'ltr' ? 0 : remaining
  return direction === 'ltr' ? remaining : 0
}

function materializeLine(context: NativeShapingContext, paragraphID: string, ordinal: number, atoms: FragmentAtom[], available: number, startOffset: number, alignment: NativeDocxShapedParagraphV1['alignment'], direction: 'ltr' | 'rtl', paragraphMarkMetrics?: ScaledLineMetrics, hardBreakRunID?: string, softWrapped = false): NativeDocxShapedLineV1 | null {
  let naturalAdvance = 0
  if (context.fragmentCount + atoms.length > DOCX_SHAPED_LINES_LIMITS.maxFragments) {
    context.resourceExceeded = true
    addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: paragraphID, message: `Complete line would exceed ${DOCX_SHAPED_LINES_LIMITS.maxFragments} fragments and was refused without partial emission` })
    return null
  }
  const atomWidths = atoms.map((atom) => {
    const width = atomAdvance(atom, naturalAdvance, context.request.tab_interval_millipoints)
    naturalAdvance += width
    return width
  })
  const order = reorderNativeBidiLineV1(atoms.map((atom) => atom.bidiLevel), direction === 'rtl' ? 1 : 0, atoms.map((atom) => BIDI_TRAILING_RE.test(atom.text)))
  if (!order.ok) {
    addDiagnostic(context, { code: 'bidi-resolution-refusal', severity: 'unsupported', scope_id: paragraphID, message: `Pinned Unicode bidi line ordering refused line ${ordinal}: ${order.code}: ${order.message}` })
    return null
  }
  const expansions = atoms.map(() => 0)
  const shouldJustify = alignment === 'both' && softWrapped && naturalAdvance <= available
  const remaining = Math.max(0, available - naturalAdvance)
  if (shouldJustify && remaining > 0) {
    const firstContent = atoms.findIndex((atom) => !atom.whitespace)
    let lastContent = -1
    for (let index = atoms.length - 1; index >= 0; index--) if (!atoms[index]!.whitespace) { lastContent = index; break }
    const opportunities = order.value.visualToLogical.filter((logical) => logical > firstContent && logical < lastContent && atoms[logical]!.text === ' ' && !atoms[logical]!.dynamicTab)
    if (opportunities.length === 0) {
      addDiagnostic(context, { code: 'justification-unsupported', severity: 'unsupported', scope_id: paragraphID, message: `Soft-wrapped justified line ${ordinal} has no bounded U+0020 expansion opportunity` })
      return null
    }
    if (opportunities.some((logical) => atoms[logical]!.glyphs.length === 0)) {
      addDiagnostic(context, { code: 'justification-unsupported', severity: 'unsupported', scope_id: paragraphID, message: `Soft-wrapped justified line ${ordinal} has an unpaintable U+0020 expansion opportunity` })
      return null
    }
    const quotient = Math.floor(remaining / opportunities.length)
    let remainder = remaining % opportunities.length
    for (const logical of opportunities) expansions[logical] = quotient + (remainder-- > 0 ? 1 : 0)
  }
  const fragments = order.value.visualToLogical.map((logicalIndex, visualIndex): NativeDocxLineFragmentV1 => {
    const atom = atoms[logicalIndex]!
    const expansion = expansions[logicalIndex]!
    const glyphs = atom.glyphs.map((glyph) => ({ ...glyph }))
    if (expansion > 0) {
      glyphs[glyphs.length - 1]!.advance_x_millipoints += expansion
    }
    context.fragmentCount += 1
    const underline = atom.sourceKind === 'list-marker'
      ? context.paragraphs.get(paragraphID)?.numbering?.marker_properties.underline
      : context.runs.get(atom.sourceID)?.properties.underline
    return {
      id: `fragment:${paragraphID}:${ordinal}:${visualIndex}`,
      source_kind: atom.sourceKind,
      source_id: atom.sourceID,
      start_utf16: atom.startUtf16,
      end_utf16: atom.endUtf16,
      text: atom.text,
      direction: atom.direction,
      bidi_level: atom.bidiLevel,
      logical_order: logicalIndex,
      script: atom.script,
      language: atom.language,
      ...(atom.faceID ? { face_id: atom.faceID } : {}),
      ...(atom.scriptTransform ? { script_transform: atom.scriptTransform } : {}),
      whitespace: atom.whitespace,
      advance_inline_millipoints: atomWidths[logicalIndex]! + expansion,
      justification_expansion_millipoints: expansion,
      ascent_millipoints: atom.metrics.ascentMilliPoints,
      descent_millipoints: atom.metrics.descentMilliPoints,
      line_gap_millipoints: atom.metrics.lineGapMilliPoints,
      ...(underline && underline !== 'none' && atom.metrics.underlinePositionMilliPoints !== undefined && atom.metrics.underlineThicknessMilliPoints !== undefined
        ? { underline_position_millipoints: atom.metrics.underlinePositionMilliPoints, underline_thickness_millipoints: atom.metrics.underlineThicknessMilliPoints } : {}),
      glyphs,
    }
  })
  const advance = naturalAdvance + expansions.reduce((sum, value) => sum + value, 0)
  const metrics = naturalLineMetrics(atoms, paragraphMarkMetrics)
  context.lineCount += 1
  return {
    id: `line:${paragraphID}:${ordinal}`,
    ordinal,
    available_width_millipoints: available,
    inline_offset_millipoints: startOffset + alignmentOffset(alignment === 'both' && !shouldJustify ? 'start' : alignment, direction, available, advance),
    advance_inline_millipoints: advance,
    ascent_millipoints: metrics.ascent,
    descent_millipoints: metrics.descent,
    line_gap_millipoints: metrics.gap,
    line_height_millipoints: metrics.height,
    justified: shouldJustify,
    logical_to_visual: [...order.value.logicalToVisual],
    fragments,
    ...(hardBreakRunID ? { hard_break_after: { source_run_id: hardBreakRunID, control: 'line-break' as const } } : {}),
  }
}

function wrapEventGroup(context: NativeShapingContext, paragraphID: string, atoms: FragmentAtom[], lines: NativeDocxShapedLineV1[], baseWidth: number, start: number, end: number, firstDelta: number, alignment: NativeDocxShapedParagraphV1['alignment'], direction: 'ltr' | 'rtl', paragraphMarkMetrics?: ScaledLineMetrics, hardBreakRunID?: string, firstLineStart?: number): void {
  if (context.lineCount >= DOCX_SHAPED_LINES_LIMITS.maxLines || context.resourceExceeded) {
    context.resourceExceeded = true
    addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: paragraphID, message: `Shaped lines exceed ${DOCX_SHAPED_LINES_LIMITS.maxLines}` })
    return
  }
  let offset = 0
  if (atoms.length === 0) {
    const first = lines.length === 0
    const ordinaryStart = direction === 'ltr' ? start + (first ? firstDelta : 0) : end
    const startOffset = first && firstLineStart !== undefined ? firstLineStart : ordinaryStart
    const width = first && firstLineStart !== undefined ? baseWidth - startOffset - end : baseWidth - start - end - (first ? firstDelta : 0)
    const line = materializeLine(context, paragraphID, lines.length, [], Math.max(0, width), startOffset, alignment, direction, paragraphMarkMetrics, hardBreakRunID)
    if (line) lines.push(line)
    return
  }
  while (offset < atoms.length && !context.resourceExceeded) {
    if (context.lineCount >= DOCX_SHAPED_LINES_LIMITS.maxLines) {
      context.resourceExceeded = true
      addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: paragraphID, message: `Shaped lines exceed ${DOCX_SHAPED_LINES_LIMITS.maxLines}` })
      return
    }
    const first = lines.length === 0
    const ordinaryStart = direction === 'ltr' ? start + (first ? firstDelta : 0) : end
    const startOffset = first && firstLineStart !== undefined ? firstLineStart : ordinaryStart
    const width = first && firstLineStart !== undefined ? baseWidth - startOffset - end : baseWidth - start - end - (first ? firstDelta : 0)
    if (width <= 0) {
      addDiagnostic(context, { code: 'cluster-overflow', severity: 'unsupported', scope_id: paragraphID, message: 'Paragraph indents leave no positive inline width' })
      return
    }
    const lineEnd = chooseLineEnd(atoms, offset, width, context.request.tab_interval_millipoints)
    const lineAtoms = atoms.slice(offset, lineEnd)
    const line = materializeLine(context, paragraphID, lines.length, lineAtoms, width, startOffset, first && firstLineStart !== undefined ? 'left' : alignment, direction, paragraphMarkMetrics, lineEnd === atoms.length ? hardBreakRunID : undefined, lineEnd < atoms.length)
    if (!line) return
    if (line.advance_inline_millipoints > width) addDiagnostic(context, { code: 'cluster-overflow', severity: 'deferred', scope_id: paragraphID, source_id: lineAtoms[0]?.sourceID, message: 'One cluster exceeds the offered width and remains intact on an overfull line' })
    lines.push(line)
    offset = lineEnd
  }
}

function finalizeLineHeights(context: NativeShapingContext, paragraphID: string, lines: NativeDocxShapedLineV1[], properties: NativeDocxResolvedParagraphPropertiesV1): void {
  for (const line of lines) {
    const height = resolvedLineHeight(line.line_height_millipoints, properties)
    if (height === null) {
      addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: paragraphID, message: `Resolved line height exceeds ${MAX_PARAGRAPH_MEASUREMENT_MILLIPOINTS} milli-points` })
      return
    }
    line.line_height_millipoints = height
    if (line.line_height_millipoints === 0) addDiagnostic(context, { code: 'empty-line-metrics-unresolved', severity: 'unsupported', scope_id: paragraphID, message: 'An empty line has no safely resolved font metrics' })
  }
}

function needsParagraphMarkMetrics(events: ParagraphEvent[]): boolean {
  if (events.length === 0) return true
  let groupHasMetrics = false
  for (const event of events) {
    if (event.kind === 'atom') {
      if (event.atom.metrics.lineHeightMilliPoints > 0) groupHasMetrics = true
      else return true
      continue
    }
    if (!groupHasMetrics) return true
    groupHasMetrics = false
  }
  return !groupHasMetrics
}

async function shapeParagraph(context: NativeShapingContext, story: NativeDocxStoryV1, paragraph: NativeDocxParagraphV1): Promise<NativeDocxShapedParagraphV1 | null> {
  const resolved = context.paragraphs.get(paragraph.id)!
  const direction = resolved.properties.bidi ? 'rtl' : 'ltr'
  const alignment = resolved.properties.alignment ?? 'start'
  const indents = logicalIndents(resolved.properties, direction)
  const lineCountSnapshot = context.lineCount
  const fragmentCountSnapshot = context.fragmentCount
  context.activeParagraphID = paragraph.id
  context.activeParagraphFailed = false
  let sourceBlocked = reportBlockingDiagnostics(context, paragraph.id, paragraph.id)
  const diagnosticScopes = [paragraph.id, ...paragraph.runs.map((run) => run.id)]
  for (const run of paragraph.runs) sourceBlocked = reportBlockingDiagnostics(context, run.id, paragraph.id, run.id) || sourceBlocked
  const numberingAffected = resolved.numbering !== undefined || hasConcreteNumberingFailure(context, diagnosticScopes)
  const refuse = (): null => {
    if (numberingAffected) context.numberingFailed = true
    context.lineCount = lineCountSnapshot
    context.fragmentCount = fragmentCountSnapshot
    context.activeParagraphID = undefined
    context.activeParagraphFailed = false
    return null
  }
  if (context.documentBlocked || sourceBlocked) return refuse()
  const bidiPlan = paragraphBidiPlan(context, paragraph, direction)
  if (!bidiPlan || context.activeParagraphFailed) return refuse()
  if (resolved.numbering && alignment !== 'start' && (direction === 'ltr' ? alignment !== 'left' : alignment !== 'right')) {
    addDiagnostic(context, { code: 'list-marker-alignment-deferred', severity: 'unsupported', scope_id: paragraph.id, message: 'Exact list-marker anchoring with non-leading paragraph alignment is unsupported' })
    return refuse()
  }
  const marker = await markerEvents(context, resolved, direction)
  const events = marker.events
  if (context.activeParagraphFailed || context.resourceExceeded) return refuse()
  for (const key of ['keep_next', 'keep_lines', 'page_break_before', 'widow_control'] as const) {
    if (resolved.properties[key] !== undefined) addDiagnostic(context, { code: 'page-control-deferred', severity: 'deferred', scope_id: paragraph.id, message: `${key} is retained in resolved layout for the future paginator and does not alter line shaping` })
  }
  for (const run of paragraph.runs) {
    if (context.activeParagraphFailed || context.resourceExceeded) break
    const resolvedRun = context.runs.get(run.id)!
    const runEvents = await shapeAuthoredRun(context, paragraph.id, run, resolvedRun, bidiPlan)
    events.push(...runEvents)
  }
  if (context.activeParagraphFailed || context.resourceExceeded) return refuse()
  const needsMarkMetrics = needsParagraphMarkMetrics(events)
  const paragraphMarkMetrics = needsMarkMetrics ? (await resolveParagraphMarkMetrics(context, resolved, direction) ?? undefined) : undefined
  if (context.activeParagraphFailed || context.resourceExceeded || (needsMarkMetrics && !paragraphMarkMetrics)) return refuse()
  if (alignment === 'distribute') {
    addDiagnostic(context, { code: 'justification-unsupported', severity: 'unsupported', scope_id: paragraph.id, message: 'Distributed character expansion is outside the deterministic U+0020 justification slice' })
    return refuse()
  }
  const lines: NativeDocxShapedLineV1[] = []
  let group: FragmentAtom[] = []
  for (const event of events) {
    if (event.kind === 'atom') {
      group.push(event.atom)
      continue
    }
    wrapEventGroup(context, paragraph.id, group, lines, context.activeAvailableWidthMilliPoints ?? context.request.available_width_millipoints, indents.start, indents.end, indents.firstDelta, alignment, direction, paragraphMarkMetrics, event.runID, marker.firstLineStart)
    group = []
  }
  wrapEventGroup(context, paragraph.id, group, lines, context.activeAvailableWidthMilliPoints ?? context.request.available_width_millipoints, indents.start, indents.end, indents.firstDelta, alignment, direction, paragraphMarkMetrics, undefined, marker.firstLineStart)
  finalizeLineHeights(context, paragraph.id, lines, resolved.properties)
  if (context.activeParagraphFailed || context.resourceExceeded) return refuse()
  const before = twipsToMilliPoints(resolved.properties.spacing_before_twips ?? 0)
  const after = twipsToMilliPoints(resolved.properties.spacing_after_twips ?? 0)
  const blockAdvance = before + after + lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)
  if (!Number.isSafeInteger(blockAdvance) || blockAdvance < 0 || blockAdvance > MAX_PARAGRAPH_BLOCK_MILLIPOINTS) {
    addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: paragraph.id, message: `Paragraph block advance exceeds ${MAX_PARAGRAPH_BLOCK_MILLIPOINTS} milli-points` })
    return refuse()
  }
  const output: NativeDocxShapedParagraphV1 = {
    paragraph_id: paragraph.id,
    story_id: story.id,
    story_kind: story.kind,
    direction,
    alignment,
    spacing_before_millipoints: before,
    spacing_after_millipoints: after,
    indent_start_millipoints: indents.start,
    indent_end_millipoints: indents.end,
    first_line_delta_millipoints: indents.firstDelta,
    ...(marker.marker ? { list_marker: marker.marker } : {}),
    block_advance_millipoints: blockAdvance,
    lines,
  }
  context.activeParagraphID = undefined
  context.activeParagraphFailed = false
  return output
}

function blockingDiagnostics(resolved: NativeDocxResolvedLayoutInputV1, paragraphIDs: Set<string>, runIDs: Set<string>, tableIDs: Set<string>): Map<string, NativeDocxResolvedLayoutInputV1['diagnostics']> {
  const result = new Map<string, NativeDocxResolvedLayoutInputV1['diagnostics']>()
  for (const diagnostic of resolved.diagnostics) {
    if ((PAINT_ONLY_RESOLUTION_DIAGNOSTICS.has(diagnostic.code) && (paragraphIDs.has(diagnostic.scope_id) || runIDs.has(diagnostic.scope_id))) || (TABLE_ONLY_RESOLUTION_DIAGNOSTICS.has(diagnostic.code) && tableIDs.has(diagnostic.scope_id))) continue
    const diagnostics = result.get(diagnostic.scope_id) ?? []
    diagnostics.push(diagnostic)
    result.set(diagnostic.scope_id, diagnostics)
  }
  return result
}

function hasConcreteNumberingFailure(context: NativeShapingContext, scopeIDs: readonly string[]): boolean {
  const source = context.request.resolved_layout.numbering_source
  if (!source) return false
  for (const scopeID of scopeIDs) for (const diagnostic of context.blockingDiagnostics.get(scopeID) ?? []) {
    if (diagnostic.part_name === source.part_name || /NUMBER|PICTURE_BULLET|CUSTOM_NUMBER_FORMAT/.test(diagnostic.code)) return true
  }
  return false
}

function propagateAllowedResolutionDiagnostics(context: NativeShapingContext, paragraphIDs: Set<string>, runIDs: Set<string>): void {
  for (const diagnostic of context.request.resolved_layout.diagnostics) {
    if (!PAINT_ONLY_RESOLUTION_DIAGNOSTICS.has(diagnostic.code) || (!paragraphIDs.has(diagnostic.scope_id) && !runIDs.has(diagnostic.scope_id))) continue
    addDiagnostic(context, {
      code: 'paint-diagnostic-preserved',
      severity: 'deferred',
      scope_id: diagnostic.scope_id,
      source_diagnostic_code: diagnostic.code,
      source_diagnostic_message: diagnostic.message,
      message: `Paint-only resolved-layout diagnostic does not change shaping advances: ${diagnostic.code}: ${diagnostic.message}`,
    })
  }
}

function resolvedFontAliases(resolved: NativeDocxResolvedLayoutInputV1): Map<string, string> {
  const aliases = new Map<string, string>()
  for (const font of resolved.fonts) if (font.alt_name) aliases.set(normalizeFontFamilyName(font.name), font.alt_name)
  return aliases
}

function tableDiagnosticScopes(table: NativeDocxTableV1): Set<string> {
  const scopes = new Set<string>([table.id])
  for (const row of table.rows) for (const cell of row.cells) for (const paragraph of cell.paragraphs) {
    scopes.add(paragraph.id)
    for (const run of paragraph.runs) scopes.add(run.id)
  }
  return scopes
}

/**
 * Validate both Go wire projections, verify every durable-id join, and shape
 * source-ordered paragraphs with injected native font providers.
 */
async function shapeNativeDocxLinesCoreV1(value: unknown, providers: NativeDocxShapingProviders, paragraphWidths?: ReadonlyMap<string, number>): Promise<ShapeNativeDocxLinesResult> {
  if (!UNICODE_13_TABLES_RUNTIME_MATCH) return { ok: false, issues: [validationIssue('INVALID_VALUE', '', 'pinned Unicode 13 classification tables failed their runtime digest')] }
  const initial = validateRequest(value)
  if (!initial.ok) return initial
  let wireSnapshot: unknown
  try {
    wireSnapshot = deepFreezeWire(structuredClone(initial.value))
  } catch {
    return { ok: false, issues: [validationIssue('INVALID_VALUE', '', 'shaping request must be a cloneable JSON wire value')] }
  }
  const decoded = validateRequest(wireSnapshot)
  if (!decoded.ok) return decoded
  const providerBoundary = snapshotProviderBoundary(providers)
  if (!providerBoundary.ok) return providerBoundary
  const request = decoded.value
  const fontManifest = snapshotFontManifest(request.font_manifest)
  const inventory = inventoryDocument(request.document)
  const qualifiedNotes = qualifyNoteNumbers(request.document)
  const context: NativeShapingContext = {
    request,
    providers: providerBoundary.value,
    fontManifest,
    paragraphs: new Map(request.resolved_layout.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph])),
    runs: new Map(request.resolved_layout.runs.map((run) => [run.run_id, run])),
    noteNumbers: qualifiedNotes.numbers,
    fontAliases: resolvedFontAliases(request.resolved_layout),
    blockingDiagnostics: blockingDiagnostics(request.resolved_layout, new Set(inventory.paragraphs.keys()), new Set(inventory.runs.keys()), new Set(inventory.tables.keys())),
    diagnostics: [],
    diagnosticKeys: new Set(),
    lineCount: 0,
    fragmentCount: 0,
    atomCount: 0,
    shapedCodeUnits: 0,
    resourceExceeded: false,
    diagnosticOverflow: false,
    documentBlocked: false,
    providerDecisionCount: 0,
    providerAttemptedFaceCount: 0,
    providerResolveCalls: 0,
    providerLoadCalls: 0,
    providerShapeCalls: 0,
    uniqueFontBytes: 0,
    providerFacingFontBytes: 0,
    fontResources: new Map(),
    providerFontResources: new Map(),
    activeParagraphFailed: false,
    numberingFailed: false,
  }
  for (const noteIssue of qualifiedNotes.issues) addDiagnostic(context, { code: 'reference-layout-unsupported', severity: 'unsupported', scope_id: noteIssue.scope, ...(noteIssue.source ? { source_id: noteIssue.source } : {}), message: noteIssue.message })
  if (qualifiedNotes.issues.length > 0) context.documentBlocked = true
  context.documentBlocked = reportBlockingDiagnostics(context, request.document.document_id, request.document.document_id) || context.documentBlocked
  propagateAllowedResolutionDiagnostics(context, new Set(inventory.paragraphs.keys()), new Set(inventory.runs.keys()))
  const paragraphs: NativeDocxShapedParagraphV1[] = []
  storyLoop:
  for (const story of nativeStories(request.document)) {
    if (story.note_role === 'continuation-separator') continue
    for (const block of story.blocks) {
      if (block.kind === 'table' && block.table) {
        if (paragraphWidths) {
          for (const cell of block.table.rows.flatMap((row) => row.cells)) for (const sourceParagraph of cell.paragraphs) {
            const width = paragraphWidths.get(sourceParagraph.id)
            if (!width) {
              addDiagnostic(context, { code: 'table-layout-unsupported', severity: 'unsupported', scope_id: block.table.id, message: 'Qualified table paragraph is missing its exact cell-content shaping width' })
              continue
            }
            if (context.resourceExceeded || paragraphs.length >= DOCX_SHAPED_LINES_LIMITS.maxParagraphs) break storyLoop
            context.activeAvailableWidthMilliPoints = width
            const paragraph = await shapeParagraph(context, story, sourceParagraph)
            context.activeAvailableWidthMilliPoints = undefined
            if (paragraph) paragraphs.push(paragraph)
          }
          continue
        }
        addDiagnostic(context, { code: 'table-layout-unsupported', severity: 'unsupported', scope_id: block.table.id, message: 'Table grid/width layout is preserved and refused by the paragraph-only line core' })
        const diagnosticScopes = tableDiagnosticScopes(block.table)
        for (const diagnostic of request.resolved_layout.diagnostics.filter((entry) => diagnosticScopes.has(entry.scope_id))) addDiagnostic(context, {
          code: 'table-layout-unsupported',
          severity: 'unsupported',
          scope_id: block.table.id,
          source_diagnostic_code: diagnostic.code,
          source_diagnostic_message: diagnostic.message,
          message: `Skipped table retains resolved-layout diagnostic ${diagnostic.code}: ${diagnostic.message}`,
        })
        const numberedTable = block.table.rows.some((row) => row.cells.some((cell) => cell.paragraphs.some((paragraph) => context.paragraphs.get(paragraph.id)?.numbering !== undefined)))
        if (numberedTable || hasConcreteNumberingFailure(context, [...diagnosticScopes])) context.numberingFailed = true
        continue
      }
      if (block.kind !== 'paragraph' || !block.paragraph) continue
      if (context.resourceExceeded || paragraphs.length >= DOCX_SHAPED_LINES_LIMITS.maxParagraphs) {
        addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: request.document.document_id, message: `Shaped paragraphs exceed ${DOCX_SHAPED_LINES_LIMITS.maxParagraphs}` })
        break storyLoop
      }
      context.activeAvailableWidthMilliPoints = paragraphWidths?.get(block.paragraph.id)
      const paragraph = await shapeParagraph(context, story, block.paragraph)
      context.activeAvailableWidthMilliPoints = undefined
      if (context.resourceExceeded) break storyLoop
      if (paragraph) paragraphs.push(paragraph)
      if (context.lineCount > DOCX_SHAPED_LINES_LIMITS.maxLines || context.fragmentCount > DOCX_SHAPED_LINES_LIMITS.maxFragments) {
        addDiagnostic(context, { code: 'resource-limit', severity: 'unsupported', scope_id: request.document.document_id, message: 'Shaped line/fragment output exceeded bounded resource limits' })
        break storyLoop
      }
    }
  }
  return {
    ok: true,
    value: {
      protocol: DOCX_SHAPED_LINES_PROTOCOL,
      version: DOCX_SHAPED_LINES_VERSION,
      document_id: request.document.document_id,
      revision: request.document.revision,
      available_width_millipoints: request.available_width_millipoints,
      tab_interval_millipoints: request.tab_interval_millipoints,
      font_manifest: { manifest_id: fontManifest.manifestId, revision: fontManifest.revision },
      providers: {
        resolver_id: context.providers.resolver.providerId,
        resolver_revision: context.providers.resolver.providerRevision,
        shaper_id: context.providers.shaper.providerId,
        shaper_revision: context.providers.shaper.providerRevision,
        bidi_id: NATIVE_BIDI_PROVIDER_ID,
        bidi_revision: NATIVE_BIDI_PROVIDER_REVISION,
        bidi_unicode_version: BIDI_UNICODE_VERSION,
        unicode13_revision: UNICODE_13_CLASSIFIER_REVISION,
      },
      ...(request.resolved_layout.numbering_source ? { numbering_source: request.resolved_layout.numbering_source } : {}),
      paragraphs: context.numberingFailed ? [] : paragraphs,
      diagnostics: context.diagnostics,
    },
  }
}

export async function shapeNativeDocxLinesV1(value: unknown, providers: NativeDocxShapingProviders): Promise<ShapeNativeDocxLinesResult> {
  return shapeNativeDocxLinesCoreV1(value, providers)
}

/** Internal canonical page-paint seam: reuses one provider/cache/budget context while shaping qualified cell widths. */
export async function shapeNativeDocxLinesWithParagraphWidthsV1(value: unknown, providers: NativeDocxShapingProviders, paragraphWidths: ReadonlyMap<string, number>): Promise<ShapeNativeDocxLinesResult> {
  return shapeNativeDocxLinesCoreV1(value, providers, paragraphWidths)
}
