/**
 * Renderer-neutral DOCX pagination, version 1.
 *
 * This module places already-shaped body lines in integer milli-points. It has
 * no DOM, HTML, CSS, canvas, browser, painting, or package-mutation dependency.
 * Semantic uncertainty refuses the whole page projection; a refused result
 * never exposes a plausible-looking partial page list.
 */

import { decodeNativeDocxApproximationEligibilityV1 } from './nativeApproximationV1.js'
import {
  DOCX_NATIVE_LIMITS,
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxHeaderFooterReferenceV1,
  type NativeDocxBlockV1,
  type NativeDocxParagraphV1,
  type NativeDocxSectionV1,
  type NativeDocxValidationIssue,
} from './nativeContract.js'
import {
  decodeNativeDocxResolvedLayout,
  type NativeDocxResolvedLayoutInputV1,
  type NativeDocxResolvedNumberingSourceV1,
  type NativeDocxResolvedParagraphV1,
} from './nativeResolvedLayout.js'
import {
  DOCX_SHAPED_LINES_PROTOCOL,
  DOCX_SHAPED_LINES_VERSION,
  type NativeDocxShapedLineV1,
  type NativeDocxShapedLinesV1,
  type NativeDocxShapedParagraphV1,
} from './nativeShapingLines.js'
import { reorderNativeBidiLineV1 } from '@injoffice/font-metrics/bidi'
import { decodeNativeDocxShapedLines } from './nativeShapedLinesContract.js'
import { resolveNativeDocxParagraphBidiPlanV1 } from './nativeBidiPlanV1.js'
import {
  DOCX_PAGINATION_SETTINGS_PROTOCOL,
  DOCX_PAGINATION_SETTINGS_VERSION,
  decodeNativeDocxPaginationSettings,
  type NativeDocxPaginationSettingsV1,
} from './nativePaginationSettings.js'
import { asciiLowerNative, compareNativeCodeUnits } from './nativeDeterminism.js'
import { layoutNativeDocxTableRowsV1, nativeDocxTableRowGroupSizeV1, qualifyNativeDocxTablesV1, type NativeDocxQualifiedTableV1, type NativeDocxTableRowGeometryV1 } from './nativeTablePagePaintV1.js'
import { qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { nativeDocxSectionsShareExactPageV1, qualifyNativeDocxSectionColumnsV1 } from './nativeSectionColumnsV1.js'
import { placeNativeDocxNotesV1 } from './nativeNotePaginationV1.js'
import { nativeDocxListSuffixTabTargetV1, positionNativeDocxListMarkerV1 } from './nativeNumberingV1.js'
import { nativeDocxRowBreakPlanV1, nativeDocxRowCutV1, type NativeDocxRowBreakPlanV1 } from './nativeTableRowBreaksV1.js'

export const DOCX_PAGINATION_REQUEST_PROTOCOL = 'injoffice.docx.pagination-request'
export const DOCX_PAGINATION_REQUEST_VERSION = 1 as const
export const DOCX_PAGINATED_LAYOUT_PROTOCOL = 'injoffice.docx.paginated-layout'
export const DOCX_PAGINATED_LAYOUT_VERSION = 1 as const
export const DOCX_PAGINATION_LIMITS = {
  maxPages: 2_048,
  maxLinePlacements: 100_000,
  maxParagraphSlices: 100_000,
  maxDiagnostics: 1_000,
  maxOutputNodes: 2_000_000,
  maxCoordinateMilliPoints: 1_000_000_000_000,
} as const

export interface NativeDocxPaginationRequestV1 {
  protocol: typeof DOCX_PAGINATION_REQUEST_PROTOCOL
  version: typeof DOCX_PAGINATION_REQUEST_VERSION
  document: NativeDocxDocumentV1
  resolved_layout: NativeDocxResolvedLayoutInputV1
  shaped_lines: NativeDocxShapedLinesV1
  pagination_settings: NativeDocxPaginationSettingsV1
}

export type DecodeNativeDocxPaginationRequestV1Result =
  | { ok: true; value: NativeDocxPaginationRequestV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export const DOCX_PAGINATION_DIAGNOSTIC_CODES = Object.freeze([
  'source-diagnostic',
  'settings-attestation-unsupported',
  'default-tab-stop-mismatch',
  'body-table-unsupported',
  'body-structure-unsupported',
  'source-control-unsupported',
  'shaped-paragraph-missing',
  'shaped-paragraph-mismatch',
  'section-map-invalid',
  'section-break-unsupported',
  'multi-column-unsupported',
  'column-geometry-invalid',
  'column-balance-ambiguous',
  'section-geometry-invalid',
  'section-width-mismatch',
  'line-geometry-invalid',
  'keep-chain-conflict',
  'keep-chain-unsatisfiable',
  'keep-lines-unsatisfiable',
  'widow-control-unsatisfiable',
  'header-footer-selection-deferred',
  'note-structure-unsupported',
  'note-reference-ambiguous',
  'note-separator-unsupported',
  'note-overflow-unsupported',
  'resource-limit',
] as const)

export type NativeDocxPaginationDiagnosticCode = typeof DOCX_PAGINATION_DIAGNOSTIC_CODES[number]

export interface NativeDocxPaginationDiagnosticV1 {
  code: NativeDocxPaginationDiagnosticCode
  severity: 'unsupported' | 'deferred'
  scope_id: string
  source_code?: string
  source_message?: string
  message: string
}

export interface NativeDocxPaginationProvenanceV1 {
  document_id: string
  revision: string
  package_sha256: string
  main_part: string
  body_story_id: string
  numbering_source?: NativeDocxResolvedNumberingSourceV1
  pagination_settings: NativeDocxPaginationSettingsV1
  shaped_lines: {
    protocol: typeof DOCX_SHAPED_LINES_PROTOCOL
    version: typeof DOCX_SHAPED_LINES_VERSION
    available_width_millipoints: number
    tab_interval_millipoints: number
  }
  font_manifest: { manifest_id: string; revision: string }
  providers: { resolver_id: string; resolver_revision: string; shaper_id: string; shaper_revision: string; bidi_id: string; bidi_revision: string; bidi_unicode_version: string; unicode13_revision: string }
}

export interface NativeDocxPageBodyBoxV1 {
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
}

export interface NativeDocxPlacedLineV1 {
  table_cell_id?: string
  repeated_table_header?: true
  id: string
  line_id: string
  paragraph_id: string
  section_id: string
  column_id: string
  column_ordinal: number
  source_line_ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
}

export interface NativeDocxParagraphSliceV1 {
  table_cell_id?: string
  repeated_table_header?: true
  id: string
  paragraph_id: string
  section_id: string
  column_id: string
  column_ordinal: number
  slice_ordinal: number
  first_line_ordinal: number
  last_line_ordinal: number
  line_ids: string[]
  top_millipoints: number
  height_millipoints: number
  space_before_millipoints: number
  continued_from_previous_page: boolean
  continues_on_next_page: boolean
  continued_from_previous_column: boolean
  continues_in_next_column: boolean
}

export interface NativeDocxPageColumnV1 {
  id: string
  section_id: string
  ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
}

export interface NativeDocxPlacedNoteStoryV1 {
  id: string
  story_id: string
  story_kind: 'footnote' | 'endnote'
  note_role: 'content' | 'separator'
  native_story_id: string
  relationship_id: string
  ordinal: number
  reference_run_id?: string
  number?: number
  section_id: string
  column_id: string
  column_ordinal: number
  top_millipoints: number
  height_millipoints: number
  lines: NativeDocxPlacedLineV1[]
}

export interface NativeDocxPaginatedPageV1 {
  table_rows?: NativeDocxPlacedTableRowFragmentV1[]
  id: string
  ordinal: number
  section_id: string
  section_ids: string[]
  section_page_ordinal: number
  kind: 'content' | 'parity-blank'
  parity_reason?: 'odd-page-section' | 'even-page-section'
  /** Filler belongs to the preceding section but is triggered by this section. */
  parity_before_section_id?: string
  width_millipoints: number
  height_millipoints: number
  body_box: NativeDocxPageBodyBoxV1
  columns: NativeDocxPageColumnV1[]
  header_refs: NativeDocxHeaderFooterReferenceV1[]
  footer_refs: NativeDocxHeaderFooterReferenceV1[]
  paragraph_slices: NativeDocxParagraphSliceV1[]
  lines: NativeDocxPlacedLineV1[]
  note_stories?: NativeDocxPlacedNoteStoryV1[]
}

export interface NativeDocxPlacedTableRowFragmentV1 {
  id: string
  table_id: string
  row_id: string
  row_ordinal: number
  fragment_ordinal: number
  section_id: string
  column_id: string
  column_ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  source_y_millipoints: number
  source_height_millipoints: number
}

export interface NativeDocxPaginatedSectionV1 {
  section_id: string
  starts_at_block_id: string
  break_type: NativeDocxSectionV1['break_type']
  column_ids: string[]
  page_ids: string[]
}

interface NativeDocxPaginatedLayoutBaseV1 {
  protocol: typeof DOCX_PAGINATED_LAYOUT_PROTOCOL
  version: typeof DOCX_PAGINATED_LAYOUT_VERSION
  provenance: NativeDocxPaginationProvenanceV1
  diagnostics: NativeDocxPaginationDiagnosticV1[]
}

export interface NativeDocxPaginatedLayoutSuccessV1 extends NativeDocxPaginatedLayoutBaseV1 {
  status: 'paginated'
  sections: NativeDocxPaginatedSectionV1[]
  pages: NativeDocxPaginatedPageV1[]
}

export interface NativeDocxPaginatedLayoutRefusedV1 extends NativeDocxPaginatedLayoutBaseV1 {
  status: 'refused'
  sections: []
  pages: []
}

export type NativeDocxPaginatedLayoutV1 = NativeDocxPaginatedLayoutSuccessV1 | NativeDocxPaginatedLayoutRefusedV1
export type PaginateNativeDocxV1Result =
  | { ok: true; value: NativeDocxPaginatedLayoutV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

const REQUEST_FIELDS = ['protocol', 'version', 'document', 'resolved_layout', 'shaped_lines', 'pagination_settings'] as const
const SETTINGS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const RELATIONSHIPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'
const PAGINATION_AFFECTING_CAPABILITIES = new Set(['sections', 'body-structure', 'pagination', 'page-layout'])
const UNSUPPORTED_CONTROLS = new Set(['page-break', 'column-break', 'soft-hyphen'])
const LAYOUT_NEUTRAL_SOURCE_UNSUPPORTED = new Set([
  'DUPLICATE_NATIVE_PARAGRAPH_ID',
  'INVALID_NATIVE_PARAGRAPH_ID',
  'HYPERLINK_SEMANTICS',
  'UNMODELED_COMMENT_MARKUP',
])
const SAFE_INTEGER_MILLI_POINT_FACTOR = 50
const BIDI_TRAILING_RE = /^[\u0009-\u000d\u001c-\u001e\u0020\u0085\u2028\u2029]+$/u

interface PaginationContext {
  approximateLegacySettings?: boolean
  request: NativeDocxPaginationRequestV1
  provenance: NativeDocxPaginationProvenanceV1
  diagnostics: NativeDocxPaginationDiagnosticV1[]
  diagnosticKeys: Set<string>
  refused: boolean
  pages: NativeDocxPaginatedPageV1[]
  sections: NativeDocxPaginatedSectionV1[]
  currentPage?: NativeDocxPaginatedPageV1
  currentSection?: NativeDocxSectionV1
  currentColumnOrdinal: number
  cursorY: number
  previousAfter: number
  sectionPageOrdinal: number
  sliceCount: number
  linePlacementCount: number
  qualifiedTables?: Map<string, NativeDocxQualifiedTableV1>
  lastSliceLocation: Map<string, { pageOrdinal: number; columnOrdinal: number }>
}

function issue(code: NativeDocxValidationIssue['code'], path: string, message: string): NativeDocxValidationIssue {
  return { code, path, message }
}

function exactObject(value: unknown, fields: readonly string[]): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const keys = Object.keys(value)
  const allowed = new Set<string>(fields)
  return keys.length === fields.length && keys.every((key) => allowed.has(key))
}

function safeSnapshot(value: unknown): unknown | undefined {
  try {
    return structuredClone(value)
  } catch {
    return undefined
  }
}

function preflightWire(value: unknown): NativeDocxValidationIssue[] {
  const issues: NativeDocxValidationIssue[] = []
  const active = new WeakSet<object>()
  let nodes = 0
  const visit = (entry: unknown, path: string, depth: number): void => {
    if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
    nodes += 1
    if (nodes > DOCX_NATIVE_LIMITS.maxNodes) {
      issues.push(issue('LIMIT_EXCEEDED', path, `pagination request traversal exceeds ${DOCX_NATIVE_LIMITS.maxNodes} values`))
      return
    }
    if (depth > DOCX_NATIVE_LIMITS.maxDepth) {
      issues.push(issue('LIMIT_EXCEEDED', path, `pagination request nesting exceeds ${DOCX_NATIVE_LIMITS.maxDepth} levels`))
      return
    }
    if (entry === null || entry === undefined) {
      issues.push(issue('INVALID_VALUE', path, 'JSON null and undefined are not permitted'))
      return
    }
    if (typeof entry === 'number' && (!Number.isFinite(entry) || Object.is(entry, -0))) {
      issues.push(issue('INVALID_VALUE', path, 'non-finite numbers and negative zero are not permitted'))
      return
    }
    if (typeof entry === 'string' && entry.length > DOCX_NATIVE_LIMITS.maxTextLength) {
      issues.push(issue('LIMIT_EXCEEDED', path, `string exceeds ${DOCX_NATIVE_LIMITS.maxTextLength} UTF-16 code units`))
      return
    }
    if (typeof entry !== 'object') {
      if (!['string', 'number', 'boolean'].includes(typeof entry)) issues.push(issue('INVALID_TYPE', path, 'must contain only JSON wire values'))
      return
    }
    if (active.has(entry)) {
      issues.push(issue('INVALID_VALUE', path, 'cyclic values are not valid JSON wire input'))
      return
    }
    active.add(entry)
    if (Array.isArray(entry)) {
      if (entry.length > 500_000) issues.push(issue('LIMIT_EXCEEDED', path, 'array exceeds the largest v1 shaped-lines collection bound'))
      const length = Math.min(entry.length, 500_000)
      for (let index = 0; index < length && nodes <= DOCX_NATIVE_LIMITS.maxNodes; index += 1) visit(entry[index], `${path}/${index}`, depth + 1)
    } else {
      for (const key of Object.keys(entry).sort()) {
        visit((entry as Record<string, unknown>)[key], `${path}/${key.replace(/~/g, '~0').replace(/\//g, '~1')}`, depth + 1)
        if (nodes > DOCX_NATIVE_LIMITS.maxNodes) break
      }
    }
    active.delete(entry)
  }
  try { visit(value, '', 0) } catch { return [issue('INVALID_VALUE', '', 'pagination request could not be safely traversed')] }
  return issues
}

function canonicalPartKey(partName: string): string {
  return asciiLowerNative(partName.split('/').map((segment) => decodeURIComponent(segment)).join('/'))
}

function asciiCaseEqual(left: string, right: string): boolean {
  if (!/^[\x00-\x7F]*$/.test(left) || !/^[\x00-\x7F]*$/.test(right)) return false
  const fold = (value: string): string => value.replace(/[A-Z]/g, (character) => String.fromCharCode(character.charCodeAt(0) + 32))
  return fold(left) === fold(right)
}

function checkedSum(...values: number[]): number | undefined {
  let sum = 0
  for (const value of values) {
    sum += value
    if (!Number.isSafeInteger(sum) || Math.abs(sum) > DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints) return undefined
  }
  return sum
}

function twips(value: number): number | undefined {
  const result = signedTwips(value)
  return result !== undefined && result >= 0 ? result : undefined
}

function signedTwips(value: number): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || Math.abs(value) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS || Math.abs(value) > Math.floor(DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints / SAFE_INTEGER_MILLI_POINT_FACTOR)) return undefined
  return value * SAFE_INTEGER_MILLI_POINT_FACTOR
}

function addDiagnostic(context: PaginationContext, diagnostic: NativeDocxPaginationDiagnosticV1): void {
  const key = `${diagnostic.code}\u0000${diagnostic.severity}\u0000${diagnostic.scope_id}\u0000${diagnostic.source_code ?? ''}\u0000${diagnostic.message}`
  if (context.diagnosticKeys.has(key)) return
  if (context.diagnostics.length >= DOCX_PAGINATION_LIMITS.maxDiagnostics) {
    context.refused = true
    if (!context.diagnostics.some((entry) => entry.code === 'resource-limit' && entry.message.includes('diagnostics'))) {
      context.diagnostics[DOCX_PAGINATION_LIMITS.maxDiagnostics - 1] = {
        code: 'resource-limit', severity: 'unsupported', scope_id: context.request.document.document_id,
        message: `Pagination diagnostics exceed ${DOCX_PAGINATION_LIMITS.maxDiagnostics}`,
      }
    }
    return
  }
  context.diagnosticKeys.add(key)
  context.diagnostics.push(diagnostic)
  if (diagnostic.severity === 'unsupported') context.refused = true
}

function refuse(context: PaginationContext, code: NativeDocxPaginationDiagnosticCode, scopeID: string, message: string, source?: { code: string; message: string }): void {
  addDiagnostic(context, { code, severity: 'unsupported', scope_id: scopeID, ...(source ? { source_code: source.code, source_message: source.message } : {}), message })
}

function provenance(request: NativeDocxPaginationRequestV1): NativeDocxPaginationProvenanceV1 {
  return {
    document_id: request.document.document_id,
    revision: request.document.revision,
    package_sha256: request.document.source.package_sha256,
    main_part: request.document.source.main_part,
    body_story_id: request.document.body.id,
    ...(request.shaped_lines.numbering_source ? { numbering_source: { ...request.shaped_lines.numbering_source } } : {}),
    pagination_settings: {
      ...request.pagination_settings,
      diagnostics: request.pagination_settings.diagnostics.map((entry) => ({ ...entry })),
    },
    shaped_lines: {
      protocol: request.shaped_lines.protocol,
      version: request.shaped_lines.version,
      available_width_millipoints: request.shaped_lines.available_width_millipoints,
      tab_interval_millipoints: request.shaped_lines.tab_interval_millipoints,
    },
    font_manifest: { ...request.shaped_lines.font_manifest },
    providers: { ...request.shaped_lines.providers },
  }
}

function nativeInventory(document: NativeDocxDocumentV1): {
  paragraphs: Map<string, { paragraph: NativeDocxParagraphV1; storyID: string; storyKind: NativeDocxShapedParagraphV1['story_kind'] }>
  runs: Map<string, string>
  tables: Set<string>
  ids: Set<string>
} {
  const paragraphs = new Map<string, { paragraph: NativeDocxParagraphV1; storyID: string; storyKind: NativeDocxShapedParagraphV1['story_kind'] }>()
  const runs = new Map<string, string>()
  const tables = new Set<string>()
  const ids = new Set<string>([document.document_id, ...document.sections.map((section) => section.id), ...document.comments.map((comment) => comment.id)])
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  for (const story of stories) {
    ids.add(story.id)
    for (const block of story.blocks) {
      ids.add(block.id)
      const collectParagraph = (paragraph: NativeDocxParagraphV1): void => {
        paragraphs.set(paragraph.id, { paragraph, storyID: story.id, storyKind: story.kind })
        ids.add(paragraph.id)
        for (const run of paragraph.runs) {
          runs.set(run.id, paragraph.id)
          ids.add(run.id)
          if (run.drawing) ids.add(run.drawing.id)
        }
      }
      if (block.paragraph) collectParagraph(block.paragraph)
      if (block.table) {
        tables.add(block.table.id)
        for (const row of block.table.rows) {
          ids.add(row.id)
          for (const cell of row.cells) {
            ids.add(cell.id)
            for (const paragraph of cell.paragraphs) collectParagraph(paragraph)
          }
        }
      }
    }
  }
  return { paragraphs, runs, tables, ids }
}

function validateAuthoritativeParagraphBidi(
  paragraph: NativeDocxShapedParagraphV1,
  native: NativeDocxParagraphV1,
  resolvedRuns: ReadonlyMap<string, { entry: NativeDocxPaginationRequestV1['resolved_layout']['runs'][number]; index: number }>,
  hasNumbering: boolean,
  paragraphIndex: number,
  issues: NativeDocxValidationIssue[],
): void {
  const resolvedByID = new Map([...resolvedRuns].map(([id, value]) => [id, value.entry]))
  const plan = resolveNativeDocxParagraphBidiPlanV1(native, resolvedByID, paragraph.direction)
  const basePath = `/shaped_lines/paragraphs/${paragraphIndex}`
  if (!plan.ok) {
    issues.push(issue('BROKEN_REFERENCE', basePath, `authoritative bidi source projection failed: ${plan.code}: ${plan.message}`))
    return
  }
  const expectedCoverage = new Array<boolean>(plan.value.paragraph.textLengthUtf16).fill(false)
  const requiredCoverage = new Array<boolean>(plan.value.paragraph.textLengthUtf16).fill(false)
  const seenCoverage = new Array<boolean>(plan.value.paragraph.textLengthUtf16).fill(false)
  for (const run of native.runs) {
    const start = plan.value.runStarts.get(run.id)
    const resolved = resolvedRuns.get(run.id)?.entry
    if (start === undefined || !resolved || resolved.properties.hidden) continue
    if (run.kind === 'text') {
      for (let offset = 0; offset < (run.text?.length ?? 0); offset++) if (run.text!.charCodeAt(offset) !== 0x0a && run.text!.charCodeAt(offset) !== 0x0d) expectedCoverage[start + offset] = true
    } else if (run.kind === 'drawing' || (run.kind === 'control' && (run.control === 'tab' || run.control === 'soft-hyphen'))) expectedCoverage[start] = requiredCoverage[start] = true
  }
  for (const [lineIndex, line] of paragraph.lines.entries()) {
    const logical: Array<{ fragment: NativeDocxShapedLineV1['fragments'][number]; visual: number; paragraphStart?: number; authoritativeLevel: number }> = []
    for (const [visualIndex, fragment] of line.fragments.entries()) {
      if (fragment.source_kind === 'list-marker') {
        if (fragment.source_id !== native.id || !hasNumbering) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}/source_id`, 'list marker must reference its exact numbered paragraph'))
        logical.push({ fragment, visual: visualIndex, authoritativeLevel: fragment.bidi_level })
        continue
      }
      const runStart: number | undefined = plan.value.runStarts.get(fragment.source_id)
      const nativeRun = native.runs.find((run) => run.id === fragment.source_id)
      const resolved = resolvedRuns.get(fragment.source_id)?.entry
      if (runStart === undefined || !nativeRun || !resolved || resolved.properties.hidden) {
        issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}/source_id`, 'fragment must replay one visible native source atom'))
        continue
      }
      let paragraphStart: number = runStart
      const scriptKind = resolved.properties.vertical_alignment
      if ((scriptKind && scriptKind !== 'baseline' ? scriptKind : undefined) !== fragment.script_transform?.kind || fragment.script_transform && nativeRun.kind !== 'text') issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}/script_transform`, 'script transform must exactly match authored vertical alignment on a text run'))
      let coveredLength = 0
      let exact = false
      if (nativeRun.kind === 'drawing') {
        coveredLength = 1
        exact = fragment.source_kind === 'image' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && fragment.text === ''
      } else if (nativeRun.kind === 'text' && nativeRun.text !== undefined) {
        paragraphStart += fragment.start_utf16
        coveredLength = fragment.end_utf16 - fragment.start_utf16
        const before = fragment.start_utf16 > 0 ? nativeRun.text.charCodeAt(fragment.start_utf16 - 1) : -1
        const atStart = fragment.start_utf16 < nativeRun.text.length ? nativeRun.text.charCodeAt(fragment.start_utf16) : -1
        const beforeEnd = fragment.end_utf16 > 0 ? nativeRun.text.charCodeAt(fragment.end_utf16 - 1) : -1
        const atEnd = fragment.end_utf16 < nativeRun.text.length ? nativeRun.text.charCodeAt(fragment.end_utf16) : -1
        const boundaries = !(before >= 0xd800 && before <= 0xdbff && atStart >= 0xdc00 && atStart <= 0xdfff) && !(beforeEnd >= 0xd800 && beforeEnd <= 0xdbff && atEnd >= 0xdc00 && atEnd <= 0xdfff)
        exact = fragment.source_kind === 'run' && coveredLength > 0 && fragment.start_utf16 >= 0 && fragment.end_utf16 <= nativeRun.text.length && boundaries && fragment.text === nativeRun.text.slice(fragment.start_utf16, fragment.end_utf16) && !/[\r\n]/.test(fragment.text)
      } else if (nativeRun.kind === 'control' && nativeRun.control === 'tab') {
        coveredLength = 1
        exact = fragment.source_kind === 'tab' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && fragment.text === '\t'
      } else if (nativeRun.kind === 'control' && nativeRun.control === 'soft-hyphen') {
        coveredLength = 1
        exact = fragment.source_kind === 'run' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && fragment.text === ''
      }
      if (!exact || coveredLength <= 0 || paragraphStart < 0 || paragraphStart + coveredLength > plan.value.paragraph.levels.length) {
        issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}`, 'fragment kind, text, and UTF-16 range must exactly replay its native source atom'))
        continue
      }
      const authoritativeLevel = plan.value.paragraph.levels[paragraphStart]!
      logical.push({ fragment, visual: visualIndex, paragraphStart, authoritativeLevel })
      for (let offset = 0; offset < coveredLength; offset++) {
        const sourceOffset = paragraphStart + offset
        if (!expectedCoverage[sourceOffset] || seenCoverage[sourceOffset]) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}`, 'fragment source range is duplicated or not paintable native text'))
        seenCoverage[sourceOffset] = true
        if (plan.value.paragraph.levels[sourceOffset] !== fragment.bidi_level) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/fragments/${visualIndex}/bidi_level`, 'must equal the pinned UAX #9 level for the exact native source cluster'))
      }
    }
    const ordered = [...logical].sort((left, right) => left.fragment.logical_order - right.fragment.logical_order)
    let previousParagraphStart = -1
    for (const entry of ordered) if (entry.paragraphStart !== undefined) {
      if (entry.paragraphStart < previousParagraphStart) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/logical_to_visual`, 'native source fragments must retain monotone logical source order'))
      previousParagraphStart = entry.paragraphStart
    }
    const reordered = reorderNativeBidiLineV1(ordered.map((entry) => entry.authoritativeLevel), paragraph.direction === 'rtl' ? 1 : 0, ordered.map((entry) => BIDI_TRAILING_RE.test(entry.fragment.text)))
    const exactOrder = ordered.length === line.fragments.length && reordered.ok
      && reordered.value.visualToLogical.every((logicalIndex, visual) => ordered[logicalIndex]!.visual === visual)
      && reordered.value.logicalToVisual.every((visual, logicalIndex) => line.logical_to_visual[logicalIndex] === visual && ordered[logicalIndex]!.fragment.logical_order === logicalIndex)
    if (!exactOrder) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines/${lineIndex}/logical_to_visual`, 'must equal pinned UAX #9 L1/L2 visual ordering for the exact native source-cluster sequence'))
  }
  if (requiredCoverage.some((expected, index) => expected && !seenCoverage[index])) issues.push(issue('BROKEN_REFERENCE', `${basePath}/lines`, 'fragments must cover every visible native control and U+FFFC image source atom exactly once'))
}

function validateRequest(value: unknown): PaginateNativeDocxV1Result | { ok: true; request: NativeDocxPaginationRequestV1 } {
  const preflightIssues = preflightWire(value)
  if (preflightIssues.length > 0) return { ok: false, issues: preflightIssues }
  const snapshot = safeSnapshot(value)
  if (snapshot === undefined) return { ok: false, issues: [issue('INVALID_VALUE', '', 'pagination request must be a cloneable JSON wire value')] }
  if (!exactObject(snapshot, REQUEST_FIELDS)) {
    if (snapshot === null || typeof snapshot !== 'object' || Array.isArray(snapshot)) return { ok: false, issues: [issue('INVALID_TYPE', '', 'pagination request must be an object')] }
    const issues: NativeDocxValidationIssue[] = []
    const allowed = new Set<string>(REQUEST_FIELDS)
    for (const key of Object.keys(snapshot).sort()) if (!allowed.has(key)) issues.push(issue('UNKNOWN_FIELD', `/${key}`, `unknown field ${JSON.stringify(key)}`))
    for (const key of REQUEST_FIELDS) if (!(key in snapshot)) issues.push(issue('REQUIRED', `/${key}`, 'field is required'))
    return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  }
  if (snapshot.protocol !== DOCX_PAGINATION_REQUEST_PROTOCOL) return { ok: false, issues: [issue('UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGINATION_REQUEST_PROTOCOL}`)] }
  if (snapshot.version !== DOCX_PAGINATION_REQUEST_VERSION) return { ok: false, issues: [issue('UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGINATION_REQUEST_VERSION}`)] }
  const document = decodeNativeDocxDocument(snapshot.document)
  const resolved = decodeNativeDocxResolvedLayout(snapshot.resolved_layout)
  const shaped = decodeNativeDocxShapedLines(snapshot.shaped_lines)
  const settings = decodeNativeDocxPaginationSettings(snapshot.pagination_settings)
  const issues: NativeDocxValidationIssue[] = []
  if (!document.ok) issues.push(...document.issues.map((entry) => ({ ...entry, path: `/document${entry.path}` })))
  if (!resolved.ok) issues.push(...resolved.issues.map((entry) => ({ ...entry, path: `/resolved_layout${entry.path}` })))
  if (!shaped.ok) issues.push(...shaped.issues.map((entry) => ({ ...entry, path: `/shaped_lines${entry.path}` })))
  if (!settings.ok) issues.push(...settings.issues.map((entry) => ({ ...entry, path: `/pagination_settings${entry.path}` })))
  if (issues.length > 0 || !document.ok || !resolved.ok || !shaped.ok || !settings.ok) return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  const IDs = [resolved.value.document_id, shaped.value.document_id, settings.value.document_id]
  if (IDs.some((entry) => entry !== document.value.document_id)) issues.push(issue('BROKEN_REFERENCE', '/document_id', 'document, resolved-layout, and shaped-lines document ids must match'))
  const revisions = [resolved.value.revision, shaped.value.revision, settings.value.revision]
  if (revisions.some((entry) => entry !== document.value.revision)) issues.push(issue('BROKEN_REFERENCE', '/revision', 'document, resolved-layout, and shaped-lines revisions must match'))
  if (canonicalPartKey(resolved.value.source_parts.main_part) !== canonicalPartKey(document.value.source.main_part)) issues.push(issue('BROKEN_REFERENCE', '/resolved_layout/source_parts/main_part', 'must identify the native document main part'))
  if (canonicalPartKey(settings.value.main_part) !== canonicalPartKey(document.value.source.main_part)) issues.push(issue('BROKEN_REFERENCE', '/pagination_settings/main_part', 'must identify the native document main part'))
  if (settings.value.package_sha256 !== document.value.source.package_sha256) issues.push(issue('BROKEN_REFERENCE', '/pagination_settings/package_sha256', 'settings attestation must bind the exact native source package fingerprint'))
  if (!samePaginationWire(resolved.value.numbering_source, shaped.value.numbering_source)) issues.push(issue('BROKEN_REFERENCE', '/shaped_lines/numbering_source', 'must exactly match resolved-layout numbering relationship, part, and model provenance'))
  if (settings.value.relationships_part) {
    const relationshipsPart = document.value.passthrough_parts.find((part) => canonicalPartKey(part.part_name) === canonicalPartKey(settings.value.relationships_part!))
    if (!relationshipsPart || !asciiCaseEqual(relationshipsPart.content_type, RELATIONSHIPS_CONTENT_TYPE) || relationshipsPart.sha256 !== settings.value.relationships_sha256) issues.push(issue('BROKEN_REFERENCE', '/pagination_settings/relationships_part', 'settings attestation must bind the exact preserved owning relationship part and fingerprint'))
  }
  if (settings.value.settings_part) {
    const settingsPart = document.value.passthrough_parts.find((part) => canonicalPartKey(part.part_name) === canonicalPartKey(settings.value.settings_part!))
    if (!settingsPart || !asciiCaseEqual(settingsPart.content_type, SETTINGS_CONTENT_TYPE) || settingsPart.sha256 !== settings.value.settings_sha256) issues.push(issue('BROKEN_REFERENCE', '/pagination_settings/settings_part', 'settings attestation must match the exact preserved native settings part, ASCII content type, and fingerprint'))
  }
  const inventory = nativeInventory(document.value)
  const resolvedParagraphs = new Map(resolved.value.paragraphs.map((entry, index) => [entry.paragraph_id, { entry, index }]))
  const resolvedRuns = new Map(resolved.value.runs.map((entry, index) => [entry.run_id, { entry, index }]))
  const resolvedTables = new Map(resolved.value.tables.map((entry, index) => [entry.table_id, index]))
  for (const paragraphID of inventory.paragraphs.keys()) if (!resolvedParagraphs.has(paragraphID)) issues.push(issue('BROKEN_REFERENCE', '/resolved_layout/paragraphs', `native paragraph ${JSON.stringify(paragraphID)} is missing from resolved layout`))
  for (const [paragraphID, resolvedParagraph] of resolvedParagraphs) if (!inventory.paragraphs.has(paragraphID)) issues.push(issue('BROKEN_REFERENCE', `/resolved_layout/paragraphs/${resolvedParagraph.index}/paragraph_id`, 'resolved paragraph must reference an exact native paragraph'))
  for (const [runID, paragraphID] of inventory.runs) {
    const resolvedRun = resolvedRuns.get(runID)
    if (!resolvedRun) issues.push(issue('BROKEN_REFERENCE', '/resolved_layout/runs', `native run ${JSON.stringify(runID)} is missing from resolved layout`))
    else if (resolvedRun.entry.paragraph_id !== paragraphID) issues.push(issue('BROKEN_REFERENCE', `/resolved_layout/runs/${resolvedRun.index}/paragraph_id`, `must equal the native owning paragraph ${JSON.stringify(paragraphID)}`))
  }
  for (const [runID, resolvedRun] of resolvedRuns) {
    if (!inventory.runs.has(runID)) issues.push(issue('BROKEN_REFERENCE', `/resolved_layout/runs/${resolvedRun.index}/run_id`, 'resolved run must reference an exact native run'))
  }
  for (const tableID of inventory.tables) if (!resolvedTables.has(tableID)) issues.push(issue('BROKEN_REFERENCE', '/resolved_layout/tables', `native table ${JSON.stringify(tableID)} is missing from resolved layout`))
  for (const [tableID, index] of resolvedTables) if (!inventory.tables.has(tableID)) issues.push(issue('BROKEN_REFERENCE', `/resolved_layout/tables/${index}/table_id`, 'resolved table must reference an exact native table'))
  for (const [paragraphIndex, paragraph] of shaped.value.paragraphs.entries()) {
    const nativeEntry = inventory.paragraphs.get(paragraph.paragraph_id)
    if (!nativeEntry || paragraph.story_id !== nativeEntry.storyID || paragraph.story_kind !== nativeEntry.storyKind) {
      issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/paragraph_id`, 'shaped paragraph must reference its exact native story and paragraph identity'))
      continue
    }
    const resolvedParagraph = resolvedParagraphs.get(paragraph.paragraph_id)?.entry
    if (!resolvedParagraph) {
      issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/paragraph_id`, 'shaped paragraph must have an exact resolved-layout paragraph'))
      continue
    }
    const properties = resolvedParagraph.properties
    const direction = properties.bidi === true ? 'rtl' : 'ltr'
    const physicalStart = direction === 'rtl' ? properties.indent_right_twips : properties.indent_left_twips
    const physicalEnd = direction === 'rtl' ? properties.indent_left_twips : properties.indent_right_twips
    const expected = {
      direction,
      alignment: properties.alignment ?? 'start',
      spacing_before_millipoints: twips(properties.spacing_before_twips ?? 0),
      spacing_after_millipoints: twips(properties.spacing_after_twips ?? 0),
      indent_start_millipoints: signedTwips(properties.indent_start_twips ?? physicalStart ?? 0),
      indent_end_millipoints: signedTwips(properties.indent_end_twips ?? physicalEnd ?? 0),
      first_line_delta_millipoints: properties.first_line_twips !== undefined
        ? signedTwips(properties.first_line_twips)
        : properties.hanging_twips !== undefined ? signedTwips(-properties.hanging_twips) : 0,
    }
    for (const key of Object.keys(expected) as Array<keyof typeof expected>) {
      if (paragraph[key] !== expected[key]) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/${key}`, `must equal the deterministic projection of resolved_layout paragraph properties (${String(expected[key])})`))
    }
    const numbering = resolvedParagraph.numbering
    const marker = paragraph.list_marker
    if ((numbering === undefined) !== (marker === undefined)) {
      issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/list_marker`, 'must appear exactly when resolved layout carries a native numbering marker'))
    } else if (numbering && marker && resolved.value.numbering_source) {
      const markerFragments = paragraph.lines[0]?.fragments
        .filter((fragment) => fragment.source_kind === 'list-marker')
        .sort((left, right) => left.logical_order - right.logical_order) ?? []
      const visibleMarkerFragments = markerFragments.filter((fragment) =>
          fragment.end_utf16 > fragment.start_utf16
          && fragment.end_utf16 <= marker.text.length
          && marker.text.slice(fragment.start_utf16, fragment.end_utf16) === fragment.text)
      const visibleMarkerAdvance = checkedSum(...visibleMarkerFragments.map((fragment) => fragment.advance_inline_millipoints))
      const expectedMarkerGeometry = visibleMarkerAdvance === undefined ? undefined : positionNativeDocxListMarkerV1(numbering, direction, visibleMarkerAdvance)
      const markerEnd = expectedMarkerGeometry === undefined || visibleMarkerAdvance === undefined
        ? undefined
        : checkedSum(expectedMarkerGeometry.marker_start_millipoints, visibleMarkerAdvance)
      const suffixFragments = markerFragments.filter((fragment) => numbering.suffix === 'tab'
        ? fragment.text === '\t' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0
        : numbering.suffix === 'space'
          ? fragment.text === ' ' && fragment.start_utf16 === marker.text.length && fragment.end_utf16 === marker.text.length + 1
          : false)
      const suffixAdvance = checkedSum(...suffixFragments.map((fragment) => fragment.advance_inline_millipoints))
      const numberingTab = numbering.numbering_tab_twips === undefined ? undefined : twips(numbering.numbering_tab_twips)
      const numberingTabValid = numbering.numbering_tab_twips === undefined || numberingTab !== undefined
      const defaultTabInterval = twips(settings.value.default_tab_stop_twips)
      const expectedTextStart = markerEnd === undefined || suffixAdvance === undefined || defaultTabInterval === undefined || !numberingTabValid
        ? undefined
        : numbering.suffix === 'nothing'
          ? suffixFragments.length === 0 ? markerEnd : undefined
          : numbering.suffix === 'space'
            ? suffixFragments.length > 0 ? checkedSum(markerEnd, suffixAdvance) : undefined
            : suffixFragments.length === 1
              ? nativeDocxListSuffixTabTargetV1(markerEnd, expectedMarkerGeometry!.body_text_start_millipoints, defaultTabInterval, numberingTab)
              : undefined
      const exactSuffixAdvance = numbering.suffix !== 'tab' || expectedTextStart === undefined || markerEnd === undefined || suffixAdvance === expectedTextStart - markerEnd
      const exact = marker.marker_id === numbering.marker_id
        && marker.definition_sha256 === numbering.definition_sha256
        && marker.numbering_part_sha256 === resolved.value.numbering_source.part_sha256
        && marker.model_sha256 === resolved.value.numbering_source.model_sha256
        && marker.num_id === numbering.num_id
        && marker.abstract_num_id === numbering.abstract_num_id
        && marker.level === numbering.level
        && marker.counter_value === numbering.counter_value
        && marker.text === numbering.resolved_text
        && marker.suffix === numbering.suffix
        && marker.alignment === numbering.alignment
        && marker.label_start_millipoints === signedTwips(numbering.label_start_twips)
        && marker.label_end_millipoints === signedTwips(numbering.label_end_twips)
        && expectedMarkerGeometry !== undefined
        && marker.marker_start_millipoints === expectedMarkerGeometry.marker_start_millipoints
        && marker.marker_advance_millipoints === visibleMarkerAdvance
        && marker.label_start_millipoints === expectedMarkerGeometry.label_start_millipoints
        && marker.label_end_millipoints === expectedMarkerGeometry.label_end_millipoints
        && marker.text_start_millipoints === expectedTextStart
        && exactSuffixAdvance
      if (!exact) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/list_marker`, 'must exactly project resolved counter, text, definition, shaped marker end, resolved suffix, exact numbering tab, attested default-tab interval, source hashes, and hanging geometry'))
    }
    const native = nativeEntry.paragraph
    const replayable = native.runs.every((run) => run.kind !== 'reference' && (run.kind !== 'drawing' || !!run.drawing && qualifyNativeDocxInlineImageV1(document.value, run.id, run.drawing).ok))
    if (replayable) validateAuthoritativeParagraphBidi(paragraph, native, resolvedRuns, resolvedParagraph.numbering !== undefined, paragraphIndex, issues)
    const runIDs = new Set(native.runs.map((run) => run.id))
    for (const [lineIndex, line] of paragraph.lines.entries()) {
      for (const [fragmentIndex, fragment] of line.fragments.entries()) {
        const validSource = fragment.source_kind === 'list-marker' ? fragment.source_id === native.id : runIDs.has(fragment.source_id)
        if (!validSource) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/lines/${lineIndex}/fragments/${fragmentIndex}/source_id`, 'fragment source must reference its native paragraph or one of that paragraph’s runs'))
        if (fragment.source_kind !== 'list-marker') {
          const resolvedRun = resolvedRuns.get(fragment.source_id)
          if (!resolvedRun || resolvedRun.entry.paragraph_id !== paragraph.paragraph_id) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/lines/${lineIndex}/fragments/${fragmentIndex}/source_id`, 'authored fragment source must reference a resolved run owned by the same paragraph'))
        }
      }
      if (line.hard_break_after) {
        const resolvedRun = resolvedRuns.get(line.hard_break_after.source_run_id)
        if (!runIDs.has(line.hard_break_after.source_run_id) || !resolvedRun || resolvedRun.entry.paragraph_id !== paragraph.paragraph_id) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/paragraphs/${paragraphIndex}/lines/${lineIndex}/hard_break_after/source_run_id`, 'hard break must reference a native and resolved run owned by the same paragraph'))
      }
    }
  }
  for (const [diagnosticIndex, diagnostic] of shaped.value.diagnostics.entries()) {
    if (!inventory.ids.has(diagnostic.scope_id)) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/diagnostics/${diagnosticIndex}/scope_id`, 'shaping diagnostic scope must reference a native identity'))
    if (diagnostic.source_id && !inventory.ids.has(diagnostic.source_id)) issues.push(issue('BROKEN_REFERENCE', `/shaped_lines/diagnostics/${diagnosticIndex}/source_id`, 'shaping diagnostic source must reference a native identity'))
  }
  if (issues.length > 0) return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  return {
    ok: true,
    request: {
      protocol: DOCX_PAGINATION_REQUEST_PROTOCOL,
      version: DOCX_PAGINATION_REQUEST_VERSION,
      document: document.value,
      resolved_layout: resolved.value,
      shaped_lines: shaped.value,
      pagination_settings: settings.value,
    },
  }
}

/** Strictly decodes and joins all four native pagination input projections. */
export function decodeNativeDocxPaginationRequestV1(value: unknown): DecodeNativeDocxPaginationRequestV1Result {
  try {
    const decoded = validateRequest(value)
    if (!decoded.ok) return decoded
    if (!('request' in decoded)) return { ok: false, issues: [issue('INVALID_VALUE', '', 'pagination request decoder reached an invalid internal state')] }
    return { ok: true, value: decoded.request }
  } catch {
    return { ok: false, issues: [issue('INVALID_VALUE', '', 'pagination request could not be safely inspected')] }
  }
}

function bodyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  return document.body.blocks.flatMap((block) => block.kind === 'paragraph' && block.paragraph
    ? [block.paragraph]
    : block.kind === 'table' && block.table ? block.table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) : [])
}

function blockParagraphs(block: NativeDocxBlockV1): NativeDocxParagraphV1[] {
  return block.paragraph ? [block.paragraph] : block.table ? block.table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) : []
}

function bodyScopeIDs(document: NativeDocxDocumentV1): Set<string> {
  const result = new Set<string>([document.document_id, document.body.id])
  for (const section of document.sections) result.add(section.id)
  for (const block of document.body.blocks) {
    result.add(block.id)
    if (block.paragraph) for (const run of block.paragraph.runs) {
      result.add(run.id)
      if (run.drawing) result.add(run.drawing.id)
    }
    if (block.table) for (const row of block.table.rows) {
      result.add(row.id)
      for (const cell of row.cells) {
        result.add(cell.id)
        for (const paragraph of cell.paragraphs) {
          result.add(paragraph.id)
          for (const run of paragraph.runs) result.add(run.id)
        }
      }
    }
  }
  return result
}

function refuseUnsupportedSource(context: PaginationContext): void {
  const { document, shaped_lines: shaped } = context.request
  const scopes = bodyScopeIDs(document)
  const settings = context.request.pagination_settings
  if (settings.profile !== 'word-modern-default') {
    if (context.approximateLegacySettings) {
      addDiagnostic(context, { code: 'settings-attestation-unsupported', severity: 'deferred', scope_id: document.document_id, message: 'Explicit approximate preview uses current layout policy instead of legacy Word layout semantics; original settings and reasons are retained in the approximate envelope' })
    } else {
      if (settings.profile === 'absent-default') refuse(context, 'settings-attestation-unsupported', document.document_id, 'Omitted compatibilityMode defaults to Word mode 12; pagination v1 requires an explicit compatibilityMode=15 attestation')
      for (const diagnostic of settings.diagnostics) refuse(context, 'settings-attestation-unsupported', document.document_id, `Native pagination settings refuse layout: ${diagnostic.code}: ${diagnostic.message}`, { code: diagnostic.code, message: diagnostic.message })
    }
  }
  const expectedTabInterval = twips(settings.default_tab_stop_twips)
  if (expectedTabInterval === undefined || expectedTabInterval !== shaped.tab_interval_millipoints) {
    refuse(context, 'default-tab-stop-mismatch', document.document_id, `Shaped tab interval ${shaped.tab_interval_millipoints} does not match the attested Word default tab stop ${settings.default_tab_stop_twips} twips`)
  }
  const qualified = qualifyNativeDocxTablesV1(document, context.request.resolved_layout, shaped)
  if (qualified.status === 'refused') for (const diagnostic of qualified.diagnostics) refuse(context, 'body-table-unsupported', diagnostic.scope_id, diagnostic.message)
  else context.qualifiedTables = new Map(qualified.tables.map((table) => [table.table.id, table]))
  for (const entry of document.unsupported) {
    if (LAYOUT_NEUTRAL_SOURCE_UNSUPPORTED.has(entry.code)) {
      addDiagnostic(context, {
        code: 'source-diagnostic', severity: 'deferred', scope_id: entry.scope_id,
        source_code: entry.code, source_message: entry.message,
        message: `Native preserve-only record is identity-, relationship-, or detached-story-only and does not alter current body advances: ${entry.code}: ${entry.message}`,
      })
      continue
    }
    if (entry.code === 'DEFAULT_SECTION_INFERRED' && entry.capability === 'sections') {
      const section = document.sections.find((candidate) => candidate.id === entry.scope_id)
      const page = section?.page
      const margins = page?.margins
      if (section && page?.width_twips === 12_240 && page.height_twips === 15_840 && page.orientation === 'portrait' && page.columns === 1 && page.column_spacing_twips === 720 && page.column_layout === 'equal-width' && page.column_definitions.length === 1 && margins?.top_twips === 1_440 && margins.right_twips === 1_440 && margins.bottom_twips === 1_440 && margins.left_twips === 1_440 && margins.header_twips === 720 && margins.footer_twips === 720 && margins.gutter_twips === 0) {
        addDiagnostic(context, {
          code: 'source-diagnostic', severity: 'deferred', scope_id: entry.scope_id,
          source_code: entry.code, source_message: entry.message,
          message: 'Schema-optional final section is absent; pagination uses the extractor’s explicit Word default geometry without guessing',
        })
        continue
      }
    }
    if (PAGINATION_AFFECTING_CAPABILITIES.has(entry.capability) || scopes.has(entry.scope_id)) {
      refuse(context, 'body-structure-unsupported', entry.scope_id, `Native unsupported record can affect body pagination: ${entry.code}: ${entry.message}`, { code: entry.code, message: entry.message })
    }
  }
  for (const paragraph of bodyParagraphs(document)) for (const run of paragraph.runs) {
    if (run.control && UNSUPPORTED_CONTROLS.has(run.control)) refuse(context, 'source-control-unsupported', run.id, `Native ${run.control} is not represented by the shaped-lines v1 pagination input`)
    if (run.drawing) {
      const image = qualifyNativeDocxInlineImageV1(document, run.id, run.drawing)
      if (!image.ok) refuse(context, 'body-structure-unsupported', run.id, `Native drawing is outside the exact inline-image slice: ${image.message}`)
    }
    if (run.reference && run.reference.kind !== 'footnote' && run.reference.kind !== 'endnote') refuse(context, 'body-structure-unsupported', run.id, `Native ${run.reference.kind} reference placement is not represented by the shaped-lines v1 pagination input`)
  }
  for (const diagnostic of shaped.diagnostics) {
    const source = { code: diagnostic.source_diagnostic_code ?? diagnostic.code, message: diagnostic.source_diagnostic_message ?? diagnostic.message }
    if (diagnostic.severity === 'unsupported' && scopes.has(diagnostic.scope_id)) {
      refuse(context, 'source-diagnostic', diagnostic.scope_id, `Body shaping diagnostic refuses pagination: ${diagnostic.code}: ${diagnostic.message}`, source)
    } else {
      addDiagnostic(context, {
        code: 'source-diagnostic', severity: 'deferred', scope_id: diagnostic.scope_id,
        source_code: source.code, source_message: source.message,
        message: `Shaping diagnostic retained by pagination: ${diagnostic.code}: ${diagnostic.message}`,
      })
    }
  }
}

interface SectionGroup {
  section: NativeDocxSectionV1
  blocks: NativeDocxBlockV1[]
}

function sectionGroups(context: PaginationContext): SectionGroup[] {
  const blocks = context.request.document.body.blocks
  const sections = context.request.document.sections
  const indexes = new Map(blocks.map((block, index) => [block.id, index]))
  const starts: number[] = []
  for (const [index, section] of sections.entries()) {
    const start = indexes.get(section.starts_at_block_id)
    if (start === undefined || (index === 0 && start !== 0) || (index > 0 && start <= (starts[index - 1] ?? -1))) {
      refuse(context, 'section-map-invalid', section.id, 'Sections must start at strictly increasing body block ids and the first section must start at the first body block')
      return []
    }
    starts.push(start)
  }
  return sections.map((section, index) => ({
    section,
    blocks: blocks.slice(starts[index], starts[index + 1] ?? blocks.length),
  }))
}

function sectionBodyBox(context: PaginationContext, section: NativeDocxSectionV1): NativeDocxPageBodyBoxV1 | undefined {
  const qualified = qualifyNativeDocxSectionColumnsV1(section)
  if (!qualified.ok) {
    refuse(context, qualified.code === 'section-geometry-invalid' ? 'section-geometry-invalid' : 'column-geometry-invalid', section.id, qualified.message)
    return undefined
  }
  const columnWidth = qualified.value.columns[0]?.width_millipoints
  if (columnWidth !== context.request.shaped_lines.available_width_millipoints) {
    refuse(context, 'section-width-mismatch', section.id, `Section column width ${columnWidth} does not match shaped width ${context.request.shaped_lines.available_width_millipoints}`)
    return undefined
  }
  return {
    x_millipoints: qualified.value.body_x_millipoints,
    y_millipoints: qualified.value.body_y_millipoints,
    width_millipoints: qualified.value.body_width_millipoints,
    height_millipoints: qualified.value.body_height_millipoints,
  }
}

function qualifiedPageColumns(context: PaginationContext, section: NativeDocxSectionV1): NativeDocxPageColumnV1[] | undefined {
  const qualified = qualifyNativeDocxSectionColumnsV1(section)
  if (!qualified.ok) {
    refuse(context, qualified.code === 'section-geometry-invalid' ? 'section-geometry-invalid' : 'column-geometry-invalid', section.id, qualified.message)
    return undefined
  }
  return qualified.value.columns.map((column) => ({ ...column, section_id: section.id }))
}

function ensurePageSectionColumns(context: PaginationContext, section: NativeDocxSectionV1): boolean {
  const page = context.currentPage
  if (!page) return false
  if (!page.section_ids.includes(section.id)) page.section_ids.push(section.id)
  if (page.columns.some((column) => column.section_id === section.id)) return true
  const columns = qualifiedPageColumns(context, section)
  if (!columns) return false
  page.columns.push(...columns)
  return true
}

function newPage(context: PaginationContext, section: NativeDocxSectionV1, kind: NativeDocxPaginatedPageV1['kind'], parityReason?: NativeDocxPaginatedPageV1['parity_reason'], parityBeforeSectionID?: string): NativeDocxPaginatedPageV1 | undefined {
  if (context.pages.length >= DOCX_PAGINATION_LIMITS.maxPages) {
    refuse(context, 'resource-limit', section.id, `Pagination exceeds ${DOCX_PAGINATION_LIMITS.maxPages} pages`)
    return undefined
  }
  const box = sectionBodyBox(context, section)
  const columns = qualifiedPageColumns(context, section)
  const width = twips(section.page.width_twips)
  const height = twips(section.page.height_twips)
  if (!box || !columns || width === undefined || height === undefined) return undefined
  const ordinal = context.pages.length
  const sectionOrdinal = context.sectionPageOrdinal
  const id = kind === 'parity-blank' ? `page:parity-before:${parityBeforeSectionID}` : `page:${section.id}:${sectionOrdinal}`
  const page: NativeDocxPaginatedPageV1 = {
    id,
    ordinal,
    section_id: section.id,
    section_ids: [section.id],
    section_page_ordinal: sectionOrdinal,
    kind,
    ...(parityReason ? { parity_reason: parityReason } : {}),
    ...(parityBeforeSectionID ? { parity_before_section_id: parityBeforeSectionID } : {}),
    width_millipoints: width,
    height_millipoints: height,
    body_box: box,
    columns,
    header_refs: kind === 'parity-blank' ? [] : section.header_refs.map((entry) => ({ ...entry })),
    footer_refs: kind === 'parity-blank' ? [] : section.footer_refs.map((entry) => ({ ...entry })),
    paragraph_slices: [],
    lines: [],
  }
  context.sectionPageOrdinal += 1
  context.pages.push(page)
  context.sections.at(-1)?.page_ids.push(page.id)
  context.currentPage = page
  context.currentSection = section
  context.currentColumnOrdinal = 0
  context.cursorY = 0
  context.previousAfter = 0
  return page
}

function currentColumn(context: PaginationContext): NativeDocxPageColumnV1 | undefined {
  return context.currentPage?.columns.find((column) => column.section_id === context.currentSection?.id && column.ordinal === context.currentColumnOrdinal)
}

function columnHasContent(context: PaginationContext): boolean {
  const column = currentColumn(context)
  if (!column || !context.currentPage) return false
  const pageColumns = new Map(context.currentPage.columns.map((entry) => [entry.id, entry]))
  return context.currentPage.lines.some((line) => {
    const occupied = pageColumns.get(line.column_id)
    return occupied !== undefined && occupied.ordinal === column.ordinal && occupied.x_millipoints === column.x_millipoints && occupied.y_millipoints === column.y_millipoints && occupied.width_millipoints === column.width_millipoints && occupied.height_millipoints === column.height_millipoints
  })
}

function pageHasContent(context: PaginationContext): boolean {
  return (context.currentPage?.lines.length ?? 0) > 0
}

function remainingHeight(context: PaginationContext): number {
  return (currentColumn(context)?.height_millipoints ?? 0) - context.cursorY
}

function startSection(context: PaginationContext, section: NativeDocxSectionV1, first: boolean): void {
  const previousSection = context.currentSection
  if (!first && (section.break_type === 'odd-page' || section.break_type === 'even-page')) {
    const nextPageNumber = context.pages.length + 1
    const wantsOdd = section.break_type === 'odd-page'
    const parityMatches = wantsOdd ? nextPageNumber % 2 === 1 : nextPageNumber % 2 === 0
    if (!parityMatches && context.currentSection) newPage(context, context.currentSection, 'parity-blank', wantsOdd ? 'odd-page-section' : 'even-page-section', section.id)
  }
  context.sections.push({ section_id: section.id, starts_at_block_id: section.starts_at_block_id, break_type: section.break_type, column_ids: section.page.column_definitions.map((column) => column.id), page_ids: [] })
  context.sectionPageOrdinal = 0
  if (!first && (section.break_type === 'continuous' || section.break_type === 'next-column')) {
    if (section.title_page) {
      refuse(context, 'section-geometry-invalid', section.id, `${section.break_type} with title-page headers or footers has no unambiguous first-page selection on a shared physical page`)
      return
    }
    if (!previousSection || !context.currentPage || !nativeDocxSectionsShareExactPageV1(previousSection, section)) {
      refuse(context, 'section-geometry-invalid', section.id, `${section.break_type} requires identical page, column, margin, and header/footer geometry across the shared physical page`)
      return
    }
    const priorColumn = context.currentColumnOrdinal
    const priorRemainingHeight = remainingHeight(context)
    context.currentSection = section
    if (section.break_type === 'continuous') {
      if (priorRemainingHeight === 0) {
        context.currentPage = undefined
        context.currentColumnOrdinal = 0
        if (!context.refused) newPage(context, section, 'content')
        return
      }
      if (!ensurePageSectionColumns(context, section)) return
      if (section.page.columns !== 1) {
        refuse(context, 'column-balance-ambiguous', section.id, 'Continuous multi-column transition requires a unique preceding-section balance plan outside the bounded v1 slice')
        return
      }
      context.currentColumnOrdinal = priorColumn
      context.sections.at(-1)!.page_ids.push(context.currentPage.id)
      context.sectionPageOrdinal = 1
      return
    }
    context.currentColumnOrdinal = priorColumn + 1
    context.cursorY = 0
    context.previousAfter = 0
    if (context.currentColumnOrdinal < section.page.columns) {
      if (!ensurePageSectionColumns(context, section)) return
      context.sections.at(-1)!.page_ids.push(context.currentPage.id)
      context.sectionPageOrdinal = 1
      return
    }
    context.currentPage = undefined
    context.currentColumnOrdinal = 0
    if (!context.refused) newPage(context, section, 'content')
    return
  }
  context.currentPage = undefined
  context.currentSection = section
  if (!context.refused) newPage(context, section, 'content')
  if ((section.header_refs.length > 0 || section.footer_refs.length > 0) && !context.refused) addDiagnostic(context, {
    code: 'header-footer-selection-deferred', severity: 'deferred', scope_id: section.id,
    message: `Header/footer references are retained on pages and even/odd selection is attested as ${context.request.pagination_settings.even_and_odd_headers}; v1 does not select variants or place story content because title-page policy and shaped story heights are not modeled`,
  })
}

function startNextContentPage(context: PaginationContext): void {
  if (!context.currentSection || context.refused) return
  newPage(context, context.currentSection, 'content')
}

function startNextFlowColumn(context: PaginationContext): void {
  if (!context.currentSection || context.refused) return
  if (context.currentColumnOrdinal + 1 < context.currentSection.page.columns) {
    context.currentColumnOrdinal += 1
    context.cursorY = 0
    context.previousAfter = 0
    return
  }
  startNextContentPage(context)
}

function paragraphGap(context: PaginationContext, paragraph: NativeDocxShapedParagraphV1, continuation: boolean): number {
  if (continuation) return 0
  if (!columnHasContent(context)) return context.currentPage?.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 ? paragraph.spacing_before_millipoints : 0
  return Math.max(context.previousAfter, paragraph.spacing_before_millipoints)
}

function lineGeometryValid(context: PaginationContext, paragraph: NativeDocxShapedParagraphV1, line: NativeDocxShapedLineV1): boolean {
  const column = currentColumn(context)
  if (!column) return false
  const left = checkedSum(column.x_millipoints, line.inline_offset_millipoints)
  const right = left === undefined ? undefined : checkedSum(left, line.advance_inline_millipoints)
  const columnRight = checkedSum(column.x_millipoints, column.width_millipoints)
  if (line.available_width_millipoints > column.width_millipoints || left === undefined || right === undefined || columnRight === undefined || left < column.x_millipoints || right > columnRight) {
    refuse(context, 'line-geometry-invalid', line.id, `Shaped line ${line.id} escapes its exact section column or advertises a wider shaping box`)
    return false
  }
  if (line.line_height_millipoints > column.height_millipoints) {
    refuse(context, 'line-geometry-invalid', paragraph.paragraph_id, 'A shaped line is taller than its section column')
    return false
  }
  return true
}

function placeSlice(context: PaginationContext, paragraph: NativeDocxShapedParagraphV1, start: number, count: number, spaceBefore: number): void {
  const page = context.currentPage
  const section = context.currentSection
  const column = currentColumn(context)
  if (!page || !section || !column || count <= 0 || context.refused) return
  if (context.sliceCount >= DOCX_PAGINATION_LIMITS.maxParagraphSlices || context.linePlacementCount + count > DOCX_PAGINATION_LIMITS.maxLinePlacements) {
    refuse(context, 'resource-limit', paragraph.paragraph_id, 'Pagination placement output exceeds its bounded slice/line budget')
    return
  }
  const topRelative = checkedSum(context.cursorY, spaceBefore)
  if (topRelative === undefined) {
    refuse(context, 'resource-limit', paragraph.paragraph_id, 'Paragraph placement coordinate exceeds the bounded integer range')
    return
  }
  let lineY = topRelative
  const placed: NativeDocxPlacedLineV1[] = []
  for (let offset = 0; offset < count; offset += 1) {
    const line = paragraph.lines[start + offset]!
    if (!lineGeometryValid(context, paragraph, line)) return
    const x = checkedSum(column.x_millipoints, line.inline_offset_millipoints)
    const y = checkedSum(column.y_millipoints, lineY)
    if (x === undefined || y === undefined) {
      refuse(context, 'resource-limit', line.id, 'Line placement coordinate exceeds the bounded integer range')
      return
    }
    placed.push({
      id: `placed:${line.id}`,
      line_id: line.id,
      paragraph_id: paragraph.paragraph_id,
      section_id: section.id,
      column_id: column.id,
      column_ordinal: column.ordinal,
      source_line_ordinal: line.ordinal,
      x_millipoints: x,
      y_millipoints: y,
      width_millipoints: line.advance_inline_millipoints,
      height_millipoints: line.line_height_millipoints,
    })
    const nextY = checkedSum(lineY, line.line_height_millipoints)
    if (nextY === undefined) {
      refuse(context, 'resource-limit', line.id, 'Line block coordinate exceeds the bounded integer range')
      return
    }
    lineY = nextY
  }
  const height = lineY - topRelative
  const sliceOrdinal = context.sliceCountForParagraph?.get(paragraph.paragraph_id) ?? 0
  context.sliceCountForParagraph?.set(paragraph.paragraph_id, sliceOrdinal + 1)
  const previousLocation = context.lastSliceLocation.get(paragraph.paragraph_id)
  const hasContinuation = start > 0
  const continues = start + count < paragraph.lines.length
  const crossesPageNext = continues && column.ordinal + 1 >= section.page.columns
  const slice: NativeDocxParagraphSliceV1 = {
    id: `slice:${paragraph.paragraph_id}:${sliceOrdinal}`,
    paragraph_id: paragraph.paragraph_id,
    section_id: section.id,
    column_id: column.id,
    column_ordinal: column.ordinal,
    slice_ordinal: sliceOrdinal,
    first_line_ordinal: start,
    last_line_ordinal: start + count - 1,
    line_ids: placed.map((entry) => entry.line_id),
    top_millipoints: checkedSum(column.y_millipoints, topRelative)!,
    height_millipoints: height,
    space_before_millipoints: spaceBefore,
    continued_from_previous_page: hasContinuation && previousLocation?.pageOrdinal !== page.ordinal,
    continues_on_next_page: crossesPageNext,
    continued_from_previous_column: hasContinuation && previousLocation?.pageOrdinal === page.ordinal && previousLocation.columnOrdinal !== column.ordinal,
    continues_in_next_column: continues && !crossesPageNext,
  }
  page.paragraph_slices.push(slice)
  page.lines.push(...placed)
  context.sliceCount += 1
  context.linePlacementCount += placed.length
  context.cursorY = lineY
  context.lastSliceLocation.set(paragraph.paragraph_id, { pageOrdinal: page.ordinal, columnOrdinal: column.ordinal })
}

// Per-paragraph slice ordinals stay stable even when a preceding paragraph repaginates.
interface PaginationContext { sliceCountForParagraph?: Map<string, number> }

function sumLineHeights(lines: readonly NativeDocxShapedLineV1[], start = 0, end = lines.length): number | undefined {
  let result = 0
  for (let index = start; index < end; index += 1) {
    const next = checkedSum(result, lines[index]!.line_height_millipoints)
    if (next === undefined) return undefined
    result = next
  }
  return result
}

function maxLinesThatFit(lines: readonly NativeDocxShapedLineV1[], start: number, available: number): number {
  let height = 0
  let count = 0
  for (let index = start; index < lines.length; index += 1) {
    const next = checkedSum(height, lines[index]!.line_height_millipoints)
    if (next === undefined || next > available) break
    height = next
    count += 1
  }
  return count
}

function paginateParagraph(context: PaginationContext, paragraph: NativeDocxShapedParagraphV1, resolved: NativeDocxResolvedParagraphV1): void {
  if (!context.currentPage || context.refused) return
  if (resolved.properties.page_break_before && pageHasContent(context)) startNextContentPage(context)
  if (!context.currentPage || context.refused) return
  const totalHeight = sumLineHeights(paragraph.lines)
  if (totalHeight === undefined) {
    refuse(context, 'resource-limit', paragraph.paragraph_id, 'Paragraph line-height sum exceeds the bounded integer range')
    return
  }
  const initialGap = paragraphGap(context, paragraph, false)
  const keepLines = resolved.properties.keep_lines === true
  if (keepLines) {
    const required = checkedSum(initialGap, totalHeight)
    if (required === undefined || totalHeight > (currentColumn(context)?.height_millipoints ?? 0)) {
      refuse(context, 'keep-lines-unsatisfiable', paragraph.paragraph_id, 'keep_lines paragraph cannot fit in one section column')
      return
    }
    if (required > remainingHeight(context)) {
      if (columnHasContent(context) || (context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 && initialGap > 0)) startNextFlowColumn(context)
      else {
        refuse(context, 'keep-lines-unsatisfiable', paragraph.paragraph_id, 'keep_lines paragraph cannot fit in one section column')
        return
      }
    }
    const gap = paragraphGap(context, paragraph, false)
    placeSlice(context, paragraph, 0, paragraph.lines.length, gap)
    context.previousAfter = paragraph.spacing_after_millipoints
    return
  }
  let start = 0
  while (start < paragraph.lines.length && !context.refused) {
    const continuation = start > 0
    const gap = paragraphGap(context, paragraph, continuation)
    const available = remainingHeight(context) - gap
    let fit = maxLinesThatFit(paragraph.lines, start, available)
    if (fit === 0) {
      if (columnHasContent(context)) {
        startNextFlowColumn(context)
        continue
      }
      if (context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 && gap > 0) {
        startNextFlowColumn(context)
        continue
      }
      refuse(context, 'line-geometry-invalid', paragraph.paragraph_id, 'The next shaped line plus required paragraph spacing cannot fit an empty section body box')
      return
    }
    const remainingLines = paragraph.lines.length - start
    if (fit < remainingLines && (resolved.properties.widow_control ?? true)) {
      if (fit === 1) {
        if (columnHasContent(context)) {
          startNextFlowColumn(context)
          continue
        }
        if (context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 && gap > 0) {
          startNextFlowColumn(context)
          continue
        }
        refuse(context, 'widow-control-unsatisfiable', paragraph.paragraph_id, 'Widow/orphan control cannot place at least two lines in an empty section body box')
        return
      }
      if (remainingLines - fit === 1) fit -= 1
      if (fit < 2 || remainingLines - fit < 2) {
        if (!columnHasContent(context) && context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 && gap > 0) {
          startNextFlowColumn(context)
          continue
        }
        refuse(context, 'widow-control-unsatisfiable', paragraph.paragraph_id, 'Widow/orphan control has no valid cluster-safe page split')
        return
      }
    }
    placeSlice(context, paragraph, start, fit, gap)
    start += fit
    if (start < paragraph.lines.length) startNextFlowColumn(context)
  }
  context.previousAfter = paragraph.spacing_after_millipoints
}

function placeTableRow(context: PaginationContext, table: NativeDocxQualifiedTableV1, geometry: NativeDocxTableRowGeometryV1, shaped: Map<string, NativeDocxShapedParagraphV1>, repeated = false): void {
  const page = context.currentPage
  const column = currentColumn(context)
  const section = context.currentSection
  if (!page || !column || !section || context.refused) {
    refuse(context, 'line-geometry-invalid', table.table.id, 'Qualified table row geometry could not be derived from exact shaped cell paragraphs')
    return
  }
  if (table.x_millipoints + table.width_millipoints > column.width_millipoints || geometry.height_millipoints > column.height_millipoints) {
    refuse(context, 'line-geometry-invalid', geometry.row_id, 'Table row exceeds the exact section column')
    return
  }
  if (geometry.height_millipoints > remainingHeight(context)) startNextFlowColumn(context)
  const target = context.currentPage
  const targetColumn = currentColumn(context)
  if (!target || !targetColumn || geometry.height_millipoints > remainingHeight(context)) {
    refuse(context, 'line-geometry-invalid', geometry.row_id, 'Indivisible table row cannot fit on an empty section column')
    return
  }
  const rowTop = context.cursorY
  for (const [cellIndex, sourceCell] of table.rows[geometry.row_ordinal]!.cells.entries()) {
    const cell = geometry.cells[cellIndex]!
    let localY = cell.content_y_millipoints
    let previousAfter = 0
    for (const [paragraphIndex, sourceParagraph] of sourceCell.cell.paragraphs.entries()) {
      const paragraph = shaped.get(sourceParagraph.id)
      if (!paragraph) { refuse(context, 'shaped-paragraph-missing', sourceParagraph.id, 'Qualified table cell paragraph has no shaped lines'); return }
      if (context.sliceCount >= DOCX_PAGINATION_LIMITS.maxParagraphSlices || context.linePlacementCount + paragraph.lines.length > DOCX_PAGINATION_LIMITS.maxLinePlacements) {
        refuse(context, 'resource-limit', sourceParagraph.id, 'Table paragraph placement exceeds the bounded slice/line budget')
        return
      }
      const gap = paragraphIndex === 0 ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)
      localY += gap
      const placed: NativeDocxPlacedLineV1[] = []
      for (const line of paragraph.lines) {
        const x = checkedSum(targetColumn.x_millipoints, cell.content_x_millipoints, line.inline_offset_millipoints)
        const y = checkedSum(targetColumn.y_millipoints, rowTop, localY)
        if (x === undefined || y === undefined || line.available_width_millipoints !== cell.content_width_millipoints || line.advance_inline_millipoints + line.inline_offset_millipoints > cell.content_width_millipoints || localY + line.line_height_millipoints > cell.height_millipoints - cell.content_y_millipoints) {
          refuse(context, 'line-geometry-invalid', line.id, 'Cell line escapes its exact qualified content box')
          return
        }
        placed.push({ table_cell_id: cell.cell_id, ...(repeated ? { repeated_table_header: true as const } : {}), id: repeated ? `placed:${line.id}:table-header:${target.id}` : `placed:${line.id}`, line_id: line.id, paragraph_id: paragraph.paragraph_id, section_id: section.id, column_id: targetColumn.id, column_ordinal: targetColumn.ordinal, source_line_ordinal: line.ordinal, x_millipoints: x, y_millipoints: y, width_millipoints: line.advance_inline_millipoints, height_millipoints: line.line_height_millipoints })
        localY += line.line_height_millipoints
      }
      const sliceOrdinal = context.sliceCountForParagraph?.get(paragraph.paragraph_id) ?? 0
      context.sliceCountForParagraph?.set(paragraph.paragraph_id, sliceOrdinal + 1)
      target.paragraph_slices.push({
        table_cell_id: cell.cell_id,
        ...(repeated ? { repeated_table_header: true as const } : {}),
        id: `slice:${paragraph.paragraph_id}:${sliceOrdinal}`, paragraph_id: paragraph.paragraph_id, section_id: section.id, column_id: targetColumn.id, column_ordinal: targetColumn.ordinal, slice_ordinal: sliceOrdinal,
        first_line_ordinal: 0, last_line_ordinal: paragraph.lines.length - 1, line_ids: placed.map((line) => line.line_id),
        top_millipoints: placed[0]?.y_millipoints ?? checkedSum(target.body_box.y_millipoints, rowTop, localY)!,
        height_millipoints: placed.reduce((sum, line) => sum + line.height_millipoints, 0), space_before_millipoints: gap,
        continued_from_previous_page: false, continues_on_next_page: false, continued_from_previous_column: false, continues_in_next_column: false,
      })
      target.lines.push(...placed)
      context.sliceCount += 1
      context.linePlacementCount += placed.length
      context.lastSliceLocation.set(paragraph.paragraph_id, { pageOrdinal: target.ordinal, columnOrdinal: targetColumn.ordinal })
      previousAfter = paragraph.spacing_after_millipoints
    }
  }
  context.cursorY += geometry.height_millipoints
  context.previousAfter = 0
}

function paginateTable(context: PaginationContext, table: NativeDocxQualifiedTableV1, shaped: Map<string, NativeDocxShapedParagraphV1>): void {
  const rows = layoutNativeDocxTableRowsV1(table, context.request.shaped_lines)
  if (!rows) { refuse(context, 'line-geometry-invalid', table.table.id, 'Qualified table row geometry could not be derived from exact shaped cell paragraphs'); return }
  if (table.table.rows.some((row) => row.cant_split !== true)) { paginateSplittableTable(context, table, rows, shaped); return }
  let index = 0
  const headerCount = table.table.rows.findIndex((row) => !row.repeat_header)
  const headers = headerCount < 0 ? rows.length : headerCount
  const headerHeight = rows.slice(0, headers).reduce((sum, row) => sum + row.height_millipoints, 0)
  while (index < rows.length && !context.refused) {
    // Keep the initial header prefix with the first body row; never emit an
    // orphan header page or repeatedly retry a body row that cannot fit.
    const group = index === 0 && headers > 0 ? Math.min(rows.length, headers + 1) : nativeDocxTableRowGroupSizeV1(table, index)
    let groupHeight = 0
    for (let offset = 0; offset < group; offset += 1) {
      groupHeight = checkedSum(groupHeight, rows[index + offset]!.height_millipoints) ?? Number.MAX_SAFE_INTEGER
    }
    const column = currentColumn(context)
    if (!column || groupHeight > column.height_millipoints) {
      refuse(context, 'line-geometry-invalid', rows[index]!.row_id, 'Merged table rows exceed the exact section column')
      return
    }
    if (groupHeight > remainingHeight(context)) {
      if (index > 0 && headers > 0 && headerHeight + groupHeight > column.height_millipoints) {
        refuse(context, 'line-geometry-invalid', rows[index]!.row_id, 'Repeated headers and the next indivisible row cannot fit on an empty page')
        return
      }
      startNextFlowColumn(context)
      if (index > 0) for (let header = 0; header < headers && !context.refused; header += 1) placeTableRow(context, table, rows[header]!, shaped, true)
    }
    if (!currentColumn(context) || groupHeight > remainingHeight(context)) {
      refuse(context, 'line-geometry-invalid', rows[index]!.row_id, 'Indivisible merged table rows cannot fit on an empty section column')
      return
    }
    for (let offset = 0; offset < group && !context.refused; offset += 1) placeTableRow(context, table, rows[index + offset]!, shaped)
    index += group
  }
}

function placeRowFragment(context: PaginationContext, table: NativeDocxQualifiedTableV1, row: NativeDocxTableRowGeometryV1, plan: NativeDocxRowBreakPlanV1, start: number, end: number, ordinal: number, previousSlices: Map<string, NativeDocxParagraphSliceV1>): void {
  const page = context.currentPage, column = currentColumn(context), section = context.currentSection
  if (!page || !column || !section || end <= start || end-start > remainingHeight(context)) { refuse(context,'line-geometry-invalid',row.row_id,'Invalid bounded table row fragment'); return }
  const pageY = column.y_millipoints + context.cursorY
  // Query only lines intersecting this fragment, rather than rescanning every
  // cell paragraph for each continuation page. Restore source paragraph order
  // inside the fragment after querying the vertical event index.
  let low = 0, high = plan.events.length
  while (low < high) { const middle = Math.floor((low + high) / 2); if (plan.events[middle]!.y < start) low = middle + 1; else high = middle }
  const selections = new Map<number, Array<{ line: NativeDocxShapedParagraphV1['lines'][number]; y: number }>>()
  for (let index = low; index < plan.events.length && plan.events[index]!.y < end; index += 1) {
    const event = plan.events[index]!, line = plan.paragraphs[event.paragraph]!.paragraph.lines[event.line]!
    if (event.y + line.line_height_millipoints > end) { refuse(context, 'line-geometry-invalid', row.row_id, 'Row cut bisects a cell line'); return }
    const selected = selections.get(event.paragraph) ?? []
    selected.push({ line, y: event.y }); selections.set(event.paragraph, selected)
  }
  for (const [paragraphIndex, selected] of [...selections].sort(([a], [b]) => a - b)) {
    const entry = plan.paragraphs[paragraphIndex]!
    const paragraph = entry.paragraph
    if (!selected.length) continue
    if (context.sliceCount >= DOCX_PAGINATION_LIMITS.maxParagraphSlices || context.linePlacementCount + selected.length > DOCX_PAGINATION_LIMITS.maxLinePlacements) { refuse(context,'resource-limit',row.row_id,'Split row exceeds the bounded paragraph/line placement budget'); return }
    const placed: NativeDocxPlacedLineV1[] = []
    for (const { line,y } of selected) {
      if (line.available_width_millipoints !== entry.content_width || line.inline_offset_millipoints + line.advance_inline_millipoints > entry.content_width) { refuse(context,'line-geometry-invalid',line.id,'Split cell line escapes the qualified content width'); return }
      placed.push({ id:`placed:${line.id}`,line_id:line.id,paragraph_id:paragraph.paragraph_id,table_cell_id:entry.cell_id,section_id:section.id,column_id:column.id,column_ordinal:column.ordinal,source_line_ordinal:line.ordinal,x_millipoints:column.x_millipoints+entry.content_x+line.inline_offset_millipoints,y_millipoints:pageY+y-start,width_millipoints:line.advance_inline_millipoints,height_millipoints:line.line_height_millipoints })
    }
    const sliceOrdinal = context.sliceCountForParagraph?.get(paragraph.paragraph_id) ?? 0
    const previous = previousSlices.get(paragraph.paragraph_id)
    if (previous) previous.continues_on_next_page = true
    const slice: NativeDocxParagraphSliceV1 = { id:`slice:${paragraph.paragraph_id}:${sliceOrdinal}`,paragraph_id:paragraph.paragraph_id,table_cell_id:entry.cell_id,section_id:section.id,column_id:column.id,column_ordinal:column.ordinal,slice_ordinal:sliceOrdinal,first_line_ordinal:selected[0]!.line.ordinal,last_line_ordinal:selected.at(-1)!.line.ordinal,line_ids:placed.map(line=>line.line_id),top_millipoints:placed[0]!.y_millipoints,height_millipoints:placed.reduce((sum,line)=>sum+line.height_millipoints,0),space_before_millipoints:0,continued_from_previous_page:previous!==undefined,continues_on_next_page:false,continued_from_previous_column:false,continues_in_next_column:false }
    page.paragraph_slices.push(slice); page.lines.push(...placed)
    previousSlices.set(paragraph.paragraph_id,slice)
    context.sliceCountForParagraph?.set(paragraph.paragraph_id,sliceOrdinal+1)
    context.sliceCount += 1; context.linePlacementCount += placed.length
  }
  ;(page.table_rows ??= []).push({ id:`table-row:${table.table.id}:${row.row_id}:${ordinal}`,table_id:table.table.id,row_id:row.row_id,row_ordinal:row.row_ordinal,fragment_ordinal:ordinal,section_id:section.id,column_id:column.id,column_ordinal:column.ordinal,x_millipoints:column.x_millipoints+table.x_millipoints,y_millipoints:pageY,width_millipoints:table.width_millipoints,height_millipoints:end-start,source_y_millipoints:start,source_height_millipoints:row.height_millipoints })
  context.cursorY += end-start; context.previousAfter = 0
}

function paginateSplittableTable(context: PaginationContext, table: NativeDocxQualifiedTableV1, rows: NativeDocxTableRowGeometryV1[], shaped: Map<string, NativeDocxShapedParagraphV1>): void {
  const resolved = new Map(context.request.resolved_layout.paragraphs.map(paragraph=>[paragraph.paragraph_id,paragraph]))
  const plans = rows.map(row=>table.table.rows[row.row_ordinal]!.cant_split === true ? undefined : nativeDocxRowBreakPlanV1(table,row,shaped,resolved))
  if (rows.some((row,index)=>table.table.rows[row.row_ordinal]!.cant_split !== true && !plans[index])) { refuse(context,'line-geometry-invalid',table.table.id,'Natural row cannot derive bounded line-safe fragmentation'); return }
  const firstBody = table.table.rows.findIndex(row=>!row.repeat_header), headers = firstBody < 0 ? rows.length : firstBody
  const headerHeight = rows.slice(0,headers).reduce((sum,row)=>sum+row.height_millipoints,0)
  const freshPage = () => {
    startNextFlowColumn(context)
    for(let index=0;index<headers&&!context.refused;index+=1)placeTableRow(context,table,rows[index]!,shaped,true)
  }
  const firstMinimum = plans[headers]?.protected[0]?.end ?? rows[headers]?.height_millipoints ?? 0
  // Never strand an initial/repeated heading without the next complete source
  // line group. A group taller than the remaining body refuses atomically.
  const column = currentColumn(context)
  if (!column || table.x_millipoints+table.width_millipoints>column.width_millipoints || headerHeight+firstMinimum>column.height_millipoints) { refuse(context,'line-geometry-invalid',table.table.id,'Header prefix and first indivisible cell line group cannot fit on an empty page'); return }
  if(headerHeight+firstMinimum>remainingHeight(context))startNextFlowColumn(context)
  for(let index=0;index<headers&&!context.refused;index+=1)placeTableRow(context,table,rows[index]!,shaped)
  for(let index=headers;index<rows.length&&!context.refused;index+=1) {
    const row=rows[index]!,plan=plans[index]
    if(!plan) {
      if(row.height_millipoints>remainingHeight(context))freshPage()
      if(context.refused)return
      if(row.height_millipoints>remainingHeight(context)){refuse(context,'line-geometry-invalid',row.row_id,'Indivisible row cannot fit after repeated headers');return}
      placeTableRow(context,table,row,shaped);continue
    }
    let start=0,ordinal=0
    const previousSlices=new Map<string,NativeDocxParagraphSliceV1>()
    while(start<plan.height&&!context.refused) {
      let end=nativeDocxRowCutV1(plan,start,remainingHeight(context))
      if(end===start){freshPage();if(context.refused)return;end=nativeDocxRowCutV1(plan,start,remainingHeight(context))}
      if(end===start){refuse(context,'line-geometry-invalid',row.row_id,'Cell line/keep/widow group cannot fit after repeated headers');return}
      placeRowFragment(context,table,row,plan,start,end,ordinal,previousSlices)
      start=end;ordinal+=1
      if(start<plan.height&&!context.refused)freshPage()
    }
  }
}

interface KeepChainPlan {
  ends: number[]
  contentHeights: Array<number | undefined>
  pageBreakConflicts: Array<number | undefined>
  atomic: boolean[]
}

// One reverse pass makes every keep-chain query O(1). Each paragraph and
// shaped line contributes to at most one planning step, so a maximal hostile
// keep_next chain cannot trigger the former quadratic forward rescans.
function planKeepChains(paragraphs: readonly NativeDocxParagraphV1[], resolved: Map<string, NativeDocxResolvedParagraphV1>, shaped: Map<string, NativeDocxShapedParagraphV1>): KeepChainPlan {
  const ends = new Array<number>(paragraphs.length)
  const contentHeights = new Array<number | undefined>(paragraphs.length)
  const pageBreakConflicts = new Array<number | undefined>(paragraphs.length)
  const atomic = new Array<boolean>(paragraphs.length)
  const lineHeights = new Array<number | undefined>(paragraphs.length)
  for (let index = 0; index < paragraphs.length; index += 1) {
    const paragraph = shaped.get(paragraphs[index]!.id)
    lineHeights[index] = paragraph ? sumLineHeights(paragraph.lines) : undefined
  }
  for (let index = paragraphs.length - 1; index >= 0; index -= 1) {
    const keepsNext = index + 1 < paragraphs.length && resolved.get(paragraphs[index]!.id)?.properties.keep_next === true
    if (!keepsNext) {
      ends[index] = index
      contentHeights[index] = lineHeights[index]
      pageBreakConflicts[index] = undefined
      atomic[index] = (shaped.get(paragraphs[index]!.id)?.lines.length ?? 0) === 1 || resolved.get(paragraphs[index]!.id)?.properties.keep_lines === true
      continue
    }
    ends[index] = ends[index + 1]!
    pageBreakConflicts[index] = resolved.get(paragraphs[index + 1]!.id)?.properties.page_break_before === true ? index + 1 : pageBreakConflicts[index + 1]
    atomic[index] = ((shaped.get(paragraphs[index]!.id)?.lines.length ?? 0) === 1 || resolved.get(paragraphs[index]!.id)?.properties.keep_lines === true) && atomic[index + 1] === true
    const paragraph = shaped.get(paragraphs[index]!.id)
    const nextParagraph = shaped.get(paragraphs[index + 1]!.id)
    const lineHeight = lineHeights[index]
    const tail = contentHeights[index + 1]
    contentHeights[index] = paragraph && nextParagraph && lineHeight !== undefined && tail !== undefined
      ? checkedSum(lineHeight, Math.max(paragraph.spacing_after_millipoints, nextParagraph.spacing_before_millipoints), tail)
      : undefined
  }
  return { ends, contentHeights, pageBreakConflicts, atomic }
}

function validateAndIndexParagraphs(context: PaginationContext, groups: readonly SectionGroup[]): {
  resolved: Map<string, NativeDocxResolvedParagraphV1>
  shaped: Map<string, NativeDocxShapedParagraphV1>
} {
  const resolved = new Map(context.request.resolved_layout.paragraphs.map((entry) => [entry.paragraph_id, entry]))
  const shaped = new Map(context.request.shaped_lines.paragraphs.map((entry) => [entry.paragraph_id, entry]))
  const bodyIDs = new Set(bodyParagraphs(context.request.document).map((entry) => entry.id))
  for (const paragraph of context.request.shaped_lines.paragraphs) if (paragraph.story_kind === 'body' && !bodyIDs.has(paragraph.paragraph_id)) {
    refuse(context, 'shaped-paragraph-mismatch', paragraph.paragraph_id, 'Shaped body paragraph does not exist in the native body story')
  }
  for (const group of groups) for (const paragraph of group.blocks.flatMap(blockParagraphs)) {
    const resolvedParagraph = resolved.get(paragraph.id)
    const shapedParagraph = shaped.get(paragraph.id)
    if (!resolvedParagraph) refuse(context, 'shaped-paragraph-missing', paragraph.id, 'Native body paragraph has no resolved-layout paragraph')
    if (!shapedParagraph) refuse(context, 'shaped-paragraph-missing', paragraph.id, 'Native body paragraph has no shaped-lines paragraph')
    if (shapedParagraph && (shapedParagraph.story_kind !== 'body' || shapedParagraph.story_id !== context.request.document.body.id)) {
      refuse(context, 'shaped-paragraph-mismatch', paragraph.id, 'Shaped body paragraph has the wrong story identity')
    }
  }
  return { resolved, shaped }
}

/**
 * Word normally balances a section's terminal multi-column fragment. V1 only
 * accepts the uniquely provable subset: one uniform-height, zero-spacing line
 * per paragraph. The quotient/remainder distribution puts at most one extra
 * line in each earlier column, yielding a unique minimum-height spread.
 */
function paginateExactlyBalancedGroup(
  context: PaginationContext,
  group: SectionGroup,
  resolved: Map<string, NativeDocxResolvedParagraphV1>,
  shaped: Map<string, NativeDocxShapedParagraphV1>,
): void {
  const entries = group.blocks.flatMap((block) => block.paragraph ? [{ native: block.paragraph, resolved: resolved.get(block.paragraph.id)!, shaped: shaped.get(block.paragraph.id)! }] : [])
  const lineHeight = entries[0]?.shaped.lines[0]?.line_height_millipoints
  const exact = lineHeight !== undefined && lineHeight > 0 && entries.every((entry) =>
    entry.shaped.lines.length === 1 && entry.shaped.lines[0]!.line_height_millipoints === lineHeight &&
    entry.shaped.spacing_before_millipoints === 0 && entry.shaped.spacing_after_millipoints === 0 &&
    entry.resolved.properties.keep_next !== true && entry.resolved.properties.page_break_before !== true)
  const column = currentColumn(context)
  const capacity = lineHeight === undefined || !column ? 0 : Math.floor(column.height_millipoints / lineHeight)
  if (!exact || capacity < 1) {
    refuse(context, 'column-balance-ambiguous', group.section.id, 'Multi-column balancing is exact only for uniform one-line, zero-spacing paragraphs without cross-paragraph break constraints')
    return
  }
  const columns = group.section.page.columns
  const firstColumns = columns - context.currentColumnOrdinal
  let index = 0
  let firstPage = true
  while (index < entries.length && !context.refused) {
    const participating = firstPage ? firstColumns : columns
    const pageLines = Math.min(entries.length - index, participating * capacity)
    const baseQuota = Math.floor(pageLines / participating)
    const extraColumns = pageLines % participating
    for (let columnIndex = 0; columnIndex < participating && index < entries.length; columnIndex += 1) {
      const quota = baseQuota + (columnIndex < extraColumns ? 1 : 0)
      for (let offset = 0; offset < quota; offset += 1) {
        const entry = entries[index++]!
        placeSlice(context, entry.shaped, 0, 1, 0)
        context.previousAfter = 0
      }
      if (index < entries.length) startNextFlowColumn(context)
    }
    firstPage = false
  }
}

function paginateGroups(context: PaginationContext, groups: readonly SectionGroup[], resolved: Map<string, NativeDocxResolvedParagraphV1>, shaped: Map<string, NativeDocxShapedParagraphV1>): void {
  for (const [groupIndex, group] of groups.entries()) {
    sectionBodyBox(context, group.section)
    if (context.refused) return
    const finalParagraph = group.blocks.at(-1)?.paragraph
    if (groupIndex < groups.length - 1 && finalParagraph && resolved.get(finalParagraph.id)?.properties.keep_next) {
      refuse(context, 'keep-chain-conflict', finalParagraph.id, 'keep_next crosses a modeled section transition; v1 refuses rather than guessing whether Word moves content across the break')
      return
    }
    startSection(context, group.section, groupIndex === 0)
    if (context.refused) return
    const hasTable = group.blocks.some((block) => block.table !== undefined)
    if (hasTable && group.section.page.columns > 1) {
      refuse(context, 'body-table-unsupported', group.section.id, 'Table pagination is exact only in single-column sections; multi-column table geometry is outside pagination v1')
      return
    }
    const outgoingBreak = groups[groupIndex + 1]?.section.break_type
    if (group.section.page.columns > 1 && outgoingBreak !== 'next-column') {
      paginateExactlyBalancedGroup(context, group, resolved, shaped)
      continue
    }
    for (let blockIndex = 0; blockIndex < group.blocks.length && !context.refused;) {
      const block = group.blocks[blockIndex]!
      if (block.table) {
        const table = context.qualifiedTables?.get(block.table.id)
        if (!table) { refuse(context, 'body-table-unsupported', block.table.id, 'Table was not present in the qualified table inventory'); return }
        paginateTable(context, table, shaped)
        blockIndex += 1
        continue
      }
      const run: NativeDocxParagraphV1[] = []
      while (blockIndex < group.blocks.length && group.blocks[blockIndex]!.paragraph) run.push(group.blocks[blockIndex++]!.paragraph!)
      const keepPlan = planKeepChains(run, resolved, shaped)
      for (let index = 0; index < run.length && !context.refused; index += 1) {
        const nativeParagraph = run[index]!
        const resolvedParagraph = resolved.get(nativeParagraph.id)!
        const shapedParagraph = shaped.get(nativeParagraph.id)!
        const chainEnd = keepPlan.ends[index]!
        if (chainEnd > index) {
          if (!keepPlan.atomic[index]) {
            refuse(context, 'keep-chain-unsatisfiable', nativeParagraph.id, 'A keep_next chain contains a multiline paragraph that is not keep_lines; v1 refuses because keep-with-next constrains the boundary and requires boundary-aware split planning')
            return
          }
          const conflict = keepPlan.pageBreakConflicts[index]
          if (conflict !== undefined) refuse(context, 'keep-chain-conflict', run[conflict]!.id, 'page_break_before conflicts with an incoming keep_next chain')
          if (context.refused || !context.currentPage) return
          const contentHeight = keepPlan.contentHeights[index]
          const onPopulatedColumn = columnHasContent(context)
          const initialGap = onPopulatedColumn
            ? Math.max(context.previousAfter, shapedParagraph.spacing_before_millipoints)
            : context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 ? shapedParagraph.spacing_before_millipoints : 0
          const chainHeight = contentHeight === undefined ? undefined : checkedSum(initialGap, contentHeight)
          const freshChainHeight = contentHeight
          if (freshChainHeight === undefined || chainHeight === undefined) {
            refuse(context, 'resource-limit', nativeParagraph.id, 'keep_next chain height exceeds the bounded integer range')
            return
          }
          if (freshChainHeight > (currentColumn(context)?.height_millipoints ?? 0)) {
            refuse(context, 'keep-chain-unsatisfiable', nativeParagraph.id, 'keep_next chain cannot fit in one section column; v1 refuses rather than breaking the keep constraint')
            return
          }
          if (chainHeight > remainingHeight(context)) {
            if (onPopulatedColumn || (context.currentPage.section_page_ordinal === 0 && context.currentColumnOrdinal === 0 && initialGap > 0)) startNextFlowColumn(context)
            else {
              refuse(context, 'keep-chain-unsatisfiable', nativeParagraph.id, 'keep_next chain cannot fit in one section column')
              return
            }
          }
        }
        paginateParagraph(context, shapedParagraph, resolvedParagraph)
      }
    }
  }
}

function samePaginationWire(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left) && Array.isArray(right) && left.length === right.length && left.every((entry, index) => samePaginationWire(entry, right[index]))
  }
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') return false
  const leftRecord = left as Record<string, unknown>
  const rightRecord = right as Record<string, unknown>
  const leftKeys = Object.keys(leftRecord).sort()
  const rightKeys = Object.keys(rightRecord).sort()
  return leftKeys.length === rightKeys.length && leftKeys.every((key, index) => key === rightKeys[index] && samePaginationWire(leftRecord[key], rightRecord[key]))
}

function expectedSectionGeometry(section: NativeDocxSectionV1): { width: number; height: number; body: NativeDocxPageBodyBoxV1 } | undefined {
  const width = twips(section.page.width_twips)
  const height = twips(section.page.height_twips)
  const top = twips(section.page.margins.top_twips)
  const right = twips(section.page.margins.right_twips)
  const bottom = twips(section.page.margins.bottom_twips)
  const left = twips(section.page.margins.left_twips)
  const gutter = twips(section.page.margins.gutter_twips)
  if ([width, height, top, right, bottom, left, gutter].some((entry) => entry === undefined)) return undefined
  const x = checkedSum(left!, gutter!)
  const bodyWidth = checkedSum(width!, -left!, -right!, -gutter!)
  const bodyHeight = checkedSum(height!, -top!, -bottom!)
  if (x === undefined || bodyWidth === undefined || bodyHeight === undefined || bodyWidth <= 0 || bodyHeight <= 0) return undefined
  return { width: width!, height: height!, body: { x_millipoints: x, y_millipoints: top!, width_millipoints: bodyWidth, height_millipoints: bodyHeight } }
}

/** Deterministic core for an already decoded and joined pagination request. */
function paginateDecodedNativeDocxV1(request: NativeDocxPaginationRequestV1, approximateLegacySettings = false): NativeDocxPaginatedLayoutV1 {
  const context: PaginationContext = {
    approximateLegacySettings,
    request,
    provenance: provenance(request),
    diagnostics: [],
    diagnosticKeys: new Set(),
    refused: false,
    pages: [],
    sections: [],
    cursorY: 0,
    previousAfter: 0,
    sectionPageOrdinal: 0,
    currentColumnOrdinal: 0,
    sliceCount: 0,
    linePlacementCount: 0,
    sliceCountForParagraph: new Map(),
    lastSliceLocation: new Map(),
  }
  refuseUnsupportedSource(context)
  const groups = sectionGroups(context)
  const indexed = validateAndIndexParagraphs(context, groups)
  if (!context.refused) paginateGroups(context, groups, indexed.resolved, indexed.shaped)
  if (!context.refused) {
    const noteFailure = placeNativeDocxNotesV1({
      protocol: DOCX_PAGINATED_LAYOUT_PROTOCOL,
      version: DOCX_PAGINATED_LAYOUT_VERSION,
      status: 'paginated',
      provenance: context.provenance,
      diagnostics: context.diagnostics,
      sections: context.sections,
      pages: context.pages,
    }, request.document, request.resolved_layout, request.shaped_lines)
    if (noteFailure) refuse(context, noteFailure.code, noteFailure.scope_id, noteFailure.message)
  }
  context.diagnostics.sort((left, right) => compareNativeCodeUnits(left.scope_id, right.scope_id) || compareNativeCodeUnits(left.code, right.code) || compareNativeCodeUnits(left.message, right.message))
  if (context.refused) return {
    protocol: DOCX_PAGINATED_LAYOUT_PROTOCOL,
    version: DOCX_PAGINATED_LAYOUT_VERSION,
    status: 'refused',
    provenance: context.provenance,
    diagnostics: context.diagnostics,
    sections: [],
    pages: [],
  }
  return {
    protocol: DOCX_PAGINATED_LAYOUT_PROTOCOL,
    version: DOCX_PAGINATED_LAYOUT_VERSION,
    status: 'paginated',
    provenance: context.provenance,
    diagnostics: context.diagnostics,
    sections: context.sections,
    pages: context.pages,
  }
}

/**
 * Proves that a structurally valid page projection is complete for one exact
 * decoded pagination request. This is the source join that a standalone output
 * cannot infer from self-consistency alone.
 */
export function validateNativeDocxPaginatedLayoutSourceV1(output: NativeDocxPaginatedLayoutV1, request: NativeDocxPaginationRequestV1): NativeDocxValidationIssue[] {
  return validatePaginatedLayoutSource(output, request, false)
}

/** @internal Explicit current-policy source join; this never attests strict fidelity. */
export function validateNativeDocxApproximatePaginatedLayoutSourceV1(output: NativeDocxPaginatedLayoutV1, request: NativeDocxPaginationRequestV1, eligibilityValue: unknown): NativeDocxValidationIssue[] {
  const eligibility = decodeNativeDocxApproximationEligibilityV1(eligibilityValue, request.pagination_settings)
  if (eligibility.status !== 'eligible') return [issue('BROKEN_REFERENCE', '', 'ineligible source settings for current-policy layout')]
  return validatePaginatedLayoutSource(output, request, true)
}

function validatePaginatedLayoutSource(output: NativeDocxPaginatedLayoutV1, request: NativeDocxPaginationRequestV1, approximateLegacySettings: boolean): NativeDocxValidationIssue[] {
  const issues: NativeDocxValidationIssue[] = []
  const add = (code: NativeDocxValidationIssue['code'], path: string, message: string): void => {
    if (issues.length < DOCX_NATIVE_LIMITS.maxIssues) issues.push(issue(code, path, message))
  }
  const expectedOutput = paginateDecodedNativeDocxV1(request, approximateLegacySettings)
  if (!samePaginationWire(output, expectedOutput)) add('BROKEN_REFERENCE', '', 'paginated output must exactly equal the deterministic placement projection for this decoded request')
  if (!samePaginationWire(output.provenance, provenance(request))) add('BROKEN_REFERENCE', '/provenance', 'paginated provenance must exactly match the joined pagination request')
  if (output.status === 'refused') return issues

  const semanticContext: PaginationContext = {
    approximateLegacySettings,
    request, provenance: provenance(request), diagnostics: [], diagnosticKeys: new Set(), refused: false,
    pages: [], sections: [], cursorY: 0, previousAfter: 0, sectionPageOrdinal: 0, currentColumnOrdinal: 0,
    sliceCount: 0, linePlacementCount: 0, sliceCountForParagraph: new Map(), lastSliceLocation: new Map(),
  }
  refuseUnsupportedSource(semanticContext)
  if (semanticContext.refused) add('BROKEN_REFERENCE', '/status', 'paginated output cannot be complete for source semantics that require native pagination refusal')

  if (output.sections.length !== request.document.sections.length) add('BROKEN_REFERENCE', '/sections', 'paginated sections must exactly cover native source sections')
  const sourceSections = new Map(request.document.sections.map((section) => [section.id, section]))
  output.sections.forEach((section, index) => {
    const source = request.document.sections[index]
    if (!source || section.section_id !== source.id || section.starts_at_block_id !== source.starts_at_block_id || section.break_type !== source.break_type) add('BROKEN_REFERENCE', `/sections/${index}`, 'paginated section identity, start block, and break must match source order exactly')
    const geometry = source ? expectedSectionGeometry(source) : undefined
    const qualified = source ? qualifyNativeDocxSectionColumnsV1(source) : undefined
    if (source && (!geometry || !qualified?.ok || qualified.value.columns.some((column) => column.width_millipoints !== request.shaped_lines.available_width_millipoints))) add('BROKEN_REFERENCE', `/sections/${index}`, 'paginated source section must have exact bounded columns at the attested shaping width')
  })
  output.pages.forEach((page, index) => {
    const source = sourceSections.get(page.section_id)
    const geometry = source ? expectedSectionGeometry(source) : undefined
    if (!source || !geometry) {
      add('BROKEN_REFERENCE', `/pages/${index}/section_id`, 'page must derive from an exact source section with bounded geometry')
      return
    }
    if (page.width_millipoints !== geometry.width || page.height_millipoints !== geometry.height || !samePaginationWire(page.body_box, geometry.body)) add('BROKEN_REFERENCE', `/pages/${index}/body_box`, 'page size and body geometry must exactly derive from its source section')
    const expectedHeaders = page.kind === 'parity-blank' ? [] : source.header_refs
    const expectedFooters = page.kind === 'parity-blank' ? [] : source.footer_refs
    if (!samePaginationWire(page.header_refs, expectedHeaders) || !samePaginationWire(page.footer_refs, expectedFooters)) add('BROKEN_REFERENCE', `/pages/${index}/header_refs`, 'page header/footer provenance must exactly derive from its source section and page kind')
  })

  const nativeParagraphs = bodyParagraphs(request.document)
  const nativeIDs = new Set(nativeParagraphs.map((paragraph) => paragraph.id))
  const paragraphSections = new Map<string, string>()
  const blockIndexes = new Map(request.document.body.blocks.map((block, index) => [block.id, index]))
  request.document.sections.forEach((section, index) => {
    const start = blockIndexes.get(section.starts_at_block_id)
    const next = request.document.sections[index + 1]
    const end = next ? blockIndexes.get(next.starts_at_block_id) : request.document.body.blocks.length
    if (start === undefined || end === undefined) return
    for (const block of request.document.body.blocks.slice(start, end)) for (const paragraph of blockParagraphs(block)) paragraphSections.set(paragraph.id, section.id)
  })
  const resolved = new Map(request.resolved_layout.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const shaped = new Map(request.shaped_lines.paragraphs.filter((paragraph) => paragraph.story_kind === 'body').map((paragraph) => [paragraph.paragraph_id, paragraph]))
  for (const paragraph of nativeParagraphs) {
    if (!resolved.has(paragraph.id)) add('BROKEN_REFERENCE', '/resolved_layout/paragraphs', `body paragraph ${paragraph.id} is missing from resolved layout input`)
    if (!shaped.has(paragraph.id)) add('BROKEN_REFERENCE', '/shaped_lines/paragraphs', `body paragraph ${paragraph.id} is missing from shaped lines input`)
  }
  for (const paragraph of request.shaped_lines.paragraphs) if (paragraph.story_kind === 'body' && !nativeIDs.has(paragraph.paragraph_id)) add('BROKEN_REFERENCE', '/shaped_lines/paragraphs', `shaped body paragraph ${paragraph.paragraph_id} is outside the paginated native body sequence`)

  const expectedLines = nativeParagraphs.flatMap((paragraph) => {
    const shapedParagraph = shaped.get(paragraph.id)
    return shapedParagraph ? shapedParagraph.lines.map((line) => ({ paragraph, line, sectionID: paragraphSections.get(paragraph.id) })) : []
  })
  // Exact replay above validates every repeated placement. The independent
  // source coverage audit below counts original body lines only.
  const actualLines = output.pages.flatMap((page, pageIndex) => page.lines.filter((line) => !line.repeated_table_header).map((line) => ({ page, pageIndex, line })))
  const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
  const cellContentX = new Map<string, number>()
  if (qualified.status === 'qualified') for (const table of qualified.tables) for (const row of table.rows) for (const cell of row.cells) for (const paragraph of cell.cell.paragraphs) cellContentX.set(paragraph.id, cell.content_x_millipoints)
  if (actualLines.length !== expectedLines.length) add('BROKEN_REFERENCE', '/pages', 'placed lines must exactly cover every shaped body line once')
  const expectedByLineID = new Map(expectedLines.map(entry=>[entry.line.id,entry]))
  const seenSourceLines = new Set<string>()
  for (let index = 0; index < Math.max(actualLines.length, expectedLines.length) && issues.length < DOCX_NATIVE_LIMITS.maxIssues; index += 1) {
    const actual = actualLines[index]
    // Cell flows interleave across pages when a row fragments; exact replay
    // above checks placement order while this independent audit checks coverage.
    if (!actual) continue
    const expected = expectedByLineID.get(actual.line.line_id)
    if (!expected || seenSourceLines.has(actual.line.line_id)) { add('BROKEN_REFERENCE',`/pages/${actual.pageIndex}/lines`,'source body line is unknown or duplicated'); continue }
    seenSourceLines.add(actual.line.line_id)
    const column = actual.page.columns.find((candidate) => candidate.id === actual.line.column_id)
    const expectedX = column ? checkedSum(column.x_millipoints, cellContentX.get(expected.paragraph.id) ?? 0, expected.line.inline_offset_millipoints) : undefined
    if (!expected.sectionID || actual.line.section_id !== expected.sectionID || !column || column.section_id !== expected.sectionID || actual.line.column_ordinal !== column.ordinal || actual.line.id !== `placed:${expected.line.id}` || actual.line.line_id !== expected.line.id || actual.line.paragraph_id !== expected.paragraph.id || actual.line.source_line_ordinal !== expected.line.ordinal || actual.line.width_millipoints !== expected.line.advance_inline_millipoints || actual.line.height_millipoints !== expected.line.line_height_millipoints || expectedX === undefined || actual.line.x_millipoints !== expectedX) {
      add('BROKEN_REFERENCE', `/pages/${actual.pageIndex}/lines`, 'placed line identity, source order, section/column origin, width, and height must match the exact shaped body line')
    }
  }
  return issues
}

/**
 * Strictly validate and paginate one native DOCX body. Semantic uncertainty is
 * an `ok: true`, `status: refused` result with no pages; malformed wire input is
 * an `ok: false` validation result.
 */
export function paginateNativeDocxV1(value: unknown): PaginateNativeDocxV1Result {
  let decoded: ReturnType<typeof validateRequest>
  try {
    decoded = validateRequest(value)
  } catch {
    return { ok: false, issues: [issue('INVALID_VALUE', '', 'pagination request could not be safely inspected')] }
  }
  if (!decoded.ok || !('request' in decoded)) return decoded
  const outputValue = paginateDecodedNativeDocxV1(decoded.request)
  const sourceIssues = validateNativeDocxPaginatedLayoutSourceV1(outputValue, decoded.request)
  return sourceIssues.length > 0 ? { ok: false, issues: sourceIssues } : { ok: true, value: outputValue }
}

/** @internal A separate read-only projection; never a strict V1 attestation. */
export function paginateNativeDocxApproximateLegacyV1(value: unknown, eligibilityValue: unknown): { fidelity: 'approximate'; layout: NativeDocxPaginatedLayoutV1 } {
  const decoded = validateRequest(value)
  if (!decoded.ok || !('request' in decoded)) throw new TypeError('approximate pagination request failed strict structural/source validation')
  const eligibility = decodeNativeDocxApproximationEligibilityV1(eligibilityValue, decoded.request.pagination_settings)
  if (eligibility.status !== 'eligible') throw new TypeError('source settings are ineligible for approximate legacy pagination')
  const layout = paginateDecodedNativeDocxV1(decoded.request, true)
  if (validatePaginatedLayoutSource(layout, decoded.request, true).length) throw new TypeError('approximate pagination failed deterministic policy/source validation')
  return { fidelity: 'approximate', layout }
}
