/**
 * Exact, renderer-neutral header/footer selection and line placement.
 *
 * The Go native extractor remains relationship authority. This layer resolves
 * Word's per-kind section inheritance, selects title/even/default variants
 * from already-paginated physical pages, and places already-shaped static
 * story lines. Any ambiguity refuses the complete plan without partial pages.
 */

import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import type {
  NativeDocxDocumentV1,
  NativeDocxHeaderFooterReferenceV1,
  NativeDocxSectionV1,
  NativeDocxStoryV1,
} from './nativeContract.js'
import { DOCX_MAX_TWIPS_FOR_MILLIPOINTS } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1, NativeDocxShapedParagraphV1 } from './nativeShapingLines.js'
import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import type { NativeDocxPaginatedLayoutV1, NativeDocxPaginatedPageV1 } from './nativePaginationV1.js'
import { compareNativeCodeUnits } from './nativeDeterminism.js'
import { qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { validateNativeDocxPageFieldVariantsV1 } from './nativePageFieldsV1.js'

export const DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL = 'injoffice.docx.header-footer-layout'
export const DOCX_HEADER_FOOTER_LAYOUT_VERSION = 1 as const

export type NativeDocxHeaderFooterRegionV1 = 'header' | 'footer'

export type NativeDocxHeaderFooterDiagnosticCodeV1 =
  | 'relationship-ambiguous'
  | 'selected-story-missing'
  | 'selected-story-kind-mismatch'
  | 'selected-story-unsupported'
  | 'selected-story-table'
  | 'selected-story-shape'
  | 'selected-story-reference'
  | 'selected-story-field'
  | 'selected-story-diagnostic'
  | 'section-geometry-invalid'
  | 'selected-paragraph-missing'
  | 'selected-paragraph-unsupported'
  | 'selected-line-invalid'
  | 'header-band-overflow'
  | 'footer-band-overflow'
  | 'resource-limit'

export interface NativeDocxHeaderFooterDiagnosticV1 {
  code: NativeDocxHeaderFooterDiagnosticCodeV1
  severity: 'unsupported'
  scope_id: string
  message: string
}

export interface NativeDocxPlacedHeaderFooterLineV1 {
  id: string
  region: NativeDocxHeaderFooterRegionV1
  story_id: string
  relationship_id: string
  reference_kind: NativeDocxHeaderFooterReferenceV1['kind']
  line_id: string
  paragraph_id: string
  source_line_ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
}

export interface NativeDocxHeaderFooterPageLayoutV1 {
  page_id: string
  header_ref?: NativeDocxHeaderFooterReferenceV1
  footer_ref?: NativeDocxHeaderFooterReferenceV1
  lines: NativeDocxPlacedHeaderFooterLineV1[]
}

interface NativeDocxHeaderFooterLayoutBaseV1 {
  protocol: typeof DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL
  version: typeof DOCX_HEADER_FOOTER_LAYOUT_VERSION
  sha256: string
  diagnostics: NativeDocxHeaderFooterDiagnosticV1[]
}

export interface NativeDocxHeaderFooterLayoutSuccessV1 extends NativeDocxHeaderFooterLayoutBaseV1 {
  status: 'placed'
  pages: NativeDocxHeaderFooterPageLayoutV1[]
}

export interface NativeDocxHeaderFooterLayoutRefusedV1 extends NativeDocxHeaderFooterLayoutBaseV1 {
  status: 'refused'
  pages: []
}

export type NativeDocxHeaderFooterLayoutV1 = NativeDocxHeaderFooterLayoutSuccessV1 | NativeDocxHeaderFooterLayoutRefusedV1
type NativeDocxHeaderFooterLayoutHashInputV1 = NativeDocxHeaderFooterLayoutV1 extends infer T ? T extends NativeDocxHeaderFooterLayoutV1 ? Omit<T, 'sha256'> : never : never

export interface NativeDocxHeaderFooterLayoutInputV1 {
  document: NativeDocxDocumentV1
  resolved_layout: NativeDocxResolvedLayoutInputV1
  shaped_lines: NativeDocxShapedLinesV1
  pagination_settings: NativeDocxPaginationSettingsV1
  paginated_layout: NativeDocxPaginatedLayoutV1
  page_field_variants?: Array<{ page_id: string; shaped_lines: NativeDocxShapedLinesV1 }>
}

type VariantMap = Partial<Record<NativeDocxHeaderFooterReferenceV1['kind'], NativeDocxHeaderFooterReferenceV1>>
interface EffectiveSection { header: VariantMap; footer: VariantMap }

const MAX_PLACED_LINES = 100_000
const MAX_COORDINATE = 1_000_000_000_000

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalize)
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value as Record<string, unknown>).sort(compareNativeCodeUnits).map((key) => [key, canonicalize((value as Record<string, unknown>)[key])]))
  }
  return value
}

function canonicalHash(value: unknown): string {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(canonicalize(value)))))}`
}

/** Hashes the complete plan while excluding only its self-referential hash. */
export function nativeDocxHeaderFooterLayoutSha256V1(value: NativeDocxHeaderFooterLayoutHashInputV1 | NativeDocxHeaderFooterLayoutV1): string {
  const { sha256: _ignored, ...payload } = value as NativeDocxHeaderFooterLayoutV1
  return canonicalHash(payload)
}

function diagnostic(code: NativeDocxHeaderFooterDiagnosticCodeV1, scopeID: string, message: string): NativeDocxHeaderFooterDiagnosticV1 {
  return { code, severity: 'unsupported', scope_id: scopeID, message }
}

function safeSum(...values: number[]): number | undefined {
  let result = 0
  for (const value of values) {
    result += value
    if (!Number.isSafeInteger(result) || Math.abs(result) > MAX_COORDINATE) return undefined
  }
  return result
}

function twips(value: number): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || Math.abs(value) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS || Math.abs(value) > Math.floor(MAX_COORDINATE / 50)) return undefined
  return value * 50
}

function scopes(story: NativeDocxStoryV1): Set<string> {
  const result = new Set<string>([story.id])
  for (const block of story.blocks) {
    result.add(block.id)
    if (block.paragraph) {
      result.add(block.paragraph.id)
      for (const run of block.paragraph.runs) {
        result.add(run.id)
        if (run.drawing) result.add(run.drawing.id)
      }
    }
    if (block.table) {
      result.add(block.table.id)
      for (const row of block.table.rows) for (const cell of row.cells) for (const paragraph of cell.paragraphs) {
        result.add(row.id)
        result.add(cell.id)
        result.add(paragraph.id)
        for (const run of paragraph.runs) result.add(run.id)
      }
    }
  }
  return result
}

function effectiveSections(document: NativeDocxDocumentV1, diagnostics: NativeDocxHeaderFooterDiagnosticV1[]): Map<string, EffectiveSection> {
  const result = new Map<string, EffectiveSection>()
  const relationshipTargets = new Map<string, { region: NativeDocxHeaderFooterRegionV1; storyID: string }>()
  let previous: EffectiveSection = { header: {}, footer: {} }
  for (const section of document.sections) {
    const current: EffectiveSection = { header: { ...previous.header }, footer: { ...previous.footer } }
    for (const [region, references] of [['header', section.header_refs], ['footer', section.footer_refs]] as const) {
      for (const reference of references) {
        const known = relationshipTargets.get(reference.relationship_id)
        if (known && (known.region !== region || known.storyID !== reference.story_id)) diagnostics.push(diagnostic('relationship-ambiguous', section.id, `Relationship ${reference.relationship_id} is reused for inconsistent native header/footer targets`))
        else relationshipTargets.set(reference.relationship_id, { region, storyID: reference.story_id })
        current[region][reference.kind] = { ...reference }
      }
    }
    result.set(section.id, current)
    previous = current
  }
  return result
}

function selectedKind(section: NativeDocxSectionV1, page: NativeDocxPaginatedPageV1, settings: NativeDocxPaginationSettingsV1): NativeDocxHeaderFooterReferenceV1['kind'] {
  if (section.title_page && page.section_page_ordinal === 0) return 'first'
  if (settings.even_and_odd_headers && (page.ordinal + 1) % 2 === 0) return 'even'
  return 'default'
}

function selectedStory(document: NativeDocxDocumentV1, region: NativeDocxHeaderFooterRegionV1, reference: NativeDocxHeaderFooterReferenceV1, diagnostics: NativeDocxHeaderFooterDiagnosticV1[]): NativeDocxStoryV1 | undefined {
  const collection = region === 'header' ? document.headers : document.footers
  const story = collection.find((candidate) => candidate.id === reference.story_id)
  if (!story) diagnostics.push(diagnostic('selected-story-missing', reference.story_id, `Selected ${region} reference does not resolve to a modeled story`))
  else if (story.kind !== region) diagnostics.push(diagnostic('selected-story-kind-mismatch', story.id, `Selected ${region} reference resolves to a ${story.kind} story`))
  return story
}

function validateSelectedStory(input: NativeDocxHeaderFooterLayoutInputV1, story: NativeDocxStoryV1, diagnostics: NativeDocxHeaderFooterDiagnosticV1[]): void {
  const selectedScopes = scopes(story)
  for (const block of story.blocks) {
    if (block.kind === 'table') diagnostics.push(diagnostic('selected-story-table', block.id, 'Selected header/footer tables require native table grid layout and are refused'))
    if (block.paragraph) for (const run of block.paragraph.runs) {
      if (run.drawing) {
        const image = qualifyNativeDocxInlineImageV1(input.document, run.id, run.drawing)
        if (!image.ok) diagnostics.push(diagnostic('selected-story-shape', run.id, `Selected header/footer drawing is outside the exact inline raster subset: ${image.message}`))
      }
      if (run.reference) diagnostics.push(diagnostic('selected-story-reference', run.id, 'Selected header/footer note/comment references have no exact display advance'))
      if (run.control && run.control !== 'tab' && run.control !== 'line-break') diagnostics.push(diagnostic('selected-paragraph-unsupported', run.id, `Selected header/footer ${run.control} control is not supported`))
    }
  }
  for (const unsupported of input.document.unsupported) if (selectedScopes.has(unsupported.scope_id)) {
    const code = unsupported.code === 'FIELD_SEMANTICS' ? 'selected-story-field' : 'selected-story-unsupported'
    diagnostics.push(diagnostic(code, unsupported.scope_id, `Selected header/footer source is preserve-only: ${unsupported.code}: ${unsupported.message}`))
  }
  for (const entry of input.resolved_layout.diagnostics) if (selectedScopes.has(entry.scope_id)) diagnostics.push(diagnostic('selected-story-diagnostic', entry.scope_id, `Selected header/footer resolved layout is not exact: ${entry.code}: ${entry.message}`))
  for (const entry of input.shaped_lines.diagnostics) if (selectedScopes.has(entry.scope_id) || entry.source_id && selectedScopes.has(entry.source_id)) diagnostics.push(diagnostic('selected-story-diagnostic', entry.scope_id, `Selected header/footer shaping is not exact: ${entry.code}: ${entry.message}`))
}

function storyLineOffsets(story: NativeDocxStoryV1, shaped: Map<string, NativeDocxShapedParagraphV1>, resolved: NativeDocxResolvedLayoutInputV1, diagnostics: NativeDocxHeaderFooterDiagnosticV1[]): { lines: { paragraph: NativeDocxShapedParagraphV1; lineIndex: number; y: number }[]; height: number } | undefined {
  const result: { paragraph: NativeDocxShapedParagraphV1; lineIndex: number; y: number }[] = []
  const resolvedParagraphs = new Map(resolved.paragraphs.map((entry) => [entry.paragraph_id, entry]))
  let cursor = 0
  let previousAfter = 0
  let first = true
  for (const block of story.blocks) {
    if (block.kind !== 'paragraph' || !block.paragraph) continue
    const paragraph = shaped.get(block.paragraph.id)
    const properties = resolvedParagraphs.get(block.paragraph.id)?.properties
    if (!paragraph || paragraph.story_id !== story.id || paragraph.story_kind !== story.kind) {
      diagnostics.push(diagnostic('selected-paragraph-missing', block.paragraph.id, 'Selected native header/footer paragraph has no exact shaped paragraph'))
      continue
    }
    if (!properties || properties.bidi || properties.keep_next === true || properties.keep_lines === true || properties.page_break_before === true || paragraph.direction !== 'ltr' || paragraph.alignment === 'both' || paragraph.alignment === 'distribute') {
      diagnostics.push(diagnostic('selected-paragraph-unsupported', paragraph.paragraph_id, 'Selected header/footer paragraph uses pagination, bidi, or alignment semantics outside the exact static subset'))
      continue
    }
    const gap = first ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)
    const withGap = safeSum(cursor, gap)
    if (withGap === undefined) { diagnostics.push(diagnostic('resource-limit', paragraph.paragraph_id, 'Header/footer paragraph spacing exceeds the bounded coordinate range')); continue }
    cursor = withGap
    for (const [lineIndex, line] of paragraph.lines.entries()) {
      if (line.ordinal !== lineIndex || line.line_height_millipoints !== line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints || line.advance_inline_millipoints < 0) {
        diagnostics.push(diagnostic('selected-line-invalid', line.id, 'Selected header/footer line lacks exact natural LTR geometry'))
        continue
      }
      result.push({ paragraph, lineIndex, y: cursor })
      const next = safeSum(cursor, line.line_height_millipoints)
      if (next === undefined) { diagnostics.push(diagnostic('resource-limit', line.id, 'Header/footer line placement exceeds the bounded coordinate range')); continue }
      cursor = next
    }
    previousAfter = paragraph.spacing_after_millipoints
    first = false
  }
  const height = safeSum(cursor, previousAfter)
  if (height === undefined) diagnostics.push(diagnostic('resource-limit', story.id, 'Header/footer story height exceeds the bounded coordinate range'))
  return diagnostics.length === 0 && height !== undefined ? { lines: result, height } : undefined
}

function placeStory(input: NativeDocxHeaderFooterLayoutInputV1, page: NativeDocxPaginatedPageV1, section: NativeDocxSectionV1, region: NativeDocxHeaderFooterRegionV1, reference: NativeDocxHeaderFooterReferenceV1, story: NativeDocxStoryV1, diagnostics: NativeDocxHeaderFooterDiagnosticV1[]): NativeDocxPlacedHeaderFooterLineV1[] {
  validateSelectedStory(input, story, diagnostics)
  const shaped = new Map((input.page_field_variants?.find((variant) => variant.page_id === page.id)?.shaped_lines ?? input.shaped_lines).paragraphs.map((entry) => [entry.paragraph_id, entry]))
  const offsets = storyLineOffsets(story, shaped, input.resolved_layout, diagnostics)
  if (!offsets || diagnostics.length > 0) return []
  const storyHeight = offsets.height
  const distance = twips(region === 'header' ? section.page.margins.header_twips : section.page.margins.footer_twips)
  const bodyBottom = safeSum(page.body_box.y_millipoints, page.body_box.height_millipoints)
  if (distance === undefined || bodyBottom === undefined) { diagnostics.push(diagnostic('resource-limit', section.id, 'Header/footer margin conversion exceeds the bounded coordinate range')); return [] }
  const originY = region === 'header' ? distance : safeSum(page.height_millipoints, -distance, -storyHeight)
  if (originY === undefined) { diagnostics.push(diagnostic('resource-limit', story.id, 'Header/footer origin exceeds the bounded coordinate range')); return [] }
  const storyBottom = safeSum(originY, storyHeight)
  if (storyBottom === undefined) { diagnostics.push(diagnostic('resource-limit', story.id, 'Header/footer story bounds exceed the bounded coordinate range')); return [] }
  if (region === 'header' && storyBottom > page.body_box.y_millipoints) diagnostics.push(diagnostic('header-band-overflow', story.id, 'Selected header story does not fit between the exact header and top body margins'))
  if (region === 'footer' && originY < bodyBottom) diagnostics.push(diagnostic('footer-band-overflow', story.id, 'Selected footer story does not fit between the exact body bottom and footer margin'))
  if (diagnostics.length > 0) return []
  return offsets.lines.map(({ paragraph, lineIndex, y }) => {
    const line = paragraph.lines[lineIndex]!
    const x = safeSum(page.body_box.x_millipoints, line.inline_offset_millipoints)
    const absoluteY = safeSum(originY, y)
    const lineRight = x === undefined ? undefined : safeSum(x, line.advance_inline_millipoints)
    const bodyRight = safeSum(page.body_box.x_millipoints, page.body_box.width_millipoints)
    if (x === undefined || absoluteY === undefined || lineRight === undefined || bodyRight === undefined || x < page.body_box.x_millipoints || lineRight > bodyRight) {
      diagnostics.push(diagnostic('selected-line-invalid', line.id, 'Selected header/footer line escapes the exact horizontal story box'))
    }
    return {
      id: `placed:${region}:${page.id}:${line.id}`,
      region, story_id: story.id, relationship_id: reference.relationship_id, reference_kind: reference.kind,
      line_id: line.id, paragraph_id: paragraph.paragraph_id, source_line_ordinal: line.ordinal,
      x_millipoints: x ?? 0, y_millipoints: absoluteY ?? 0,
      width_millipoints: line.advance_inline_millipoints, height_millipoints: line.line_height_millipoints,
    }
  })
}

export function layoutNativeDocxHeadersFootersV1(input: NativeDocxHeaderFooterLayoutInputV1): NativeDocxHeaderFooterLayoutV1 {
  const diagnostics: NativeDocxHeaderFooterDiagnosticV1[] = []
  const pages: NativeDocxHeaderFooterPageLayoutV1[] = []
  try { validateNativeDocxPageFieldVariantsV1({ protocol: 'injoffice.docx.pagination-request', version: 1, document: input.document, resolved_layout: input.resolved_layout, shaped_lines: input.shaped_lines, pagination_settings: input.pagination_settings }, input.paginated_layout, input.page_field_variants) }
  catch (error) { diagnostics.push(diagnostic('selected-story-field', input.document.document_id, error instanceof Error ? error.message : 'Invalid page-field source')) }
  if (input.paginated_layout.status !== 'paginated') {
    diagnostics.push(diagnostic('selected-story-unsupported', input.document.document_id, 'Header/footer placement requires a complete paginated body'))
  } else {
    const effective = effectiveSections(input.document, diagnostics)
    const sections = new Map(input.document.sections.map((section) => [section.id, section]))
    let placedCount = 0
    for (const page of input.paginated_layout.pages) {
      if (page.kind === 'parity-blank') { pages.push({ page_id: page.id, lines: [] }); continue }
      const section = sections.get(page.section_id)
      const variants = effective.get(page.section_id)
      if (!section || !variants) { diagnostics.push(diagnostic('selected-story-missing', page.section_id, 'Paginated page has no exact native section')); continue }
      const sectionUnsupported = input.document.unsupported.filter((entry) => entry.scope_id === section.id && entry.capability === 'sections')
      for (const entry of sectionUnsupported) diagnostics.push(diagnostic('section-geometry-invalid', section.id, `Exact header/footer page geometry is unavailable: ${entry.code}: ${entry.message}`))
      if (section.page.orientation === 'portrait' ? section.page.width_twips > section.page.height_twips : section.page.width_twips < section.page.height_twips) diagnostics.push(diagnostic('section-geometry-invalid', section.id, 'Section orientation contradicts its exact page width and height'))
      const kind = selectedKind(section, page, input.pagination_settings)
      const pageLayout: NativeDocxHeaderFooterPageLayoutV1 = { page_id: page.id, lines: [] }
      for (const region of ['header', 'footer'] as const) {
        const reference = variants[region][kind]
        if (!reference) continue
        if (region === 'header') pageLayout.header_ref = { ...reference }
        else pageLayout.footer_ref = { ...reference }
        const story = selectedStory(input.document, region, reference, diagnostics)
        if (!story) continue
        const placed = placeStory(input, page, section, region, reference, story, diagnostics)
        placedCount += placed.length
        if (placedCount > MAX_PLACED_LINES) diagnostics.push(diagnostic('resource-limit', input.document.document_id, `Header/footer placements exceed ${MAX_PLACED_LINES}`))
        pageLayout.lines.push(...placed)
      }
      pages.push(pageLayout)
    }
  }
  diagnostics.sort((left, right) => compareNativeCodeUnits(left.scope_id, right.scope_id) || compareNativeCodeUnits(left.code, right.code) || compareNativeCodeUnits(left.message, right.message))
  const payload: NativeDocxHeaderFooterLayoutHashInputV1 = diagnostics.length > 0
    ? { protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL, version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, status: 'refused' as const, diagnostics, pages: [] as [] }
    : { protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL, version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, status: 'placed' as const, diagnostics, pages }
  return { ...payload, sha256: nativeDocxHeaderFooterLayoutSha256V1(payload) }
}
