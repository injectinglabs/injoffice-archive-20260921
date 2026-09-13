import type { NativeDocxFootnoteReservationProfileV1, NativeDocxFootnoteAreaMeasurementV1 } from './nativeFootnoteReservationV1.js'
/**
 * Exact bounded footnote/endnote placement over canonical native page output.
 *
 * This module has no package I/O or rendering authority. It admits whole notes and bounded paragraph-boundary
 * endnote continuation; any ambiguity returns
 * one refusal and leaves the caller responsible for discarding every page.
 */

import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxStoryV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1, NativeDocxShapedParagraphV1 } from './nativeShapingLines.js'
import { layoutNativeDocxTableRowsV1, qualifyNativeDocxTablesV1, type NativeDocxQualifiedTableV1 } from './nativeTablePagePaintV1.js'
import type {
  NativeDocxPaginatedLayoutSuccessV1,
  NativeDocxPaginatedPageV1,
  NativeDocxPageColumnV1,
  NativeDocxPlacedLineV1,
  NativeDocxPlacedNoteStoryV1,
} from './nativePaginationV1.js'

export const DOCX_NOTE_PAGINATION_LIMITS = {
  maxNotes: 10_000,
  maxNoteLines: 100_000,
} as const

export interface NativeDocxNotePaginationRefusalV1 {
  scope_id: string
  code: 'note-structure-unsupported' | 'note-reference-ambiguous' | 'note-separator-unsupported' | 'note-continuation-shaping-required' | 'note-overflow-unsupported' | 'resource-limit'
  message: string
}

type NoteKind = 'footnote' | 'endnote'

interface NoteReference {
  kind: NoteKind
  story: NativeDocxStoryV1
  runID: string
  number: number
  pageOrdinal: number
  sectionID: string
  columnID: string
  columnOrdinal: number
}

interface NoteLabel {
  paragraphID: string
  runID: string
}

interface PlacedLineIndex {
  page: NativeDocxPaginatedPageV1
  line: NativeDocxPlacedLineV1
}

interface PlacementBudget {
  lines: number
}

interface TableRowBottomIndexEntry {
  rowID: string
  contentOffsetMilliPoints: number
  heightMilliPoints: number
}

interface ContentBottomIndex {
  shapedByParagraph: ReadonlyMap<string, NativeDocxShapedParagraphV1>
  bodyParagraphIDs: ReadonlySet<string>
  tableRowByFirstParagraph: ReadonlyMap<string, TableRowBottomIndexEntry>
}

function bodyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  return document.body.blocks.flatMap((block) => block.kind === 'paragraph' && block.paragraph ? [block.paragraph] : [])
}

function storyParagraphs(story: NativeDocxStoryV1): NativeDocxParagraphV1[] | undefined {
  const result: NativeDocxParagraphV1[] = []
  for (const block of story.blocks) {
    if (block.kind !== 'paragraph' || !block.paragraph) return undefined
    result.push(block.paragraph)
  }
  return result
}

function noteScopes(document: NativeDocxDocumentV1): Set<string> {
  const scopes = new Set<string>()
  for (const story of document.notes) {
    // Continuation separators are dormant package sentinels until a note
    // actually crosses a page. The continuation path validates activated
    // scopes separately; dormant diagnostics cannot affect fitting notes.
    if (story.note_role === 'continuation-separator') continue
    scopes.add(story.id)
    for (const block of story.blocks) {
      scopes.add(block.id)
      if (block.paragraph) for (const run of block.paragraph.runs) scopes.add(run.id)
      if (block.table) scopes.add(block.table.id)
    }
  }
  return scopes
}

function addChecked(...values: number[]): number | undefined {
  let total = 0
  for (const value of values) {
    total += value
    if (!Number.isSafeInteger(total) || total < 0) return undefined
  }
  return total
}

function placedStory(
  story: NativeDocxStoryV1,
  shapedByParagraph: Map<string, NativeDocxShapedParagraphV1>,
  page: Pick<NativeDocxPaginatedPageV1, 'ordinal'>,
  column: NativeDocxPageColumnV1,
  top: number,
  ordinal: number,
  referenceRunID?: string,
  number?: number,
): NativeDocxPlacedNoteStoryV1 | NativeDocxNotePaginationRefusalV1 {
  const paragraphs = storyParagraphs(story)
  if (!paragraphs || paragraphs.length === 0) return { scope_id: story.id, code: 'note-structure-unsupported', message: 'Exact note placement requires one or more paragraph blocks and no nested tables' }
  const lines: NativeDocxPlacedLineV1[] = []
  let cursor = top
  let previousAfter = 0
  for (const paragraph of paragraphs) {
    const shaped = shapedByParagraph.get(paragraph.id)
    if (!shaped || shaped.story_id !== story.id || shaped.story_kind !== story.kind) return { scope_id: paragraph.id, code: 'note-structure-unsupported', message: 'Note paragraph is missing its exact shaped story projection' }
    if (shaped.lines.length === 0) return { scope_id: paragraph.id, code: 'note-structure-unsupported', message: 'Note paragraph has no exact shaped line placement' }
    const gap = lines.length === 0 ? shaped.spacing_before_millipoints : Math.max(previousAfter, shaped.spacing_before_millipoints)
    const afterGap = addChecked(cursor, gap)
    if (afterGap === undefined) return { scope_id: story.id, code: 'resource-limit', message: 'Note vertical geometry exceeds safe integer bounds' }
    cursor = afterGap
    for (const line of shaped.lines) {
      const x = addChecked(column.x_millipoints, line.inline_offset_millipoints)
      const right = x === undefined ? undefined : addChecked(x, line.advance_inline_millipoints)
      const bodyRight = addChecked(column.x_millipoints, column.width_millipoints)
      if (x === undefined || right === undefined || bodyRight === undefined) return { scope_id: line.id, code: 'resource-limit', message: 'Note horizontal geometry exceeds safe integer bounds' }
      if (line.available_width_millipoints > column.width_millipoints || x < column.x_millipoints || right > bodyRight || line.line_height_millipoints > column.height_millipoints) return { scope_id: line.id, code: 'note-structure-unsupported', message: 'Shaped note line escapes its exact page column geometry' }
      lines.push({
        id: `placed-note:${page.ordinal}:${ordinal}:${line.id}`,
        line_id: line.id,
        paragraph_id: paragraph.id,
        section_id: column.section_id,
        column_id: column.id,
        column_ordinal: column.ordinal,
        source_line_ordinal: line.ordinal,
        x_millipoints: x,
        y_millipoints: cursor,
        width_millipoints: line.advance_inline_millipoints,
        height_millipoints: line.line_height_millipoints,
      })
      const next = addChecked(cursor, line.line_height_millipoints)
      if (next === undefined) return { scope_id: line.id, code: 'resource-limit', message: 'Note vertical geometry exceeds safe integer bounds' }
      cursor = next
    }
    previousAfter = shaped.spacing_after_millipoints
  }
  const bottom = addChecked(cursor, previousAfter)
  if (bottom === undefined) return { scope_id: story.id, code: 'resource-limit', message: 'Note vertical geometry exceeds safe integer bounds' }
  return {
    id: `note-placement:${page.ordinal}:${story.id}`,
    story_id: story.id,
    story_kind: story.kind as NoteKind,
    note_role: (story.note_role ?? 'content') as NativeDocxPlacedNoteStoryV1['note_role'],
    native_story_id: story.native_story_id!,
    relationship_id: story.relationship_id!,
    ordinal,
    ...(referenceRunID ? { reference_run_id: referenceRunID } : {}),
    ...(number !== undefined ? { number } : {}),
    section_id: column.section_id,
    column_id: column.id,
    column_ordinal: column.ordinal,
    top_millipoints: top,
    height_millipoints: bottom - top,
    lines,
  }
}

function buildContentBottomIndex(
  document: NativeDocxDocumentV1,
  tables: readonly NativeDocxQualifiedTableV1[],
  shaped: NativeDocxShapedLinesV1,
): ContentBottomIndex | undefined {
  const shapedByParagraph = new Map(shaped.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const tableRowByFirstParagraph = new Map<string, TableRowBottomIndexEntry>()
  for (const table of tables) {
    const rows = layoutNativeDocxTableRowsV1(table, shaped)
    if (!rows) return undefined
    for (const [rowIndex, row] of rows.entries()) {
      const firstParagraph = table.rows[rowIndex]!.cells[0]!.cell.paragraphs[0]!
      const firstShaped = shapedByParagraph.get(firstParagraph.id)
      const contentOffsetMilliPoints = firstShaped ? addChecked(table.table.cell_margins!.top_twips * 50, firstShaped.spacing_before_millipoints) : undefined
      if (contentOffsetMilliPoints === undefined || tableRowByFirstParagraph.has(firstParagraph.id)) return undefined
      tableRowByFirstParagraph.set(firstParagraph.id, { rowID: row.row_id, contentOffsetMilliPoints, heightMilliPoints: row.height_millipoints })
    }
  }
  return { shapedByParagraph, bodyParagraphIDs: new Set(bodyParagraphs(document).map((paragraph) => paragraph.id)), tableRowByFirstParagraph }
}

function contentBottom(page: NativeDocxPaginatedPageV1, column: NativeDocxPageColumnV1, index: ContentBottomIndex): number | undefined {
  let bottom = column.y_millipoints
  const seenRows = new Set<string>()
  for (const line of page.lines) {
    if (line.column_id !== column.id) continue
    const lineBottom = addChecked(line.y_millipoints, line.height_millipoints)
    if (lineBottom === undefined) return undefined
    bottom = Math.max(bottom, lineBottom)
    const shaped = index.shapedByParagraph.get(line.paragraph_id)
    if (index.bodyParagraphIDs.has(line.paragraph_id) && shaped && line.source_line_ordinal === shaped.lines.length - 1) {
      const paragraphBottom = addChecked(lineBottom, shaped.spacing_after_millipoints)
      if (paragraphBottom === undefined) return undefined
      bottom = Math.max(bottom, paragraphBottom)
    }
    const row = index.tableRowByFirstParagraph.get(line.paragraph_id)
    if (row && !seenRows.has(row.rowID)) {
      seenRows.add(row.rowID)
      const rowTop = line.y_millipoints - row.contentOffsetMilliPoints
      const rowBottom = addChecked(rowTop, row.heightMilliPoints)
      if (!Number.isSafeInteger(rowTop) || rowTop < column.y_millipoints || rowBottom === undefined) return undefined
      bottom = Math.max(bottom, rowBottom)
    }
  }
  return bottom
}

function appendPlacement(page: NativeDocxPaginatedPageV1, placement: NativeDocxPlacedNoteStoryV1): void {
  ;(page.note_stories ??= []).push(placement)
}

function qualifiedNoteColumn(page: NativeDocxPaginatedPageV1, reference?: Pick<NoteReference, 'sectionID' | 'columnID' | 'columnOrdinal'>): NativeDocxPageColumnV1 | NativeDocxNotePaginationRefusalV1 {
  if (page.kind !== 'content' || page.section_ids.length !== 1 || page.columns.length !== 1) return {
    scope_id: page.id,
    code: 'note-structure-unsupported',
    message: 'Note-bearing pages require one content section grid and one exact column; multicolumn and shared continuous boundaries remain supported only on note-free pages',
  }
  const column = page.columns[0]!
  if (column.section_id !== page.section_id || page.section_ids[0] !== page.section_id || reference && (reference.sectionID !== column.section_id || reference.columnID !== column.id || reference.columnOrdinal !== column.ordinal)) return {
    scope_id: page.id,
    code: 'note-structure-unsupported',
    message: 'Note reference and placement must preserve one exact page section and column provenance',
  }
  return column
}

function shapedRunTextIndex(shaped: NativeDocxShapedLinesV1): Map<string, string> {
  const result = new Map<string, string>()
  for (const paragraph of shaped.paragraphs) for (const line of paragraph.lines) for (const fragment of line.fragments) {
    if (fragment.source_kind === 'run') result.set(fragment.source_id, (result.get(fragment.source_id) ?? '') + fragment.text)
  }
  return result
}

function exactInstructionSentinelProjection(story: NativeDocxStoryV1): boolean {
  return story.blocks.length === 1 && story.blocks[0]?.kind === 'paragraph' && story.blocks[0].paragraph?.runs.length === 0
}

function measureNoteGroup(
  page: Pick<NativeDocxPaginatedPageV1, 'ordinal'>,
  column: NativeDocxPageColumnV1,
  stories: Array<{ story: NativeDocxStoryV1; reference?: NoteReference }>,
  shapedByParagraph: Map<string, NativeDocxShapedParagraphV1>,
  budget: PlacementBudget,
): number | NativeDocxNotePaginationRefusalV1 {
  let height = 0
  let measuredLineCount = 0
  for (const [ordinal, entry] of stories.entries()) {
    const result = placedStory(entry.story, shapedByParagraph, page, column, 0, ordinal, entry.reference?.runID, entry.reference?.number)
    if ('code' in result) return result
    height += result.height_millipoints
    if (!Number.isSafeInteger(height)) return { scope_id: entry.story.id, code: 'resource-limit', message: 'Combined note height exceeds safe integer bounds' }
    measuredLineCount += result.lines.length
    const nextLineCount = budget.lines + measuredLineCount
    if (!Number.isSafeInteger(nextLineCount) || nextLineCount > DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines) return { scope_id: entry.story.id, code: 'resource-limit', message: `Note lines exceed ${DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines}` }
  }
  return height
}

/** Internal reservation probe uses the same note geometry and resource checks
 * as final placement; all source/label/relationship checks still run before
 * final output is committed by placeNativeDocxNotesV1. No probe is published. */
export function measureNativeDocxFootnoteAreaForReservationV1(profile: NativeDocxFootnoteReservationProfileV1): NativeDocxFootnoteAreaMeasurementV1 | undefined {
  const column = { ...profile.column, section_id: profile.section_id }
  const reference: NoteReference = { kind: 'footnote', story: profile.note, runID: profile.reference_run_id, number: 1, pageOrdinal: 0, sectionID: profile.section_id, columnID: column.id, columnOrdinal: column.ordinal }
  const height = measureNoteGroup({ ordinal: 0 }, column, [{ story: profile.separator }, { story: profile.note, reference }], new Map(profile.request.shaped_lines.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph])), { lines: 0 })
  if (typeof height !== 'number') return undefined
  return { section_id: profile.section_id, column_id: column.id, reference_run_id: profile.reference_run_id, note_story_id: profile.note.id, separator_story_id: profile.separator.id, height_millipoints: height }
}

function placeGroup(
  page: NativeDocxPaginatedPageV1,
  column: NativeDocxPageColumnV1,
  stories: Array<{ story: NativeDocxStoryV1; reference?: NoteReference }>,
  shapedByParagraph: Map<string, NativeDocxShapedParagraphV1>,
  budget: PlacementBudget,
  occupiedBottom: number,
  alignment: 'bottom' | 'flow',
): NativeDocxNotePaginationRefusalV1 | undefined {
  const height = measureNoteGroup(page, column, stories, shapedByParagraph, budget)
  if (typeof height !== 'number') return height
  const bodyBottom = column.y_millipoints + column.height_millipoints
  const top = alignment === 'bottom' ? bodyBottom - height : occupiedBottom
  const groupBottom = addChecked(top, height)
  if (!Number.isSafeInteger(top) || top < occupiedBottom || groupBottom === undefined || groupBottom > bodyBottom) return { scope_id: stories.at(-1)?.story.id ?? page.id, code: 'note-overflow-unsupported', message: 'Whole note area does not fit without continuation or body reflow' }
  let cursor = top
  for (const [index, entry] of stories.entries()) {
    const result = placedStory(entry.story, shapedByParagraph, page, column, cursor, index, entry.reference?.runID, entry.reference?.number)
    if ('code' in result) return result
    appendPlacement(page, result)
    budget.lines += result.lines.length
    cursor += result.height_millipoints
  }
  return undefined
}

/** Only an already overflowing single endnote can activate a continuation sentinel.
 * Complete source paragraphs express the slices; deterministic replay verifies
 * their order and exact-once line coverage without synthesizing note labels. */
function continueEndnote(
  layout: NativeDocxPaginatedLayoutSuccessV1,
  document: NativeDocxDocumentV1,
  resolved: NativeDocxResolvedLayoutInputV1,
  shaped: NativeDocxShapedLinesV1,
  reference: NoteReference,
  separator: NativeDocxStoryV1,
  shapedByParagraph: Map<string, NativeDocxShapedParagraphV1>,
  budget: PlacementBudget,
  occupiedBottom: number,
): NativeDocxNotePaginationRefusalV1 | undefined {
  const story = reference.story
  const fail = (message: string): NativeDocxNotePaginationRefusalV1 => ({ scope_id: story.id, code: 'note-overflow-unsupported', message })
  if (document.sections.length !== 1 || document.notes.some((note) => note.kind === 'footnote') ||
    document.notes.filter((note) => (note.note_role ?? 'content') === 'content').length !== 1) return fail('Endnote continuation requires one endnote, one section, and no footnotes')
  const paragraphs = storyParagraphs(story)
  if (!paragraphs || paragraphs.length < 2) return fail('Endnote continuation requires at least two complete paragraphs')
  const resolvedByID = new Map(resolved.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  for (const paragraph of paragraphs) {
    const lines = shapedByParagraph.get(paragraph.id)
    const properties = resolvedByID.get(paragraph.id)?.properties
    if (!lines || !properties || lines.spacing_before_millipoints !== 0 || lines.spacing_after_millipoints !== 0 ||
      properties.keep_next === true || properties.page_break_before === true ||
      (lines.lines.length > 1 && properties.keep_lines !== true) || paragraph.runs.some((run) => run.page_field !== undefined || run.layout_page_field !== undefined)) {
      return fail('Continued endnote paragraphs require zero spacing, no fields or cross-paragraph breaks, and keep_lines for multiline content')
    }
  }
  const continuations = document.notes.filter((note) => note.kind === 'endnote' && note.note_role === 'continuation-separator')
  const continuation = continuations[0]
  if (continuations.length !== 1 || !continuation || !exactInstructionSentinelProjection(continuation) ||
    continuation.relationship_id !== story.relationship_id || continuation.part_name !== story.part_name) {
    return { scope_id: story.id, code: 'note-separator-unsupported', message: 'Endnote continuation requires one exact source-bound continuation separator' }
  }
  const activeScopes = new Set([continuation.id, ...continuation.blocks.flatMap((block) => [block.id, ...(block.paragraph?.runs.map((run) => run.id) ?? [])])])
  if (document.unsupported.some((entry) => activeScopes.has(entry.scope_id)) || resolved.diagnostics.some((entry) => activeScopes.has(entry.scope_id)) ||
    shaped.diagnostics.some((entry) => activeScopes.has(entry.scope_id) || entry.source_id !== undefined && activeScopes.has(entry.source_id))) {
    return { scope_id: continuation.id, code: 'note-separator-unsupported', message: 'Activated endnote continuation separator has unsupported source, layout, or shaping semantics' }
  }
  const continuationParagraph = continuation.blocks[0]!.paragraph!
  if (!shapedByParagraph.has(continuationParagraph.id)) return {
    scope_id: continuation.id, code: 'note-continuation-shaping-required', message: 'Endnote continuation requires activated separator shaping',
  }
  for (const sentinel of [separator, continuation]) {
    const paragraph = shapedByParagraph.get(sentinel.blocks[0]!.id)
    if (!paragraph || paragraph.lines.length !== 1 || paragraph.lines[0]!.fragments.length !== 0 ||
      paragraph.spacing_before_millipoints !== 0 || paragraph.spacing_after_millipoints !== 0) {
      return { scope_id: sentinel.id, code: 'note-separator-unsupported', message: 'Continued endnotes require exact one-line, zero-spacing instruction separators' }
    }
  }
  let page = layout.pages.at(-1)!
  let column = qualifiedNoteColumn(page)
  if ('code' in column) return column
  const section = layout.sections.at(-1)!
  const template = page
  const appendPage = (): NativeDocxNotePaginationRefusalV1 | undefined => {
    if (layout.pages.length >= 2_048) return { scope_id: story.id, code: 'resource-limit', message: 'Continued endnotes exceed the page budget' }
    const ordinal = section.page_ids.length
    page = { ...structuredClone(template), id: `page:${section.section_id}:${ordinal}`, ordinal: layout.pages.length, section_page_ordinal: ordinal, paragraph_slices: [], lines: [], note_stories: [] }
    layout.pages.push(page)
    section.page_ids.push(page.id)
    occupiedBottom = column.y_millipoints
    return undefined
  }
  let start = 0
  while (start < paragraphs.length) {
    const sentinel = start === 0 ? separator : continuation
    const measuredSeparator = placedStory(sentinel, shapedByParagraph, page, column, 0, 0)
    if ('code' in measuredSeparator) return measuredSeparator
    let height = measuredSeparator.height_millipoints
    let end = start
    while (end < paragraphs.length) {
      const candidate = { ...story, blocks: story.blocks.slice(end, end + 1) }
      const measured = placedStory(candidate, shapedByParagraph, page, column, 0, 1)
      if ('code' in measured) return measured
      if (occupiedBottom + height + measured.height_millipoints > column.y_millipoints + column.height_millipoints) break
      height += measured.height_millipoints
      end += 1
    }
    if (end === start) {
      if (occupiedBottom === column.y_millipoints) return fail('One complete endnote paragraph plus its separator cannot fit an empty page')
      const failure = appendPage()
      if (failure) return failure
      continue
    }
    const failure = placeGroup(page, column, [{ story: sentinel }, { story: { ...story, blocks: story.blocks.slice(start, end) }, reference }], shapedByParagraph, budget, occupiedBottom, 'flow')
    if (failure) return failure
    start = end
    if (start < paragraphs.length) {
      const failure = appendPage()
      if (failure) return failure
    }
  }
  return undefined
}

/** Mutates only the caller-owned canonical success snapshot. */
export function placeNativeDocxNotesV1(
  layout: NativeDocxPaginatedLayoutSuccessV1,
  document: NativeDocxDocumentV1,
  resolved: NativeDocxResolvedLayoutInputV1,
  shaped: NativeDocxShapedLinesV1,
): NativeDocxNotePaginationRefusalV1 | undefined {
  const staged = structuredClone(layout)
  for (const page of staged.pages) page.note_stories = []
  if (document.notes.length === 0) {
    layout.pages.splice(0, layout.pages.length, ...staged.pages)
    layout.sections.splice(0, layout.sections.length, ...staged.sections)
    return undefined
  }
  const scopes = noteScopes(document)
  const resolvedParagraphs = new Map(resolved.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const sourceFailure = document.unsupported.find((entry) => scopes.has(entry.scope_id))
  if (sourceFailure) return { scope_id: sourceFailure.scope_id, code: 'note-structure-unsupported', message: `Unsupported note source semantics: ${sourceFailure.code}: ${sourceFailure.message}` }
  if (resolved.diagnostics.some((entry) => scopes.has(entry.scope_id))) return { scope_id: document.document_id, code: 'note-structure-unsupported', message: 'Resolved note layout contains unsupported or ambiguous style semantics' }

  const content = new Map(document.notes.filter((story) => (story.note_role ?? 'content') === 'content').map((story) => [story.id, story]))
  const separators = new Map<NoteKind, NativeDocxStoryV1>()
  const relationshipByKind = new Map<NoteKind, string>()
  const partByKind = new Map<NoteKind, string>()
  const kindByRelationship = new Map<string, NoteKind>()
  const labelsByStory = new Map<string, NoteLabel>()
  for (const story of document.notes) {
    if ((story.note_role === 'separator' || story.note_role === 'continuation-separator') && !exactInstructionSentinelProjection(story)) return { scope_id: story.id, code: 'note-separator-unsupported', message: 'Reserved note sentinels must project exactly one instruction-only paragraph and no visible runs' }
    if (story.note_role === 'continuation-separator') continue
    if (!story.relationship_id) return { scope_id: story.id, code: 'note-reference-ambiguous', message: 'Note story lacks its resolved relationship identity' }
    const relationshipKind = story.kind as NoteKind
    const knownRelationship = relationshipByKind.get(relationshipKind)
    if (knownRelationship !== undefined && knownRelationship !== story.relationship_id) return { scope_id: story.id, code: 'note-reference-ambiguous', message: 'One note kind resolves through more than one relationship identity' }
    const knownKind = kindByRelationship.get(story.relationship_id)
    if (knownKind !== undefined && knownKind !== relationshipKind) return { scope_id: story.id, code: 'note-reference-ambiguous', message: 'One relationship identity cannot resolve both footnotes and endnotes' }
    relationshipByKind.set(relationshipKind, story.relationship_id)
    kindByRelationship.set(story.relationship_id, relationshipKind)
    const knownPart = partByKind.get(relationshipKind)
    if (knownPart !== undefined && knownPart !== story.part_name) return { scope_id: story.id, code: 'note-reference-ambiguous', message: 'One note kind resolves to more than one package part' }
    partByKind.set(relationshipKind, story.part_name)
    const paragraphs = storyParagraphs(story)
    if (!paragraphs) return { scope_id: story.id, code: 'note-structure-unsupported', message: 'Nested note tables are unsupported' }
    for (const paragraph of paragraphs) {
      const properties = resolvedParagraphs.get(paragraph.id)?.properties
      if (properties?.keep_next === true || properties?.page_break_before === true) return { scope_id: paragraph.id, code: 'note-structure-unsupported', message: 'Note cross-paragraph keep and page-break constraints are outside bounded placement' }
    }
    for (const paragraph of paragraphs) for (const run of paragraph.runs) {
      if (run.drawing) return { scope_id: run.id, code: 'note-structure-unsupported', message: 'Note drawings are outside bounded native note pagination' }
      if (run.control && run.control !== 'tab' && run.control !== 'line-break') return { scope_id: run.id, code: 'note-structure-unsupported', message: `Note ${run.control} controls are outside bounded native note pagination` }
    }
    if (story.note_role === 'separator') {
      if (separators.has(story.kind as NoteKind)) return { scope_id: story.id, code: 'note-separator-unsupported', message: 'Duplicate separator stories are ambiguous' }
      separators.set(story.kind as NoteKind, story)
    }
    if ((story.note_role ?? 'content') === 'content') {
      let labels = 0
      for (const block of story.blocks) {
        if (block.kind !== 'paragraph' || !block.paragraph) return { scope_id: block.id, code: 'note-structure-unsupported', message: 'Nested note tables are unsupported' }
        for (const run of block.paragraph.runs) if (run.reference) {
          if (run.reference.role === 'label' && run.reference.kind === story.kind && run.reference.target_id === story.id) {
            labels += 1
            labelsByStory.set(story.id, { paragraphID: block.paragraph.id, runID: run.id })
          }
          else return { scope_id: run.id, code: 'note-reference-ambiguous', message: 'Note-to-note references, mismatched labels, and cycles are unsupported' }
        }
      }
      if (labels !== 1) return { scope_id: story.id, code: 'note-reference-ambiguous', message: 'Each content note requires exactly one owning label' }
    }
  }
  const shapedByParagraph = new Map(shaped.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const shapedTextByRunID = shapedRunTextIndex(shaped)
  const shapedFailure = shaped.diagnostics.find((entry) => {
    if (!scopes.has(entry.scope_id) && !(entry.source_id !== undefined && scopes.has(entry.source_id))) return false
    const properties = resolvedParagraphs.get(entry.scope_id)?.properties
    // Whole paragraphs satisfy keep_lines/widow control. Run page controls and
    // cross-paragraph constraints still need layout semantics outside this slice.
    return !(entry.code === 'page-control-deferred' && entry.severity === 'deferred' && entry.source_id === undefined && entry.source_diagnostic_code === undefined && entry.source_diagnostic_message === undefined &&
      properties && properties.keep_next !== true && properties.page_break_before !== true &&
      (['keep_next', 'keep_lines', 'page_break_before', 'widow_control'] as const).some((key) => properties[key] !== undefined &&
        entry.message === `${key} is retained in resolved layout for the future paginator and does not alter line shaping`))
  })
  if (shapedFailure) return { scope_id: shapedFailure.scope_id, code: 'note-structure-unsupported', message: `Note shaping is not exact: ${shapedFailure.code}: ${shapedFailure.message}` }
  const budget: PlacementBudget = { lines: staged.pages.reduce((count, page) => count + page.lines.length, 0) }
  if (budget.lines > DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines) return { scope_id: document.document_id, code: 'resource-limit', message: `Placed body lines exceed ${DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines}` }
  const placedByLine = new Map<string, PlacedLineIndex>()
  for (const page of staged.pages) for (const line of page.lines) placedByLine.set(line.line_id, { page, line })
  const references: NoteReference[] = []
  const seen = new Set<string>()
  const counters = new Map<NoteKind, number>([['footnote', 0], ['endnote', 0]])
  for (const paragraph of bodyParagraphs(document)) {
    const shapedParagraph = shapedByParagraph.get(paragraph.id)
    if (!shapedParagraph) continue
    const runPlacements = new Map<string, Set<PlacedLineIndex>>()
    for (const line of shapedParagraph.lines) {
      const placement = placedByLine.get(line.id)
      if (placement) for (const fragment of line.fragments) {
        const placements = runPlacements.get(fragment.source_id) ?? new Set<PlacedLineIndex>()
        placements.add(placement)
        runPlacements.set(fragment.source_id, placements)
      }
    }
    for (const run of paragraph.runs) {
      const ref = run.reference
      if (!ref || ref.role === 'label' || (ref.kind !== 'footnote' && ref.kind !== 'endnote')) continue
      const story = content.get(ref.target_id)
      const number = (counters.get(ref.kind) ?? 0) + 1
      const marker = String(number)
      const placements = runPlacements.get(run.id)
      const placement = placements?.size === 1 ? [...placements][0] : undefined
      const label = story ? labelsByStory.get(story.id) : undefined
      if (!story || story.kind !== ref.kind || !placement || seen.has(story.id) || shapedTextByRunID.get(run.id) !== marker || !label || shapedTextByRunID.get(label.runID) !== marker) return { scope_id: run.id, code: 'note-reference-ambiguous', message: 'Note anchors and labels must uniquely resolve to their exact shaped decimal marker and placed section/column in body reference order' }
      const column = qualifiedNoteColumn(placement.page, { sectionID: placement.line.section_id, columnID: placement.line.column_id, columnOrdinal: placement.line.column_ordinal })
      if ('code' in column) return column
      seen.add(story.id)
      counters.set(ref.kind, number)
      references.push({ kind: ref.kind, story, runID: run.id, number, pageOrdinal: placement.page.ordinal, sectionID: column.section_id, columnID: column.id, columnOrdinal: column.ordinal })
    }
  }
  if (seen.size !== content.size) return { scope_id: document.document_id, code: 'note-reference-ambiguous', message: 'Every content note must be referenced exactly once' }
  if (references.length > DOCX_NOTE_PAGINATION_LIMITS.maxNotes) return { scope_id: document.document_id, code: 'resource-limit', message: `Notes exceed ${DOCX_NOTE_PAGINATION_LIMITS.maxNotes}` }

  const qualifiedTables = qualifyNativeDocxTablesV1(document, resolved)
  if (qualifiedTables.status !== 'qualified') return { scope_id: qualifiedTables.diagnostics[0]?.scope_id ?? document.document_id, code: 'note-structure-unsupported', message: qualifiedTables.diagnostics[0]?.message ?? 'Table geometry is unavailable for exact note placement' }
  const contentBottomIndex = buildContentBottomIndex(document, qualifiedTables.tables, shaped)
  if (!contentBottomIndex) return { scope_id: document.document_id, code: 'note-structure-unsupported', message: 'Body/table geometry cannot be indexed for exact note placement' }
  const footnotesByPage = new Map<number, NoteReference[]>()
  const endnotes: NoteReference[] = []
  for (const reference of references) {
    if (reference.kind === 'endnote') endnotes.push(reference)
    else {
      const pageReferences = footnotesByPage.get(reference.pageOrdinal) ?? []
      pageReferences.push(reference)
      footnotesByPage.set(reference.pageOrdinal, pageReferences)
    }
  }
  for (const page of staged.pages) {
    const footnotes = footnotesByPage.get(page.ordinal) ?? []
    if (footnotes.length === 0) continue
    const separator = separators.get('footnote')
    if (!separator) return { scope_id: page.id, code: 'note-separator-unsupported', message: 'Footnote placement requires one exact ordinary separator story' }
    const column = qualifiedNoteColumn(page, footnotes[0])
    if ('code' in column) return column
    if (footnotes.some((entry) => entry.sectionID !== column.section_id || entry.columnID !== column.id || entry.columnOrdinal !== column.ordinal)) return { scope_id: page.id, code: 'note-structure-unsupported', message: 'One footnote group cannot cross section or column provenance' }
    const occupiedBottom = contentBottom(page, column, contentBottomIndex)
    if (occupiedBottom === undefined) return { scope_id: page.id, code: 'note-structure-unsupported', message: 'Body/table bottom does not exact-join qualified page geometry' }
    const failure = placeGroup(page, column, [{ story: separator }, ...footnotes.map((reference) => ({ story: reference.story, reference }))], shapedByParagraph, budget, occupiedBottom, 'bottom')
    if (failure) return failure
  }

  if (endnotes.length > 0) {
    const separator = separators.get('endnote')
    if (!separator) return { scope_id: document.document_id, code: 'note-separator-unsupported', message: 'Endnote placement requires one exact ordinary separator story' }
    let page = staged.pages.at(-1)!
    let column = qualifiedNoteColumn(page)
    if ('code' in column) return column
    let occupiedBottom = contentBottom(page, column, contentBottomIndex)
    if (occupiedBottom === undefined) return { scope_id: page.id, code: 'note-structure-unsupported', message: 'Body/table bottom does not exact-join qualified page geometry' }
    let failure = (page.note_stories?.length ?? 0) > 0
      ? { scope_id: page.id, code: 'note-overflow-unsupported' as const, message: 'Endnotes require a distinct final page after a footnote-bearing final content page' }
      : placeGroup(page, column, [{ story: separator }, ...endnotes.map((reference) => ({ story: reference.story, reference }))], shapedByParagraph, budget, occupiedBottom, 'flow')
    if (failure?.code === 'note-overflow-unsupported' && endnotes.length === 1) {
      const whole = placedStory(endnotes[0]!.story, shapedByParagraph, page, column, 0, 1)
      const rule = placedStory(separator, shapedByParagraph, page, column, 0, 0)
      if (!('code' in whole) && !('code' in rule) && whole.height_millipoints + rule.height_millipoints > column.height_millipoints) {
        const continued = continueEndnote(staged, document, resolved, shaped, endnotes[0]!, separator, shapedByParagraph, budget, occupiedBottom)
        if (continued) return continued
        failure = undefined
      }
    }
    if (failure?.code === 'note-overflow-unsupported') {
      const section = staged.sections.at(-1)!
      let template: NativeDocxPaginatedPageV1 | undefined
      for (let index = staged.pages.length - 1; index >= 0; index -= 1) {
        const candidate = staged.pages[index]!
        if (candidate.section_id === section.section_id && candidate.kind === 'content') { template = candidate; break }
      }
      if (!template) return { scope_id: section.section_id, code: 'note-structure-unsupported', message: 'Endnotes have no owning content-page geometry' }
      if (staged.pages.length >= 2_048) return { scope_id: document.document_id, code: 'resource-limit', message: 'Endnote page would exceed the page budget' }
      const sectionOrdinal = section.page_ids.length
      page = { ...structuredClone(template), id: `page:${section.section_id}:${sectionOrdinal}`, ordinal: staged.pages.length, section_page_ordinal: sectionOrdinal, paragraph_slices: [], lines: [], note_stories: [] }
      staged.pages.push(page)
      section.page_ids.push(page.id)
      column = qualifiedNoteColumn(page)
      if ('code' in column) return column
      occupiedBottom = column.y_millipoints
      failure = placeGroup(page, column, [{ story: separator }, ...endnotes.map((reference) => ({ story: reference.story, reference }))], shapedByParagraph, budget, occupiedBottom, 'flow')
    }
    if (failure) return failure
  }
  layout.pages.splice(0, layout.pages.length, ...staged.pages)
  layout.sections.splice(0, layout.sections.length, ...staged.sections)
  return undefined
}
