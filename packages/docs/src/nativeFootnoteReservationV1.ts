/** Source-bound whole-footnote reservation for the internal body paginator. */
import { decodeNativeDocxPaginationRequestV1, type NativeDocxPaginationRequestV1 } from './nativePaginationV1.js'
import type { NativeDocxParagraphV1, NativeDocxStoryV1 } from './nativeContract.js'
import { qualifyNativeDocxSectionColumnsV1, type NativeDocxQualifiedColumnV1 } from './nativeSectionColumnsV1.js'

export interface NativeDocxFootnoteReferenceV1 {
  reference_run_id: string
  reference_paragraph_id: string
  note: NativeDocxStoryV1
  number: number
}
export interface NativeDocxFootnoteReservationProfileV1 {
  request: NativeDocxPaginationRequestV1
  section_id: string
  column: NativeDocxQualifiedColumnV1
  references: NativeDocxFootnoteReferenceV1[]
  separator: NativeDocxStoryV1
  body_paragraphs: NativeDocxParagraphV1[]
}
export interface NativeDocxFootnoteAreaMeasurementV1 {
  section_id: string
  column_id: string
  reference_run_ids: string[]
  note_story_ids: string[]
  separator_story_id: string
  height_millipoints: number
}
export interface NativeDocxFootnoteReservationV1 extends NativeDocxFootnoteReservationProfileV1 {
  groups: NativeDocxFootnoteAreaMeasurementV1[]
}

/** Internal trusted note-core callback, never request-supplied data. Every
 * source/label/relationship/diagnostic join runs before final output commits. */
export type NativeDocxMeasureFootnoteAreaV1 = (profile: NativeDocxFootnoteReservationProfileV1, references: NativeDocxFootnoteReferenceV1[]) => NativeDocxFootnoteAreaMeasurementV1 | undefined

export function qualifyNativeDocxFootnoteReservationV1(value: unknown): NativeDocxFootnoteReservationProfileV1 | undefined {
  const decoded = decodeNativeDocxPaginationRequestV1(value)
  if (!decoded.ok) return undefined
  const request = decoded.value
  if ('column_shaped_lines' in request) return undefined
  const { document, resolved_layout: resolved, shaped_lines: shaped, pagination_settings: settings } = request
  if (('no_column_balance' in settings && settings.no_column_balance === true) || settings.profile !== 'word-modern-default' || settings.diagnostics.length || settings.even_and_odd_headers || settings.mirror_margins || settings.gutter_at_top) return undefined
  if (document.sections.length !== 1 || document.headers.length || document.footers.length || document.comments.length || document.comment_stories.length || document.unsupported.length || resolved.tables.length || resolved.numbering_source) return undefined
  const section = document.sections[0]!
  if (section.page.columns !== 1 || section.header_refs.length || section.footer_refs.length || section.title_page || section.page_number_start !== undefined || section.break_type !== 'next-page') return undefined
  const geometry = qualifyNativeDocxSectionColumnsV1(section)
  if (!geometry.ok) return undefined
  const column = geometry.value.columns[0]!
  if (shaped.available_width_millipoints !== column.width_millipoints || section.starts_at_block_id !== document.body.blocks[0]?.id) return undefined
  const notes = document.notes.filter((story) => (story.note_role ?? 'content') === 'content')
  const separators = document.notes.filter((story) => story.note_role === 'separator')
  if ((notes.length !== 1 && notes.length !== 2) || separators.length !== 1 || document.notes.some((story) => story.kind !== 'footnote') || document.notes.filter((story) => story.note_role === 'continuation-separator').length > 1) return undefined
  const separator = separators[0]!
  if (notes.some((note) => note.blocks.length < 1 || note.blocks.length > 16 || note.blocks.some((block) => block.kind !== 'paragraph' || !block.paragraph)) || separator.blocks.length !== 1 || !separator.blocks[0]?.paragraph || separator.blocks[0].paragraph.runs.length !== 0) return undefined
  const body: NativeDocxParagraphV1[] = []
  if (!document.body.blocks.length || document.body.blocks.length > 256) return undefined
  for (const block of document.body.blocks) { if (block.kind !== 'paragraph' || !block.paragraph) return undefined; body.push(block.paragraph) }
  const noteParagraphs = notes.flatMap((note) => note.blocks.map((block) => block.paragraph!))
  const owningNote = new Map(notes.flatMap((note) => note.blocks.map((block) => [block.id, note] as const)))
  const references: NativeDocxFootnoteReferenceV1[] = []
  const labels = new Map<string, number>()
  let textLength = 0
  for (const paragraph of [...body, ...noteParagraphs, separator.blocks[0].paragraph]) {
    const properties = resolved.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)
    const shape = shaped.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)
    if (!properties || !shape || properties.numbering || paragraph.properties.numbering || properties.properties.bidi || properties.properties.page_break_before) return undefined
    const owner = owningNote.get(paragraph.id)
    const chained = owner !== undefined && owner.blocks.at(-1)!.id !== paragraph.id
    if (chained ? properties.properties.keep_next !== true : properties.properties.keep_next === true) return undefined
    if ((properties.properties.line_rule ?? 'auto') !== 'auto' || (properties.properties.line ?? 240) !== 240 || shape.spacing_before_millipoints || shape.spacing_after_millipoints || shape.indent_start_millipoints || shape.indent_end_millipoints || shape.first_line_delta_millipoints || shape.direction !== 'ltr') return undefined
    if (paragraph === separator.blocks[0].paragraph && shape.lines.length !== 1) return undefined
    if (shape.lines.some((line) => line.line_height_millipoints !== line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints) || shape.block_advance_millipoints !== shape.lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)) return undefined
    if (!shape.lines.length || shape.lines.length > 256 || shape.block_advance_millipoints > column.height_millipoints || (shape.lines.length > 1 && properties.properties.keep_lines !== true)) return undefined
    if (shape.lines.some((line) => line.available_width_millipoints !== column.width_millipoints || line.hard_break_after || line.exclusion_start_millipoints !== undefined || line.advance_inline_millipoints > line.available_width_millipoints)) return undefined
    for (const run of paragraph.runs) {
      if (run.kind === 'text') {
        if (run.page_field || run.layout_page_field || !/^[\x20-\x7e]*$/.test(run.text ?? '')) return undefined
        textLength += (run.text ?? '').length
      } else if (run.kind === 'reference' && run.reference?.kind === 'footnote') {
        const note = notes.find((entry) => entry.id === run.reference!.target_id)
        if (!note) return undefined
        if (owner) {
          if (owner !== note || paragraph.id !== owner.blocks[0]!.id || run.reference.role !== 'label') return undefined
          labels.set(note.id, (labels.get(note.id) ?? 0) + 1)
        } else {
          if (!body.includes(paragraph) || run.reference.role === 'label' || references.some((reference) => reference.note.id === note.id || reference.reference_paragraph_id === paragraph.id)) return undefined
          references.push({ reference_run_id: run.id, reference_paragraph_id: paragraph.id, note, number: references.length + 1 })
        }
      } else return undefined
    }
  }
  if (references.length !== notes.length || notes.some((note) => labels.get(note.id) !== 1) || textLength > 100000 || resolved.runs.some((entry) => entry.properties.rtl || entry.properties.hidden)) return undefined
  return { request, section_id: section.id, column, references, separator, body_paragraphs: body }
}

/** Does not place pages or expose a partial reservation on any failed join. */
export function measureNativeDocxFootnoteReservationV1(value: unknown, measureArea: NativeDocxMeasureFootnoteAreaV1): NativeDocxFootnoteReservationV1 | undefined {
  const profile = qualifyNativeDocxFootnoteReservationV1(value)
  if (!profile) return undefined
  const groups: NativeDocxFootnoteAreaMeasurementV1[] = []
  const candidates = profile.references.map((reference) => [reference])
  if (profile.references.length === 2) candidates.push(profile.references)
  for (const references of candidates) {
    const measured = measureArea(profile, references)
    if (!measured || measured.section_id !== profile.section_id || measured.column_id !== profile.column.id || measured.separator_story_id !== profile.separator.id ||
      !Array.isArray(measured.reference_run_ids) || !Array.isArray(measured.note_story_ids) ||
      measured.reference_run_ids.length !== references.length || measured.reference_run_ids.some((id, index) => id !== references[index]!.reference_run_id) ||
      measured.note_story_ids.length !== references.length || measured.note_story_ids.some((id, index) => id !== references[index]!.note.id) ||
      !Number.isSafeInteger(measured.height_millipoints) || measured.height_millipoints <= 0) return undefined
    // A combined group may exceed the page: its second reference then advances.
    // Each individually activated reference/group pair must fit an empty page.
    if (references.length === 1) {
      const reference = profile.request.shaped_lines.paragraphs.find((paragraph) => paragraph.paragraph_id === references[0]!.reference_paragraph_id)!
      const pair = reference.block_advance_millipoints + measured.height_millipoints
      if (!Number.isSafeInteger(pair) || pair > profile.column.height_millipoints) return undefined
    }
    groups.push(measured)
  }
  return { ...profile, groups }
}
