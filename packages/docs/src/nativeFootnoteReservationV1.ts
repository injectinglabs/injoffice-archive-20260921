/** Internal preparation only; body reflow is not wired until the shared paginator integration. */
import { decodeNativeDocxPaginationRequestV1, type NativeDocxPaginationRequestV1 } from './nativePaginationV1.js'
import type { NativeDocxParagraphV1, NativeDocxStoryV1 } from './nativeContract.js'
import { qualifyNativeDocxSectionColumnsV1, type NativeDocxQualifiedColumnV1 } from './nativeSectionColumnsV1.js'

export interface NativeDocxFootnoteReservationProfileV1 {
  request: NativeDocxPaginationRequestV1
  section_id: string
  column: NativeDocxQualifiedColumnV1
  reference_run_id: string
  reference_paragraph_id: string
  note: NativeDocxStoryV1
  separator: NativeDocxStoryV1
  body_paragraphs: NativeDocxParagraphV1[]
}
export interface NativeDocxFootnoteAreaMeasurementV1 {
  section_id: string
  column_id: string
  reference_run_id: string
  note_story_id: string
  separator_story_id: string
  height_millipoints: number
}
export interface NativeDocxFootnoteReservationV1 extends NativeDocxFootnoteReservationProfileV1 {
  height_millipoints: number
  reference_height_millipoints: number
}

/** The callback is internal trusted note-core code, never request-supplied data.
 * Integration must lift the existing placedStory/placeGroup measurement once,
 * preserving every note source/label/relationship/diagnostic check. */
export type NativeDocxMeasureFootnoteAreaV1 = (profile: NativeDocxFootnoteReservationProfileV1) => NativeDocxFootnoteAreaMeasurementV1 | undefined

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
  if (notes.length !== 1 || separators.length !== 1 || document.notes.some((story) => story.kind !== 'footnote') || document.notes.filter((story) => story.note_role === 'continuation-separator').length > 1) return undefined
  const note = notes[0]!, separator = separators[0]!
  if (note.blocks.length !== 1 || !note.blocks[0]?.paragraph || separator.blocks.length !== 1 || !separator.blocks[0]?.paragraph || separator.blocks[0].paragraph.runs.length !== 0) return undefined
  const body: NativeDocxParagraphV1[] = []
  if (!document.body.blocks.length || document.body.blocks.length > 256) return undefined
  for (const block of document.body.blocks) { if (block.kind !== 'paragraph' || !block.paragraph) return undefined; body.push(block.paragraph) }
  const noteParagraph = note.blocks[0].paragraph
  let referenceRun: string | undefined, referenceParagraph: string | undefined, labels = 0, textLength = 0
  for (const paragraph of [...body, noteParagraph, separator.blocks[0].paragraph]) {
    const properties = resolved.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)
    const shape = shaped.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)
    if (!properties || !shape || properties.numbering || paragraph.properties.numbering || properties.properties.bidi || properties.properties.keep_next || properties.properties.page_break_before) return undefined
    if ((properties.properties.line_rule ?? 'auto') !== 'auto' || (properties.properties.line ?? 240) !== 240 || shape.spacing_before_millipoints || shape.spacing_after_millipoints || shape.indent_start_millipoints || shape.indent_end_millipoints || shape.first_line_delta_millipoints || shape.direction !== 'ltr') return undefined
    if (paragraph === separator.blocks[0].paragraph && shape.lines.length !== 1) return undefined
    if (shape.lines.some((line) => line.line_height_millipoints !== line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints) || shape.block_advance_millipoints !== shape.lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)) return undefined
    if (!shape.lines.length || shape.lines.length > 256 || shape.block_advance_millipoints > column.height_millipoints || (shape.lines.length > 1 && properties.properties.keep_lines !== true)) return undefined
    if (shape.lines.some((line) => line.available_width_millipoints !== column.width_millipoints || line.hard_break_after || line.exclusion_start_millipoints !== undefined || line.advance_inline_millipoints > line.available_width_millipoints)) return undefined
    for (const run of paragraph.runs) {
      if (run.kind === 'text') {
        if (run.page_field || run.layout_page_field || !/^[\x20-\x7e]*$/.test(run.text ?? '')) return undefined
        textLength += (run.text ?? '').length
      } else if (run.kind === 'reference' && run.reference?.kind === 'footnote' && run.reference.target_id === note.id) {
        if (paragraph === noteParagraph) { if (run.reference.role !== 'label') return undefined; labels += 1 }
        else { if (run.reference.role === 'label' || referenceRun) return undefined; referenceRun = run.id; referenceParagraph = paragraph.id }
      } else return undefined
    }
  }
  if (!referenceRun || !referenceParagraph || labels !== 1 || textLength > 100000 || resolved.runs.some((entry) => entry.properties.rtl || entry.properties.hidden)) return undefined
  return { request, section_id: section.id, column, reference_run_id: referenceRun, reference_paragraph_id: referenceParagraph, note, separator, body_paragraphs: body }
}

/** Does not place pages or expose a partial reservation on any failed join. */
export function measureNativeDocxFootnoteReservationV1(value: unknown, measureArea: NativeDocxMeasureFootnoteAreaV1): NativeDocxFootnoteReservationV1 | undefined {
  const profile = qualifyNativeDocxFootnoteReservationV1(value)
  if (!profile) return undefined
  const measured = measureArea(profile)
  if (!measured || measured.section_id !== profile.section_id || measured.column_id !== profile.column.id || measured.reference_run_id !== profile.reference_run_id || measured.note_story_id !== profile.note.id || measured.separator_story_id !== profile.separator.id || !Number.isSafeInteger(measured.height_millipoints) || measured.height_millipoints <= 0 || measured.height_millipoints > profile.column.height_millipoints) return undefined
  const reference = profile.request.shaped_lines.paragraphs.find((paragraph) => paragraph.paragraph_id === profile.reference_paragraph_id)!
  const pair = reference.block_advance_millipoints + measured.height_millipoints
  if (!Number.isSafeInteger(pair) || pair > profile.column.height_millipoints) return undefined
  return { ...profile, height_millipoints: measured.height_millipoints, reference_height_millipoints: reference.block_advance_millipoints }
}
