import { measureNativeDocxFootnoteReservationV1, type NativeDocxFootnoteReservationProfileV1, type NativeDocxFootnoteAreaMeasurementV1 } from './nativeFootnoteReservationV1.js'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import {
  DOCX_NATIVE_LIMITS,
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  DOCX_NATIVE_PROTOCOL,
  DOCX_NATIVE_VERSION,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxParagraphV1,
  type NativeDocxSectionV1,
} from './nativeContract.js'
import {
  DOCX_RESOLVED_LAYOUT_PROTOCOL,
  DOCX_RESOLVED_LAYOUT_VERSION,
  decodeNativeDocxResolvedLayout,
  nativeDocxResolvedNumberingDefinitionSha256V1,
  nativeDocxResolvedNumberingModelSha256V1,
  type NativeDocxResolvedLayoutInputV1,
} from './nativeResolvedLayout.js'
import {
  DOCX_SHAPED_LINES_PROTOCOL,
  DOCX_SHAPED_LINES_VERSION,
  twipsToMilliPoints,
  type NativeDocxShapedLinesV1,
} from './nativeShapingLines.js'
import { decodeNativeDocxShapedLines } from './nativeShapedLinesContract.js'
import { decodeNativeDocxPaginatedLayout, decodeNativeDocxPaginatedLayoutForRequest } from './nativePaginatedLayoutContract.js'
import {
  DOCX_DEFAULT_TAB_STOP_TWIPS,
  DOCX_PAGINATION_SETTINGS_PROTOCOL,
  DOCX_PAGINATION_SETTINGS_VERSION,
  DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS,
  decodeNativeDocxPaginationSettings,
  type NativeDocxPaginationSettingsV1,
} from './nativePaginationSettings.js'
import {
  DOCX_PAGINATED_LAYOUT_PROTOCOL,
  DOCX_PAGINATION_LIMITS,
  DOCX_PAGINATION_REQUEST_PROTOCOL,
  DOCX_PAGINATION_REQUEST_VERSION,
  decodeNativeDocxPaginationRequestV1,
  paginateNativeDocxV1,
  type NativeDocxPaginationRequestV1,
} from './nativePaginationV1.js'
import { asciiLowerNative, asciiUpperNative, compareNativeCodeUnits } from './nativeDeterminism.js'
import { placeNativeDocxNotesV1 } from './nativeNotePaginationV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const NUMBERING_PART = 'word/numbering.xml'
const NUMBERING_HASH = `sha256:${'c'.repeat(64)}`

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

function paragraph(id: string, index: number): NativeDocxParagraphV1 {
  return {
    id,
    anchor: anchor(`/w:document[1]/w:body[1]/w:p[${index + 1}]`, 100 + index * 100, 190 + index * 100),
    edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Test source remains authoritative.', preservation: 'refuse-mutation' } },
    properties: {},
    runs: [{ kind: 'text', id: `run:${id}`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[${index + 1}]/w:r[1]`, 110 + index * 100, 180 + index * 100), text: id }],
  }
}

interface FixtureOptions {
  lineCounts?: number[]
  lineHeight?: number
  bodyHeight?: number
  spacings?: Array<{ before?: number; after?: number }>
  properties?: Array<NativeDocxResolvedLayoutInputV1['paragraphs'][number]['properties']>
  sections?: Array<{ start: number; breakType?: NativeDocxSectionV1['break_type']; bodyWidth?: number; bodyHeight?: number; columns?: number }>
}

function fixture(options: FixtureOptions = {}): NativeDocxPaginationRequestV1 {
  const counts = options.lineCounts ?? [1]
  const bodyHeight = options.bodyHeight ?? 40_000
  const bodyWidth = 40_000
  const lineHeight = options.lineHeight ?? 10_000
  const paragraphs = counts.map((_, index) => paragraph(`paragraph:${index + 1}`, index))
  const sectionOptions = options.sections ?? [{ start: 0, bodyHeight, bodyWidth }]
  const sections: NativeDocxSectionV1[] = sectionOptions.map((entry, index) => {
    const width = entry.bodyWidth ?? bodyWidth
    const height = entry.bodyHeight ?? bodyHeight
    return {
      id: `section:${index + 1}`,
      anchor: anchor(`/w:document[1]/w:body[1]/w:sectPr[${index + 1}]`, 2_000 + index * 100, 2_090 + index * 100),
      starts_at_block_id: paragraphs[entry.start]!.id,
      break_type: entry.breakType ?? 'next-page',
      title_page: false,
      page: {
        width_twips: width / 50 + 200,
        height_twips: height / 50 + 200,
        orientation: width <= height ? 'portrait' : 'landscape',
        margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100, header_twips: 50, footer_twips: 50, gutter_twips: 0 },
        columns: entry.columns ?? 1,
        column_spacing_twips: 100,
        column_layout: 'equal-width',
        column_definitions: Array.from({ length: entry.columns ?? 1 }, (_, ordinal) => ({ id: `column:section:${index + 1}:${ordinal}`, ordinal })),
      },
      header_refs: [],
      footer_refs: [],
    }
  })
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL,
    version: DOCX_NATIVE_VERSION,
    document_id: 'document:test',
    revision: 'revision:1',
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: {
      id: 'story:body', kind: 'body', part_name: 'word/document.xml',
      anchor: anchor('/w:document[1]/w:body[1]', 1, Math.max(4_000, 200 + paragraphs.length * 100)),
      blocks: paragraphs.map((entry) => ({ kind: 'paragraph', id: entry.id, paragraph: entry })),
    },
    sections,
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [], passthrough_parts: [
      { part_name: SETTINGS_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: RELATIONSHIPS_PART, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: RELATIONSHIPS_HASH, policy: 'preserve-verbatim' },
    ], unsupported: [],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL,
    version: DOCX_RESOLVED_LAYOUT_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    source_parts: { main_part: 'WORD/document.xml' },
    paragraphs: paragraphs.map((entry, index) => ({
      paragraph_id: entry.id,
      applied_styles: [],
      properties: {
        ...(options.spacings?.[index]?.before !== undefined ? { spacing_before_twips: options.spacings[index]!.before! / 50 } : {}),
        ...(options.spacings?.[index]?.after !== undefined ? { spacing_after_twips: options.spacings[index]!.after! / 50 } : {}),
        ...(options.properties?.[index] ?? {}),
      },
      paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 },
    })),
    runs: paragraphs.map((entry) => ({ run_id: entry.runs[0]!.id, paragraph_id: entry.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } })),
    tables: [], fonts: [{ name: 'Test' }], diagnostics: [],
  }
  const shaped: NativeDocxShapedLinesV1 = {
    protocol: DOCX_SHAPED_LINES_PROTOCOL,
    version: DOCX_SHAPED_LINES_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    available_width_millipoints: bodyWidth,
    tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
    font_manifest: { manifest_id: 'manifest:test', revision: 'manifest-revision:1' },
    providers: { resolver_id: 'resolver:test', resolver_revision: '1', shaper_id: 'shaper:test', shaper_revision: '1', bidi_id: NATIVE_BIDI_PROVIDER_ID, bidi_revision: NATIVE_BIDI_PROVIDER_REVISION, bidi_unicode_version: BIDI_UNICODE_VERSION, unicode13_revision: UNICODE_13_CLASSIFIER_REVISION },
    paragraphs: paragraphs.map((entry, paragraphIndex) => {
      const before = options.spacings?.[paragraphIndex]?.before ?? 0
      const after = options.spacings?.[paragraphIndex]?.after ?? 0
      const lines = Array.from({ length: counts[paragraphIndex]! }, (_, lineIndex) => ({
        id: `line:${entry.id}:${lineIndex}`,
        ordinal: lineIndex,
        available_width_millipoints: bodyWidth,
        inline_offset_millipoints: 0,
        advance_inline_millipoints: 5_000,
        ascent_millipoints: 8_000,
        descent_millipoints: -2_000,
        line_gap_millipoints: 0,
        line_height_millipoints: lineHeight,
        justified: false,
        logical_to_visual: [],
        fragments: [],
      }))
      return {
        paragraph_id: entry.id,
        story_id: document.body.id,
        story_kind: 'body' as const,
        direction: 'ltr' as const,
        alignment: 'start' as const,
        spacing_before_millipoints: before,
        spacing_after_millipoints: after,
        indent_start_millipoints: 0,
        indent_end_millipoints: 0,
        first_line_delta_millipoints: 0,
        block_advance_millipoints: before + counts[paragraphIndex]! * lineHeight + after,
        lines,
      }
    }),
    diagnostics: [],
  }
  const paginationSettings: NativeDocxPaginationSettingsV1 = {
    protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL,
    version: DOCX_PAGINATION_SETTINGS_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    package_sha256: document.source.package_sha256,
    main_part: document.source.main_part,
    relationships_part: RELATIONSHIPS_PART,
    relationships_sha256: RELATIONSHIPS_HASH,
    relationship_id: 'rIdSettings',
    settings_part: SETTINGS_PART,
    settings_sha256: HASH,
    profile: 'word-modern-default',
    default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS,
    mirror_margins: false,
    gutter_at_top: false,
    even_and_odd_headers: false,
    compatibility_mode: 15,
    diagnostics: [],
  }
  return { protocol: DOCX_PAGINATION_REQUEST_PROTOCOL, version: DOCX_PAGINATION_REQUEST_VERSION, document, resolved_layout: resolved, shaped_lines: shaped, pagination_settings: paginationSettings }
}

function noteAnchor(partName: string, path: string, start: number, end: number) {
  return { part_name: partName, path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

function addFootnote(request: NativeDocxPaginationRequestV1, nativeID = '1', suffix = '1', lineHeight = 5_000): void {
  const part = 'word/footnotes.xml'
  const relationship = 'rIdFootnotes'
  const displayNumber = String(request.document.notes.filter((story) => (story.note_role ?? 'content') === 'content' && story.kind === 'footnote').length + 1)
  const separatorParagraph = paragraph(`paragraph:separator:${suffix}`, 20 + Number(suffix))
  separatorParagraph.anchor = noteAnchor(part, `/w:footnotes[1]/w:footnote[1]/w:p[1]`, 10, 50)
  // The extractor consumes the OOXML separator instruction itself. Its native
  // projection is exactly one paragraph with no visible runs.
  separatorParagraph.runs = []
  const separator = {
    id: `story:footnote:separator:${suffix}`, kind: 'footnote' as const, native_story_id: '-1', relationship_id: relationship, note_role: 'separator' as const, part_name: part,
    anchor: noteAnchor(part, `/w:footnotes[1]/w:footnote[1]`, 1, 60), blocks: [{ kind: 'paragraph' as const, id: separatorParagraph.id, paragraph: separatorParagraph }],
  }
  const noteParagraph = paragraph(`paragraph:footnote:${suffix}`, 30 + Number(suffix))
  noteParagraph.anchor = noteAnchor(part, `/w:footnotes[1]/w:footnote[2]/w:p[1]`, 70, 140)
  noteParagraph.runs = [{
    kind: 'reference', id: `run:footnote-label:${suffix}`, anchor: noteAnchor(part, `/w:footnotes[1]/w:footnote[2]/w:p[1]/w:r[1]`, 80, 100),
    reference: { kind: 'footnote', target_id: `story:footnote:${suffix}`, role: 'label' },
  }]
  const note = {
    id: `story:footnote:${suffix}`, kind: 'footnote' as const, native_story_id: nativeID, relationship_id: relationship, note_role: 'content' as const, part_name: part,
    anchor: noteAnchor(part, `/w:footnotes[1]/w:footnote[2]`, 61, 150), blocks: [{ kind: 'paragraph' as const, id: noteParagraph.id, paragraph: noteParagraph }],
  }
  const addSeparator = !request.document.notes.some((story) => story.note_role === 'separator' && story.kind === 'footnote')
  if (addSeparator) request.document.notes.push(separator)
  request.document.notes.push(note)
  const bodyParagraph = request.document.body.blocks[Math.max(0, Number(suffix) - 1)]!.paragraph!
  const bodyRun = bodyParagraph.runs[0]!
  bodyRun.kind = 'reference'; delete bodyRun.text
  bodyRun.reference = { kind: 'footnote', target_id: note.id }
  const bodyLine = request.shaped_lines.paragraphs.find((entry) => entry.paragraph_id === bodyParagraph.id)!.lines[0]!
  bodyLine.fragments = [{
    id: `fragment:${bodyParagraph.id}:0:0`, source_kind: 'run', source_id: bodyRun.id, start_utf16: 0, end_utf16: displayNumber.length, text: displayNumber,
    direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Zyyy', language: 'und', whitespace: false, justification_expansion_millipoints: 0, advance_inline_millipoints: bodyLine.advance_inline_millipoints,
    ascent_millipoints: 8_000, descent_millipoints: -2_000, line_gap_millipoints: 0, glyphs: [],
  }]
  bodyLine.logical_to_visual = [0]
  const addResolved = (entry: NativeDocxParagraphV1) => {
    request.resolved_layout.paragraphs.push({ paragraph_id: entry.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } })
    for (const run of entry.runs) request.resolved_layout.runs.push({ run_id: run.id, paragraph_id: entry.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } })
  }
  if (addSeparator) addResolved(separatorParagraph)
  addResolved(noteParagraph)
  const shapedStory = (storyID: string, entry: NativeDocxParagraphV1, labelText?: string) => ({
    paragraph_id: entry.id, story_id: storyID, story_kind: 'footnote' as const, direction: 'ltr' as const, alignment: 'start' as const,
    spacing_before_millipoints: 0, spacing_after_millipoints: 0, indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0,
    block_advance_millipoints: lineHeight, lines: [{ id: `line:${entry.id}:0`, ordinal: 0, available_width_millipoints: 40_000, inline_offset_millipoints: 0, advance_inline_millipoints: 5_000, ascent_millipoints: Math.max(1, lineHeight - 1_000), descent_millipoints: -1_000, line_gap_millipoints: 0, line_height_millipoints: lineHeight, justified: false, logical_to_visual: labelText === undefined ? [] : [0], fragments: labelText === undefined ? [] : [{ id: `fragment:${entry.id}:0:0`, source_kind: 'run' as const, source_id: entry.runs[0]!.id, start_utf16: 0, end_utf16: labelText.length, text: labelText, direction: 'ltr' as const, bidi_level: 0, logical_order: 0, script: 'Zyyy', language: 'und', whitespace: false, justification_expansion_millipoints: 0, advance_inline_millipoints: 5_000, ascent_millipoints: Math.max(1, lineHeight - 1_000), descent_millipoints: -1_000, line_gap_millipoints: 0, glyphs: [] }] }],
  })
  if (addSeparator) request.shaped_lines.paragraphs.push(shapedStory(separator.id, separatorParagraph))
  request.shaped_lines.paragraphs.push(shapedStory(note.id, noteParagraph, displayNumber))
}


describe('dedicated footnote reservation preparation', () => {
  function input() {
    const request = fixture()
    addFootnote(request)
    return request
  }
  function measured(profile: NativeDocxFootnoteReservationProfileV1): NativeDocxFootnoteAreaMeasurementV1 {
    return { section_id: profile.section_id, column_id: profile.column.id, reference_run_ids: profile.references.map((entry) => entry.reference_run_id), note_story_ids: profile.references.map((entry) => entry.note.id), separator_story_id: profile.separator.id, height_millipoints: 10000 }
  }
  it('qualifies one source reference and invokes only the authoritative measurement hook', () => {
    const request = input(), before = structuredClone(request)
    let calls = 0
    const result = measureNativeDocxFootnoteReservationV1(request, (profile) => { calls += 1; return measured(profile) })
    expect(result?.groups[0]?.height_millipoints).toBe(10000)
    expect(result?.references[0]?.reference_paragraph_id).toBe(request.document.body.blocks[0]!.id)
    expect(calls).toBe(1)
    expect(request).toEqual(before)
  })
  it('refuses measurement identity, bounded arithmetic and empty-page pair overflow', () => {
    for (const change of [
      { section_id: 'wrong' }, { column_id: 'wrong' }, { reference_run_ids: ['wrong'] }, { note_story_ids: ['wrong'] }, { separator_story_id: 'wrong' },
      { height_millipoints: NaN }, { height_millipoints: Infinity }, { height_millipoints: -1 }, { height_millipoints: 0 }, { height_millipoints: 0.5 }, { height_millipoints: 35000 },
    ]) expect(measureNativeDocxFootnoteReservationV1(input(), (profile) => ({ ...measured(profile), ...change }))).toBeUndefined()
    expect(measureNativeDocxFootnoteReservationV1(input(), () => undefined)).toBeUndefined()
  })
  it('refuses unsupported source profiles before calling measurement', () => {
    for (const mutate of [
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs[0]!.properties.keep_next = true },
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs[0]!.properties.page_break_before = true },
      (r: NativeDocxPaginationRequestV1) => { r.document.notes.push(structuredClone(r.document.notes[1]!)) },
      (r: NativeDocxPaginationRequestV1) => { r.pagination_settings.profile = 'unsupported' },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.available_width_millipoints -= 50 },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.paragraphs[0]!.lines[0]!.available_width_millipoints -= 50 },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.paragraphs[1]!.spacing_after_millipoints = 50; r.shaped_lines.paragraphs[1]!.block_advance_millipoints += 50; r.resolved_layout.paragraphs[1]!.properties.spacing_after_twips = 1 },
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs[1]!.properties.page_break_before = true },
      (r: NativeDocxPaginationRequestV1) => { r.document.unsupported.push({id:'unsupported:1',code:'NOTE_POSITION',capability:'notes',scope_id:r.document.notes[1]!.id,preservation:'refuse-mutation',message:'position'}) },
    ]) {
      const request=input(); mutate(request); let called=false
      expect(measureNativeDocxFootnoteReservationV1(request, (profile) => {called=true;return measured(profile)})).toBeUndefined()
      expect(called).toBe(false)
    }
  })
})


function extendFootnoteChain(request: NativeDocxPaginationRequestV1, count: number): void {
  const note = request.document.notes.find((story) => story.note_role === 'content')!
  const first = note.blocks[0]!.paragraph!
  const firstShape = request.shaped_lines.paragraphs.find((entry) => entry.paragraph_id === first.id)!
  for (let index = 1; index < count; index += 1) {
    const paragraph = structuredClone(first)
    paragraph.id = `${first.id}:member:${index}`
    paragraph.runs = [{ ...structuredClone(first.runs[0]!), id: `run:${paragraph.id}`, kind: 'text', text: '1' }]
    delete paragraph.runs[0]!.reference
    note.blocks.push({ kind: 'paragraph', id: paragraph.id, paragraph })
    request.resolved_layout.paragraphs.push({ ...structuredClone(request.resolved_layout.paragraphs.find((entry) => entry.paragraph_id === first.id)!), paragraph_id: paragraph.id })
    request.resolved_layout.runs.push({ ...structuredClone(request.resolved_layout.runs.find((entry) => entry.run_id === first.runs[0]!.id)!), run_id: paragraph.runs[0]!.id, paragraph_id: paragraph.id })
    const shape = structuredClone(firstShape)
    shape.paragraph_id = paragraph.id
    for (const line of shape.lines) {
      line.id = `line:${paragraph.id}:${line.ordinal}`
      for (const [ordinal, fragment] of line.fragments.entries()) { fragment.id = `fragment:${paragraph.id}:${line.ordinal}:${ordinal}`; fragment.source_id = paragraph.runs[0]!.id }
    }
    request.shaped_lines.paragraphs.push(shape)
  }
  for (const [index, block] of note.blocks.entries()) {
    request.resolved_layout.paragraphs.find((entry) => entry.paragraph_id === block.id)!.properties.keep_next = index < count - 1
    request.shaped_lines.diagnostics.push({ code: 'page-control-deferred', severity: 'deferred', scope_id: block.id, message: 'keep_next is retained in resolved layout for the future paginator and does not alter line shaping' })
  }
}

describe('multi-paragraph reservation source qualification', () => {
  function input(count = 2) { const request = fixture(); addFootnote(request); extendFootnoteChain(request, count); return request }
  it('measures only a complete authored chain with its label in the first member', () => {
    const request = input()
    expect(measureNativeDocxFootnoteReservationV1(request, (profile) => ({ section_id: profile.section_id, column_id: profile.column.id, reference_run_ids: profile.references.map((entry) => entry.reference_run_id), note_story_ids: profile.references.map((entry) => entry.note.id), separator_story_id: profile.separator.id, height_millipoints: 15000 }))?.references[0]?.note.blocks).toHaveLength(2)
  })
  it('refuses broken/final chains, misplaced labels and multiline members without keepLines before measurement', () => {
    for (const mutate of [
      (r: NativeDocxPaginationRequestV1) => { delete r.resolved_layout.paragraphs.find((p) => p.paragraph_id === r.document.notes[1]!.blocks[0]!.id)!.properties.keep_next },
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs.at(-1)!.properties.keep_next = true },
      (r: NativeDocxPaginationRequestV1) => { const note = r.document.notes[1]!; note.blocks.reverse() },
      (r: NativeDocxPaginationRequestV1) => { const shape = r.shaped_lines.paragraphs.at(-1)!; const line = structuredClone(shape.lines[0]!); line.id += ':second'; line.ordinal = 1; line.fragments = []; line.logical_to_visual = []; shape.lines.push(line); shape.block_advance_millipoints *= 2 },
    ]) {
      const request = input(); mutate(request); let called = false
      expect(measureNativeDocxFootnoteReservationV1(request, () => { called = true; return undefined })).toBeUndefined()
      expect(called).toBe(false)
    }
  })
})

describe('two-note reservation qualification', () => {
  function input() { const request = fixture({ lineCounts: [1, 1] }); addFootnote(request); addFootnote(request, '2', '2'); return request }
  it('measures each empty-page pair and the ordered group with one shared separator', () => {
    const seen: string[][] = []
    const result = measureNativeDocxFootnoteReservationV1(input(), (profile, references) => {
      seen.push(references.map((entry) => entry.reference_run_id))
      return { section_id: profile.section_id, column_id: profile.column.id, reference_run_ids: references.map((entry) => entry.reference_run_id), note_story_ids: references.map((entry) => entry.note.id), separator_story_id: profile.separator.id, height_millipoints: 5000 + references.length * 5000 }
    })
    expect(seen).toEqual([['run:paragraph:1'], ['run:paragraph:2'], ['run:paragraph:1', 'run:paragraph:2']])
    expect(result?.groups.map((group) => group.height_millipoints)).toEqual([10000, 10000, 15000])
    expect(result?.references.map((entry) => entry.number)).toEqual([1, 2])
  })
  it('refuses a measured group that omits or reverses source reference identities', () => {
    expect(measureNativeDocxFootnoteReservationV1(input(), (profile, references) => ({ section_id: profile.section_id, column_id: profile.column.id, reference_run_ids: references.map((entry) => entry.reference_run_id).reverse(), note_story_ids: references.map((entry) => entry.note.id), separator_story_id: profile.separator.id, height_millipoints: 15000 }))).toBeUndefined()
  })
  it('refuses two source references in the same paragraph before measurement', () => {
    const request = input(), first = request.document.body.blocks[0]!.paragraph!, second = request.document.body.blocks[1]!.paragraph!
    const moved = second.runs.pop()!; moved.anchor = structuredClone(first.runs[0]!.anchor); first.runs.push(moved)
    request.resolved_layout.runs.find((run) => run.run_id === moved.id)!.paragraph_id = first.id
    const firstLine = request.shaped_lines.paragraphs[0]!.lines[0]!, secondLine = request.shaped_lines.paragraphs[1]!.lines[0]!
    const fragment = secondLine.fragments.pop()!; fragment.logical_order = 1; fragment.id = `fragment:${first.id}:0:1`; firstLine.fragments.push(fragment); firstLine.logical_to_visual.push(1)
    firstLine.advance_inline_millipoints += fragment.advance_inline_millipoints
    secondLine.logical_to_visual = []; secondLine.advance_inline_millipoints = 0
    expect(decodeNativeDocxPaginationRequestV1(request).ok, JSON.stringify(decodeNativeDocxPaginationRequestV1(request))).toBe(true)
    let called = false
    expect(measureNativeDocxFootnoteReservationV1(request, () => { called = true; return undefined })).toBeUndefined()
    expect(called).toBe(false)
  })
})
