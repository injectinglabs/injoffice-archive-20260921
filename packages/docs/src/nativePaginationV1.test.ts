import { measureNativeDocxFootnoteAreaForReservationV1 } from './nativeNotePaginationV1.js'
import { measureNativeDocxFootnoteReservationV1 } from './nativeFootnoteReservationV1.js'
import { planNativeDocxColumnParagraphFlowV1 } from './nativeColumnParagraphFlowV1.js'
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
  paginateNativeDocxApproximateLegacyV1,
  DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING,
  nativeDocxApproximatePaginationPolicyReasonsV1,
  type NativeDocxPaginationRequestV1,
} from './nativePaginationV1.js'
import { asciiLowerNative, asciiUpperNative, compareNativeCodeUnits } from './nativeDeterminism.js'
import { placeNativeDocxNotesV1, DOCX_NOTE_PAGINATION_LIMITS } from './nativeNotePaginationV1.js'

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

function numberedFixture(options: FixtureOptions = {}): NativeDocxPaginationRequestV1 {
  const value = fixture(options)
  value.document.passthrough_parts.push({ part_name: NUMBERING_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', byte_length: 1, sha256: NUMBERING_HASH, policy: 'preserve-verbatim' })
  value.resolved_layout.source_parts.numbering_part = NUMBERING_PART
  const numberingSourceBase = {
    relationships_part: RELATIONSHIPS_PART, relationships_sha256: RELATIONSHIPS_HASH,
    relationship_id: 'rIdNumbering', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', relationship_target: 'numbering.xml',
    part_name: NUMBERING_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' as const, part_sha256: NUMBERING_HASH,
  }
  value.resolved_layout.paragraphs.forEach((paragraph, index) => {
    const native = value.document.body.blocks[index]!.paragraph!
    native.properties.numbering = { num_id: '7', abstract_num_id: '3', level: 0 }
    paragraph.properties = { ...paragraph.properties, indent_start_twips: 200, hanging_twips: 200 }
    const numbering = {
      marker_id: `marker:${paragraph.paragraph_id}`, definition_sha256: HASH, num_id: '7', abstract_num_id: '3', level: 0, start: 1,
      format: 'decimal' as const, text: '%1.', suffix: 'tab' as const, alignment: 'start' as const, never_restart: true,
      counter_value: index + 1, counter_values: [{ level: 0, value: index + 1, format: 'decimal' as const }], resolved_text: `${index + 1}.`,
      label_start_twips: 0, label_end_twips: 200, text_start_twips: 200, numbering_tab_twips: 240,
      marker_properties: { font_family: 'Test', font_size_half_points: 20, language: 'en-US' },
    }
    numbering.definition_sha256 = nativeDocxResolvedNumberingDefinitionSha256V1(numbering, NUMBERING_HASH)
    paragraph.numbering = numbering
    const shaped = value.shaped_lines.paragraphs[index]!
    shaped.indent_start_millipoints = 10_000
    shaped.first_line_delta_millipoints = -10_000
    shaped.list_marker = {
      marker_id: numbering.marker_id, definition_sha256: numbering.definition_sha256, numbering_part_sha256: NUMBERING_HASH, model_sha256: HASH,
      num_id: '7', abstract_num_id: '3', level: 0, counter_value: index + 1, text: `${index + 1}.`, suffix: 'tab', alignment: 'start',
      label_start_millipoints: 0, label_end_millipoints: 10_000, marker_start_millipoints: 0, marker_advance_millipoints: 4_000, text_start_millipoints: 12_000,
    }
    const line = shaped.lines[0]!
    line.inline_offset_millipoints = 0
    line.advance_inline_millipoints = 12_000
    line.logical_to_visual = [0, 1]
    line.fragments = [
      {
        id: `fragment:${paragraph.paragraph_id}:0:0`, source_kind: 'list-marker', source_id: paragraph.paragraph_id,
        start_utf16: 0, end_utf16: `${index + 1}.`.length, text: `${index + 1}.`, direction: 'ltr', bidi_level: 0, logical_order: 0,
        script: 'Zyyy', language: 'en-US', face_id: 'face:test', whitespace: false, advance_inline_millipoints: 4_000, justification_expansion_millipoints: 0,
        ascent_millipoints: 8_000, descent_millipoints: -2_000, line_gap_millipoints: 0,
        glyphs: [{ glyph_id: 1, advance_x_millipoints: 4_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 }],
      },
      {
        id: `fragment:${paragraph.paragraph_id}:0:1`, source_kind: 'list-marker', source_id: paragraph.paragraph_id,
        start_utf16: 0, end_utf16: 0, text: '\t', direction: 'ltr', bidi_level: 0, logical_order: 1,
        script: 'Zyyy', language: 'en-US', whitespace: true, advance_inline_millipoints: 8_000, justification_expansion_millipoints: 0,
        ascent_millipoints: 0, descent_millipoints: 0, line_gap_millipoints: 0, glyphs: [],
      },
    ]
  })
  const source = { ...numberingSourceBase, model_sha256: nativeDocxResolvedNumberingModelSha256V1(value.resolved_layout.paragraphs, numberingSourceBase) }
  value.resolved_layout.numbering_source = source
  value.shaped_lines.numbering_source = source
  for (const paragraph of value.shaped_lines.paragraphs) paragraph.list_marker!.model_sha256 = source.model_sha256
  return value
}

function paginated(request: NativeDocxPaginationRequestV1) {
  const result = paginateNativeDocxV1(request)
  expect(result.ok, JSON.stringify(result)).toBe(true)
  if (!result.ok) throw new Error(JSON.stringify(result.issues))
  if (result.value.status !== 'paginated') throw new Error(JSON.stringify(result.value.diagnostics))
  expect(result.value.status).toBe('paginated')
  const decoded = decodeNativeDocxPaginatedLayoutForRequest(result.value, request)
  if (!decoded.ok) throw new Error(JSON.stringify(decoded.issues))
  return result.value
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

function convertFootnotesToEndnotes(request: NativeDocxPaginationRequestV1): void {
  for (const story of request.document.notes) {
    story.kind = 'endnote'
    story.id = story.id.replace('footnote', 'endnote')
    story.part_name = 'word/endnotes.xml'
    story.relationship_id = 'rIdEndnotes'
    story.anchor.part_name = story.part_name
    for (const block of story.blocks) if (block.paragraph) {
      block.paragraph.anchor.part_name = story.part_name
      for (const run of block.paragraph.runs) {
        run.anchor.part_name = story.part_name
        if (run.reference) { run.reference.kind = 'endnote'; run.reference.target_id = story.id }
      }
    }
  }
  for (const block of request.document.body.blocks) if (block.paragraph) for (const run of block.paragraph.runs) {
    if (run.reference?.kind === 'footnote') { run.reference.kind = 'endnote'; run.reference.target_id = run.reference.target_id.replace('footnote', 'endnote') }
  }
  for (const paragraph of request.shaped_lines.paragraphs) if (paragraph.story_kind === 'footnote') {
    paragraph.story_kind = 'endnote'
    paragraph.story_id = paragraph.story_id.replace('footnote', 'endnote')
  }
}

function continuedEndnoteFixture(bodyLines = 1, bodyHeight = 40_000): NativeDocxPaginationRequestV1 {
  const request = fixture({ bodyHeight, lineCounts: [bodyLines] })
  addFootnote(request, '1', '1', 10_000)
  convertFootnotesToEndnotes(request)
  const separator = request.document.notes[0]!
  const note = request.document.notes[1]!
  const separatorParagraph = separator.blocks[0]!.paragraph!
  const separatorShaped = request.shaped_lines.paragraphs.find((entry) => entry.paragraph_id === separatorParagraph.id)!
  const separatorResolved = request.resolved_layout.paragraphs.find((entry) => entry.paragraph_id === separatorParagraph.id)!
  const continuation = structuredClone(separator)
  continuation.id = 'story:endnote:continuation'
  continuation.native_story_id = '0'
  continuation.note_role = 'continuation-separator'
  const continuationParagraph = continuation.blocks[0]!.paragraph!
  continuationParagraph.id = 'paragraph:endnote:continuation'
  continuation.blocks[0]!.id = continuationParagraph.id
  request.document.notes.push(continuation)
  request.resolved_layout.paragraphs.push({ ...structuredClone(separatorResolved), paragraph_id: continuationParagraph.id })
  request.shaped_lines.paragraphs.push({ ...structuredClone(separatorShaped), paragraph_id: continuationParagraph.id, story_id: continuation.id,
    lines: [{ ...structuredClone(separatorShaped.lines[0]!), id: `line:${continuationParagraph.id}:0` }],
  })
  for (let ordinal = 2; ordinal <= 6; ordinal += 1) {
    const paragraph = structuredClone(note.blocks[0]!.paragraph!)
    paragraph.runs = []
    paragraph.id = `paragraph:endnote:${ordinal}`
    note.blocks.push({ kind: 'paragraph', id: paragraph.id, paragraph })
    request.resolved_layout.paragraphs.push({ ...structuredClone(separatorResolved), paragraph_id: paragraph.id })
    request.shaped_lines.paragraphs.push({ ...structuredClone(separatorShaped), paragraph_id: paragraph.id, story_id: note.id,
      lines: [{ ...structuredClone(separatorShaped.lines[0]!), id: `line:${paragraph.id}:0` }],
    })
  }
  return request
}

describe('native DOCX shaped-lines wire decoder', () => {
  it('accepts the exact shaped projection and rejects unknown, null, negative zero, and inconsistent block advances', () => {
    const shaped = fixture().shaped_lines
    expect(decodeNativeDocxShapedLines(shaped).ok).toBe(true)
    for (const mutate of [
      (value: any) => { value.extra = true },
      (value: any) => { value.paragraphs[0].lines[0].fragments = [null] },
      (value: any) => { value.paragraphs[0].lines[0].ordinal = -0 },
      (value: any) => { value.paragraphs[0].block_advance_millipoints += 1 },
    ]) {
      const invalid = structuredClone(shaped) as any
      mutate(invalid)
      expect(decodeNativeDocxShapedLines(invalid).ok).toBe(false)
    }
  })

  it('bounds every accepted twip before exact x50 conversion and retains negative-zero refusal', () => {
    const resolved = structuredClone(fixture().resolved_layout)
    resolved.paragraphs[0]!.properties.indent_start_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS
    expect(decodeNativeDocxResolvedLayout(resolved).ok).toBe(true)
    expect(twipsToMilliPoints(DOCX_MAX_TWIPS_FOR_MILLIPOINTS)).toBe(DOCX_MAX_TWIPS_FOR_MILLIPOINTS * 50)
    for (const unsafe of [DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1, -0]) {
      const candidate = structuredClone(resolved)
      candidate.paragraphs[0]!.properties.indent_start_twips = unsafe
      expect(decodeNativeDocxResolvedLayout(candidate).ok).toBe(false)
      expect(() => twipsToMilliPoints(unsafe)).toThrow(RangeError)
    }

    const document = structuredClone(fixture().document)
    document.sections[0]!.page.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS
    expect(decodeNativeDocxDocument(document).ok).toBe(true)
    document.sections[0]!.page.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1
    expect(decodeNativeDocxDocument(document).ok).toBe(false)
    document.sections[0]!.page.width_twips = 1_000
    document.sections[0]!.page.margins.left_twips = -0
    expect(decodeNativeDocxDocument(document).ok).toBe(false)
  })
})

describe('native DOCX pagination v1', () => {
  function setEqualColumns(request: NativeDocxPaginationRequestV1, count: number, spacingTwips = 100): number {
    const bodyWidth = 40_000
    const spacing = spacingTwips * 50
    const width = (bodyWidth - (count - 1) * spacing) / count
    if (!Number.isSafeInteger(width)) throw new Error('test column width must be integral')
    request.document.sections.forEach((section, sectionIndex) => {
      section.page.columns = count
      section.page.column_spacing_twips = spacingTwips
      section.page.column_layout = 'equal-width'
      section.page.column_definitions = Array.from({ length: count }, (_, ordinal) => ({ id: `column:section:${sectionIndex + 1}:${ordinal}`, ordinal }))
    })
    request.shaped_lines.available_width_millipoints = width
    request.shaped_lines.paragraphs.forEach((paragraph) => paragraph.lines.forEach((line) => { line.available_width_millipoints = width }))
    return width
  }

  it('accepts the shared modern Word settings attestation and binds tab shaping', () => {
    const shared = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/pagination-settings-v1.json', import.meta.url), 'utf8'))
    expect(Object.keys(shared).sort()).toEqual([...DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS.SettingsV1].filter((field) => field !== 'no_column_balance').sort())
    expect(decodeNativeDocxPaginationSettings(shared).ok).toBe(true)
    const request = fixture()
    request.document.passthrough_parts.find((part) => part.part_name === SETTINGS_PART)!.sha256 = shared.settings_sha256
    request.pagination_settings = shared
    expect(paginated(request).provenance.pagination_settings).toEqual(shared)
  })

  it('strictly validates settings and refuses an unbound shaped tab interval', () => {
    const settings = fixture().pagination_settings
    for (const mutate of [
      (value: any) => { value.extra = true },
      (value: any) => { value.default_tab_stop_twips = -0 },
      (value: any) => { value.diagnostics = [null] },
    ]) {
      const invalid = structuredClone(settings) as any
      mutate(invalid)
      expect(decodeNativeDocxPaginationSettings(invalid).ok).toBe(false)
    }
    const request = fixture()
    request.shaped_lines.tab_interval_millipoints = 18_000
    const result = paginateNativeDocxV1(request)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'default-tab-stop-mismatch' })]) }))

    const orphan = structuredClone(fixture()) as any
    delete orphan.pagination_settings.relationships_part
    delete orphan.pagination_settings.relationships_sha256
    delete orphan.pagination_settings.relationship_id
    expect(decodeNativeDocxPaginationSettings(orphan.pagination_settings).ok).toBe(false)

    const missingModernMode = structuredClone(settings) as any
    delete missingModernMode.compatibility_mode
    expect(decodeNativeDocxPaginationSettings(missingModernMode).ok).toBe(false)

    const absent = fixture() as any
    absent.pagination_settings.profile = 'absent-default'
    delete absent.pagination_settings.settings_part
    delete absent.pagination_settings.settings_sha256
    delete absent.pagination_settings.relationships_part
    delete absent.pagination_settings.relationships_sha256
    delete absent.pagination_settings.relationship_id
    delete absent.pagination_settings.compatibility_mode
    expect(decodeNativeDocxPaginationSettings({ ...absent.pagination_settings, no_column_balance: true }).ok).toBe(false)
    expect(decodeNativeDocxPaginationSettings({ ...absent.pagination_settings, no_column_balance: false }).ok).toBe(false)
    expect(decodeNativeDocxPaginationSettings({ ...settings, no_column_balance: false }).ok).toBe(true)
    expect(decodeNativeDocxPaginationSettings({ ...settings, no_column_balance: 'true' }).ok).toBe(false)
    const absentResult = paginateNativeDocxV1(absent)
    expect(absentResult).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'settings-attestation-unsupported' })]) }) }))
  })

  it('recomputes marker text start and tab suffix from authoritative numbering and the attested default interval', () => {
    const request = numberedFixture()
    expect(paginated(request).pages).toHaveLength(1)

    const tampered = structuredClone(request)
    const paragraph = tampered.shaped_lines.paragraphs[0]!
    paragraph.list_marker!.text_start_millipoints += 1_000
    paragraph.lines[0]!.fragments[1]!.advance_inline_millipoints += 1_000
    paragraph.lines[0]!.advance_inline_millipoints += 1_000
    expect(decodeNativeDocxShapedLines(tampered.shaped_lines).ok).toBe(true)
    expect(paginateNativeDocxV1(tampered)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/list_marker' })]) }))
  })

  it('binds settings to the exact package and relationship closure', () => {
    for (const mutate of [
      (value: any) => { value.pagination_settings.package_sha256 = `sha256:${'c'.repeat(64)}` },
      (value: any) => { value.pagination_settings.relationships_sha256 = `sha256:${'c'.repeat(64)}` },
      (value: any) => { value.pagination_settings.relationships_part = 'word/_rels/orphan.xml.rels' },
    ]) {
      const request = fixture() as any
      mutate(request)
      expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))
    }
  })

  it('keeps settings identity validation in exact parity with the Go native id grammar', () => {
    for (const field of ['document_id', 'revision', 'relationship_id'] as const) {
      const invalid = structuredClone(fixture().pagination_settings) as any
      invalid[field] = `${invalid[field]}/not-a-native-id`
      expect(decodeNativeDocxPaginationSettings(invalid)).toEqual(expect.objectContaining({
        ok: false,
        issues: expect.arrayContaining([expect.objectContaining({ path: `/${field}` })]),
      }))
    }
  })

  it('rejects shaped paragraph semantics that do not exactly project resolved layout', () => {
    const spacing = fixture() as any
    spacing.shaped_lines.paragraphs[0].spacing_before_millipoints = 50
    spacing.shaped_lines.paragraphs[0].block_advance_millipoints += 50
    expect(paginateNativeDocxV1(spacing)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/spacing_before_millipoints' })]) }))

    const direction = fixture() as any
    direction.shaped_lines.paragraphs[0].direction = 'rtl'
    direction.shaped_lines.paragraphs[0].lines[0].inline_offset_millipoints = direction.shaped_lines.paragraphs[0].lines[0].available_width_millipoints - direction.shaped_lines.paragraphs[0].lines[0].advance_inline_millipoints
    expect(paginateNativeDocxV1(direction)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/direction' })]) }))

    const indent = fixture() as any
    indent.shaped_lines.paragraphs[0].indent_start_millipoints = 50
    indent.shaped_lines.paragraphs[0].lines[0].inline_offset_millipoints = 50
    expect(paginateNativeDocxV1(indent)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/indent_start_millipoints' })]) }))
  })

  it('caps cross-contract ownership issues at the public validation bound', () => {
    const request = fixture() as any
    request.shaped_lines.diagnostics = Array.from({ length: DOCX_NATIVE_LIMITS.maxIssues + 25 }, (_, index) => ({
      code: `diagnostic:${index}`,
      severity: 'deferred',
      scope_id: `orphan:${index}`,
      message: `orphan diagnostic ${index}`,
    }))
    const result = paginateNativeDocxV1(request)
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.issues).toHaveLength(DOCX_NATIVE_LIMITS.maxIssues)
      expect(result.issues.every((entry) => entry.code === 'BROKEN_REFERENCE')).toBe(true)
    }
  })

  it('requires exact resolved paragraph, run, and table completeness and ownership', () => {
    const extraParagraph = fixture() as any
    extraParagraph.resolved_layout.paragraphs.push({
      ...extraParagraph.resolved_layout.paragraphs[0],
      paragraph_id: 'paragraph:orphan',
    })
    expect(paginateNativeDocxV1(extraParagraph)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/resolved_layout/paragraphs/1/paragraph_id' })]),
    }))

    const extraTable = fixture() as any
    extraTable.resolved_layout.tables.push({ table_id: 'table:orphan' })
    expect(paginateNativeDocxV1(extraTable)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/resolved_layout/tables/0/table_id' })]),
    }))

    const fragment = fixture() as any
    fragment.shaped_lines.paragraphs[0].lines[0].fragments = [{
      id: 'fragment:paragraph:1:0:0', source_kind: 'tab', source_id: 'run:paragraph:1',
      start_utf16: 0, end_utf16: 0, text: '\t', direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Zyyy', language: 'und', whitespace: true,
      advance_inline_millipoints: 5_000, justification_expansion_millipoints: 0, ascent_millipoints: 0, descent_millipoints: 0, line_gap_millipoints: 0, glyphs: [],
    }]
    fragment.shaped_lines.paragraphs[0].lines[0].logical_to_visual = [0]
    fragment.resolved_layout.runs = []
    expect(paginateNativeDocxV1(fragment)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/lines/0/fragments/0/source_id' })]),
    }))

    const hardBreak = fixture() as any
    hardBreak.shaped_lines.paragraphs[0].lines[0].hard_break_after = { source_run_id: 'run:paragraph:1', control: 'line-break' }
    hardBreak.resolved_layout.runs = []
    expect(paginateNativeDocxV1(hardBreak)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/lines/0/hard_break_after/source_run_id' })]),
    }))

    const wrongOwner = fixture({ lineCounts: [1, 1] }) as any
    wrongOwner.resolved_layout.runs[0].paragraph_id = 'paragraph:2'
    expect(paginateNativeDocxV1(wrongOwner)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/resolved_layout/runs/0/paragraph_id' })]),
    }))

    const extra = fixture() as any
    extra.resolved_layout.runs.push({ ...extra.resolved_layout.runs[0], run_id: 'run:orphan' })
    expect(paginateNativeDocxV1(extra)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ path: '/resolved_layout/runs/1/run_id' })]),
    }))
  })

  it('uses percent-decoded ASCII-only OPC part equivalence without Unicode locale folding', () => {
    const encoded = fixture()
    encoded.document.passthrough_parts[0] = { part_name: 'WORD/%53ettings.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' }
    encoded.pagination_settings = {
      ...encoded.pagination_settings, profile: 'word-modern-default', settings_part: 'word/Settings.xml', settings_sha256: HASH,
    }
    expect(paginated(encoded).status).toBe('paginated')

    const dottedI = fixture()
    dottedI.document.passthrough_parts.push({ part_name: 'word/%C4%B0.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' })
    dottedI.pagination_settings = { ...dottedI.pagination_settings, profile: 'word-modern-default', settings_part: 'word/%C4%B0.xml', settings_sha256: HASH }
    expect(paginated(dottedI).status).toBe('paginated')
    dottedI.pagination_settings.settings_part = 'word/i.xml'
    expect(paginateNativeDocxV1(dottedI)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))

    const ASCIIContentType = fixture()
    ASCIIContentType.document.passthrough_parts[0] = { part_name: 'word/settings.xml', content_type: 'APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.WORDPROCESSINGML.SETTINGS+XML', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' }
    ASCIIContentType.pagination_settings = { ...ASCIIContentType.pagination_settings, profile: 'word-modern-default', settings_part: 'word/settings.xml', settings_sha256: HASH }
    expect(paginated(ASCIIContentType).status).toBe('paginated')
    ASCIIContentType.document.passthrough_parts[0]!.content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settingſ+xml'
    expect(paginateNativeDocxV1(ASCIIContentType)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))

    const output = paginated(fixture())
    output.provenance.main_part = 'WORD/%64ocument.xml'
    output.provenance.pagination_settings.main_part = 'word/document.XML'
    expect(decodeNativeDocxPaginatedLayout(output).ok).toBe(true)
    output.provenance.main_part = 'word/%C4%B0.xml'
    output.provenance.pagination_settings.main_part = 'word/i.xml'
    expect(decodeNativeDocxPaginatedLayout(output).ok).toBe(false)
  })

  it('places deterministic integer lines and uses max interparagraph spacing', () => {
    const value = paginated(fixture({ lineCounts: [1, 1], spacings: [{ before: 100, after: 300 }, { before: 200 }] }))
    expect(value.protocol).toBe(DOCX_PAGINATED_LAYOUT_PROTOCOL)
    expect(value.pages).toHaveLength(1)
    expect(value.pages[0]!.lines.map((line) => line.y_millipoints)).toEqual([5_100, 15_400])
    expect(value.pages[0]!.paragraph_slices.map((slice) => slice.space_before_millipoints)).toEqual([100, 300])
    expect(value.pages[0]!.lines.every((line) => Number.isSafeInteger(line.y_millipoints))).toBe(true)
    expect(decodeNativeDocxPaginatedLayout(value).ok).toBe(true)
  })

  it('honors before-spacing only at the first content page of each section', () => {
    const automatic = paginated(fixture({ lineCounts: [2, 1], bodyHeight: 20_000, spacings: [{}, { before: 5_000 }] }))
    expect(automatic.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)
    expect(automatic.pages[1]!.lines[0]!.y_millipoints).toBe(5_000)

    const explicit = paginated(fixture({ lineCounts: [1, 1], bodyHeight: 20_000, spacings: [{}, { before: 5_000 }], properties: [{}, { page_break_before: true }] }))
    expect(explicit.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    const kept = paginated(fixture({ lineCounts: [2, 1, 1], bodyHeight: 30_000, spacings: [{}, { before: 5_000 }, {}], properties: [{}, { keep_next: true }, {}] }))
    expect(kept.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    const widowMoved = paginated(fixture({ lineCounts: [1, 3], bodyHeight: 30_000, spacings: [{}, { before: 5_000 }] }))
    expect(widowMoved.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    const section = paginated(fixture({ lineCounts: [1, 1], spacings: [{}, { before: 5_000 }], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, bodyWidth: 40_000, bodyHeight: 40_000 },
    ] }))
    expect(section.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(5_000)
  })

  it('advances past an empty first section page when only retained top spacing prevents a fit', () => {
    const paragraph = paginated(fixture({ lineCounts: [1], bodyHeight: 10_000, spacings: [{ before: 5_000 }] }))
    expect(paragraph.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([[], ['paragraph:1']])
    expect(paragraph.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    const keepLines = paginated(fixture({ lineCounts: [1], bodyHeight: 10_000, spacings: [{ before: 5_000 }], properties: [{ keep_lines: true }] }))
    expect(keepLines.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([[], ['paragraph:1']])
    expect(keepLines.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    const keepNext = paginated(fixture({
      lineCounts: [1, 1], bodyHeight: 20_000,
      spacings: [{ before: 5_000 }, {}], properties: [{ keep_next: true }, {}],
    }))
    expect(keepNext.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([[], ['paragraph:1', 'paragraph:2']])
    expect(keepNext.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)

    for (const before of [5_000, 20_000]) {
      const widow = paginated(fixture({ lineCounts: [3], bodyHeight: 30_000, spacings: [{ before }] }))
      expect(widow.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([[], ['paragraph:1', 'paragraph:1', 'paragraph:1']])
      expect(widow.pages[1]!.paragraph_slices[0]!.space_before_millipoints).toBe(0)
    }
  })

  it('splits a paragraph without breaking shaped clusters and preserves line provenance', () => {
    const value = paginated(fixture({ lineCounts: [4], bodyHeight: 20_000 }))
    expect(value.pages).toHaveLength(2)
    expect(value.pages.map((page) => page.lines.map((line) => line.line_id))).toEqual([
      ['line:paragraph:1:0', 'line:paragraph:1:1'],
      ['line:paragraph:1:2', 'line:paragraph:1:3'],
    ])
    expect(value.pages[0]!.paragraph_slices[0]).toEqual(expect.objectContaining({ continued_from_previous_page: false, continues_on_next_page: true }))
    expect(value.pages[1]!.paragraph_slices[0]).toEqual(expect.objectContaining({ continued_from_previous_page: true, continues_on_next_page: false }))
  })

  it('accounts for vertical square-wrap exclusions while selecting a page slice', () => {
    const request = fixture({ lineCounts: [3, 1], bodyHeight: 30_000, properties: [{}, { widow_control: false }] })
    request.shaped_lines.paragraphs[1]!.lines[0]!.exclusion_start_millipoints = 0
    request.shaped_lines.paragraphs[1]!.lines[0]!.exclusion_end_millipoints = 15_000
    const output = paginated(request)
    expect(output.pages.map((page) => page.lines.map((line) => line.y_millipoints))).toEqual([[5_000, 15_000, 25_000], [20_000]])
  })

  it('honors page_break_before and keep_lines', () => {
    const pageBreak = paginated(fixture({ lineCounts: [1, 1], properties: [{}, { page_break_before: true }] }))
    expect(pageBreak.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([['paragraph:1'], ['paragraph:2']])

    const kept = paginated(fixture({ lineCounts: [2, 2], bodyHeight: 30_000, properties: [{}, { keep_lines: true }] }))
    expect(kept.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([
      ['paragraph:1', 'paragraph:1'],
      ['paragraph:2', 'paragraph:2'],
    ])
  })

  it('moves a keep_next chain together when it fits a fresh page', () => {
    const value = paginated(fixture({ lineCounts: [2, 1, 1], bodyHeight: 30_000, properties: [{}, { keep_next: true }, {}] }))
    expect(value.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([
      ['paragraph:1', 'paragraph:1'],
      ['paragraph:2', 'paragraph:3'],
    ])

    const largePriorSpacing = paginated(fixture({
      lineCounts: [1, 1, 1], bodyHeight: 30_000,
      spacings: [{ after: 50_000 }, {}, {}], properties: [{}, { keep_next: true }, {}],
    }))
    expect(largePriorSpacing.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([
      ['paragraph:1'], ['paragraph:2', 'paragraph:3'],
    ])
  })

  it('refuses splittable multiline keep_next chains but accepts explicitly indivisible members', () => {
    const splittable = fixture({ lineCounts: [1, 3, 1], bodyHeight: 40_000, properties: [{}, { keep_next: true }, {}] })
    const refused = paginateNativeDocxV1(splittable)
    expect(refused).toEqual(expect.objectContaining({
      ok: true,
      value: expect.objectContaining({
        status: 'refused', pages: [], sections: [],
        diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'keep-chain-unsatisfiable' })]),
      }),
    }))

    const indivisible = paginated(fixture({
      lineCounts: [1, 2, 1], bodyHeight: 30_000,
      properties: [{}, { keep_next: true, keep_lines: true }, {}],
    }))
    expect(indivisible.pages.map((page) => page.lines.map((line) => line.paragraph_id))).toEqual([
      ['paragraph:1'], ['paragraph:2', 'paragraph:2', 'paragraph:3'],
    ])
  })

  it('plans long keep_next chains in a single bounded reverse pass', () => {
    const paragraphCount = 500
    const request = fixture({
      lineCounts: Array.from({ length: paragraphCount }, () => 1),
      lineHeight: 1,
      bodyHeight: paragraphCount,
      properties: Array.from({ length: paragraphCount }, (_, index) => index + 1 < paragraphCount ? { keep_next: true } : {}),
    })
    expect(paginated(request).pages[0]!.lines).toHaveLength(paragraphCount)
    const source = readFileSync(new URL('./nativePaginationV1.ts', import.meta.url), 'utf8')
    expect(source).toContain('function planKeepChains(')
    expect(source).not.toContain('function keepChainHeight(')
    expect(source).not.toContain('function keepChainEnd(')
  })

  it('refuses unsatisfiable keep-lines, keep-next, and widow/orphan constraints without partial pages', () => {
    for (const request of [
      fixture({ lineCounts: [3], bodyHeight: 20_000, properties: [{ keep_lines: true }] }),
      fixture({ lineCounts: [3], bodyHeight: 20_000 }),
      fixture({ lineCounts: [2, 2], bodyHeight: 30_000, properties: [{ keep_next: true }, {}] }),
    ]) {
      const result = paginateNativeDocxV1(request)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value).toEqual(expect.objectContaining({ status: 'refused', pages: [], sections: [] }))
      expect(result.value.diagnostics.some((entry) => entry.severity === 'unsupported')).toBe(true)
    }
  })

  it('starts page sections and inserts deterministic odd/even parity blanks', () => {
    const odd = paginated(fixture({ lineCounts: [1, 1], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, breakType: 'odd-page', bodyWidth: 40_000, bodyHeight: 40_000 },
    ] }))
    expect(odd.pages.map((page) => [page.kind, page.section_id, page.ordinal])).toEqual([
      ['content', 'section:1', 0],
      ['parity-blank', 'section:1', 1],
      ['content', 'section:2', 2],
    ])
    expect(odd.sections[0]!.page_ids).toEqual(['page:section:1:0', 'page:parity-before:section:2'])
    expect(odd.sections[1]!.page_ids).toEqual(['page:section:2:0'])
    expect(odd.pages[1]).toEqual(expect.objectContaining({ parity_before_section_id: 'section:2', header_refs: [], footer_refs: [] }))

    const even = paginated(fixture({ lineCounts: [1, 1], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, breakType: 'even-page', bodyWidth: 40_000, bodyHeight: 40_000 },
    ] }))
    expect(even.pages.map((page) => page.kind)).toEqual(['content', 'content'])

    const evenWithBlank = paginated(fixture({ lineCounts: [4, 1], bodyHeight: 20_000, sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 20_000 },
      { start: 1, breakType: 'even-page', bodyWidth: 40_000, bodyHeight: 20_000 },
    ] }))
    expect(evenWithBlank.pages.map((page) => [page.kind, page.ordinal])).toEqual([
      ['content', 0], ['content', 1], ['parity-blank', 2], ['content', 3],
    ])

    const next = paginated(fixture({ lineCounts: [1, 1], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, breakType: 'next-page', bodyWidth: 40_000, bodyHeight: 40_000 },
    ] }))
    expect(next.pages.map((page) => page.kind)).toEqual(['content', 'content'])
  })

  it('carries continuous section and column identity on one exact shared page', () => {
    const request = fixture({ lineCounts: [1, 1], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'continuous' },
    ] })
    const output = paginated(request)
    expect(output.pages).toHaveLength(1)
    expect(output.pages[0]!.section_ids).toEqual(['section:1', 'section:2'])
    expect(output.sections.map((section) => section.page_ids)).toEqual([[output.pages[0]!.id], [output.pages[0]!.id]])
    expect(output.pages[0]!.lines.map((line) => [line.section_id, line.column_id, line.column_ordinal])).toEqual([
      ['section:1', 'column:section:1:0', 0],
      ['section:2', 'column:section:2:0', 0],
    ])

    const spaced = fixture({ lineCounts: [1, 1], spacings: [{ after: 4_000 }, { before: 6_000 }], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'continuous' },
    ] })
    expect(paginated(spaced).pages[0]!.lines.map((line) => line.y_millipoints)).toEqual([5_000, 21_000])
  })

  it('starts an exactly exhausted continuous section on its own ordinal-zero page', () => {
    const request = fixture({ lineCounts: [4, 1], bodyHeight: 40_000, sections: [
      { start: 0, breakType: 'next-page', bodyHeight: 40_000 },
      { start: 1, breakType: 'continuous', bodyHeight: 40_000 },
    ] })
    const output = paginated(request)
    expect(output.pages.map((page) => [page.id, page.section_ids, page.section_page_ordinal])).toEqual([
      ['page:section:1:0', ['section:1'], 0],
      ['page:section:2:0', ['section:2'], 0],
    ])
    expect(output.sections.map((section) => section.page_ids)).toEqual([['page:section:1:0'], ['page:section:2:0']])
  })

  it('balances a terminal fragment across exact equal-width columns', () => {
    const request = fixture({ lineCounts: [1, 1, 1, 1], bodyHeight: 40_000 })
    const columnWidth = setEqualColumns(request, 2)
    const output = paginated(request)
    expect(output.pages).toHaveLength(1)
    expect(output.pages[0]!.columns.map((column) => [column.id, column.x_millipoints, column.width_millipoints])).toEqual([
      ['column:section:1:0', 5_000, columnWidth],
      ['column:section:1:1', 27_500, columnWidth],
    ])
    expect(output.pages[0]!.lines.map((line) => [line.column_ordinal, line.y_millipoints])).toEqual([
      [0, 5_000], [0, 15_000], [1, 5_000], [1, 15_000],
    ])
    expect(JSON.stringify(paginated(request))).toBe(JSON.stringify(output))
  })

  it('balances multiline paragraphs and preserves exact source slices across columns and pages', () => {
    const request = fixture({ lineCounts: [10], bodyHeight: 40_000, properties: [{ widow_control: false }] })
    setEqualColumns(request, 2)
    const output = paginated(request)
    expect(output.pages.map((page) => page.lines.map((line) => line.source_line_ordinal))).toEqual([
      [0, 1, 2, 3, 4, 5, 6, 7], [8, 9],
    ])
    expect(output.pages.flatMap((page) => page.paragraph_slices.map((slice) => [
      slice.first_line_ordinal, slice.last_line_ordinal, slice.continued_from_previous_page,
      slice.continues_on_next_page, slice.continued_from_previous_column, slice.continues_in_next_column,
    ]))).toEqual([
      [0, 3, false, false, false, true], [4, 7, false, true, true, false],
      [8, 8, true, false, false, true], [9, 9, false, false, true, false],
    ])
    expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)
    expect(paginated(request)).toEqual(output)
  })

  it('balances mixed paragraphs when the ideal plan respects keep and default widow constraints', () => {
    const request = fixture({ lineCounts: [2, 4, 2], bodyHeight: 40_000, properties: [{ keep_lines: true }, {}, { keep_lines: true }] })
    setEqualColumns(request, 2)
    const output = paginated(request)
    expect(output.pages[0]!.paragraph_slices.map((slice) => [slice.paragraph_id, slice.column_ordinal, slice.first_line_ordinal, slice.last_line_ordinal])).toEqual([
      ['paragraph:1', 0, 0, 1], ['paragraph:2', 0, 0, 1], ['paragraph:2', 1, 2, 3], ['paragraph:3', 1, 0, 1],
    ])
  })

  it('balances default-widow multiline continuations and a next-column terminal section', () => {
    const request = fixture({ lineCounts: [12], bodyHeight: 40_000 })
    setEqualColumns(request, 2)
    expect(paginated(request).pages.map((page) => page.paragraph_slices.map((slice) => slice.line_ids.length))).toEqual([[4, 4], [2, 2]])

    const nextColumn = fixture({ lineCounts: [1, 4], sections: [
      { start: 0, breakType: 'next-page' }, { start: 1, breakType: 'next-column' },
    ] })
    setEqualColumns(nextColumn, 2)
    expect(paginated(nextColumn).pages[0]!.lines.map((line) => line.column_ordinal)).toEqual([0, 1, 1, 1, 1])
  })

  it('refuses constrained ideal splits and accepts explicitly disabled widow control', () => {
    for (const properties of [{}, { widow_control: true }, { widow_control: false, keep_lines: true }]) {
      const request = fixture({ lineCounts: [3], properties: [properties] })
      setEqualColumns(request, 2)
      expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({
        status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'column-balance-ambiguous' })]),
      }) }))
    }
    const request = fixture({ lineCounts: [3], properties: [{ widow_control: false }] })
    setEqualColumns(request, 2)
    expect(paginated(request).pages[0]!.lines.map((line) => line.column_ordinal)).toEqual([0, 0, 1])
  })

  it('retains note-free multicolumn support and refuses only note-bearing ambiguous grids atomically', () => {
    const multicolumn = fixture({ lineCounts: [1, 1, 1, 1], bodyHeight: 40_000 })
    setEqualColumns(multicolumn, 2)
    expect(paginated(structuredClone(multicolumn)).pages[0]!.columns).toHaveLength(2)
    addFootnote(multicolumn)
    expect(paginateNativeDocxV1(multicolumn)).toEqual(expect.objectContaining({
      ok: true,
      value: expect.objectContaining({
        status: 'refused', pages: [], sections: [],
        diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'note-structure-unsupported', message: expect.stringContaining('note-free pages') })]),
      }),
    }))

    const sharedContinuous = fixture({ lineCounts: [1, 1], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'continuous' },
    ] })
    expect(paginated(structuredClone(sharedContinuous)).pages[0]!.section_ids).toEqual(['section:1', 'section:2'])
    addFootnote(sharedContinuous)
    expect(paginateNativeDocxV1(sharedContinuous)).toEqual(expect.objectContaining({
      ok: true,
      value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'note-structure-unsupported' })]) }),
    }))
  })

  it('uses a unique quotient/remainder terminal balance after sequential full pages', () => {
    const remainder = fixture({ lineCounts: [1, 1, 1, 1, 1], bodyHeight: 40_000 })
    setEqualColumns(remainder, 2)
    expect(paginated(remainder).pages[0]!.lines.map((line) => line.column_ordinal)).toEqual([0, 0, 0, 1, 1])

    const multiPage = fixture({ lineCounts: Array.from({ length: 9 }, () => 1), bodyHeight: 40_000 })
    setEqualColumns(multiPage, 2)
    expect(paginated(multiPage).pages.map((page) => page.lines.map((line) => line.column_ordinal))).toEqual([
      [0, 0, 0, 0, 1, 1, 1, 1],
      [0],
    ])

    const sparse = fixture({ lineCounts: [1], bodyHeight: 40_000 })
    setEqualColumns(sparse, 3)
    expect(paginated(sparse).pages[0]!.lines.map((line) => line.column_ordinal)).toEqual([0])
  })

  it('places next-column on the following column and advances to a new page after column exhaustion', () => {
    const middle = fixture({ lineCounts: [1, 1], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'next-column' },
    ] })
    setEqualColumns(middle, 2)
    const middleOutput = paginated(middle)
    expect(middleOutput.pages).toHaveLength(1)
    expect(middleOutput.pages[0]!.lines.map((line) => [line.section_id, line.column_ordinal])).toEqual([['section:1', 0], ['section:2', 1]])

    const expanded = fixture({ lineCounts: [1, 1, 1, 1, 1, 1], lineHeight: 10_000, bodyHeight: 40_000, sections: [
      { start: 0, breakType: 'next-page', bodyHeight: 40_000 },
      { start: 5, breakType: 'next-column', bodyHeight: 40_000 },
    ] })
    setEqualColumns(expanded, 2)
    const exhaustedOutput = paginated(expanded)
    expect(exhaustedOutput.pages).toHaveLength(2)
    expect(exhaustedOutput.pages[1]!.lines[0]).toEqual(expect.objectContaining({ section_id: 'section:2', column_ordinal: 0 }))
  })

  it('integrates concrete numbered markers with multicolumn balance, continuous, and next-column sections', () => {
    const balanced = numberedFixture({ lineCounts: [1, 1, 1, 1], bodyHeight: 40_000 })
    setEqualColumns(balanced, 2)
    const balancedOutput = paginated(balanced)
    expect(balancedOutput.pages[0]!.lines.map((line) => [line.paragraph_id, line.column_ordinal])).toEqual([
      ['paragraph:1', 0], ['paragraph:2', 0], ['paragraph:3', 1], ['paragraph:4', 1],
    ])
    expect(balanced.shaped_lines.paragraphs.map((paragraph) => paragraph.list_marker?.counter_value)).toEqual([1, 2, 3, 4])

    const continuous = numberedFixture({ lineCounts: [1, 1], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'continuous' },
    ] })
    const continuousOutput = paginated(continuous)
    expect(continuousOutput.pages).toHaveLength(1)
    expect(continuousOutput.pages[0]!.section_ids).toEqual(['section:1', 'section:2'])
    expect(continuousOutput.pages[0]!.lines.map((line) => [line.section_id, line.column_ordinal])).toEqual([['section:1', 0], ['section:2', 0]])

    const nextColumn = numberedFixture({ lineCounts: [1, 1], sections: [
      { start: 0, breakType: 'next-page' },
      { start: 1, breakType: 'next-column' },
    ] })
    setEqualColumns(nextColumn, 2)
    const nextColumnOutput = paginated(nextColumn)
    expect(nextColumnOutput.pages).toHaveLength(1)
    expect(nextColumnOutput.pages[0]!.lines.map((line) => [line.section_id, line.column_ordinal])).toEqual([['section:1', 0], ['section:2', 1]])
  })

  it('supports explicit equal widths and refuses unequal, ambiguous, or incompatible geometry atomically', () => {
    const explicit = fixture({ lineCounts: [1, 1, 1, 1] })
    const width = setEqualColumns(explicit, 2)
    explicit.document.sections[0]!.page.column_layout = 'explicit'
    explicit.document.sections[0]!.page.column_spacing_twips = 0
    explicit.document.sections[0]!.page.column_definitions = [
      { id: 'column:section:1:0', ordinal: 0, width_twips: width / 50, space_after_twips: 100 },
      { id: 'column:section:1:1', ordinal: 1, width_twips: width / 50, space_after_twips: 0 },
    ]
    expect(paginated(explicit).pages[0]!.columns.map((column) => column.width_millipoints)).toEqual([width, width])

    const ambiguousBalance = fixture({ lineCounts: [3] })
    setEqualColumns(ambiguousBalance, 2)
    expect(paginateNativeDocxV1(ambiguousBalance)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'column-balance-ambiguous' })]) }) }))

    const lateAmbiguousBalance = fixture({ lineCounts: [1, 2], sections: [
      { start: 0, bodyWidth: 17_500, bodyHeight: 40_000 },
      { start: 1, breakType: 'next-page', bodyWidth: 40_000, bodyHeight: 40_000, columns: 2 },
    ] })
    lateAmbiguousBalance.shaped_lines.available_width_millipoints = 17_500
    lateAmbiguousBalance.shaped_lines.paragraphs.forEach((paragraph) => paragraph.lines.forEach((line) => { line.available_width_millipoints = 17_500 }))
    expect(paginateNativeDocxV1(lateAmbiguousBalance)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'column-balance-ambiguous' })]) }) }))

    const unequal = structuredClone(explicit)
    unequal.document.sections[0]!.page.column_definitions[0]!.width_twips! -= 1
    unequal.document.sections[0]!.page.column_definitions[1]!.width_twips! += 1
    expect(paginateNativeDocxV1(unequal)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))

    const continuousGeometry = fixture({ lineCounts: [1, 1], sections: [{ start: 0 }, { start: 1, breakType: 'continuous' }] })
    continuousGeometry.document.sections[1]!.page.margins.right_twips += 1
    expect(paginateNativeDocxV1(continuousGeometry)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))

    const headerGeometry = fixture({ lineCounts: [1, 1], sections: [{ start: 0 }, { start: 1, breakType: 'continuous' }] })
    headerGeometry.document.sections[1]!.page.margins.header_twips += 1
    expect(paginateNativeDocxV1(headerGeometry)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))

    const nondivisibleTwips = fixture({ lineCounts: [1, 1] })
    setEqualColumns(nondivisibleTwips, 2, 99)
    expect(paginateNativeDocxV1(nondivisibleTwips)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'column-geometry-invalid' })]) }) }))

    for (const breakType of ['continuous', 'next-column'] as const) {
      const titlePage = fixture({ lineCounts: [1, 1], sections: [{ start: 0 }, { start: 1, breakType }] })
      titlePage.document.sections[1]!.title_page = true
      expect(paginateNativeDocxV1(titlePage)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'section-geometry-invalid' })]) }) }))
    }
  })

  // A paragraph vanishes from shaped lines because a blocking resolved-layout
  // diagnostic made shaping drop it. Reporting only the absence discards the
  // cause and makes the refusal unactionable.
  it('names the resolved-layout cause when a paragraph is missing from shaped lines', () => {
    const request = fixture({}) as any
    const droppedID = request.shaped_lines.paragraphs[0].paragraph_id
    request.shaped_lines.paragraphs = []
    request.shaped_lines.diagnostics = [{
      code: 'unresolved-layout-diagnostic', severity: 'unsupported', scope_id: droppedID,
      source_diagnostic_code: 'VERTICAL_ALIGNMENT_UNSUPPORTED',
      source_diagnostic_message: 'Vertical alignment requires a future paginator',
      message: 'Resolved layout diagnostic VERTICAL_ALIGNMENT_UNSUPPORTED blocks native shaping',
    }]
    const result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (!result.ok) return
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({
      code: 'shaped-paragraph-missing',
      scope_id: droppedID,
      source_code: 'VERTICAL_ALIGNMENT_UNSUPPORTED',
      source_message: 'Vertical alignment requires a future paginator',
    })]))
  })

  // paint-diagnostic-preserved also carries a source code, but it is emitted for
  // codes that explicitly do not change shaping advances. Attributing a refusal
  // to one names an innocent code and sends the reader at the wrong subsystem.
  it('ignores non-blocking paint-only diagnostics when naming the cause', () => {
    const request = fixture({}) as any
    const droppedID = request.shaped_lines.paragraphs[0].paragraph_id
    request.shaped_lines.paragraphs = []
    request.shaped_lines.diagnostics = [
      {
        code: 'paint-diagnostic-preserved', severity: 'deferred', scope_id: droppedID,
        source_diagnostic_code: 'THEME_COLOR_PRESERVED',
        source_diagnostic_message: 'Automatic color requires presentation context',
        message: 'Paint-only resolved-layout diagnostic does not change shaping advances',
      },
      {
        code: 'unresolved-layout-diagnostic', severity: 'unsupported', scope_id: droppedID,
        source_diagnostic_code: 'UNSUPPORTED_NUMBER_FORMAT',
        source_diagnostic_message: 'Numbering format is not modelled',
        message: 'Resolved layout diagnostic UNSUPPORTED_NUMBER_FORMAT blocks native shaping',
      },
    ]
    const result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (!result.ok) return
    const entry = result.value.diagnostics.find(d => d.code === 'shaped-paragraph-missing' && d.scope_id === droppedID)
    expect(entry).toMatchObject({ source_code: 'UNSUPPORTED_NUMBER_FORMAT' })
  })

  // A paint-only diagnostic alone never blocked shaping, so it must not be named.
  it('names no cause when only a paint-only diagnostic is present', () => {
    const request = fixture({}) as any
    const droppedID = request.shaped_lines.paragraphs[0].paragraph_id
    request.shaped_lines.paragraphs = []
    request.shaped_lines.diagnostics = [{
      code: 'paint-diagnostic-preserved', severity: 'deferred', scope_id: droppedID,
      source_diagnostic_code: 'THEME_COLOR_PRESERVED',
      source_diagnostic_message: 'Automatic color requires presentation context',
      message: 'Paint-only resolved-layout diagnostic does not change shaping advances',
    }]
    const result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (!result.ok) return
    const entry = result.value.diagnostics.find(d => d.code === 'shaped-paragraph-missing' && d.scope_id === droppedID)
    expect(entry).toBeDefined()
    expect(entry).not.toHaveProperty('source_code')
  })

  // With no shaping diagnostic to attribute it to, the refusal stays as it was.
  it('still refuses an unexplained missing shaped paragraph without inventing a cause', () => {
    const request = fixture({}) as any
    const droppedID = request.shaped_lines.paragraphs[0].paragraph_id
    request.shaped_lines.paragraphs = []
    const result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (!result.ok) return
    const entry = result.value.diagnostics.find(d => d.code === 'shaped-paragraph-missing' && d.scope_id === droppedID)
    expect(entry).toBeDefined()
    expect(entry).not.toHaveProperty('source_code')
  })

  it.each([
    ['section width mismatch', (value: any) => { value.document.sections[0].page.margins.right_twips += 1 }, 'section-width-mismatch'],
    ['unsupported settings semantics', (value: any) => {
      value.pagination_settings = {
        ...value.pagination_settings, profile: 'unsupported',
        default_tab_stop_twips: 720, mirror_margins: true, gutter_at_top: false, even_and_odd_headers: false,
        diagnostics: [{ code: 'MIRROR_MARGINS_UNSUPPORTED', severity: 'unsupported', part_name: 'word/settings.xml', path: '/w:settings[1]/w:mirrorMargins[1]', preservation: 'preserve-verbatim', message: 'Mirror margins alter page geometry.' }],
      }
    }, 'settings-attestation-unsupported'],
    ['unsupported source control', (value: any) => { const run = value.document.body.blocks[0].paragraph.runs[0]; delete run.text; run.kind = 'control'; run.control = 'page-break' }, 'source-control-unsupported'],
    ['missing shaped paragraph', (value: any) => { value.shaped_lines.paragraphs = [] }, 'shaped-paragraph-missing'],
    ['keep-next across section', (value: any) => { value.resolved_layout.paragraphs[0].properties.keep_next = true }, 'keep-chain-conflict'],
  ])('refuses %s rather than approximating', (_name, mutate, code) => {
    const twoSections = _name === 'keep-next across section'
    const request = fixture(twoSections ? { lineCounts: [1, 1], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, breakType: 'next-page', bodyWidth: 40_000, bodyHeight: 40_000 },
    ] } : {}) as any
    mutate(request)
    const result = paginateNativeDocxV1(request)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.status).toBe('refused')
    expect(result.value.pages).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code })]))
  })

  it('omits an unshaped comment/drawing paragraph in approximate layout and keeps the sibling paragraph', () => {
    const request = fixture({ lineCounts: [1, 1] })
    const dropped = request.document.body.blocks[0]!.paragraph!
    dropped.runs = [{
      kind: 'reference', id: 'run:comment-start',
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:commentRangeStart[1]', 110, 120),
      reference: { kind: 'comment-range-start', target_id: 'comment:1' },
    }, {
      kind: 'drawing', id: 'run:comment-drawing',
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 121, 180),
      drawing: {
        id: 'drawing:1',
        anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 130, 170),
        placement: 'inline', width_emu: 914_400, height_emu: 914_400,
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'DRAWING_EFFECTS_UNSUPPORTED', message: 'Preserve the original drawing.', preservation: 'refuse-mutation' } },
      },
    }]
    const commentAnchor = { part_name: 'word/comments.xml', path: '/w:comments[1]/w:comment[1]', start_byte: 10, end_byte: 400, xml_sha256: HASH }
    const storyAnchor = { part_name: 'word/comments.xml', path: '/w:comments[1]/w:comment[1]/w:p[1]', start_byte: 20, end_byte: 300, xml_sha256: HASH }
    const commentBody = paragraph('paragraph:comment-body', 9)
    commentBody.anchor = storyAnchor
    commentBody.runs[0]!.anchor = { ...storyAnchor, path: `${storyAnchor.path}/w:r[1]`, start_byte: 30, end_byte: 80 }
    request.document.comment_stories = [{
      id: 'story:comment:1', kind: 'comment', part_name: 'word/comments.xml', native_story_id: '1',
      anchor: storyAnchor, blocks: [{ kind: 'paragraph', id: commentBody.id, paragraph: commentBody }],
    }]
    request.document.comments = [{ id: 'comment:1', native_comment_id: '1', author: 't', anchor: commentAnchor, body_story_id: 'story:comment:1' }]
    request.resolved_layout.paragraphs.push({ paragraph_id: commentBody.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } })
    request.resolved_layout.runs = request.resolved_layout.runs.filter(run => run.run_id !== `run:${dropped.id}`)
    request.resolved_layout.runs.push(
      { run_id: 'run:comment-start', paragraph_id: dropped.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } },
      { run_id: 'run:comment-drawing', paragraph_id: dropped.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } },
      { run_id: commentBody.runs[0]!.id, paragraph_id: commentBody.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } },
    )
    request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== dropped.id)
    request.pagination_settings.profile = 'unsupported'
    delete request.pagination_settings.compatibility_mode
    request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
    const strict = paginateNativeDocxV1(request)
    expect(strict, JSON.stringify(strict)).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (strict.ok) expect(strict.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'shaped-paragraph-missing', scope_id: dropped.id })]))
    const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(approximate.layout.status).toBe('paginated')
    expect(approximate.layout.pages.flatMap(page => page.lines.map(line => line.paragraph_id))).toEqual(['paragraph:2'])
  })

  it('omits an unshaped picture or partial-run paragraph in approximate layout and keeps the sibling paragraph', () => {
    for (const code of ['PICTURE_GRAPHIC_REQUIRED', 'PARTIAL_RUN_PROPERTIES'] as const) {
      const request = fixture({ lineCounts: [1, 1] })
      const dropped = request.document.body.blocks[0]!.paragraph!
      dropped.runs = []
      request.resolved_layout.runs = request.resolved_layout.runs.filter(run => run.paragraph_id !== dropped.id)
      request.document.unsupported.push({ id: `unsupported:${code}`, code, capability: 'drawings', scope_id: dropped.id, preservation: 'refuse-mutation', message: code })
      request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== dropped.id)
      request.pagination_settings.profile = 'unsupported'
      delete request.pagination_settings.compatibility_mode
      request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
      const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
      const strict = paginateNativeDocxV1(request)
      expect(strict, code).toMatchObject({ ok: true, value: { status: 'refused' } })
      if (strict.ok) expect(strict.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'shaped-paragraph-missing', scope_id: dropped.id })]))
      const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
      expect(approximate.layout.status, code).toBe('paginated')
      expect(approximate.layout.pages.flatMap(page => page.lines.map(line => line.paragraph_id))).toEqual(['paragraph:2'])
    }
  })

  it('paints around a tracked-move marker that states no content or formatting', () => {
    // A dragged table row leaves content-free range endpoints between rows and a
    // CT_TrackChange annotation on each paragraph mark. Neither states content
    // or formatting, so the approximate tier paints the paragraphs around them.
    for (const code of ['NON_VISUAL_RANGE_MARKER', 'TRACKED_MARK_REVISION_PRESERVED'] as const) {
      const request = fixture({ lineCounts: [1, 1] })
      const marked = request.document.body.blocks[0]!.paragraph!
      request.document.unsupported.push({ id: `unsupported:${code}`, code, capability: 'table-structure', scope_id: marked.id, preservation: 'refuse-mutation', message: code })
      request.pagination_settings.profile = 'unsupported'
      delete request.pagination_settings.compatibility_mode
      request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
      const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
      const strict = paginateNativeDocxV1(request)
      expect(strict, code).toMatchObject({ ok: true, value: { status: 'refused' } })
      if (strict.ok) expect(strict.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'body-structure-unsupported', scope_id: marked.id })]))
      const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
      expect(approximate.layout.status, code).toBe('paginated')
      expect(approximate.layout.pages.flatMap(page => page.lines.map(line => line.paragraph_id))).toEqual([marked.id, 'paragraph:2'])
    }
  })

  it('omits an unshaped empty-run sibling paragraph in approximate layout and keeps the sibling paragraph', () => {
    const request = fixture({ lineCounts: [1, 1] })
    const dropped = request.document.body.blocks[0]!.paragraph!
    dropped.runs = []
    request.resolved_layout.runs = request.resolved_layout.runs.filter(run => run.paragraph_id !== dropped.id)
    request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== dropped.id)
    request.pagination_settings.profile = 'unsupported'
    delete request.pagination_settings.compatibility_mode
    request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
    const strict = paginateNativeDocxV1(request)
    expect(strict).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (strict.ok) expect(strict.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'shaped-paragraph-missing', scope_id: dropped.id })]))
    const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(approximate.layout.status).toBe('paginated')
    expect(approximate.layout.pages.flatMap(page => page.lines.map(line => line.paragraph_id))).toEqual(['paragraph:2'])
  })

  it('still refuses an unshaped text-only paragraph in approximate layout', () => {
    const request = fixture({ lineCounts: [1, 1] })
    const dropped = request.document.body.blocks[0]!.paragraph!
    request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== dropped.id)
    request.pagination_settings.profile = 'unsupported'
    delete request.pagination_settings.compatibility_mode
    request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
    const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(approximate.layout.status).toBe('refused')
    expect(approximate.layout.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'shaped-paragraph-missing', scope_id: dropped.id })]))
  })

  it('still refuses an unshaped paragraph that keeps remaining text beside comment markers', () => {
    const request = fixture({ lineCounts: [1, 1] })
    const dropped = request.document.body.blocks[0]!.paragraph!
    dropped.runs = [{
      kind: 'reference', id: 'run:comment-start',
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:commentRangeStart[1]', 110, 120),
      reference: { kind: 'comment-range-start', target_id: 'comment:1' },
    }, dropped.runs[0]!]
    const commentAnchor = { part_name: 'word/comments.xml', path: '/w:comments[1]/w:comment[1]', start_byte: 10, end_byte: 400, xml_sha256: HASH }
    const storyAnchor = { part_name: 'word/comments.xml', path: '/w:comments[1]/w:comment[1]/w:p[1]', start_byte: 20, end_byte: 300, xml_sha256: HASH }
    const commentBody = paragraph('paragraph:comment-body', 9)
    commentBody.anchor = storyAnchor
    commentBody.runs[0]!.anchor = { ...storyAnchor, path: `${storyAnchor.path}/w:r[1]`, start_byte: 30, end_byte: 80 }
    request.document.comment_stories = [{
      id: 'story:comment:1', kind: 'comment', part_name: 'word/comments.xml', native_story_id: '1',
      anchor: storyAnchor, blocks: [{ kind: 'paragraph', id: commentBody.id, paragraph: commentBody }],
    }]
    request.document.comments = [{ id: 'comment:1', native_comment_id: '1', author: 't', anchor: commentAnchor, body_story_id: 'story:comment:1' }]
    request.resolved_layout.paragraphs.push({ paragraph_id: commentBody.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } })
    request.resolved_layout.runs.push(
      { run_id: 'run:comment-start', paragraph_id: dropped.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } },
      { run_id: commentBody.runs[0]!.id, paragraph_id: commentBody.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } },
    )
    request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== dropped.id)
    request.pagination_settings.profile = 'unsupported'
    delete request.pagination_settings.compatibility_mode
    request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 14 as const, reasons: ['Legacy mode 14 uses current layout'] }
    const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(approximate.layout.status).toBe('refused')
    expect(approximate.layout.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'shaped-paragraph-missing', scope_id: dropped.id })]))
  })

  it('refuses body tables, drawings, and note references from source authority even when shaped output omits them', () => {
    const table = fixture() as any
    table.document.body.blocks = [{
      kind: 'table', id: 'table:1', table: {
        id: 'table:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 100, 190),
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Table placement unavailable.', preservation: 'refuse-mutation' } },
        rows: [],
      },
    }]
    table.document.sections[0].starts_at_block_id = 'table:1'
    table.resolved_layout.paragraphs = []
    table.resolved_layout.runs = []
    table.resolved_layout.tables = [{ table_id: 'table:1' }]
    table.shaped_lines.paragraphs = []

    const drawing = fixture() as any
    const drawingRun = drawing.document.body.blocks[0].paragraph.runs[0]
    delete drawingRun.text
    drawingRun.kind = 'drawing'
    drawingRun.drawing = {
      id: 'drawing:1',
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 120, 170),
      placement: 'floating', width_emu: 914400, height_emu: 914400, x_emu: 0, y_emu: 0,
      horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'square',
      edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Drawing placement unavailable.', preservation: 'refuse-mutation' } },
    }

    const note = fixture() as any
    const noteRun = note.document.body.blocks[0].paragraph.runs[0]
    delete noteRun.text
    noteRun.kind = 'reference'
    noteRun.reference = { kind: 'footnote', target_id: 'story:footnote:1' }
    note.document.notes = [{
      id: 'story:footnote:1', kind: 'footnote', native_story_id: '1', relationship_id: 'rIdFootnotes', note_role: 'content', part_name: 'word/footnotes.xml',
      anchor: { ...anchor('/w:footnotes[1]/w:footnote[1]', 1, 90), part_name: 'word/footnotes.xml' }, blocks: [],
    }]

    for (const [request, code] of [[table, 'body-table-unsupported'], [drawing, 'body-structure-unsupported'], [note, 'note-reference-ambiguous']] as const) {
      const result = paginateNativeDocxV1(request)
      expect(result).toEqual(expect.objectContaining({
        ok: true,
        value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code })]) }),
      }))
    }
  })

  it('derives guttered landscape geometry exactly and rejects contradictory or escaping geometry atomically', () => {
    const landscape = fixture()
    landscape.document.sections[0]!.page = {
      width_twips: 2_000, height_twips: 1_500, orientation: 'landscape', columns: 1, column_spacing_twips: 100, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
      margins: { top_twips: 100, right_twips: 300, bottom_twips: 200, left_twips: 200, header_twips: 50, footer_twips: 50, gutter_twips: 100 },
    }
    landscape.shaped_lines.available_width_millipoints = 70_000
    landscape.shaped_lines.paragraphs[0]!.lines[0]!.available_width_millipoints = 70_000
    const placed = paginated(landscape)
    expect(placed.pages[0]).toEqual(expect.objectContaining({
      width_millipoints: 100_000, height_millipoints: 75_000,
      body_box: { x_millipoints: 15_000, y_millipoints: 5_000, width_millipoints: 70_000, height_millipoints: 60_000 },
    }))

    const adversaries: Array<[NativeDocxPaginationRequestV1, string]> = []
    const orientation = fixture(); orientation.document.sections[0]!.page.height_twips += 1; orientation.document.sections[0]!.page.orientation = 'landscape'; adversaries.push([orientation, 'section-geometry-invalid'])
    const margins = fixture(); margins.document.sections[0]!.page.margins.right_twips = 1_000; adversaries.push([margins, 'section-geometry-invalid'])
    const coordinateOverflow = fixture(); coordinateOverflow.document.sections[0]!.page.width_twips = DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints / 50 + 1; adversaries.push([coordinateOverflow, 'section-geometry-invalid'])
    const leftEscape = fixture(); leftEscape.shaped_lines.paragraphs[0]!.lines[0]!.inline_offset_millipoints = -1
    expect(paginateNativeDocxV1(leftEscape)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/shaped_lines/paragraphs/0/lines/0/inline_offset_millipoints' })]) }))
    const rightEscape = fixture(); rightEscape.shaped_lines.paragraphs[0]!.lines[0]!.advance_inline_millipoints = 40_001; adversaries.push([rightEscape, 'line-geometry-invalid'])
    for (const [request, code] of adversaries) {
      const result = paginateNativeDocxV1(request)
      expect(result).toEqual(expect.objectContaining({
        ok: true,
        value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code })]) }),
      }))
    }
  })

  it('rejects malformed request keys and cross-contract identity drift', () => {
    const unknown = { ...fixture(), css_width: 10 }
    expect(paginateNativeDocxV1(unknown)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'UNKNOWN_FIELD' })]) }))
    const drift = fixture()
    drift.shaped_lines.revision = 'revision:other'
    expect(paginateNativeDocxV1(drift)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))
  })

  it('strictly validates paginated output identities, joins, geometry, exclusivity, null, and negative zero', () => {
    const output = paginated(fixture({ lineCounts: [2] }))
    const mutations = [
      (value: any) => { value.pages[0].extra = true },
      (value: any) => { value.pages[0].lines[0].id = 'placed:forged' },
      (value: any) => { value.pages[0].paragraph_slices[0].line_ids[0] = 'line:missing:0' },
      (value: any) => { value.pages[0].body_box.x_millipoints = -0 },
      (value: any) => { value.pages[0].lines[0].y_millipoints = null },
      (value: any) => { value.provenance.shaped_lines.tab_interval_millipoints = 18_000 },
      (value: any) => { value.provenance.pagination_settings.package_sha256 = `sha256:${'c'.repeat(64)}` },
      (value: any) => {
        value.pages[0].columns[0].id = 'column:forged'
        value.pages[0].lines.forEach((line: any) => { line.column_id = 'column:forged' })
        value.pages[0].paragraph_slices.forEach((slice: any) => { slice.column_id = 'column:forged' })
      },
      (value: any) => {
        value.provenance.pagination_settings.profile = 'absent-default'
        for (const field of ['relationships_part', 'relationships_sha256', 'relationship_id', 'settings_part', 'settings_sha256', 'compatibility_mode']) delete value.provenance.pagination_settings[field]
      },
      (value: any) => { value.status = 'refused'; value.diagnostics = []; },
      (value: any) => { value.diagnostics.push({ code: 'not-enumerated', severity: 'deferred', scope_id: 'document:test', message: 'forged' }) },
    ]
    for (const mutate of mutations) {
      const invalid = structuredClone(output) as any
      mutate(invalid)
      expect(decodeNativeDocxPaginatedLayout(invalid).ok).toBe(false)
    }
  })

  it('caps merged paginated-output and nested-settings issues at the public validation bound', () => {
    const output = structuredClone(paginated(fixture())) as any
    output.extra = true
    output.provenance.pagination_settings.diagnostics = Array.from({ length: DOCX_NATIVE_LIMITS.maxIssues + 1 }, () => null)
    const decoded = decodeNativeDocxPaginatedLayout(output)
    expect(decoded.ok).toBe(false)
    if (!decoded.ok) expect(decoded.issues).toHaveLength(DOCX_NATIVE_LIMITS.maxIssues)
  })

  it('requires exact global line coverage and contiguous paragraph slices across pages', () => {
    const onePage = paginated(fixture({ lineCounts: [2] }))
    const extraLine = structuredClone(onePage) as any
    extraLine.pages[0].lines.push({ ...extraLine.pages[0].lines[1], id: 'placed:line:extra', line_id: 'line:extra', source_line_ordinal: 2, y_millipoints: 25_000 })
    expect(decodeNativeDocxPaginatedLayout(extraLine).ok).toBe(false)

    const duplicateSlice = structuredClone(onePage) as any
    duplicateSlice.pages[0].paragraph_slices[0].continues_on_next_page = true
    duplicateSlice.pages[0].paragraph_slices.push({
      ...duplicateSlice.pages[0].paragraph_slices[0], id: 'slice:paragraph:1:1', slice_ordinal: 1,
      line_ids: [duplicateSlice.pages[0].lines[1].line_id], first_line_ordinal: 1, last_line_ordinal: 1,
      top_millipoints: duplicateSlice.pages[0].lines[1].y_millipoints, height_millipoints: duplicateSlice.pages[0].lines[1].height_millipoints,
      continued_from_previous_page: true, continues_on_next_page: false,
    })
    expect(decodeNativeDocxPaginatedLayout(duplicateSlice).ok).toBe(false)

    const reorderedSlices = structuredClone(paginated(fixture({ lineCounts: [1, 1] }))) as any
    reorderedSlices.pages[0].paragraph_slices.reverse()
    expect(decodeNativeDocxPaginatedLayout(reorderedSlices).ok).toBe(false)
    const overlappingSlices = structuredClone(paginated(fixture({ lineCounts: [1, 1] }))) as any
    overlappingSlices.pages[0].paragraph_slices[1].top_millipoints = overlappingSlices.pages[0].paragraph_slices[0].top_millipoints
    expect(decodeNativeDocxPaginatedLayout(overlappingSlices).ok).toBe(false)

    const split = paginated(fixture({ lineCounts: [4], bodyHeight: 20_000 }))
    for (const mutate of [
      (value: any) => { value.pages[1].paragraph_slices[0].slice_ordinal = 2; value.pages[1].paragraph_slices[0].id = 'slice:paragraph:1:2' },
      (value: any) => { value.pages[1].paragraph_slices[0].first_line_ordinal = 3 },
      (value: any) => { value.pages[0].paragraph_slices[0].continues_on_next_page = false },
      (value: any) => { value.pages[1].paragraph_slices[0].continued_from_previous_page = false },
      (value: any) => { value.pages[0].lines.reverse() },
    ]) {
      const invalid = structuredClone(split) as any
      mutate(invalid)
      expect(decodeNativeDocxPaginatedLayout(invalid).ok).toBe(false)
    }
  })

  it('joins paginated output to the complete shaped source, sections, geometry, and provenance', () => {
    const request = fixture({ lineCounts: [1, 1] })
    const output = paginated(request)
    expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)

    const omitted = structuredClone(output) as any
    omitted.pages[0].lines.splice(1, 1)
    omitted.pages[0].paragraph_slices.splice(1, 1)
    expect(decodeNativeDocxPaginatedLayout(omitted).ok).toBe(true)
    expect(decodeNativeDocxPaginatedLayoutForRequest(omitted, request).ok).toBe(false)

    const substituted = structuredClone(output) as any
    const line = substituted.pages[0].lines[1]
    line.id = 'placed:line:forged:0'; line.line_id = 'line:forged:0'; line.paragraph_id = 'paragraph:forged'; line.source_line_ordinal = 0
    const slice = substituted.pages[0].paragraph_slices[1]
    slice.id = 'slice:paragraph:forged:0'; slice.paragraph_id = 'paragraph:forged'; slice.slice_ordinal = 0; slice.first_line_ordinal = 0; slice.last_line_ordinal = 0; slice.line_ids = ['line:forged:0']
    expect(decodeNativeDocxPaginatedLayout(substituted).ok).toBe(true)
    expect(decodeNativeDocxPaginatedLayoutForRequest(substituted, request).ok).toBe(false)

    for (const mutate of [
      (value: any) => { value.sections[0].starts_at_block_id = 'paragraph:forged' },
      (value: any) => { value.sections[0].break_type = 'even-page' },
      (value: any) => { value.pages[0].width_millipoints += 100; value.pages[0].body_box.width_millipoints += 100 },
      (value: any) => { value.provenance.font_manifest.revision = 'manifest-revision:forged' },
    ]) {
      const forged = structuredClone(output) as any
      mutate(forged)
      expect(decodeNativeDocxPaginatedLayout(forged).ok).toBe(true)
      expect(decodeNativeDocxPaginatedLayoutForRequest(forged, request).ok).toBe(false)
    }

    const verticalGap = structuredClone(output) as any
    verticalGap.pages[0].lines[1].y_millipoints += 1_000
    verticalGap.pages[0].paragraph_slices[1].top_millipoints += 1_000
    expect(decodeNativeDocxPaginatedLayout(verticalGap).ok).toBe(true)
    expect(decodeNativeDocxPaginatedLayoutForRequest(verticalGap, request).ok).toBe(false)

    const forgedSpacing = structuredClone(output) as any
    forgedSpacing.pages[0].paragraph_slices[1].space_before_millipoints += 1_000
    expect(decodeNativeDocxPaginatedLayout(forgedSpacing).ok).toBe(true)
    expect(decodeNativeDocxPaginatedLayoutForRequest(forgedSpacing, request).ok).toBe(false)

    const breakRequest = fixture({ lineCounts: [6], bodyHeight: 40_000 })
    const alteredBreak = structuredClone(paginated(breakRequest)) as any
    const movedLine = alteredBreak.pages[0].lines.pop()
    alteredBreak.pages[0].paragraph_slices[0].line_ids.pop()
    alteredBreak.pages[0].paragraph_slices[0].last_line_ordinal = 2
    alteredBreak.pages[0].paragraph_slices[0].height_millipoints = 30_000
    movedLine.y_millipoints = 5_000
    alteredBreak.pages[1].lines.forEach((line: any) => { line.y_millipoints += 10_000 })
    alteredBreak.pages[1].lines.unshift(movedLine)
    alteredBreak.pages[1].paragraph_slices[0].first_line_ordinal = 3
    alteredBreak.pages[1].paragraph_slices[0].line_ids.unshift(movedLine.line_id)
    alteredBreak.pages[1].paragraph_slices[0].height_millipoints = 30_000
    expect(decodeNativeDocxPaginatedLayout(alteredBreak).ok).toBe(true)
    expect(decodeNativeDocxPaginatedLayoutForRequest(alteredBreak, breakRequest).ok).toBe(false)
  })

  it('binds parity fillers to the immediately following odd/even section break', () => {
    const output = paginated(fixture({ lineCounts: [1, 1], sections: [
      { start: 0, bodyWidth: 40_000, bodyHeight: 40_000 },
      { start: 1, breakType: 'odd-page', bodyWidth: 40_000, bodyHeight: 40_000 },
    ] }))
    const wrongReason = structuredClone(output) as any
    wrongReason.pages[1].parity_reason = 'even-page-section'
    expect(decodeNativeDocxPaginatedLayout(wrongReason).ok).toBe(false)
    const wrongBreak = structuredClone(output) as any
    wrongBreak.sections[1].break_type = 'next-page'
    expect(decodeNativeDocxPaginatedLayout(wrongBreak).ok).toBe(false)
    const notImmediate = structuredClone(output) as any
    ;[notImmediate.pages[1], notImmediate.pages[2]] = [notImmediate.pages[2], notImmediate.pages[1]]
    notImmediate.pages.forEach((page: any, index: number) => { page.ordinal = index })
    expect(decodeNativeDocxPaginatedLayout(notImmediate).ok).toBe(false)
    const missingFiller = structuredClone(output) as any
    missingFiller.pages.splice(1, 1)
    missingFiller.pages.forEach((page: any, index: number) => { page.ordinal = index })
    missingFiller.sections[0].page_ids = ['page:section:1:0']
    expect(decodeNativeDocxPaginatedLayout(missingFiller).ok).toBe(false)
  })

  it('uses explicit schema-default section geometry and retains layout-neutral native diagnostics', () => {
    const request = fixture()
    request.document.sections[0]!.page = {
      width_twips: 12_240, height_twips: 15_840, orientation: 'portrait', columns: 1, column_spacing_twips: 720, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
      margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
    }
    request.shaped_lines.available_width_millipoints = 468_000
    for (const line of request.shaped_lines.paragraphs[0]!.lines) line.available_width_millipoints = 468_000
    request.document.unsupported.push({
      id: 'unsupported:default-section', code: 'DEFAULT_SECTION_INFERRED', capability: 'sections', scope_id: 'section:1',
      preservation: 'refuse-mutation', message: 'Word default page geometry is exposed.',
    })
    request.document.unsupported.push({
      id: 'unsupported:identity', code: 'INVALID_NATIVE_PARAGRAPH_ID', capability: 'identity', scope_id: 'paragraph:1',
      preservation: 'refuse-mutation', message: 'Identity hint was ignored.',
    })
    const value = paginated(request)
    expect(value.pages[0]!.body_box).toEqual(expect.objectContaining({ width_millipoints: 468_000, height_millipoints: 648_000 }))
    expect(value.diagnostics.filter((entry) => entry.code === 'source-diagnostic')).toHaveLength(2)
  })

  // A body-level w:sectPr the extractor omitted because it governs no block
  // reaches no page: nothing falls in its range. Pagination reads the sections
  // that do hold blocks, and the omission stays visible as a deferred record.
  it('defers an omitted empty trailing section and refuses one that still claims a section', () => {
    const request = fixture()
    request.document.unsupported.push({
      id: 'unsupported:empty-trailing-section', code: 'EMPTY_TRAILING_SECTION_OMITTED', capability: 'sections', scope_id: 'story:body',
      anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[9]', 9_000, 9_090),
      preservation: 'refuse-mutation', message: 'The final section properties govern no block.',
    })
    const value = paginated(request)
    expect(value.status).toBe('paginated')
    expect(value.diagnostics.filter((entry) => entry.code === 'source-diagnostic')).toHaveLength(1)

    const claimed = fixture()
    claimed.document.unsupported.push({
      id: 'unsupported:empty-trailing-section', code: 'EMPTY_TRAILING_SECTION_OMITTED', capability: 'sections', scope_id: 'story:body',
      anchor: anchor(claimed.document.sections[0]!.anchor.path, 9_000, 9_090),
      preservation: 'refuse-mutation', message: 'The final section properties govern no block.',
    })
    expect(paginateNativeDocxV1(claimed)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'body-structure-unsupported' })]) }) }))
  })

  it('fails safely for cyclic input and never mutates caller-owned wire values', () => {
    const cyclic: any = fixture()
    cyclic.document.loop = cyclic
    expect(paginateNativeDocxV1(cyclic)).toEqual(expect.objectContaining({ ok: false }))

    const request = fixture({ lineCounts: [4], bodyHeight: 20_000 })
    const before = JSON.stringify(request)
    const first = paginateNativeDocxV1(request)
    const second = paginateNativeDocxV1(request)
    expect(first).toEqual(second)
    expect(JSON.stringify(request)).toBe(before)
  })

  it('uses explicit UTF-16 code-unit ordering for wire-visible diagnostics', () => {
    expect(['Å', 'Z', '😀', 'A'].sort(compareNativeCodeUnits)).toEqual(['A', 'Z', 'Å', '😀'])
    expect(asciiLowerNative('Aİ😀Z')).toBe('aİ😀z')
    expect(asciiUpperNative('aı😀z')).toBe('Aı😀Z')
    const request = fixture()
    request.document.unsupported.push(
      { id: 'unsupported:z', code: 'INVALID_NATIVE_PARAGRAPH_ID', capability: 'identity', scope_id: 'paragraph:1', preservation: 'refuse-mutation', message: 'Å diagnostic' },
      { id: 'unsupported:a', code: 'INVALID_NATIVE_PARAGRAPH_ID', capability: 'identity', scope_id: 'paragraph:1', preservation: 'refuse-mutation', message: 'Z diagnostic' },
    )
    const first = paginated(request)
    const second = paginated(structuredClone(request))
    expect(first).toEqual(second)
    expect(first.diagnostics.map((entry) => entry.source_message)).toEqual(['Z diagnostic', 'Å diagnostic'])
  })

  it('stops output object traversal immediately at the depth bound', () => {
    const output: any = paginated(fixture())
    let deep: any = { terminal: true }
    for (let index = 0; index < 70; index += 1) deep = { a: deep }
    const hostile: any = { a_deep: deep }
    Object.defineProperty(hostile, 'z_after_limit', { enumerable: true, get: () => { throw new Error('preflight walked beyond its bound') } })
    output.extra = hostile
    expect(() => decodeNativeDocxPaginatedLayout(output)).not.toThrow()
    expect(decodeNativeDocxPaginatedLayout(output)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'LIMIT_EXCEEDED' })]) }))
  })

  it('refuses before exposing partial output when the page budget is exhausted', () => {
    const request = fixture({
      lineCounts: [DOCX_PAGINATION_LIMITS.maxPages + 1], bodyHeight: 10_000,
      properties: [{ widow_control: false }],
    })
    const result = paginateNativeDocxV1(request)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.status).toBe('refused')
    expect(result.value.pages).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'resource-limit' })]))
    expect(decodeNativeDocxPaginatedLayout(result.value).ok).toBe(true)
  })

  it('refuses atomically when a parity filler consumes the final page budget slot', () => {
    const request = fixture({
      lineCounts: [DOCX_PAGINATION_LIMITS.maxPages - 1, 1], bodyHeight: 10_000,
      properties: [{ widow_control: false }, {}],
      sections: [
        { start: 0, bodyHeight: 10_000 },
        { start: 1, breakType: 'odd-page', bodyHeight: 10_000 },
      ],
    })
    const result = paginateNativeDocxV1(request)
    expect(result).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'resource-limit' })]) }) }))
  })

  it('places relationship-bound footnotes at the page bottom in body reference order deterministically', () => {
    const request = fixture({ lineCounts: [1, 1], bodyHeight: 60_000 })
    addFootnote(request, '9', '1')
    addFootnote(request, '3', '2')
    const separator = request.document.notes.find((story) => story.note_role === 'separator')!
    const note1 = request.document.notes.find((story) => story.id === 'story:footnote:1')!
    const note2 = request.document.notes.find((story) => story.id === 'story:footnote:2')!
    request.document.notes = [separator, note2, note1]
    const first = paginated(request)
    const second = paginated(structuredClone(request))
    expect(first).toEqual(second)
    expect(first.pages[0]!.note_stories?.map((entry) => [entry.note_role, entry.story_id, entry.number, entry.relationship_id])).toEqual([
      ['separator', separator.id, undefined, 'rIdFootnotes'],
      ['content', note1.id, 1, 'rIdFootnotes'],
      ['content', note2.id, 2, 'rIdFootnotes'],
    ])
    const area = first.pages[0]!.note_stories!
    expect(area.every((entry) => entry.section_id === 'section:1' && entry.column_id === 'column:section:1:0' && entry.column_ordinal === 0 && entry.lines.every((line) => line.section_id === entry.section_id && line.column_id === entry.column_id && line.column_ordinal === entry.column_ordinal))).toBe(true)
    expect(area[0]!.top_millipoints).toBeGreaterThanOrEqual(first.pages[0]!.lines.at(-1)!.y_millipoints + first.pages[0]!.lines.at(-1)!.height_millipoints)
    expect(area.at(-1)!.top_millipoints + area.at(-1)!.height_millipoints).toBe(first.pages[0]!.body_box.y_millipoints + first.pages[0]!.body_box.height_millipoints)
  })

  it('reserves final body paragraph after-spacing before admitting footnotes', () => {
    const request = fixture({ bodyHeight: 20_000, spacings: [{ after: 5_000 }] })
    addFootnote(request, '1', '1', 5_000)
    const result = paginateNativeDocxV1(request)
    expect(result).toEqual(expect.objectContaining({
      ok: true,
      value: expect.objectContaining({
        status: 'refused', pages: [], sections: [],
        diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'note-overflow-unsupported' })]),
      }),
    }))
  })

  it('repeats separator paint sources with placement-unique identities across footnote pages', () => {
    const request = fixture({ lineCounts: [1, 1], bodyHeight: 30_000, properties: [{}, { page_break_before: true }] })
    addFootnote(request, '9', '1')
    addFootnote(request, '3', '2')
    const output = paginated(request)
    expect(output.pages).toHaveLength(2)
    const placements = output.pages.flatMap((page) => page.note_stories ?? [])
    expect(placements.filter((entry) => entry.note_role === 'separator')).toHaveLength(2)
    const placedLineIDs = placements.flatMap((entry) => entry.lines.map((line) => line.id))
    expect(new Set(placedLineIDs).size).toBe(placedLineIDs.length)
    expect(decodeNativeDocxPaginatedLayout(output).ok).toBe(true)
  })

  it('admits only exact satisfied paragraph-control diagnostics and independently checks note constraints', () => {
    for (const mutation of ['exact', 'severity', 'run-source', 'message', 'origin'] as const) {
      const request = fixture()
      addFootnote(request)
      const paragraphID = 'paragraph:footnote:1'
      request.resolved_layout.paragraphs.find((paragraph) => paragraph.paragraph_id === paragraphID)!.properties.keep_lines = true
      request.shaped_lines.diagnostics.push({
        code: 'page-control-deferred', severity: mutation === 'severity' ? 'unsupported' : 'deferred', scope_id: paragraphID,
        ...(mutation === 'run-source' ? { source_id: 'run:footnote-label:1' } : {}),
        ...(mutation === 'origin' ? { source_diagnostic_code: 'FONT_MISSING' } : {}),
        message: mutation === 'message' ? 'Unrecognized deferred layout semantics' : 'keep_lines is retained in resolved layout for the future paginator and does not alter line shaping',
      })
      const result = paginateNativeDocxV1(request)
      expect(result.ok).toBe(true)
      if (result.ok) expect(result.value.status).toBe(mutation === 'exact' ? 'paginated' : 'refused')
    }
    for (const key of ['keep_next', 'page_break_before'] as const) {
      const request = fixture()
      addFootnote(request)
      request.resolved_layout.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:footnote:1')!.properties[key] = true
      expect(request.shaped_lines.diagnostics).toEqual([])
      const result = paginateNativeDocxV1(request)
      expect(result.ok).toBe(true)
      if (result.ok) expect(result.value).toEqual(expect.objectContaining({ status: 'refused', pages: [], sections: [] }))
    }
  })

  it('continues one endnote at complete paragraph boundaries with source-bound separators and one label', () => {
    const request = continuedEndnoteFixture()
    const snapshot = structuredClone(request)
    const output = paginated(request)
    expect(output.pages.map((page) => page.note_stories?.map((story) => story.note_role))).toEqual([
      ['separator', 'content'], ['continuation-separator', 'content'], ['continuation-separator', 'content'],
    ])
    expect(output.pages.map((page) => page.note_stories?.[1]?.lines.length)).toEqual([2, 3, 1])
    const placed = output.pages.flatMap((page) => page.note_stories?.filter((story) => story.note_role === 'content').flatMap((story) => story.lines) ?? [])
    expect(new Set(placed.map((line) => line.line_id)).size).toBe(6)
    expect(placed.filter((line) => line.paragraph_id === 'paragraph:footnote:1')).toHaveLength(1)
    expect(request).toEqual(snapshot)
    expect(paginated(request)).toEqual(output)
    for (const mutation of ['drop', 'duplicate', 'separator', 'reference'] as const) {
      const forged = structuredClone(output)
      if (mutation === 'drop') forged.pages[2]!.note_stories![1]!.lines = []
      if (mutation === 'duplicate') forged.pages[2]!.note_stories![1]!.lines[0]!.line_id = placed[0]!.line_id
      if (mutation === 'separator') forged.pages[1]!.note_stories![0]!.note_role = 'separator'
      if (mutation === 'reference') forged.pages[1]!.note_stories![1]!.reference_run_id = 'run:forged'
      expect(decodeNativeDocxPaginatedLayoutForRequest(forged, request).ok).toBe(false)
    }
  })

  it.each(['endnote', 'footnote'])('honors keep-lines and default widow control for two-line %s paragraphs', (kind) => {
    for (const keepLines of [true, false, undefined]) {
      const request = kind === 'endnote' ? continuedEndnoteFixture() : continuedFootnoteFixture()
      const paragraph = request.shaped_lines.paragraphs.find((entry) => entry.paragraph_id === `paragraph:${kind}:2`)!
      paragraph.lines.push({ ...structuredClone(paragraph.lines[0]!), id: `line:${paragraph.paragraph_id}:1`, ordinal: 1 })
      paragraph.block_advance_millipoints *= 2
      const properties = request.resolved_layout.paragraphs.find((entry) => entry.paragraph_id === paragraph.paragraph_id)!.properties
      if (keepLines !== undefined) properties.keep_lines = keepLines
      const result = paginateNativeDocxV1(request)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.status).toBe('paginated')
      if (result.value.status === 'paginated') {
        const pages = result.value.pages.filter((page) => page.note_stories?.some((story) => story.lines.some((line) => line.paragraph_id === paragraph.paragraph_id)))
        expect(pages).toHaveLength(1)
        expect(pages[0]!.note_stories![1]!.lines.filter((line) => line.paragraph_id === paragraph.paragraph_id)).toHaveLength(2)
        expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, request).ok).toBe(true)
      }
    }
  })

  it('retries an unfitting first note paragraph on a fresh page with an ordinary separator', () => {
    const output = paginated(continuedEndnoteFixture(3))
    expect(output.pages[0]!.note_stories).toEqual([])
    expect(output.pages[1]!.note_stories![0]!.note_role).toBe('separator')
    expect(output.pages[2]!.note_stories![0]!.note_role).toBe('continuation-separator')
  })

  it('discards all pages when endnote continuation exceeds the page budget', () => {
    const result = paginateNativeDocxV1(continuedEndnoteFixture(4_096, 20_000))
    expect(result.ok, JSON.stringify(result)).toBe(true)
    if (result.ok) expect(result.value).toEqual(expect.objectContaining({
      status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'resource-limit' })]),
    }))
  })

  it('refuses activated missing, unsupported, or drifting continuation sentinels atomically', () => {
    for (const mutation of ['missing', 'unshaped', 'unsupported', 'relationship', 'spacing', 'keep-next', 'oversized'] as const) {
      const request = continuedEndnoteFixture()
      const continuation = request.document.notes.at(-1)!
      if (mutation === 'missing') { request.document.notes.pop(); request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter((paragraph) => paragraph.story_id !== continuation.id); request.resolved_layout.paragraphs = request.resolved_layout.paragraphs.filter((paragraph) => paragraph.paragraph_id !== continuation.blocks[0]!.id) }
      if (mutation === 'unshaped') request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter((paragraph) => paragraph.story_id !== continuation.id)
      if (mutation === 'unsupported') request.document.unsupported.push({ id: 'unsupported:continuation', code: 'UNMODELED_NOTE_MARKUP', capability: 'notes', scope_id: continuation.id, preservation: 'refuse-mutation', message: 'Unsupported activated instruction.' })
      if (mutation === 'relationship') continuation.relationship_id = 'rIdDrift'
      if (mutation === 'spacing') { const paragraph = request.shaped_lines.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:endnote:2')!; paragraph.spacing_after_millipoints = 50; paragraph.block_advance_millipoints += 50; request.resolved_layout.paragraphs.find((entry) => entry.paragraph_id === paragraph.paragraph_id)!.properties.spacing_after_twips = 1 }
      if (mutation === 'keep-next') request.resolved_layout.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:endnote:2')!.properties.keep_next = true
      if (mutation === 'oversized') {
        const paragraph = request.shaped_lines.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:endnote:2')!
        paragraph.lines[0]!.line_height_millipoints = 40_000
        paragraph.block_advance_millipoints = 40_000
      }
      const result = paginateNativeDocxV1(request)
      expect(result.ok, `${mutation}: ${JSON.stringify(result)}`).toBe(mutation !== 'relationship')
      if (result.ok) expect(result.value).toEqual(expect.objectContaining({ status: 'refused', pages: [], sections: [] }))
    }
  })

  it('moves one whole endnote group to one fresh bounded final page', () => {
    const request = fixture({ bodyHeight: 20_000 })
    addFootnote(request, '1', '1', 6_000)
    convertFootnotesToEndnotes(request)
    const output = paginated(request)
    expect(output.pages).toHaveLength(2)
    expect(output.pages[0]!.note_stories).toEqual([])
    expect(output.pages[1]!.lines).toEqual([])
    expect(output.pages[1]!.note_stories?.map((entry) => [entry.story_kind, entry.note_role])).toEqual([['endnote', 'separator'], ['endnote', 'content']])
    expect(output.pages[1]!.note_stories?.[0]!.top_millipoints).toBe(output.pages[1]!.body_box.y_millipoints)
    expect(output.sections[0]!.page_ids).toEqual(output.pages.map((page) => page.id))
  })

  it('flows a fitting endnote group immediately after final body content', () => {
    const request = fixture({ bodyHeight: 60_000 })
    addFootnote(request, '1', '1', 5_000)
    convertFootnotesToEndnotes(request)
    const output = paginated(request)
    expect(output.pages).toHaveLength(1)
    const bodyBottom = output.pages[0]!.lines.at(-1)!.y_millipoints + output.pages[0]!.lines.at(-1)!.height_millipoints
    expect(output.pages[0]!.note_stories?.[0]!.top_millipoints).toBe(bodyBottom)
  })

  it('keeps mixed final-page footnotes and endnotes disjoint and preserves header provenance on the added page', () => {
    const request = fixture({ lineCounts: [1, 1], bodyHeight: 40_000 })
    addFootnote(request, '1', '1', 5_000)
    convertFootnotesToEndnotes(request)
    addFootnote(request, '2', '2', 5_000)
    const headerPart = 'word/header1.xml'
    request.document.headers.push({ id: 'story:header:1', kind: 'header', part_name: headerPart, anchor: noteAnchor(headerPart, '/w:hdr[1]', 1, 20), blocks: [] })
    request.document.sections[0]!.header_refs.push({ kind: 'default', story_id: 'story:header:1', relationship_id: 'rIdHeader' })
    const output = paginated(request)
    expect(output.pages).toHaveLength(2)
    expect(output.pages[0]!.note_stories?.map((entry) => entry.story_kind)).toEqual(['footnote', 'footnote'])
    expect(output.pages[1]!.note_stories?.map((entry) => entry.story_kind)).toEqual(['endnote', 'endnote'])
    expect(output.pages[1]!.header_refs).toEqual([{ kind: 'default', story_id: 'story:header:1', relationship_id: 'rIdHeader' }])
  })

  it('fails closed for duplicate references, note cycles, relationship drift, marker drift, and only actual continuation', () => {
    const duplicate = fixture({ lineCounts: [1, 1], bodyHeight: 60_000 })
    addFootnote(duplicate)
    const secondParagraph = duplicate.document.body.blocks[1]!.paragraph!
    const secondRun = secondParagraph.runs[0]!
    secondRun.kind = 'reference'; delete secondRun.text
    secondRun.reference = { kind: 'footnote', target_id: 'story:footnote:1' }
    duplicate.shaped_lines.paragraphs[1]!.lines[0]!.fragments = [{ ...duplicate.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!, id: 'fragment:paragraph:2:0:0', source_id: secondRun.id }]
    duplicate.shaped_lines.paragraphs[1]!.lines[0]!.logical_to_visual = [0]

    const cycle = fixture({ bodyHeight: 60_000 })
    addFootnote(cycle)
    cycle.document.notes.find((story) => story.id === 'story:footnote:1')!.blocks[0]!.paragraph!.runs[0]!.reference!.role = 'anchor'

    const relationship = fixture({ bodyHeight: 60_000 })
    addFootnote(relationship)
    relationship.document.notes.find((story) => story.id === 'story:footnote:1')!.relationship_id = 'rIdDrifted'
    expect(decodeNativeDocxPaginationRequestV1(relationship)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ path: '/document/notes/1/relationship_id' })]) }))

    const marker = fixture({ bodyHeight: 60_000 })
    addFootnote(marker)
    marker.shaped_lines.paragraphs.find((paragraph) => paragraph.story_kind === 'footnote' && paragraph.story_id === 'story:footnote:1')!.lines[0]!.fragments[0]!.text = '9'

    const overflow = fixture({ bodyHeight: 20_000 })
    addFootnote(overflow, '1', '1', 5_000)
    overflow.document.notes.push({
      id: 'story:footnote:continuation', kind: 'footnote', native_story_id: '0', relationship_id: 'rIdFootnotes', note_role: 'continuation-separator', part_name: 'word/footnotes.xml',
      anchor: noteAnchor('word/footnotes.xml', '/w:footnotes[1]/w:footnote[3]', 151, 190), blocks: [{
        kind: 'paragraph', id: 'paragraph:footnote:continuation', paragraph: {
          ...paragraph('paragraph:footnote:continuation', 40),
          anchor: noteAnchor('word/footnotes.xml', '/w:footnotes[1]/w:footnote[3]/w:p[1]', 160, 180),
          runs: [],
        },
      }],
    })
    overflow.resolved_layout.paragraphs.push({
      paragraph_id: 'paragraph:footnote:continuation', applied_styles: [], properties: {},
      paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 },
    })
    overflow.document.unsupported.push({
      id: 'unsupported:inert-continuation', code: 'UNMODELED_CONTINUATION_SENTINEL', capability: 'notes', scope_id: 'story:footnote:continuation',
      preservation: 'refuse-mutation', message: 'Dormant continuation markup is retained but cannot affect a fitting note.',
    })

    const tooTall = fixture({ bodyHeight: 20_000 })
    addFootnote(tooTall, '1', '1', 30_000)

    const fittingWithContinuationSentinel = paginateNativeDocxV1(overflow)
    expect(fittingWithContinuationSentinel).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'paginated' }) }))

    const continuationRequired = fixture({ bodyHeight: 20_000 })
    addFootnote(continuationRequired, '1', '1', 15_000)
    continuationRequired.document.notes.push(structuredClone(overflow.document.notes.at(-1)!))
    continuationRequired.resolved_layout.paragraphs.push(structuredClone(overflow.resolved_layout.paragraphs.at(-1)!))

    for (const [request, code] of [[duplicate, 'note-reference-ambiguous'], [cycle, 'note-reference-ambiguous'], [marker, 'note-reference-ambiguous'], [continuationRequired, 'note-overflow-unsupported'], [tooTall, 'note-structure-unsupported']] as const) {
      const result = paginateNativeDocxV1(request)
      expect(result).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code })]) }) }))
    }
  })

  it('fails closed before placement for note drawings and escaped shaped geometry', () => {
    const drawing = fixture({ bodyHeight: 60_000 }) as any
    addFootnote(drawing)
    const noteParagraph = drawing.document.notes.find((story: any) => story.id === 'story:footnote:1').blocks[0].paragraph
    noteParagraph.runs.push({
      kind: 'drawing', id: 'run:note-drawing', anchor: { ...noteParagraph.anchor, path: `${noteParagraph.anchor.path}/w:r[2]`, start_byte: 101, end_byte: 130 },
      drawing: {
        id: 'drawing:note', anchor: { ...noteParagraph.anchor, path: `${noteParagraph.anchor.path}/w:r[2]/w:drawing[1]`, start_byte: 105, end_byte: 125 },
        placement: 'floating', width_emu: 914400, height_emu: 914400, x_emu: 0, y_emu: 0,
        horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'square',
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Drawing placement unavailable.', preservation: 'refuse-mutation' } },
      },
    })
    drawing.resolved_layout.runs.push({ run_id: 'run:note-drawing', paragraph_id: noteParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } })
    const drawingResult = paginateNativeDocxV1(drawing)
    expect(drawingResult).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'note-structure-unsupported' })]) }) }))

    const geometry = fixture({ bodyHeight: 60_000 })
    addFootnote(geometry)
    const noteLine = geometry.shaped_lines.paragraphs.find((paragraph) => paragraph.story_id === 'story:footnote:1')!.lines[0]!
    noteLine.inline_offset_millipoints = geometry.shaped_lines.available_width_millipoints
    const geometryResult = paginateNativeDocxV1(geometry)
    expect(geometryResult).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'INVALID_VALUE', path: expect.stringContaining('/inline_offset_millipoints') })]) }))
  })

  it('stages direct note placement transactionally when a later page overflows', () => {
    const request = fixture({ lineCounts: [1, 1], bodyHeight: 30_000, properties: [{}, { page_break_before: true }] })
    const layout = paginated(structuredClone(request))
    addFootnote(request, '1', '1', 5_000)
    addFootnote(request, '2', '2', 30_000)
    const before = structuredClone(layout)
    const failure = placeNativeDocxNotesV1(layout, request.document, request.resolved_layout, request.shaped_lines)
    expect(failure).toEqual(expect.objectContaining({ code: 'note-overflow-unsupported' }))
    expect(layout).toEqual(before)
  })

  it('rejects tampered note placement identities through deterministic source replay', () => {
    const request = fixture({ bodyHeight: 60_000 })
    addFootnote(request)
    const output = paginated(request)
    const drifted = structuredClone(output)
    drifted.pages[0]!.note_stories![1]!.relationship_id = 'rIdDrifted'
    expect(decodeNativeDocxPaginatedLayoutForRequest(drifted, request)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))
    const columnDrift = structuredClone(output)
    columnDrift.pages[0]!.note_stories![1]!.lines[0]!.column_id = 'column:drifted'
    expect(decodeNativeDocxPaginatedLayoutForRequest(columnDrift, request)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))
  })

  it('contains no DOM, browser, HTML, canvas, or legacy paginator dependency', () => {
    const source = [
      readFileSync(new URL('./nativePaginationV1.ts', import.meta.url), 'utf8'),
      readFileSync(new URL('./nativePaginationSettings.ts', import.meta.url), 'utf8'),
      readFileSync(new URL('./nativeShapedLinesContract.ts', import.meta.url), 'utf8'),
      readFileSync(new URL('./nativePaginatedLayoutContract.ts', import.meta.url), 'utf8'),
      readFileSync(new URL('./nativeNotePaginationV1.ts', import.meta.url), 'utf8'),
    ].join('\n')
    expect(source).not.toMatch(/\bglobalThis\s*\.\s*(?:document|window)\b|(?:^|[^\w.])window\s*[.[]/m)
    expect(source).not.toMatch(/\b(?:HTMLElement|DOMParser|OffscreenCanvas|CanvasRenderingContext2D|getBoundingClientRect|measureText|mammoth)\b\s*[.(]/)
    expect(source).not.toMatch(/from ['"]\.\/paginate(?:\.js)?['"]/)
  })
})

function unequalColumnFixture(): NativeDocxPaginationRequestV1 {
  const request = fixture({ lineCounts: [3, 2, 3, 2, 1], properties: Array.from({ length: 5 }, () => ({ keep_lines: true })) })
  const section = request.document.sections[0]!
  section.page.columns = 2
  section.page.column_layout = 'explicit'
  section.page.column_spacing_twips = 0
  section.page.column_definitions = [
    { id: 'column:section:1:0', ordinal: 0, width_twips: 300, space_after_twips: 100 },
    { id: 'column:section:1:1', ordinal: 1, width_twips: 400, space_after_twips: 0 },
  ]
  request.document.unsupported = [{ id: 'unsupported:unequal', code: 'UNEQUAL_SECTION_COLUMNS', capability: 'sections', scope_id: section.id, anchor: structuredClone(section.anchor), preservation: 'refuse-mutation', message: 'Unequal column widths require per-column shaping outside the exact v1 slice' }]
  request.pagination_settings.no_column_balance = true
  const candidates = [15000, 20000].map((width) => {
    const shaped = structuredClone(request.shaped_lines)
    shaped.available_width_millipoints = width
    shaped.paragraphs.forEach((paragraph) => paragraph.lines.forEach((line) => { line.available_width_millipoints = width }))
    return shaped
  }) as [NativeDocxShapedLinesV1, NativeDocxShapedLinesV1]
  // A wider candidate has fewer complete lines; paragraph identities remain stable.
  candidates[1].paragraphs[1]!.lines.pop()
  candidates[1].paragraphs[1]!.block_advance_millipoints = 10000
  for (const [index, block] of request.document.body.blocks.entries()) block.paragraph!.runs[0]!.text = 'AAA'
  for (const candidate of candidates) for (const [index, paragraph] of candidate.paragraphs.entries()) {
    let start = 0
    for (const [ordinal, line] of paragraph.lines.entries()) {
      const end = ordinal + 1 === paragraph.lines.length ? 3 : start + 1
      line.logical_to_visual = [0]
      line.fragments = [{ id: `fragment:${paragraph.paragraph_id}:${ordinal}:0`, source_kind: 'run', source_id: request.document.body.blocks[index]!.paragraph!.runs[0]!.id, start_utf16: start, end_utf16: end, text: 'AAA'.slice(start, end), direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Latn', language: 'und', face_id: 'face:test', whitespace: false, advance_inline_millipoints: 5000, justification_expansion_millipoints: 0, ascent_millipoints: 8000, descent_millipoints: -2000, line_gap_millipoints: 0, glyphs: [{ glyph_id: 1, advance_x_millipoints: 5000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 }] }]
      start = end
    }
  }
  request.column_shaped_lines = candidates
  const plan = planNativeDocxColumnParagraphFlowV1(request.document, request.resolved_layout, request.pagination_settings, candidates)
  if (!plan) throw new Error('unequal fixture did not qualify')
  request.shaped_lines = plan.shaped_lines
  return request
}

describe('source-bound unequal whole-paragraph columns', () => {
  it('tries current width before advancing and retains unused final columns', () => {
    const request = unequalColumnFixture()
    const output = paginated(request)
    expect(output.pages.map((page) => page.paragraph_slices.map((slice) => [slice.paragraph_id, slice.column_ordinal, slice.line_ids.length]))).toEqual([
      [['paragraph:1', 0, 3], ['paragraph:2', 1, 1], ['paragraph:3', 1, 3]],
      [['paragraph:4', 0, 2], ['paragraph:5', 0, 1]],
    ])
    expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)
  })
  it('refuses noColumnBalance outside qualified unequal flow', () => {
    const request = fixture()
    request.pagination_settings.no_column_balance = true
    expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [] }) }))
  })
  it('rejects plausible unused-candidate omission, reordering, font and geometry forgery', () => {
    for (const mutate of [
      (r: NativeDocxPaginationRequestV1) => { const p = r.column_shaped_lines![1].paragraphs[0]!; p.lines.pop(); p.block_advance_millipoints -= 10000 },
      (r: NativeDocxPaginationRequestV1) => { const p = r.column_shaped_lines![1].paragraphs[0]!; const first = p.lines[0]!.fragments[0]!; const second = p.lines[1]!.fragments[0]!; [first.start_utf16, second.start_utf16] = [second.start_utf16, first.start_utf16]; [first.end_utf16, second.end_utf16] = [second.end_utf16, first.end_utf16] },
      (r: NativeDocxPaginationRequestV1) => { r.column_shaped_lines![1].paragraphs[0]!.lines[0]!.fragments[0]!.face_id = 'face:other' },
      (r: NativeDocxPaginationRequestV1) => { const p = r.column_shaped_lines![1].paragraphs[0]!; p.lines[0]!.line_height_millipoints -= 1; p.block_advance_millipoints -= 1 },
      (r: NativeDocxPaginationRequestV1) => { const p = r.column_shaped_lines![1].paragraphs[0]!; p.lines[0]!.fragments[0]!.ascent_millipoints -= 1; p.lines[0]!.ascent_millipoints -= 1; p.lines[0]!.line_height_millipoints -= 1; p.block_advance_millipoints -= 1 },
      (r: NativeDocxPaginationRequestV1) => { r.document.unsupported[0]!.anchor!.start_byte += 1 },
      (r: NativeDocxPaginationRequestV1) => { r.document.unsupported[0]!.message += ' changed' },
    ]) {
      const request = unequalColumnFixture()
      // paragraph 1 is painted from candidate 0; mutate only its unused candidate.
      mutate(request)
      expect(planNativeDocxColumnParagraphFlowV1(request.document, request.resolved_layout, request.pagination_settings, request.column_shaped_lines!)).toBeUndefined()
      expect(decodeNativeDocxPaginationRequestV1(request).ok).toBe(false)
    }
  })
  it('refuses placement budget exhaustion atomically', () => {
    const original = DOCX_PAGINATION_LIMITS.maxPages
    try {
      Object.assign(DOCX_PAGINATION_LIMITS, { maxPages: 1 })
      expect(paginateNativeDocxV1(unequalColumnFixture())).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))
    } finally { Object.assign(DOCX_PAGINATION_LIMITS, { maxPages: original }) }
  })
  it('rejects missing candidates and source, controls, width, provenance, or selected-shape tampering', () => {
    const missing = unequalColumnFixture()
    delete missing.column_shaped_lines
    expect(paginateNativeDocxV1(missing)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [] }) }))
    for (const mutate of [
      (r: NativeDocxPaginationRequestV1) => { r.pagination_settings.no_column_balance = false },
      (r: NativeDocxPaginationRequestV1) => { r.document.unsupported[0]!.scope_id = r.document.body.id },
      (r: NativeDocxPaginationRequestV1) => { delete r.resolved_layout.paragraphs[0]!.properties.keep_lines },
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs[0]!.properties.keep_next = true },
      (r: NativeDocxPaginationRequestV1) => { r.column_shaped_lines![1].available_width_millipoints -= 1 },
      (r: NativeDocxPaginationRequestV1) => { r.column_shaped_lines![1].providers.shaper_revision = 'tampered' },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.paragraphs[1] = structuredClone(r.column_shaped_lines![0].paragraphs[1]!) },
    ]) {
      const request = unequalColumnFixture()
      mutate(request)
      expect(decodeNativeDocxPaginationRequestV1(request).ok).toBe(false)
    }
  })
})

describe('whole-footnote body reservation', () => {
  it('reserves only the reference page and pushes later paragraphs', () => {
    const request = fixture({ lineCounts: [1, 1, 1, 1] })
    addFootnote(request)
    const output = paginated(request)
    expect(output.pages.map((page) => page.paragraph_slices.map((slice) => slice.paragraph_id))).toEqual([
      ['paragraph:1', 'paragraph:2', 'paragraph:3'], ['paragraph:4'],
    ])
    expect(output.pages.map((page) => page.note_stories?.length ?? 0)).toEqual([2, 0])
    expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)
    const forged = structuredClone(output)
    forged.pages[0]!.note_stories![0]!.height_millipoints -= 1
    expect(decodeNativeDocxPaginatedLayoutForRequest(forged, request).ok).toBe(false)
  })
  it('moves the whole reference and note together without reserving the previous page', () => {
    const request = fixture({ lineCounts: [1, 1, 1, 1, 1] })
    addFootnote(request, '1', '4')
    const before = structuredClone(request)
    const output = paginated(request)
    expect(output.pages.map((page) => page.paragraph_slices.map((slice) => slice.paragraph_id))).toEqual([
      ['paragraph:1', 'paragraph:2', 'paragraph:3'], ['paragraph:4', 'paragraph:5'],
    ])
    expect(output.pages.map((page) => page.note_stories?.length ?? 0)).toEqual([0, 2])
    expect(output.pages[1]!.note_stories![1]!.reference_run_id).toBe('run:paragraph:4')
    expect(request).toEqual(before)
  })
  it('restores full height after leaving the reference page and honors exact pair fit', () => {
    const request = fixture({ lineCounts: [1, 4], properties: [{}, { keep_lines: true }] })
    addFootnote(request)
    const output = paginated(request)
    expect(output.pages.map((page) => page.lines.length)).toEqual([1, 4])
    expect(output.pages[1]!.paragraph_slices[0]!.height_millipoints).toBe(40000)
    const exact = fixture({ lineCounts: [1, 1], bodyHeight: 20000 })
    addFootnote(exact)
    expect(paginated(exact).pages.map((page) => page.lines.length)).toEqual([1, 1])
  })
  it('retains atomic refusal for an oversized pair, invalid note authority, and page budgets', () => {
    const oversized = fixture({ bodyHeight: 15000 })
    addFootnote(oversized)
    const invalid = fixture({ lineCounts: [1, 1, 1, 1] })
    addFootnote(invalid)
    invalid.shaped_lines.diagnostics.push({ code: 'page-control-deferred', severity: 'unsupported', scope_id: invalid.document.notes[1]!.blocks[0]!.paragraph!.id, message: 'Unqualified note source' })
    for (const request of [oversized, invalid]) expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))
    const original = DOCX_PAGINATION_LIMITS.maxPages
    try {
      Object.assign(DOCX_PAGINATION_LIMITS, { maxPages: 1 })
      const request = fixture({ lineCounts: [1, 1, 1, 1] }); addFootnote(request)
      expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))
    } finally { Object.assign(DOCX_PAGINATION_LIMITS, { maxPages: original }) }
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

describe('authored whole-footnote paragraph chains', () => {
  function input(count = 2, reference = '1', height = 40000) {
    const request = fixture({ lineCounts: [1, 1, 1, 1, 1], bodyHeight: height })
    addFootnote(request, '1', reference)
    extendFootnoteChain(request, count)
    return request
  }
  it('keeps two source paragraphs with their reference and restores later page height', () => {
    for (const reference of ['1', '4']) {
      const request = input(2, reference), before = structuredClone(request)
      const output = paginated(request)
      const page = output.pages.find((entry) => entry.note_stories?.length)!
      expect(page.ordinal).toBe(reference === '1' ? 0 : 1)
      expect(page.note_stories![1]!.lines.map((paragraph) => paragraph.paragraph_id)).toEqual(request.document.notes[1]!.blocks.map((block) => block.id))
      expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)
      expect(request).toEqual(before)
      const forged = structuredClone(output)
      forged.pages[page.ordinal]!.note_stories![1]!.lines.reverse()
      expect(decodeNativeDocxPaginatedLayoutForRequest(forged, request).ok).toBe(false)
    }
  })
  it('recomputes internal chain evidence against actual placement inputs', () => {
    const request = input(), output = paginated(request)
    const reservation = measureNativeDocxFootnoteReservationV1(request, measureNativeDocxFootnoteAreaForReservationV1)!
    reservation.references[0]!.note = structuredClone(request.document.notes[0]!)
    reservation.groups[0]!.height_millipoints = 1
    expect(placeNativeDocxNotesV1(structuredClone(output), request.document, request.resolved_layout, request.shaped_lines, reservation)).toBeUndefined()
    const changed = structuredClone(request)
    changed.resolved_layout.paragraphs.find((p) => p.paragraph_id === changed.document.notes[1]!.blocks[1]!.id)!.properties.keep_next = true
    const staged = structuredClone(output), before = structuredClone(staged)
    expect(placeNativeDocxNotesV1(staged, changed.document, changed.resolved_layout, changed.shaped_lines, reservation)).toBeDefined()
    expect(staged).toEqual(before)
  })
  it('admits sixteen authored members at exact pair fit and refuses seventeen', () => {
    expect(paginated(input(16, '1', 95000)).pages[0]!.note_stories![1]!.lines).toHaveLength(16)
    expect(paginateNativeDocxV1(input(17, '1', 200000))).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [] }) }))
  })
  it('admits explicitly unchained paragraphs and refuses overflowing pairs and forged diagnostic admission', () => {
    const unchained=input()
    unchained.resolved_layout.paragraphs.find(p=>p.paragraph_id===unchained.document.notes[1]!.blocks[0]!.id)!.properties.keep_next=false
    expect(paginated(unchained).pages[0]!.note_stories![1]!.lines).toHaveLength(2)
    const cases = [input(2, '1', 20000)]
    for (const mutate of [
      (r: NativeDocxPaginationRequestV1) => { r.resolved_layout.paragraphs.at(-1)!.properties.keep_next = true },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.diagnostics.at(-1)!.severity = 'unsupported' },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.diagnostics.at(-1)!.source_id = r.document.notes[1]!.blocks[0]!.paragraph!.runs[0]!.id },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.diagnostics.at(-1)!.source_diagnostic_code = 'forged' },
      (r: NativeDocxPaginationRequestV1) => { r.shaped_lines.diagnostics.at(-1)!.message += ' changed' },
    ]) { const request = input(); mutate(request); cases.push(request) }
    for (const request of cases) expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [] }) }))
  })
})

describe('two whole-footnote page reservations', () => {
  function input(height = 40000, second = '2', lineCounts = [1, 1, 1, 1], noteHeight = 5000) {
    const request = fixture({ bodyHeight: height, lineCounts })
    addFootnote(request, '1', '1', noteHeight)
    addFootnote(request, '2', second, noteHeight)
    return request
  }
  const groups = (output: ReturnType<typeof paginated>) => output.pages.map((page) => (page.note_stories ?? []).filter((story) => story.note_role === 'content').map((story) => story.number))
  it('shares one separator at exact total fit without double-counting the first reservation', () => {
    const request = input(35000), before = structuredClone(request)
    const output = paginated(request)
    expect(groups(output)).toEqual([[1, 2], []])
    expect(output.pages[0]!.note_stories).toHaveLength(3)
    expect(output.pages[0]!.note_stories!.reduce((sum, note) => sum + note.height_millipoints, 0)).toBe(15000)
    expect(output.pages[0]!.lines).toHaveLength(2)
    expect(decodeNativeDocxPaginatedLayoutForRequest(output, request).ok).toBe(true)
    expect(request).toEqual(before)
  })
  it('moves only the second pair when combined reservation exceeds remaining space', () => {
    const request = input(34950)
    const output = paginated(request)
    expect(groups(output)).toEqual([[1], [2], []])
    expect(output.pages[0]!.paragraph_slices.map((p) => p.paragraph_id)).toEqual(['paragraph:1'])
    expect(output.pages[1]!.paragraph_slices.map((p) => p.paragraph_id)).toEqual(['paragraph:2', 'paragraph:3'])
    expect(output.pages.slice(0, 2).map((page) => page.note_stories!.filter((note) => note.note_role === 'separator').length)).toEqual([1, 1])
  })
  it('clears active membership after ordinary body advance and preserves global note numbering', () => {
    const request = input(40000, '5', [1, 1, 1, 1, 1, 1])
    const output = paginated(request)
    expect(groups(output)).toEqual([[1], [2]])
    expect(output.pages[1]!.paragraph_slices.map((p) => p.paragraph_id)).toEqual(['paragraph:4', 'paragraph:5', 'paragraph:6'])
    expect(output.pages[1]!.note_stories![1]!.reference_run_id).toBe('run:paragraph:5')
  })
  it('moves the second pair when the combined group exceeds a page but both individual pairs fit', () => {
    expect(groups(paginated(input(40000, '2', [1, 1], 15000)))).toEqual([[1], [2]])
  })
  it('uses body reference order even if native story storage order differs', () => {
    const request = input(35000)
    request.document.notes = [request.document.notes[0]!, request.document.notes[2]!, request.document.notes[1]!]
    expect(groups(paginated(request))).toEqual([[1, 2], []])
  })
  it('rejects replay with reordered notes, wrong numbers, missing lines or an extra separator', () => {
    const request = input(35000), output = paginated(request)
    for (const mutate of [
      (o: typeof output) => { const notes = o.pages[0]!.note_stories!; [notes[1], notes[2]] = [notes[2]!, notes[1]!] },
      (o: typeof output) => { o.pages[0]!.note_stories![2]!.number = 1 },
      (o: typeof output) => { o.pages[0]!.note_stories![2]!.lines.pop() },
      (o: typeof output) => { o.pages[0]!.note_stories!.push(structuredClone(o.pages[0]!.note_stories![0]!)) },
    ]) { const forged = structuredClone(output); mutate(forged); expect(decodeNativeDocxPaginatedLayoutForRequest(forged, request).ok).toBe(false) }
  })
  it('counts both notes and the shared separator against the final cumulative line budget', () => {
    const original = DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines
    try {
      Object.assign(DOCX_NOTE_PAGINATION_LIMITS, { maxNoteLines: 4 })
      const request = input(35000, '2', [1, 1])
      expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))
    } finally { Object.assign(DOCX_NOTE_PAGINATION_LIMITS, { maxNoteLines: original }) }
  })
  it('retains atomic refusal for an individually oversized pair and a wrong second label', () => {
    const oversized = input(39950, '2', [1, 1], 15000)
    const wrong = input(35000)
    wrong.shaped_lines.paragraphs.find((p) => p.story_id === wrong.document.notes[2]!.id)!.lines[0]!.fragments[0]!.text = '1'
    for (const request of [oversized, wrong]) expect(paginateNativeDocxV1(request)).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ status: 'refused', pages: [], sections: [] }) }))
  })
})

function continuedFootnoteFixture(bodyLines=1,bodyHeight=40_000):NativeDocxPaginationRequestV1 {
 return JSON.parse(JSON.stringify(continuedEndnoteFixture(bodyLines,bodyHeight)).replaceAll('endnote','footnote'))
}
describe('final-body-page footnote continuation',()=>{
 it('places whole paragraph slices at page bottom with one source label and ordered continuation rules',()=>{
  const request=continuedFootnoteFixture(),before=structuredClone(request),output=paginated(request)
  expect(output.pages.map(p=>p.note_stories?.map(n=>n.note_role))).toEqual([['separator','content'],['continuation-separator','content'],['continuation-separator','content']])
  expect(output.pages.map(p=>p.note_stories?.[1]?.lines.length)).toEqual([2,3,1])
  const lines=output.pages.flatMap(p=>p.note_stories![1]!.lines)
  expect(new Set(lines.map(l=>l.line_id)).size).toBe(6)
  expect(lines.filter(l=>l.paragraph_id==='paragraph:footnote:1')).toHaveLength(1)
  for(const page of output.pages){const note=page.note_stories!.at(-1)!;expect(note.top_millipoints+note.height_millipoints).toBe(page.body_box.y_millipoints+page.body_box.height_millipoints)}
  expect(request).toEqual(before);expect(paginated(request)).toEqual(output)
  for(const mutation of ['drop','duplicate','separator','page','reference'] as const){
   const forged=structuredClone(output)
   if(mutation==='drop')forged.pages[2]!.note_stories![1]!.lines=[]
   if(mutation==='duplicate')forged.pages[2]!.note_stories![1]!.lines[0]!.line_id=lines[0]!.line_id
   if(mutation==='separator')forged.pages[1]!.note_stories![0]!.note_role='separator'
   if(mutation==='page')forged.pages[2]!.note_stories![1]!.top_millipoints--
   if(mutation==='reference')forged.pages[1]!.note_stories![1]!.reference_run_id='run:forged'
   expect(decodeNativeDocxPaginatedLayoutForRequest(forged,request).ok).toBe(false)
  }
 })
 it('refuses when a three-line widow group cannot retain the first footnote slice',()=>{
  for(const bodyLines of [3]){
   const result=paginateNativeDocxV1(continuedFootnoteFixture(bodyLines))
   expect(result.ok).toBe(true)
   if(result.ok)expect(result.value).toEqual(expect.objectContaining({status:'refused',pages:[],sections:[]}))
  }
 })
 it('requires an exact active continuation separator and honors kept note paragraph boundaries',()=>{
  for(const mutation of ['missing','unshaped','unsupported','relationship','spacing','keep-next','oversized'] as const){
   const request=continuedFootnoteFixture(),continuation=request.document.notes.at(-1)!
   const paragraph=request.shaped_lines.paragraphs.find(p=>p.paragraph_id==='paragraph:footnote:2')!
   const properties=request.resolved_layout.paragraphs.find(p=>p.paragraph_id===paragraph.paragraph_id)!.properties
   if(mutation==='missing'){request.document.notes.pop();request.shaped_lines.paragraphs=request.shaped_lines.paragraphs.filter(p=>p.story_id!==continuation.id);request.resolved_layout.paragraphs=request.resolved_layout.paragraphs.filter(p=>p.paragraph_id!==continuation.blocks[0]!.id)}
   if(mutation==='unshaped')request.shaped_lines.paragraphs=request.shaped_lines.paragraphs.filter(p=>p.story_id!==continuation.id)
   if(mutation==='unsupported')request.document.unsupported.push({id:'unsupported:continuation',code:'UNMODELED_NOTE_MARKUP',capability:'notes',scope_id:continuation.id,preservation:'refuse-mutation',message:'Unsupported active source'})
   if(mutation==='relationship')continuation.relationship_id='rIdDrift'
   if(mutation==='spacing'){paragraph.spacing_after_millipoints=50;paragraph.block_advance_millipoints+=50;properties.spacing_after_twips=1}
   if(mutation==='keep-next')properties.keep_next=true
   if(mutation==='oversized'){paragraph.lines[0]!.line_height_millipoints=40_000;paragraph.block_advance_millipoints=40_000}
   const result=paginateNativeDocxV1(request);expect(result.ok,`${mutation}: ${JSON.stringify(result)}`).toBe(mutation!=='relationship')
   if(result.ok)expect(result.value).toEqual(expect.objectContaining({status:'refused',pages:[],sections:[]}))
  }
 })
})


describe('line-boundary note continuation',()=>{
 function multiline(kind:'footnote'|'endnote',count=7){
  const request=kind==='footnote'?continuedFootnoteFixture():continuedEndnoteFixture()
  const note=request.document.notes.find(n=>n.note_role==='content')!
  const first=note.blocks[0]!,shape=request.shaped_lines.paragraphs.find(p=>p.paragraph_id===first.id)!
  const removed=new Set(note.blocks.slice(1).map(b=>b.id));note.blocks=[first]
  request.resolved_layout.paragraphs=request.resolved_layout.paragraphs.filter(p=>!removed.has(p.paragraph_id))
  request.shaped_lines.paragraphs=request.shaped_lines.paragraphs.filter(p=>!removed.has(p.paragraph_id))
  const original=shape.lines[0]!
  shape.lines=Array.from({length:count},(_,ordinal)=>({...structuredClone(original),id:`line:${shape.paragraph_id}:${ordinal}`,ordinal,logical_to_visual:ordinal===0?original.logical_to_visual:[],fragments:ordinal===0?structuredClone(original.fragments):[]}))
  shape.block_advance_millipoints=shape.lines.reduce((sum,l)=>sum+l.line_height_millipoints,0)
  const properties=request.resolved_layout.paragraphs.find(p=>p.paragraph_id===first.id)!.properties
  properties.keep_lines=false
  return {request,shape,properties}
 }
 it.each(['footnote','endnote'] as const)('splits a single %s paragraph without duplicating source lines or its label',kind=>{
  const {request,shape}=multiline(kind),before=structuredClone(request),output=paginated(request)
  const slices=output.pages.flatMap(p=>p.note_stories?.filter(n=>n.note_role==='content')??[])
  expect(slices.map(s=>s.lines.length)).toEqual([2,3,2])
  expect(slices.flatMap(s=>s.lines.map(l=>l.source_line_ordinal))).toEqual([0,1,2,3,4,5,6])
  expect(slices.flatMap(s=>s.lines.map(l=>l.line_id))).toEqual(shape.lines.map(l=>l.id))
  expect(decodeNativeDocxPaginatedLayoutForRequest(output,request).ok).toBe(true)
  expect(request).toEqual(before);expect(paginated(request)).toEqual(output)
  for(const corrupt of [(v:typeof output)=>{v.pages[1]!.note_stories![1]!.lines[0]!.source_line_ordinal=0},(v:typeof output)=>{v.pages[1]!.note_stories![1]!.lines.reverse()},(v:typeof output)=>{v.pages[2]!.note_stories![1]!.lines.pop()}]){
   const forged=structuredClone(output);corrupt(forged);expect(decodeNativeDocxPaginatedLayoutForRequest(forged,request).ok).toBe(false)
  }
 })
 it('backs up a split to avoid a final orphan and honors explicit disabled widow control',()=>{
  const guarded=multiline('footnote',6)
  expect(paginated(guarded.request).pages.map(p=>p.note_stories![1]!.lines.length)).toEqual([2,2,2])
  const unguarded=multiline('footnote',6);unguarded.properties.widow_control=false
  expect(paginated(unguarded.request).pages.map(p=>p.note_stories![1]!.lines.length)).toEqual([2,3,1])
 })
 it('refuses an oversized kept paragraph and unsatisfiable widow group without partial pages',()=>{
  for(const count of [3,7]){
   const {request,properties}=multiline('footnote',count)
   if(count===7)properties.keep_lines=true
   const result=paginateNativeDocxV1(request)
   expect(result.ok).toBe(true)
   if(result.ok)expect(result.value).toMatchObject({status:'refused',pages:[],sections:[]})
  }
 })
})


describe('joint body and continued footnote flow',()=>{
 function input(bodyCount=8,noteCount=1){
  const request=fixture({lineCounts:Array.from({length:bodyCount},()=>1),bodyHeight:40_000})
  addFootnote(request,'1','1',10_000)
  const note=request.document.notes.find(n=>n.note_role==='content')!,shape=request.shaped_lines.paragraphs.find(p=>p.paragraph_id===note.blocks[0]!.id)!,first=shape.lines[0]!
  shape.lines=Array.from({length:6},(_,ordinal)=>({...structuredClone(first),id:`line:${shape.paragraph_id}:${ordinal}`,ordinal,fragments:ordinal?[]:structuredClone(first.fragments),logical_to_visual:ordinal?[]:[0]}))
  shape.block_advance_millipoints=60_000
  request.resolved_layout.paragraphs.find(p=>p.paragraph_id===shape.paragraph_id)!.properties.widow_control=false
  const template=continuedFootnoteFixture(),sentinel=template.document.notes.at(-1)!
  sentinel.relationship_id=request.document.notes[0]!.relationship_id
  request.document.notes.push(structuredClone(sentinel))
  request.resolved_layout.paragraphs.push(structuredClone(template.resolved_layout.paragraphs.find(p=>p.paragraph_id===sentinel.blocks[0]!.id)!))
  request.shaped_lines.paragraphs.push(structuredClone(template.shaped_lines.paragraphs.find(p=>p.paragraph_id===sentinel.blocks[0]!.id)!))
  for(let i=2;i<=noteCount;i++)addFootnote(request,String(i),String(i),10_000)
  return request
 }
 function coverage(request:NativeDocxPaginationRequestV1){
  const before=structuredClone(request),output=paginated(request)
  expect(output.pages.flatMap(p=>p.lines.map(l=>l.line_id))).toEqual(request.document.body.blocks.flatMap(b=>request.shaped_lines.paragraphs.find(p=>p.paragraph_id===b.id)!.lines.map(l=>l.id)))
  for(const note of request.document.notes.filter(n=>n.note_role==='content')){
   const expected=note.blocks.flatMap(b=>request.shaped_lines.paragraphs.find(p=>p.paragraph_id===b.id)!.lines.map(l=>l.id))
   expect(output.pages.flatMap(p=>p.note_stories?.filter(n=>n.story_id===note.id).flatMap(n=>n.lines.map(l=>l.line_id))??[])).toEqual(expected)
  }
  for(const page of output.pages){
   const bodyBottom=page.lines.reduce((y,l)=>Math.max(y,l.y_millipoints+l.height_millipoints),page.body_box.y_millipoints)
   expect(page.note_stories?.[0]?.top_millipoints??bodyBottom).toBeGreaterThanOrEqual(bodyBottom)
  }
  expect(request).toEqual(before);expect(paginated(request)).toEqual(output)
  expect(decodeNativeDocxPaginatedLayoutForRequest(output,request).ok).toBe(true)
  return output
 }
 it('shares carried note slices with later body paragraphs and restores full body height afterward',()=>{
  const output=coverage(input())
  expect(output.pages.map(p=>p.lines.length)).toEqual([1,1,1,4,1])
  expect(output.pages.slice(0,3).map(p=>p.note_stories!.map(n=>n.note_role))).toEqual([['separator','content'],['continuation-separator','content'],['continuation-separator','content']])
 })
 it('finishes a carried note and starts another under one continuation rule on its reference page',()=>{
  const request=input(8,2),output=coverage(request)
  expect(output.pages[1]!.lines).toHaveLength(0)
  expect(output.pages[2]!.note_stories!.map(n=>[n.note_role,n.number])).toEqual([['continuation-separator',undefined],['content',1],['content',2]])
  expect(output.pages[2]!.lines[0]!.paragraph_id).toBe('paragraph:2')
  for(const mutation of ['drop','order','reference','body'] as const){
   const forged=structuredClone(output)
   if(mutation==='drop')forged.pages[1]!.note_stories![1]!.lines.pop()
   if(mutation==='order')forged.pages[2]!.note_stories!.reverse()
   if(mutation==='reference')forged.pages[2]!.note_stories![2]!.reference_run_id='run:paragraph:1'
   if(mutation==='body')forged.pages[2]!.lines[0]!.y_millipoints++
   expect(decodeNativeDocxPaginatedLayoutForRequest(forged,request).ok).toBe(false)
  }
 })
 it('handles three references without a two-note limit',()=>{
  const output=coverage(input(8,3))
  expect(output.pages.flatMap(p=>p.note_stories?.filter(n=>n.note_role==='content').map(n=>n.number)??[])).toEqual([1,1,1,2,3])
 })
 it('handles large reference groups and retains the existing request traversal budget',()=>{
  const make=(count:number)=>{const r=fixture({lineCounts:Array.from({length:count},()=>1),bodyHeight:20_000_000});for(let i=1;i<=count;i++)addFootnote(r,String(i),String(i));return r}
  const output=paginated(make(200))
  expect(output.pages).toHaveLength(1)
  expect(output.pages[0]!.lines).toHaveLength(200)
  expect(output.pages[0]!.note_stories!.slice(1).map(n=>n.number)).toEqual(Array.from({length:200},(_,i)=>i+1))
  expect(paginateNativeDocxV1(make(1000))).toMatchObject({ok:false,issues:expect.arrayContaining([expect.objectContaining({code:'LIMIT_EXCEEDED'})])})
 })
 it('splits later body text while keeping every paragraph slice in source order',()=>{
  const output=coverage(continuedFootnoteFixture(6))
  expect(output.pages.flatMap(p=>p.lines.map(l=>l.source_line_ordinal))).toEqual([0,1,2,3,4,5])
  expect(output.pages.filter(p=>p.lines.length&&p.note_stories?.[0]?.note_role==='continuation-separator')).not.toHaveLength(0)
 })
 it('honors explicit later page breaks and indivisible body paragraphs',()=>{
  const broken=input()
  broken.resolved_layout.paragraphs.find(p=>p.paragraph_id==='paragraph:5')!.properties.page_break_before=true
  const output=coverage(broken)
  expect(output.pages.at(-1)!.lines[0]!.paragraph_id).toBe('paragraph:5')
  const kept=input(),shape=kept.shaped_lines.paragraphs.find(p=>p.paragraph_id==='paragraph:2')!,first=shape.lines[0]!
  shape.lines=Array.from({length:3},(_,ordinal)=>({...structuredClone(first),id:`line:${shape.paragraph_id}:${ordinal}`,ordinal,fragments:ordinal?[]:first.fragments,logical_to_visual:ordinal?[]:first.logical_to_visual}))
  shape.block_advance_millipoints=30_000
  kept.resolved_layout.paragraphs.find(p=>p.paragraph_id===shape.paragraph_id)!.properties.keep_lines=true
  const placed=coverage(kept)
  expect(placed.pages.filter(p=>p.lines.some(l=>l.paragraph_id===shape.paragraph_id))).toHaveLength(1)
 })
 it('recomputes source geometry before committing note placements',()=>{
  const request=input(),output=coverage(request)
  for(const mutate of [(v:typeof output)=>{v.pages[0]!.lines[0]!.x_millipoints++},(v:typeof output)=>{v.pages[0]!.width_millipoints++}]){
   const staged=structuredClone(output);mutate(staged);const before=structuredClone(staged)
   expect(placeNativeDocxNotesV1(staged,request.document,request.resolved_layout,request.shaped_lines,undefined,request)).toBeDefined()
   expect(staged).toEqual(before)
  }
 })
 it('refuses missing activated separators, impossible kept groups, and page/line budgets atomically',()=>{
  const requests=[input(),input()]
  const note=requests[1]!.document.notes.find(n=>n.note_role==='content')!
  requests[1]!.resolved_layout.paragraphs.find(p=>p.paragraph_id===note.blocks[0]!.id)!.properties.keep_lines=true
  const sentinel=requests[0]!.document.notes.pop()!
  requests[0]!.resolved_layout.paragraphs=requests[0]!.resolved_layout.paragraphs.filter(p=>p.paragraph_id!==sentinel.blocks[0]!.id)
  requests[0]!.shaped_lines.paragraphs=requests[0]!.shaped_lines.paragraphs.filter(p=>p.story_id!==sentinel.id)
  for(const request of requests)expect(paginateNativeDocxV1(request)).toMatchObject({ok:true,value:{status:'refused',pages:[],sections:[]}})
  for(const [limits,key,bound] of [[DOCX_PAGINATION_LIMITS,'maxPages',1],[DOCX_NOTE_PAGINATION_LIMITS,'maxNoteLines',15]] as const){
   const original=limits[key as keyof typeof limits]
   try{Object.assign(limits,{[key]:bound});expect(paginateNativeDocxV1(input())).toMatchObject({ok:true,value:{status:'refused',pages:[],sections:[]}})}finally{Object.assign(limits,{[key]:original})}
  }
 })
})

describe('approximate pagination of inert note separator stories', () => {
  function legacy(request: NativeDocxPaginationRequestV1) {
    request.pagination_settings.profile = 'unsupported'
    delete request.pagination_settings.compatibility_mode
    request.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different layout semantics' }]
    return { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.pagination_settings.document_id, revision: request.pagination_settings.revision, package_sha256: request.pagination_settings.package_sha256, settings_sha256: request.pagination_settings.settings_sha256 ?? null, status: 'eligible', legacy_compatibility_mode: 14, reasons: ['Legacy mode 14 uses current layout'] }
  }
  // Word 2010 separator stories: the instruction paragraph plus a trailing empty
  // paragraph, which the extractor reports as UNMODELED_NOTE_MARKUP and leaves unshaped.
  function addInertSeparators(request: NativeDocxPaginationRequestV1): void {
    for (const [index, role] of (['separator', 'continuation-separator'] as const).entries()) {
      const part = 'word/footnotes.xml'
      const first = paragraph(`paragraph:${role}:1`, 40 + index * 2); first.runs = []; first.anchor = noteAnchor(part, `/w:footnotes[1]/w:footnote[${index + 1}]/w:p[1]`, 10 + index * 100, 40 + index * 100)
      const second = paragraph(`paragraph:${role}:2`, 41 + index * 2); second.runs = []; second.anchor = noteAnchor(part, `/w:footnotes[1]/w:footnote[${index + 1}]/w:p[2]`, 41 + index * 100, 60 + index * 100)
      const story = { id: `story:footnote:${role}`, kind: 'footnote' as const, native_story_id: role === 'separator' ? '-1' : '0', relationship_id: 'rIdFootnotes', note_role: role, part_name: part, anchor: noteAnchor(part, `/w:footnotes[1]/w:footnote[${index + 1}]`, 1 + index * 100, 70 + index * 100), blocks: [first, second].map((entry) => ({ kind: 'paragraph' as const, id: entry.id, paragraph: entry })) }
      request.document.notes.push(story)
      request.document.unsupported.push({ id: `unsupported:${role}`, code: 'UNMODELED_NOTE_MARKUP', capability: 'notes', scope_id: story.id, anchor: story.anchor, preservation: 'preserve-verbatim', message: 'Reserved note separator stories must contain exactly one matching instruction leaf and no visible text' })
      for (const entry of [first, second]) request.resolved_layout.paragraphs.push({ paragraph_id: entry.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } })
    }
  }

  it('omits unreferenced separator stories with unmodeled markup only in approximate layout, disclosed as a declared policy', () => {
    const request = fixture({ lineCounts: [1, 1] })
    addInertSeparators(request)
    const strict = paginateNativeDocxV1(structuredClone(request))
    expect(strict).toMatchObject({ ok: true, value: { status: 'refused' } })
    if (strict.ok) expect(strict.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'note-structure-unsupported' })]))
    const eligibility = legacy(request)
    const approximate = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(approximate.layout.status).toBe('paginated')
    if (approximate.layout.status !== 'paginated') return
    expect(approximate.layout.pages.flatMap((page) => page.lines.map((line) => line.paragraph_id))).toEqual(['paragraph:1', 'paragraph:2'])
    expect(approximate.layout.pages.every((page) => (page.note_stories ?? []).length === 0)).toBe(true)
    expect(approximate.layout.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'source-diagnostic', severity: 'deferred', scope_id: 'story:footnote:separator', message: DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING }),
      expect.objectContaining({ code: 'source-diagnostic', severity: 'deferred', scope_id: 'story:footnote:continuation-separator', message: DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING }),
    ]))
    expect(nativeDocxApproximatePaginationPolicyReasonsV1(approximate.layout)).toEqual([DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING])
    const clean = fixture({ lineCounts: [1, 1] })
    expect(nativeDocxApproximatePaginationPolicyReasonsV1(paginateNativeDocxApproximateLegacyV1(clean, legacy(clean)).layout)).toEqual([])
  })

  it('still refuses unmodeled separator markup in approximate layout once a footnote reference exists', () => {
    const request = fixture({ lineCounts: [1, 1] })
    addFootnote(request)
    const separator = request.document.notes.find((story) => story.note_role === 'separator')!
    request.document.unsupported.push({ id: 'unsupported:separator', code: 'UNMODELED_NOTE_MARKUP', capability: 'notes', scope_id: separator.id, anchor: separator.anchor!, preservation: 'preserve-verbatim', message: 'Reserved note separator stories must contain exactly one matching instruction leaf and no visible text' })
    const approximate = paginateNativeDocxApproximateLegacyV1(request, legacy(request))
    expect(approximate.layout.status).toBe('refused')
    expect(approximate.layout.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'note-structure-unsupported' })]))
    expect(nativeDocxApproximatePaginationPolicyReasonsV1(approximate.layout)).toEqual([])
  })
})
