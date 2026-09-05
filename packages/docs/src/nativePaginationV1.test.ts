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
    expect(Object.keys(shared).sort()).toEqual([...DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS.SettingsV1].sort())
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

    const ambiguousBalance = fixture({ lineCounts: [2, 1] })
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
