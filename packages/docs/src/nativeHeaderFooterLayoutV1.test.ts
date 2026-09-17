import { describe, expect, it } from 'vitest'
import {
  DOCX_APPROXIMATE_HEADER_FOOTER_BAND_WARNING,
  DOCX_APPROXIMATE_HEADER_FOOTER_NONBLOCKING_SOURCE,
  DOCX_APPROXIMATE_HEADER_FOOTER_OMITTED_PARAGRAPH_WARNING,
  DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL,
  DOCX_HEADER_FOOTER_LAYOUT_VERSION,
  layoutNativeDocxHeadersFootersV1,
  nativeDocxApproximateHeaderFooterPolicyReasonsV1,
  nativeDocxHeaderFooterLayoutSha256V1,
  type NativeDocxHeaderFooterLayoutInputV1,
  type NativeDocxHeaderFooterLayoutSuccessV1,
} from './nativeHeaderFooterLayoutV1.js'
import { DOCX_APPROXIMATE_OMITTED_CONTENT_CODES, nativeDocxOmittedContentCategoryV1 } from './nativeApproximateOmittedContentV1.js'

const HASH = `sha256:${'a'.repeat(64)}`

function anchor(part: string, path: string) {
  return { part_name: part, path, start_byte: 1, end_byte: 2, xml_sha256: HASH }
}

function paragraph(id: string, part: string) {
  return {
    id, anchor: anchor(part, `/w:hdr[1]/w:p[${id}]`),
    edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'READ_ONLY', message: 'fixture', preservation: 'refuse-mutation' } },
    properties: {}, runs: [],
  }
}

function story(id: string, kind: 'header' | 'footer') {
  const part = `word/${id}.xml`
  const value = paragraph(`paragraph:${id}`, part)
  return { id: `story:${id}`, kind, part_name: part, anchor: anchor(part, kind === 'header' ? '/w:hdr[1]' : '/w:ftr[1]'), blocks: [{ kind: 'paragraph', id: value.id, paragraph: value }] }
}

function shaped(stories: ReturnType<typeof story>[]) {
  return stories.map((value) => ({
    paragraph_id: value.blocks[0]!.paragraph.id, story_id: value.id, story_kind: value.kind,
    direction: 'ltr', alignment: 'start', spacing_before_millipoints: 0, spacing_after_millipoints: 0,
    indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0,
    block_advance_millipoints: 10_000,
    lines: [{
      id: `line:${value.id}`, ordinal: 0, available_width_millipoints: 468_000, inline_offset_millipoints: 0,
      advance_inline_millipoints: 20_000, ascent_millipoints: 8_000, descent_millipoints: -2_000,
      line_gap_millipoints: 0, line_height_millipoints: 10_000, fragments: [],
    }],
  }))
}

function fixture(): NativeDocxHeaderFooterLayoutInputV1 {
  const stories = [story('header-default', 'header'), story('header-first', 'header'), story('header-even', 'header'), story('header-second-default', 'header'), story('footer-default', 'footer')]
  const bodyParagraph = paragraph('paragraph:body', 'word/document.xml')
  const page = {
    width_twips: 12_240, height_twips: 15_840, orientation: 'portrait' as const,
    margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
    columns: 1, column_spacing_twips: 720,
  }
  const ref = (kind: 'default' | 'first' | 'even', id: string) => ({ kind, story_id: `story:${id}`, relationship_id: `r:${id}` })
  const sections = [{
    id: 'section:1', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]/w:p[1]/w:pPr[1]/w:sectPr[1]'), starts_at_block_id: bodyParagraph.id,
    break_type: 'next-page' as const, title_page: true, page,
    header_refs: [ref('default', 'header-default'), ref('first', 'header-first'), ref('even', 'header-even')],
    footer_refs: [ref('default', 'footer-default'), ref('first', 'footer-default'), ref('even', 'footer-default')],
  }, {
    id: 'section:2', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]/w:sectPr[1]'), starts_at_block_id: bodyParagraph.id,
    break_type: 'next-page' as const, title_page: false, page,
    header_refs: [ref('default', 'header-second-default')], footer_refs: [],
  }]
  const makePage = (id: string, ordinal: number, sectionID: string, sectionOrdinal: number) => ({
    id, ordinal, section_id: sectionID, section_ids: [sectionID], section_page_ordinal: sectionOrdinal, kind: 'content' as const,
    width_millipoints: 612_000, height_millipoints: 792_000,
    body_box: { x_millipoints: 72_000, y_millipoints: 72_000, width_millipoints: 468_000, height_millipoints: 648_000 },
    columns: [{ id: `${sectionID}:column:0`, section_id: sectionID, ordinal: 0, x_millipoints: 72_000, y_millipoints: 72_000, width_millipoints: 468_000, height_millipoints: 648_000 }],
    header_refs: [], footer_refs: [], paragraph_slices: [], lines: [],
  })
  return {
    document: {
      protocol: 'injoffice.docx.native', version: 1, document_id: 'document:test', revision: 'revision:1', source: { package_sha256: HASH, main_part: 'word/document.xml' },
      body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]'), blocks: [{ kind: 'paragraph', id: bodyParagraph.id, paragraph: bodyParagraph }] },
      sections, headers: stories.filter((value) => value.kind === 'header'), footers: stories.filter((value) => value.kind === 'footer'), notes: [], comment_stories: [], comments: [], capabilities: [], passthrough_parts: [], unsupported: [],
    } as any,
    resolved_layout: {
      protocol: 'injoffice.docx.resolved-layout', version: 1, document_id: 'document:test', revision: 'revision:1', source_parts: { main_part: 'word/document.xml' },
      paragraphs: stories.map((value) => ({ paragraph_id: value.blocks[0]!.paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })),
      runs: [], tables: [], fonts: [], diagnostics: [],
    } as any,
    shaped_lines: {
      protocol: 'injoffice.docx.shaped-lines', version: 1, document_id: 'document:test', revision: 'revision:1', available_width_millipoints: 468_000, tab_interval_millipoints: 36_000,
      font_manifest: { manifest_id: 'manifest:test', revision: 'revision:1' }, providers: { resolver_id: 'resolver', resolver_revision: '1', shaper_id: 'shaper', shaper_revision: '1' }, paragraphs: shaped(stories), diagnostics: [],
    } as any,
    pagination_settings: {
      protocol: 'injoffice.docx.pagination-settings', version: 1, document_id: 'document:test', revision: 'revision:1', package_sha256: HASH, main_part: 'word/document.xml',
      profile: 'word-modern-default', default_tab_stop_twips: 720, mirror_margins: false, gutter_at_top: false, even_and_odd_headers: true, compatibility_mode: 15, diagnostics: [],
    },
    paginated_layout: {
      protocol: 'injoffice.docx.paginated-layout', version: 1, status: 'paginated', provenance: {} as any, diagnostics: [],
      sections: [], pages: [makePage('page:section:1:0', 0, 'section:1', 0), makePage('page:section:1:1', 1, 'section:1', 1), makePage('page:section:2:0', 2, 'section:2', 0), makePage('page:section:2:1', 3, 'section:2', 1)],
    },
  }
}

describe('native DOCX header/footer layout v1', () => {
  it('resolves independent inheritance and exact title/even/default physical-page selection', () => {
    const value = layoutNativeDocxHeadersFootersV1(fixture())
    expect(value.status).toBe('placed')
    if (value.status !== 'placed') throw new Error(JSON.stringify(value.diagnostics))
    expect(value.pages.map((page) => page.header_ref?.story_id)).toEqual(['story:header-first', 'story:header-even', 'story:header-second-default', 'story:header-even'])
    expect(value.pages.map((page) => page.footer_ref?.story_id)).toEqual(['story:footer-default', 'story:footer-default', 'story:footer-default', 'story:footer-default'])
    expect(value.pages[0]!.lines.find((line) => line.region === 'header')).toEqual(expect.objectContaining({ y_millipoints: 36_000, x_millipoints: 72_000 }))
    expect(value.pages[0]!.lines.find((line) => line.region === 'footer')).toEqual(expect.objectContaining({ y_millipoints: 746_000, x_millipoints: 72_000 }))
    expect(value.sha256).toBe(nativeDocxHeaderFooterLayoutSha256V1(value))
    expect(layoutNativeDocxHeadersFootersV1(structuredClone(fixture()))).toEqual(value)
  })

  it('hashes canonical object keys while preserving semantic array order', () => {
    const left: Omit<NativeDocxHeaderFooterLayoutSuccessV1, 'sha256'> = { protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL, version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, status: 'placed', diagnostics: [], pages: [{ page_id: 'page:1', lines: [] }] }
    const right: Omit<NativeDocxHeaderFooterLayoutSuccessV1, 'sha256'> = { pages: [{ lines: [], page_id: 'page:1' }], diagnostics: [], status: 'placed', version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL }
    expect(nativeDocxHeaderFooterLayoutSha256V1(left)).toBe(nativeDocxHeaderFooterLayoutSha256V1(right))
    expect(nativeDocxHeaderFooterLayoutSha256V1({ ...left, pages: [{ page_id: 'page:2', lines: [] }, ...left.pages] })).not.toBe(nativeDocxHeaderFooterLayoutSha256V1(left))
  })

  it('refuses selected fields, tables, shapes, and one-unit reserved-band overflow without partial pages', () => {
    const cases: Array<[string, (value: NativeDocxHeaderFooterLayoutInputV1) => void]> = [
      ['selected-story-field', (value) => { value.document.unsupported.push({ id: 'unsupported:field', code: 'FIELD_SEMANTICS', capability: 'fields', scope_id: 'paragraph:header-first', preservation: 'refuse-mutation', message: 'PAGE field is not explicitly modeled' }) }],
      ['selected-story-table', (value) => { (value.document.headers.find((entry) => entry.id === 'story:header-first')!.blocks as any[]).push({ kind: 'table', id: 'table:1', table: { id: 'table:1', rows: [] } }) }],
      ['selected-story-shape', (value) => { (value.document.headers.find((entry) => entry.id === 'story:header-first')!.blocks[0]!.paragraph!.runs as any[]).push({ kind: 'drawing', id: 'run:drawing', drawing: { id: 'drawing:1' } }) }],
      ['header-band-overflow', (value) => { value.document.sections[0]!.page.margins.top_twips = 919; value.paginated_layout.status === 'paginated' && (value.paginated_layout.pages[0]!.body_box.y_millipoints = 45_950) }],
      ['section-geometry-invalid', (value) => { value.document.unsupported.push({ id: 'unsupported:section', code: 'UNMODELED_SECTION_PROPERTY', capability: 'sections', scope_id: 'section:1', preservation: 'refuse-mutation', message: 'Ambiguous section geometry' }) }],
    ]
    for (const [code, mutate] of cases) {
      const input = fixture(); mutate(input)
      const value = layoutNativeDocxHeadersFootersV1(input)
      expect(value).toEqual(expect.objectContaining({ status: 'refused', pages: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code })]) }))
    }
  })

  it('omits unmodeled section geometry when approximate header/footer placement is requested', () => {
    const input = fixture()
    input.omit_unmodeled_section_geometry = true
    input.document.unsupported.push({ id: 'unsupported:section', code: 'UNMODELED_SECTION_PROPERTY', capability: 'sections', scope_id: 'section:1', preservation: 'refuse-mutation', message: 'Ambiguous section geometry' })
    const value = layoutNativeDocxHeadersFootersV1(input)
    expect(value.status).toBe('placed')
    expect(value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ code: 'section-geometry-invalid' })]))
  })

  it('refuses inconsistent relationship-id reuse before returning any placement', () => {
    const input = fixture()
    input.document.sections[1]!.header_refs[0]!.relationship_id = 'r:header-even'
    const value = layoutNativeDocxHeadersFootersV1(input)
    expect(value).toEqual(expect.objectContaining({ status: 'refused', pages: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'relationship-ambiguous' })]) }))
  })
})

describe('approximate header/footer placement policy', () => {
  // The production allowlist the paint passes, not a hand-built set.
  const NONBLOCKING = DOCX_APPROXIMATE_HEADER_FOOTER_NONBLOCKING_SOURCE
  const footerLine = (input: NativeDocxHeaderFooterLayoutInputV1) => input.shaped_lines.paragraphs.find((entry) => entry.paragraph_id === 'paragraph:footer-default')!.lines[0]!
  const approximate = (input: NativeDocxHeaderFooterLayoutInputV1) => { input.approximate_nonblocking_source = NONBLOCKING; return input }

  it('keeps painting a story whose only source diagnostics are approximate-nonblocking, and still refuses other codes', () => {
    const framePr = { id: 'unsupported:framePr', code: 'UNMODELED_PARAGRAPH_PROPERTY', capability: 'paragraph-properties', scope_id: 'paragraph:footer-default', preservation: 'preserve-verbatim', message: 'framePr page number frame is preserved verbatim' }
    const strict = fixture(); strict.document.unsupported.push(framePr as never)
    expect(layoutNativeDocxHeadersFootersV1(strict)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-story-unsupported', scope_id: 'paragraph:footer-default' })]) }))
    const tolerated = approximate(fixture()); tolerated.document.unsupported.push(framePr as never)
    tolerated.resolved_layout.diagnostics.push({ code: 'UNMODELED_PARAGRAPH_PROPERTY', severity: 'unsupported', scope_id: 'paragraph:footer-default', part_name: 'word/footer-default.xml', path: '/w:ftr[1]/w:p[1]/w:pPr[1]/w:framePr[1]', preservation: 'preserve-verbatim', message: 'This paragraph property is preserved and not guessed' } as never)
    const placed = layoutNativeDocxHeadersFootersV1(tolerated)
    expect(placed.status).toBe('placed')
    expect(placed.approximations).toBeUndefined()
    expect(placed.pages.some((page) => page.lines.some((line) => line.paragraph_id === 'paragraph:footer-default'))).toBe(true)
    for (const [code, expected] of [['FIELD_SEMANTICS', 'selected-story-field'], ['UNMODELED_SECTION_PROPERTY', 'selected-story-unsupported'], ['NESTED_TABLE_OR_CELL_MARKUP', 'selected-story-unsupported']] as const) {
      expect(NONBLOCKING.has(code)).toBe(false)
      const blocking = approximate(fixture())
      blocking.document.unsupported.push({ ...framePr, id: `unsupported:${code}`, code } as never)
      expect(layoutNativeDocxHeadersFootersV1(blocking)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: expected, scope_id: 'paragraph:footer-default' })]) }))
    }
  })

  it('paints a header/footer whose drawing the extractor dropped, and keeps refusing a modeled one', () => {
    // Both codes are only ever recorded when the extractor refused the drawing and left
    // no RunV1.drawing behind, so the object is absent from the model rather than
    // painted wrongly. Each is admitted here only because the approximate
    // omitted-content discloser reports it, so the drop is never silent.
    for (const code of ['UNMODELED_DRAWING', 'PICTURE_GRAPHIC_REQUIRED'] as const) {
      expect(NONBLOCKING.has(code)).toBe(true)
      expect(DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has(code)).toBe(true)
      expect(nativeDocxOmittedContentCategoryV1(code, undefined)).toBe('drawing')
      const dropped = { id: `unsupported:${code}`, code, capability: 'drawings', scope_id: 'paragraph:footer-default', preservation: 'refuse-mutation', message: 'Drawing payload is preserved verbatim' }
      const strict = fixture(); strict.document.unsupported.push(dropped as never)
      expect(layoutNativeDocxHeadersFootersV1(strict)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-story-unsupported', scope_id: 'paragraph:footer-default' })]) }))
      const tolerated = approximate(fixture()); tolerated.document.unsupported.push(dropped as never)
      const placed = layoutNativeDocxHeadersFootersV1(tolerated)
      expect(placed.status).toBe('placed')
      expect(placed.diagnostics).toEqual([])
      expect(placed.pages.some((page) => page.lines.some((line) => line.paragraph_id === 'paragraph:footer-default'))).toBe(true)
    }
    // A drawing the extractor DID model but that is outside the exact inline raster
    // subset is a different fact: it would be silently dropped, so it still refuses.
    const modeled = approximate(fixture())
    ;(modeled.document.footers.find((entry) => entry.id === 'story:footer-default')!.blocks[0]!.paragraph!.runs as any[]).push({ kind: 'drawing', id: 'run:footer-drawing', drawing: { id: 'drawing:1' } })
    expect(layoutNativeDocxHeadersFootersV1(modeled)).toEqual(expect.objectContaining({ status: 'refused', pages: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-story-shape' })]) }))
  })

  it('accepts an expanded line box under the declared line-box policy and still refuses a compressed one', () => {
    const expanded = approximate(fixture()); footerLine(expanded).line_height_millipoints = 14_000
    const value = layoutNativeDocxHeadersFootersV1(expanded)
    expect(value.status).toBe('placed')
    expect(value.pages.flatMap((page) => page.lines).find((line) => line.paragraph_id === 'paragraph:footer-default')?.height_millipoints).toBe(14_000)
    const strict = fixture(); footerLine(strict).line_height_millipoints = 14_000
    expect(layoutNativeDocxHeadersFootersV1(strict)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-line-invalid' })]) }))
    const compressed = approximate(fixture()); footerLine(compressed).line_height_millipoints = 9_000
    expect(layoutNativeDocxHeadersFootersV1(compressed)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-line-invalid' })]) }))
  })

  it('paints an overflowing footer at the authored distance only in approximate mode, disclosed as a declared policy', () => {
    const overflow = (input: NativeDocxHeaderFooterLayoutInputV1) => { for (const section of input.document.sections) section.page.margins.footer_twips = 1_440; return input }
    expect(layoutNativeDocxHeadersFootersV1(overflow(fixture()))).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'footer-band-overflow' })]) }))
    const value = layoutNativeDocxHeadersFootersV1(overflow(approximate(fixture())))
    expect(value.status).toBe('placed')
    expect(value.approximations).toEqual([{ policy: 'band-overflow', scope_id: 'story:footer-default', message: DOCX_APPROXIMATE_HEADER_FOOTER_BAND_WARNING }])
    const footer = value.pages[0]!.lines.find((line) => line.region === 'footer')!
    expect(footer.y_millipoints).toBe(792_000 - 72_000 - 10_000)
    expect(nativeDocxApproximateHeaderFooterPolicyReasonsV1(value)).toEqual([DOCX_APPROXIMATE_HEADER_FOOTER_BAND_WARNING])
    expect(nativeDocxApproximateHeaderFooterPolicyReasonsV1(layoutNativeDocxHeadersFootersV1(approximate(fixture())))).toEqual([])
    // The disclosed approximation is part of the canonical layout hash.
    expect(value.sha256).not.toBe(layoutNativeDocxHeadersFootersV1(approximate(fixture())).sha256)
  })

  it('omits an unshaped text-less story paragraph in approximate mode and still refuses one with text', () => {
    const drop = (input: NativeDocxHeaderFooterLayoutInputV1) => { input.shaped_lines.paragraphs = input.shaped_lines.paragraphs.filter((entry) => entry.paragraph_id !== 'paragraph:footer-default'); return input }
    expect(layoutNativeDocxHeadersFootersV1(drop(fixture()))).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-paragraph-missing' })]) }))
    const omitted = layoutNativeDocxHeadersFootersV1(drop(approximate(fixture())))
    expect(omitted.status).toBe('placed')
    expect(omitted.pages.every((page) => page.lines.every((line) => line.region !== 'footer'))).toBe(true)
    expect(omitted.approximations).toEqual([{ policy: 'omitted-paragraph', scope_id: 'paragraph:footer-default', message: DOCX_APPROXIMATE_HEADER_FOOTER_OMITTED_PARAGRAPH_WARNING }])
    expect(nativeDocxApproximateHeaderFooterPolicyReasonsV1(omitted)).toEqual([DOCX_APPROXIMATE_HEADER_FOOTER_OMITTED_PARAGRAPH_WARNING])
    const withText = drop(approximate(fixture()))
    const story = withText.document.footers[0]!
    ;(story.blocks[0]!.paragraph!.runs as unknown[]).push({ kind: 'text', id: 'run:footer-text', anchor: anchor(story.part_name, '/w:ftr[1]/w:p[1]/w:r[1]'), text: 'Page' })
    expect(layoutNativeDocxHeadersFootersV1(withText)).toEqual(expect.objectContaining({ status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'selected-paragraph-missing' })]) }))
  })
})
